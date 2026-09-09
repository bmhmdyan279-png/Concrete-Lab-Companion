#!/usr/bin/env python3
"""Audit a built or released workbook against the v4 product contract.

This script used to be a v2/v3-era heuristic scanner.  It graded an
artifact by counting named ranges, demanding ``calcMode == auto``, and
complaining about "too few" data validations and conditional formats —
every one of which the v4 architecture deliberately abolished: the
workbook is display-only, so there is nothing to calculate, no input
cells to validate and no formulas to name.  Run against a correct v4
artifact it reported four fake crises.

What it checks now is what the product actually promises:

* the artifact contains **zero** formula cells (the headline v4 contract,
  verified both by string prefix and by openpyxl's own ``data_type``);
* the SHA-256 sidecar and the manifest describe the same bytes;
* the manifest is complete and its version is the package version;
* every worksheet the catalogue promises is present, and no stray sheet
  was added;
* the ESTIMATION test (4-5) really carries its badge and disclaimer;
* protection state matches what the manifest claims;
* no cell still says ``TODO`` / ``soon`` / ``placeholder``.

Usage::

    python build.py --output output --no-protect
    python scripts/audit_excel.py --output output          # audit that build
    python scripts/audit_excel.py --file path/to/file.xlsx  # audit one file

Writes ``docs/qa/EXCEL_AUDIT_REPORT.json`` and exits non-zero if any
check fails, so it can be used as a release gate.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT / "src") not in sys.path:
    sys.path.insert(0, str(REPO_ROOT / "src"))

import openpyxl  # noqa: E402

from concrete_lab import __version__  # noqa: E402
from concrete_lab.cli import ARTIFACT_PATTERN  # noqa: E402
from concrete_lab.qa.structural import run_structural_checks  # noqa: E402
from concrete_lab.report.manifest import REQUIRED_MANIFEST_KEYS  # noqa: E402
from concrete_lab.specs.base import all_tests, implemented_tests  # noqa: E402
from concrete_lab.utils import compute_sha256  # noqa: E402

#: Report location (kept for continuity with earlier releases).
REPORT_PATH: Path = REPO_ROOT / "docs" / "qa" / "EXCEL_AUDIT_REPORT.json"

#: Unfinished-work markers that must never survive into a published artifact.
#:
#: Deliberately **case-sensitive**: these markers are conventional in upper
#: case, while ordinary English prose in this product may legitimately use
#: words like "placeholder" (the golden suite writes "kept as a structural
#: placeholder" for tests the engine does not implement yet).  A
#: case-insensitive match turned that honest sentence into a false crisis.
PLACEHOLDER_PATTERN = re.compile(
    r"\b(?:TODO|TBD|FIXME|XXX|HACK|LOREM)\b"   # conventional markers
    r"|\{\{[^}]*\}\}"                        # unrendered template variable
    r"|<[a-z_]+_placeholder>"                    # template slot
)

#: Badge that must accompany every ESTIMATION-scope result.
ESTIMATION_BADGE = "ESTIMATED — NOT FOR ACCEPTANCE"


class Finding:
    """One audit observation: a check name, its verdict and a detail."""

    def __init__(self, check: str, ok: bool, detail: str = "") -> None:
        self.check = check
        self.ok = ok
        self.detail = detail

    def as_dict(self) -> Dict[str, Any]:
        """JSON form for the report."""
        return {"check": self.check, "ok": self.ok, "detail": self.detail}


def find_formula_cells(workbook: openpyxl.Workbook) -> List[str]:
    """Return addresses of formula cells, by string prefix *and* data type.

    Checking ``data_type == "f"`` as well as a leading ``"="`` matters:
    openpyxl assigns the type when the workbook is loaded, so a value that
    was written as a formula is caught even if it no longer looks like one.
    """
    offenders: List[str] = []
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                value = cell.value
                is_formula = cell.data_type == "f" or (
                    isinstance(value, str) and value.startswith("=")
                )
                if is_formula:
                    offenders.append(f"{sheet.title}!{cell.coordinate}")
    return offenders


def find_placeholders(workbook: openpyxl.Workbook) -> List[str]:
    """Return addresses whose text still contains an unfinished-work marker."""
    offenders: List[str] = []
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and PLACEHOLDER_PATTERN.search(cell.value):
                    offenders.append(f"{sheet.title}!{cell.coordinate}={cell.value[:40]!r}")
    return offenders


def sheet_text(workbook: openpyxl.Workbook, title: str) -> str:
    """Concatenate every string cell of one worksheet (missing → empty)."""
    if title not in workbook.sheetnames:
        return ""
    values = [
        str(cell.value)
        for row in workbook[title].iter_rows()
        for cell in row
        if cell.value is not None
    ]
    return "\n".join(values)


def expected_sheet_names() -> List[str]:
    """Worksheet names a full build must contain, derived from the catalogue."""
    fixed = ["00_راهنما", "01_اطلاعات_آزمون", "22_گزارش", "23_داشبورد",
             "24_QA_Test", "25_خطاها_هشدارها", "_Standards"]
    per_test = [spec.sheet_name for spec in implemented_tests()]
    return fixed[:2] + sorted(per_test) + fixed[2:]


def audit(xlsx_path: Path, report_to: Optional[Path] = REPORT_PATH) -> Tuple[List[Finding], Dict[str, Any]]:
    """Run every artifact check.

    Args:
        xlsx_path: Workbook to audit.
        report_to: Where to write the JSON report (``None`` to skip).

    Returns:
        ``(findings, report)`` — ``report`` is the JSON-serialisable payload.
    """
    findings: List[Finding] = []

    def check(name: str, ok: bool, detail: str = "") -> bool:
        findings.append(Finding(name, ok, detail))
        return ok

    if not xlsx_path.is_file():
        check("artifact-exists", False, f"{xlsx_path} not found")
        return findings, {
            "generator": "scripts/audit_excel.py",
            "package_version": __version__,
            "file": xlsx_path.name,
            "passed": 0,
            "failed": len(findings),
            "checks": [finding.as_dict() for finding in findings],
        }

    manifest_path = xlsx_path.with_suffix(".json")
    sha_path = Path(str(xlsx_path) + ".sha256")

    # ── integrity ────────────────────────────────────────────────────────
    actual_sha = compute_sha256(xlsx_path)
    if sha_path.is_file():
        fields = sha_path.read_text(encoding="utf-8").split()
        recorded = fields[0] if fields else ""
        check("sha256-sidecar-matches", recorded == actual_sha,
              f"sidecar {recorded[:16]}… vs actual {actual_sha[:16]}…")
    else:
        check("sha256-sidecar-present", False, f"missing {sha_path.name}")

    manifest: Dict[str, Any] = {}
    if manifest_path.is_file():
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            check("manifest-parses", True)
        except json.JSONDecodeError as exc:
            check("manifest-parses", False, str(exc))
    else:
        check("manifest-present", False, f"missing {manifest_path.name}")

    if manifest:
        missing = [key for key in REQUIRED_MANIFEST_KEYS if key not in manifest]
        check("manifest-complete", not missing, f"missing keys: {missing}")
        check("manifest-version-current", manifest.get("application_version") == __version__,
              f"manifest {manifest.get('application_version')!r} vs package {__version__!r}")
        check("manifest-sha-matches", manifest.get("sha256") == actual_sha,
              "the manifest was written for different bytes")

    # ── the v4 contract, inside the workbook ─────────────────────────────
    workbook = openpyxl.load_workbook(str(xlsx_path))
    formulas = find_formula_cells(workbook)
    check("zero-formula-cells", not formulas,
          f"{len(formulas)} formula cell(s): {formulas[:5]}")

    placeholders = find_placeholders(workbook)
    check("no-placeholder-text", not placeholders, f"{placeholders[:5]}")

    names = list(workbook.sheetnames)
    expected = expected_sheet_names()
    absent = [name for name in expected if name not in names]
    stray = [name for name in names if name not in expected]
    check("sheet-inventory-complete", not absent, f"missing: {absent}")
    check("no-stray-sheets", not stray, f"unexpected: {stray}")
    check("catalogue-size-consistent", len(all_tests()) == 20,
          "the 20-test catalogue changed; update config.yaml and this script")

    # ── estimation scope must stay loud ──────────────────────────────────
    estimation_sheet = next((spec.sheet_name for spec in all_tests()
                             if spec.scope.value == "estimation"), None)
    if estimation_sheet:
        text = sheet_text(workbook, estimation_sheet)
        check("estimation-badge-present", ESTIMATION_BADGE in text,
              f"{estimation_sheet} lacks the {ESTIMATION_BADGE!r} badge")
        check("estimation-disclaimer-present", "NOT FOR ACCEPTANCE" in text.upper(),
              f"{estimation_sheet} lacks the C805 disclaimer")

    # ── protection must match the claim ──────────────────────────────────
    protected = sum(1 for ws in workbook.worksheets if ws.protection and ws.protection.sheet)
    claimed = bool((manifest.get("protection") or {}).get("enabled"))
    check("protection-matches-manifest", bool(protected) == claimed,
          f"{protected} protected sheet(s) but the manifest claims enabled={claimed}")

    # ── reuse the in-process structural tier so the two never diverge ─────
    if manifest_path.is_file():
        structural = run_structural_checks(xlsx_path, manifest_path, expected_sheets=tuple(expected))
        check("structural-qa-tier", structural.failed == 0, "; ".join(structural.failures[:3]))

    report = {
        "generator": "scripts/audit_excel.py",
        "package_version": __version__,
        "file": xlsx_path.name,
        "size_kb": round(xlsx_path.stat().st_size / 1024, 1),
        "sha256": actual_sha,
        "sheets": names,
        "sheet_count": len(names),
        "formula_cells": len(formulas),
        "protected_sheets": protected,
        "catalogue": {"total": len(all_tests()), "implemented": len(implemented_tests())},
        "passed": sum(1 for finding in findings if finding.ok),
        "failed": sum(1 for finding in findings if not finding.ok),
        "checks": [finding.as_dict() for finding in findings],
    }

    if report_to is not None:
        report_to.parent.mkdir(parents=True, exist_ok=True)
        report_to.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n",
                             encoding="utf-8")
    return findings, report


def resolve_artifact(args: argparse.Namespace) -> Path:
    """Find the workbook to audit from ``--file`` or ``--output``."""
    if args.file:
        return Path(args.file)
    output = Path(args.output)
    expected = output / ARTIFACT_PATTERN.format(version=__version__)
    if expected.is_file():
        return expected
    candidates = sorted(output.glob("*.xlsx"), key=lambda path: path.stat().st_mtime)
    return candidates[-1] if candidates else expected


def main(argv: Optional[List[str]] = None) -> int:
    """CLI entry point."""
    parser = argparse.ArgumentParser(
        prog="audit_excel",
        description="Audit a built workbook against the v4 display-only product contract.",
    )
    parser.add_argument("--output", default="output", help="build output directory (default: ./output)")
    parser.add_argument("--file", default=None, help="audit this specific .xlsx instead")
    parser.add_argument("--report", default=str(REPORT_PATH), help="where to write the JSON report")
    parser.add_argument("--no-report", action="store_true", help="print only; do not write a report")
    args = parser.parse_args(argv)

    xlsx_path = resolve_artifact(args)
    findings, report = audit(xlsx_path, None if args.no_report else Path(args.report))

    print(f"\n{'=' * 70}\n🔬 ARTIFACT AUDIT — {report.get('file', xlsx_path.name)}\n{'=' * 70}")
    for finding in findings:
        mark = "✅" if finding.ok else "❌"
        detail = f" — {finding.detail}" if finding.detail and not finding.ok else ""
        print(f"  {mark} {finding.check}{detail}")
    print(f"\n  {report.get('passed', 0)} passed, {report.get('failed', 0)} failed")
    if not args.no_report:
        print(f"  📄 report: {args.report}")
    return 0 if report.get("failed", 1) == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
