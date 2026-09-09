"""Structural QA: verify built artifacts after the fact.

Checks run against the emitted files (workbook, manifest, checksum) so
the build can prove — not merely claim — that the artifact is intact
and that the renderer kept its no-formula guarantee.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import List, Tuple

import openpyxl

from concrete_lab.qa.engine import QAReport
from concrete_lab.report.manifest import REQUIRED_MANIFEST_KEYS
from concrete_lab.utils import compute_sha256


def run_structural_checks(
    xlsx_path: Path,
    manifest_path: Path,
    expected_sheets: Tuple[str, ...] = (),
) -> QAReport:
    """Run all structural checks over one build artifact set.

    Args:
        xlsx_path: The generated workbook.
        manifest_path: The generated manifest JSON.
        expected_sheets: Worksheet titles that must exist.

    Returns:
        A :class:`QAReport` where each check contributes one pass or
        one failure.
    """
    passed = 0
    failures: List[str] = []

    def check(ok: bool, description: str) -> None:
        nonlocal passed
        if ok:
            passed += 1
        else:
            failures.append(description)

    # 1 — workbook file exists
    check(xlsx_path.exists(), f"workbook missing: {xlsx_path}")

    # 2 — SHA-256 sidecar matches the actual file
    sha_path = xlsx_path.parent / (xlsx_path.name + ".sha256")
    if xlsx_path.exists():
        actual = compute_sha256(xlsx_path)
        recorded = sha_path.read_text(encoding="utf-8").split()[0] if sha_path.exists() else ""
        check(bool(recorded) and recorded == actual, f"SHA-256 mismatch for {xlsx_path.name}")
    else:
        check(False, "SHA-256 check skipped (workbook missing)")

    # 3 — manifest parses and carries the required traceability keys
    required_ok = False
    if manifest_path.exists():
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            required_ok = all(key in manifest for key in REQUIRED_MANIFEST_KEYS)
        except (json.JSONDecodeError, OSError):
            required_ok = False
    check(required_ok, f"manifest missing or incomplete: {manifest_path}")

    # 4..6 — workbook content checks
    if xlsx_path.exists():
        workbook = openpyxl.load_workbook(xlsx_path)
        sheet_names = [ws.title for ws in workbook.worksheets]
        check(len(sheet_names) > 0, "workbook has no sheets")

        missing = [name for name in expected_sheets if name not in sheet_names]
        check(not missing, f"missing sheets: {missing}")

        formula_cells = _find_formula_cells(workbook)
        check(
            not formula_cells,
            f"renderer leaked formulas into {len(formula_cells)} cells "
            f"(first: {formula_cells[:3] if formula_cells else ''})",
        )
    else:
        check(False, "sheet checks skipped (workbook missing)")
        check(False, "expected-sheet check skipped (workbook missing)")
        check(False, "no-formula check skipped (workbook missing)")

    return QAReport(
        source="structural",
        passed=passed,
        failed=len(failures),
        failures=tuple(failures),
    )


def _find_formula_cells(workbook: openpyxl.Workbook) -> List[str]:
    """Return addresses of cells containing formula strings.

    The v4 renderer is display-only, so finding even one formula means
    the separation contract was broken.
    """
    offenders: List[str] = []
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                value = cell.value
                if isinstance(value, str) and value.startswith("="):
                    offenders.append(f"{sheet.title}!{cell.coordinate}")
    return offenders
