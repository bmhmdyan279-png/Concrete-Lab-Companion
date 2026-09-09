"""Tests for ``scripts/audit_excel.py`` — the release gate.

A gate that cries wolf gets ignored, and a gate that stays silent is worse
than no gate.  The previous version of this script did both: it demanded
named ranges, ``calcMode == auto``, data validations and conditional
formatting from a workbook that v4 deliberately renders with none of them,
so a *correct* artifact scored four "critical issues".  These tests pin the
v4 contract instead, including the false-positive that the rewrite had to
avoid (the golden suite legitimately writes the word "placeholder").
"""

import importlib.util
import json
from pathlib import Path
from typing import List

import openpyxl
import pytest
from openpyxl import Workbook

REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT = REPO_ROOT / "scripts" / "audit_excel.py"


def load_script():
    """Import the audit script as a module (it lives outside the package)."""
    spec = importlib.util.spec_from_file_location("audit_excel_under_test", SCRIPT)
    module = importlib.util.module_from_spec(spec)
    assert spec and spec.loader
    spec.loader.exec_module(module)
    return module


@pytest.fixture(scope="module")
def auditor():
    """The audit module, imported once per session."""
    return load_script()


@pytest.fixture(scope="module")
def real_build(tmp_path_factory) -> Path:
    """A genuine build, produced once and audited by several tests."""
    import subprocess
    import sys

    out = tmp_path_factory.mktemp("real_output")
    result = subprocess.run(
        [sys.executable, str(REPO_ROOT / "build.py"), "--output", str(out), "--no-protect"],
        cwd=str(REPO_ROOT), capture_output=True, text=True, timeout=300,
    )
    assert result.returncode == 0, result.stderr
    from concrete_lab import __version__
    from concrete_lab.cli import ARTIFACT_PATTERN

    return out / ARTIFACT_PATTERN.format(version=__version__)


def names_of(findings: List) -> dict:
    """Map each finding's check name to its verdict."""
    return {finding.check: finding.ok for finding in findings}


# ── Detection helpers ─────────────────────────────────────────────────────


class TestFormulaDetection:
    def test_detects_a_formula_by_prefix(self, auditor) -> None:
        workbook = Workbook()
        workbook.active["A1"] = "=SUM(1,2)"
        assert auditor.find_formula_cells(workbook) == ["Sheet!A1"]

    def test_detects_a_formula_by_openpyxl_data_type(self, auditor) -> None:
        """The stronger signal: openpyxl tags a formula cell with ``data_type == 'f'``."""
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "محاسبه"
        sheet["B2"] = "=IF(A2>0,1,0)"
        assert auditor.find_formula_cells(workbook)
        assert sheet["B2"].data_type == "f"

    def test_clean_workbook_reports_nothing(self, auditor) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet["A1"] = "نمایش فقط"
        sheet["A2"] = 42.5
        sheet["A3"] = None
        assert auditor.find_formula_cells(workbook) == []

    def test_text_that_merely_starts_with_an_equals_sign_is_not_a_formula(self, auditor) -> None:
        """A display string is not a formula, and must not be reported as one."""
        workbook = Workbook()
        workbook.active["A1"] = "a = b"
        assert auditor.find_formula_cells(workbook) == []


class TestPlaceholderDetection:
    @pytest.mark.parametrize("text", ["TODO", "TBD", "FIXME", "XXX", "HACK",
                                      "{{version}}", "<name_placeholder>"])
    def test_markers_are_caught(self, auditor, text: str) -> None:
        workbook = Workbook()
        workbook.active["A1"] = f"release {text} soon"
        assert auditor.find_placeholders(workbook)

    @pytest.mark.parametrize("text", [
        "golden case kept as a structural placeholder",  # honest prose, not a marker
        "todo list of tests",                            # lowercase prose
        "مقدار تخمینی — برای پذیرش قابل استناد نیست",
    ])
    def test_ordinary_prose_is_not_a_false_positive(self, auditor, text: str) -> None:
        """Regression: this exact sentence used to fail the release gate."""
        workbook = Workbook()
        workbook.active["A1"] = text
        assert auditor.find_placeholders(workbook) == []


class TestSheetInventory:
    def test_expected_names_come_from_the_live_catalogue(self, auditor) -> None:
        from concrete_lab.specs.base import implemented_tests

        names = auditor.expected_sheet_names()
        for spec in implemented_tests():
            assert spec.sheet_name in names
        assert names[0] == "00_راهنما"
        assert "_Standards" in names

    def test_a_real_build_matches_the_expected_inventory(self, auditor, real_build: Path) -> None:
        workbook = openpyxl.load_workbook(str(real_build))
        assert set(workbook.sheetnames) == set(auditor.expected_sheet_names())


# ── Whole-artifact audit ──────────────────────────────────────────────────


class TestAuditOfARealBuild:
    def test_a_correct_artifact_passes_every_check(self, auditor, real_build: Path,
                                                   tmp_path: Path) -> None:
        findings, report = auditor.audit(real_build, tmp_path / "report.json")
        failed = [f.check for f in findings if not f.ok]
        assert not failed, f"the released artifact failed: {failed}"
        assert report["failed"] == 0 and report["passed"] == len(findings)

    def test_report_is_written_and_complete(self, auditor, real_build: Path,
                                            tmp_path: Path) -> None:
        dest = tmp_path / "nested" / "report.json"
        _findings, report = auditor.audit(real_build, dest)
        assert dest.is_file()
        on_disk = json.loads(dest.read_text(encoding="utf-8"))
        assert on_disk == report
        for key in ("generator", "package_version", "file", "sha256", "sheets",
                    "formula_cells", "catalogue", "checks"):
            assert key in on_disk
        assert on_disk["formula_cells"] == 0

    def test_no_report_is_written_when_disabled(self, auditor, real_build: Path) -> None:
        _findings, report = auditor.audit(real_build, None)
        assert report["failed"] == 0


class TestAuditDetectsDefects:
    def _make(self, auditor, tmp_path: Path, *, with_formula: bool = False,
              with_placeholder: bool = False, with_manifest: bool = True,
              sha_mismatch: bool = False) -> Path:
        """Build a minimal artifact set with one injected defect."""
        from concrete_lab import __version__
        from concrete_lab.cli import ARTIFACT_PATTERN
        from concrete_lab.utils import compute_sha256

        xlsx = tmp_path / ARTIFACT_PATTERN.format(version=__version__)
        workbook = Workbook()
        for name in auditor.expected_sheet_names():
            sheet = workbook.create_sheet(title=name)
            sheet["A1"] = name
        workbook.remove(workbook["Sheet"])  # drop the implicit default sheet first
        if with_formula:
            workbook[workbook.sheetnames[0]]["Z9"] = "=1+1"
        if with_placeholder:
            workbook[workbook.sheetnames[0]]["Z8"] = "TODO finish this"
        workbook.save(str(xlsx))

        sha = "0" * 64 if sha_mismatch else compute_sha256(xlsx)
        (tmp_path / (xlsx.name + ".sha256")).write_text(f"{sha}  {xlsx.name}\n", encoding="utf-8")
        if with_manifest:
            (tmp_path / (xlsx.stem + ".json")).write_text(json.dumps({
                "application_version": __version__, "filename": xlsx.name, "sha256": sha,
                "protection": {"enabled": False},
            }), encoding="utf-8")
        return xlsx

    def test_formula_leak_fails_the_gate(self, auditor, tmp_path: Path) -> None:
        findings, report = auditor.audit(self._make(auditor, tmp_path, with_formula=True), None)
        verdicts = names_of(findings)
        assert verdicts["zero-formula-cells"] is False
        assert report["formula_cells"] == 1

    def test_unfinished_marker_fails_the_gate(self, auditor, tmp_path: Path) -> None:
        findings, _ = auditor.audit(self._make(auditor, tmp_path, with_placeholder=True), None)
        assert names_of(findings)["no-placeholder-text"] is False

    def test_checksum_mismatch_fails_the_gate(self, auditor, tmp_path: Path) -> None:
        findings, _ = auditor.audit(self._make(auditor, tmp_path, sha_mismatch=True), None)
        verdicts = names_of(findings)
        assert verdicts["sha256-sidecar-matches"] is False
        assert verdicts["manifest-sha-matches"] is False

    def test_missing_artifact_fails_immediately(self, auditor, tmp_path: Path) -> None:
        findings, report = auditor.audit(tmp_path / "absent.xlsx", None)
        assert names_of(findings)["artifact-exists"] is False
        assert report["failed"] == 1

    def test_missing_manifest_is_reported(self, auditor, tmp_path: Path) -> None:
        findings, _ = auditor.audit(self._make(auditor, tmp_path, with_manifest=False), None)
        assert names_of(findings)["manifest-present"] is False


# ── CLI ───────────────────────────────────────────────────────────────────


class TestAuditCommandLine:
    def test_exit_code_is_zero_only_when_clean(self, auditor, real_build: Path,
                                               tmp_path: Path) -> None:
        code = auditor.main(["--file", str(real_build), "--no-report"])
        assert code == 0

    def test_exit_code_is_one_for_a_defective_artifact(self, auditor, tmp_path: Path) -> None:
        bad = tmp_path / "absent.xlsx"
        assert auditor.main(["--file", str(bad), "--no-report"]) == 1

    def test_resolves_the_current_version_from_an_output_directory(
        self, auditor, real_build: Path
    ) -> None:
        class Args:
            file = None
            output = str(real_build.parent)

        assert auditor.resolve_artifact(Args()) == real_build

    def test_prefers_an_explicit_file(self, auditor, tmp_path: Path) -> None:
        class Args:
            file = str(tmp_path / "chosen.xlsx")
            output = "output"

        assert auditor.resolve_artifact(Args()) == tmp_path / "chosen.xlsx"

    def test_falls_back_to_the_newest_workbook(self, auditor, tmp_path: Path) -> None:
        import os
        import time

        first = tmp_path / "zzz_old.xlsx"
        second = tmp_path / "aaa_new.xlsx"
        first.write_bytes(b"x")
        second.write_bytes(b"y")
        now = time.time()
        os.utime(first, (now - 3600, now - 3600))
        os.utime(second, (now, now))

        class Args:
            file = None
            output = str(tmp_path)

        assert auditor.resolve_artifact(Args()) == second
