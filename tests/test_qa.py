"""Tests for the QA engine (pytest runner) and structural checks."""

import json
from pathlib import Path
from textwrap import dedent

from openpyxl import Workbook

from concrete_lab.domain.statuses import ValidationStatus
from concrete_lab.qa.engine import QAEngine, QAReport
from concrete_lab.qa.structural import run_structural_checks
from concrete_lab.report.manifest import REQUIRED_MANIFEST_KEYS
from concrete_lab.utils import compute_sha256

PYTEST_OVERRIDES = ("-o", "addopts=", "-p", "no:cacheprovider")


def _write_test_file(tmp_path: Path, name: str, body: str) -> Path:
    path = tmp_path / name
    path.write_text(dedent(body), encoding="utf-8")
    return path


# ─── QAReport ─────────────────────────────────────────────────────────────


class TestQAReport:
    def test_status_precedence(self) -> None:
        base = dict(source="t", passed=1)
        assert QAReport(failed=0, warnings=0, **base).status is ValidationStatus.PASS
        assert QAReport(failed=0, warnings=2, **base).status is ValidationStatus.WARN
        assert QAReport(failed=1, warnings=9, **base).status is ValidationStatus.FAIL

    def test_to_dict_round_trip(self) -> None:
        report = QAReport(source="t", passed=3, failed=1, warnings=2, failures=("boom",))
        data = report.to_dict()
        assert data["total"] == 6
        assert data["status"] == "fail"
        assert data["failures"] == ["boom"]

    def test_merged_reports_add_counters(self) -> None:
        merged = QAReport("a", 2, 1, 1).merged_with(QAReport("b", 3, 0, 2))
        assert (merged.passed, merged.failed, merged.warnings) == (5, 1, 3)


# ─── QAEngine.run_pytest ──────────────────────────────────────────────────


class TestQAEnginePytest:
    def test_passing_suite_reports_pass(self, tmp_path: Path) -> None:
        test_file = _write_test_file(tmp_path, "test_ok.py", """
            def test_one():
                assert 1 + 1 == 2

            def test_two():
                assert True
        """)
        report = QAEngine().run_pytest([str(test_file), *PYTEST_OVERRIDES])
        assert report.failed == 0
        assert report.passed == 2
        assert report.status is ValidationStatus.PASS

    def test_failing_test_reports_fail_with_detail(self, tmp_path: Path) -> None:
        test_file = _write_test_file(tmp_path, "test_bad.py", """
            def test_broken():
                assert 1 == 2, "deliberate failure"
        """)
        report = QAEngine().run_pytest([str(test_file), *PYTEST_OVERRIDES])
        assert report.failed == 1
        assert report.status is ValidationStatus.FAIL
        assert any("deliberate failure" in failure for failure in report.failures)

    def test_crashed_run_is_never_reported_as_pass(self, tmp_path: Path) -> None:
        """Regression: a run that executed nothing used to come back as PASS."""
        report = QAEngine().run_pytest(
            [str(tmp_path / "does_not_exist"), *PYTEST_OVERRIDES])
        assert report.passed == 0
        assert report.failed == 1
        assert report.status is ValidationStatus.FAIL
        assert "exit code" in report.failures[0]

    def test_unusable_arguments_are_reported_as_failure(self) -> None:
        report = QAEngine().run_pytest(["--not-a-real-flag", *PYTEST_OVERRIDES])
        assert report.status is ValidationStatus.FAIL
        assert "exit code" in report.failures[0]

    def test_empty_selection_does_not_claim_success(self, tmp_path: Path) -> None:
        empty = _write_test_file(tmp_path, "empty_module.py", "VALUE = 1\n")
        report = QAEngine().run_pytest([str(empty), *PYTEST_OVERRIDES])
        assert report.passed == 0
        assert report.failed >= 1 or report.notes

    def test_exit_code_is_interpreted_not_ignored(self) -> None:
        from concrete_lab.qa import engine as qa_engine

        assert qa_engine.PYTEST_OK == 0 and qa_engine.PYTEST_TESTS_FAILED == 1
        assert set(qa_engine.PYTEST_EXIT_MEANING) == {2, 3, 4, 5}

    def test_skipped_tests_count_as_warnings(self, tmp_path: Path) -> None:
        test_file = _write_test_file(tmp_path, "test_skip.py", """
            import pytest

            @pytest.mark.skip(reason="not applicable")
            def test_skipped():
                pass

            def test_ok():
                assert True
        """)
        report = QAEngine().run_pytest([str(test_file), *PYTEST_OVERRIDES])
        assert report.passed == 1
        assert report.warnings == 1
        assert report.status is ValidationStatus.WARN


# ─── Structural QA ────────────────────────────────────────────────────────


def _make_artifact(tmp_path: Path, with_formula: bool = False) -> dict:
    """Produce a minimal but valid artifact set for structural checks."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "00_راهنما"
    sheet["A1"] = "Concrete Lab Companion"
    sheet["A2"] = "=SUM(1,2)" if with_formula else "نمایش فقط"
    xlsx_path = tmp_path / "artifact.xlsx"
    workbook.save(str(xlsx_path))

    sha = compute_sha256(xlsx_path)
    (tmp_path / "artifact.xlsx.sha256").write_text(f"{sha}  artifact.xlsx\n", encoding="utf-8")

    manifest_path = tmp_path / "artifact.json"
    manifest_path.write_text(
        json.dumps(dict.fromkeys(REQUIRED_MANIFEST_KEYS, 0)), encoding="utf-8"
    )
    return {"xlsx": xlsx_path, "manifest": manifest_path}


class TestStructuralQA:
    def test_valid_artifact_passes_all_checks(self, tmp_path: Path) -> None:
        artifact = _make_artifact(tmp_path)
        report = run_structural_checks(artifact["xlsx"], artifact["manifest"],
                                       expected_sheets=("00_راهنما",))
        assert report.failed == 0, report.failures
        assert report.passed == 6

    def test_sha_mismatch_is_detected(self, tmp_path: Path) -> None:
        artifact = _make_artifact(tmp_path)
        (tmp_path / "artifact.xlsx.sha256").write_text("0" * 64 + "  artifact.xlsx\n",
                                                       encoding="utf-8")
        report = run_structural_checks(artifact["xlsx"], artifact["manifest"])
        assert any("SHA-256 mismatch" in failure for failure in report.failures)

    def test_formula_leak_is_detected(self, tmp_path: Path) -> None:
        """The renderer's no-formula contract is enforced structurally."""
        artifact = _make_artifact(tmp_path, with_formula=True)
        report = run_structural_checks(artifact["xlsx"], artifact["manifest"])
        assert any("formula" in failure for failure in report.failures)
        assert report.status is ValidationStatus.FAIL

    def test_missing_sheets_are_reported(self, tmp_path: Path) -> None:
        artifact = _make_artifact(tmp_path)
        report = run_structural_checks(artifact["xlsx"], artifact["manifest"],
                                       expected_sheets=("23_داشبورد",))
        assert any("missing sheets" in failure for failure in report.failures)

    def test_missing_workbook_fails_everything(self, tmp_path: Path) -> None:
        artifact = _make_artifact(tmp_path)
        artifact["xlsx"].unlink()
        report = run_structural_checks(artifact["xlsx"], artifact["manifest"])
        assert report.status is ValidationStatus.FAIL
        # only the manifest check survives; every workbook check fails
        assert report.passed == 1
        assert report.failed == 5
