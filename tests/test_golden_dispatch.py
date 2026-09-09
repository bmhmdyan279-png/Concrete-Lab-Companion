"""Tests for the golden-suite dispatcher (``concrete_lab.qa``).

``tests/test_golden_corpus.py`` asserts the *corpus* is well formed; these
tests assert the *dispatcher* is honest: that a wrong expectation really
fails, that a mismatch is described usefully, and that the corner cases
(no mapping, missing value, broken tolerance, an implemented test with no
checker) behave as documented instead of silently passing.
"""

import json
from pathlib import Path
from typing import Any, Dict

from concrete_lab.qa import golden
from concrete_lab.qa.golden import (
    ENGINE_CASE_CHECKERS,
    LEGACY_ERRATA_IDS,
    _check_expected,
    _close,
    _pending_note,
    _value_of,
    run_golden_suite,
)


def write_case(directory: Path, name: str, case: Dict[str, Any]) -> Path:
    """Write one golden case file."""
    path = directory / name
    path.write_text(json.dumps(case, ensure_ascii=False), encoding="utf-8")
    return path


GRADING_CASE: Dict[str, Any] = {
    "test_id": "1-1",
    "inputs": {"curve": {"9.5": 97.5, "4.75": 88.0, "2.36": 70.0, "1.18": 50.0,
                         "0.6": 33.0, "0.3": 18.0, "0.15": 6.0, "0.075": 1.0}},
    "expected": {"fineness_modulus": 3.38, "overall_status": "pass"},
    "tolerance": 0.01,
}

COMPRESSIVE_CASE: Dict[str, Any] = {
    "test_id": "4-1",
    "inputs": {"geometry": "cylinder", "diameter_mm": 150, "length_mm": 300, "load_kn": 715},
    "expected": {"ld_ratio": 2.0, "correction_factor": 1.0,
                 "gross_strength_mpa": 40.4607, "reported_strength_mpa": 40.5,
                 "overall_status": "pass"},
    "tolerance": 0.01,
}

REBOUND_CASE: Dict[str, Any] = {
    "test_id": "4-5",
    "inputs": {"readings": [40, 42, 38, 41, 39, 43, 40, 42, 38, 41],
               "correlation_id": "PRJ-X", "correlation_slope": 0.62,
               "correlation_intercept": -8.5},
    "expected": {"mean": 40.4, "retained_mean": 40.4, "outliers": [],
                 "estimated_strength_mpa": 16.548, "overall_status": "warn"},
    "tolerance": 0.01,
}


# ─── Tolerance helper ─────────────────────────────────────────────────────


class TestClose:
    def test_zero_tolerance_means_near_exact(self) -> None:
        assert _close(1.0, 1.0, 0.0)
        assert not _close(1.0, 1.0 + 1e-6, 0.0)

    def test_negative_tolerance_is_treated_as_exact(self) -> None:
        assert _close(2.0, 2.0, -1.0)
        assert not _close(2.0, 2.5, -1.0)

    def test_absolute_tolerance_is_used(self) -> None:
        assert _close(40.46, 40.5, 0.05)
        assert not _close(40.46, 40.5, 0.01)


# ─── Expected-value checking ──────────────────────────────────────────────


class TestCheckExpected:
    def result(self):
        from concrete_lab.domain.engine import calculate

        return calculate("4-1", COMPRESSIVE_CASE["inputs"])

    def test_numeric_match(self) -> None:
        assert _check_expected(self.result(), "reported_strength", 40.5, 0.01, "4-1") is None

    def test_numeric_mismatch_is_described(self) -> None:
        message = _check_expected(self.result(), "reported_strength", 41.0, 0.01, "4-1")
        assert message and "4-1" in message and "reported_strength" in message

    def test_unknown_key_is_reported(self) -> None:
        message = _check_expected(self.result(), "does_not_exist", 1.0, 0.01, "4-1")
        assert message and "no numeric value" in message

    def test_overall_status_is_compared_as_an_enum_value(self) -> None:
        assert _check_expected(self.result(), "overall_status", "pass", 0.0, "4-1") is None
        message = _check_expected(self.result(), "overall_status", "fail", 0.0, "4-1")
        assert message and "overall status" in message

    def test_warnings_min_counts_the_warnings(self) -> None:
        from concrete_lab.domain.engine import calculate

        # 900 kg/m³ is far outside the plausible fresh-density window, so the
        # engine must produce exactly one plausibility warning.
        odd = calculate("2-1", {"apparatus_mass_g": 0, "total_mass_g": 9000,
                                "volume_cm3": 10000})
        assert odd.status.value == "warn" and len(odd.warnings) == 1
        assert _check_expected(odd, "warnings_min", 1, 0.0, "2-1") is None
        message = _check_expected(odd, "warnings_min", 5, 0.0, "2-1")
        assert message and "expected >= 5 warning" in message

    def test_status_row_without_a_quantity_is_not_numeric(self) -> None:
        from concrete_lab.domain.engine import calculate

        result = calculate("1-1", GRADING_CASE["inputs"])
        assert _value_of(result, "grading_overall") is None


# ─── Per-test checkers ────────────────────────────────────────────────────


class TestCheckers:
    def test_grading_case_passes(self) -> None:
        assert ENGINE_CASE_CHECKERS["1-1"](GRADING_CASE) == (True, "")

    def test_grading_percent_passing_is_verified(self) -> None:
        case = json.loads(json.dumps(GRADING_CASE))
        case["expected"]["percent_passing"] = {"4.75": 88.0}
        assert ENGINE_CASE_CHECKERS["1-1"](case) == (True, "")

        case["expected"]["percent_passing"] = {"4.75": 12.0}
        ok, message = ENGINE_CASE_CHECKERS["1-1"](case)
        assert not ok and "passing on sieve 4.75" in message

    def test_grading_unknown_sieve_is_reported(self) -> None:
        case = json.loads(json.dumps(GRADING_CASE))
        case["expected"]["percent_passing"] = {"19.0": 100.0}
        ok, message = ENGINE_CASE_CHECKERS["1-1"](case)
        assert not ok and "no engine value for sieve" in message

    def test_compressive_case_passes(self) -> None:
        assert ENGINE_CASE_CHECKERS["4-1"](COMPRESSIVE_CASE) == (True, "")

    def test_compressive_wrong_strength_fails(self) -> None:
        case = json.loads(json.dumps(COMPRESSIVE_CASE))
        case["expected"]["reported_strength_mpa"] = 20.0
        ok, message = ENGINE_CASE_CHECKERS["4-1"](case)
        assert not ok and "reported_strength" in message

    def test_rebound_case_passes(self) -> None:
        assert ENGINE_CASE_CHECKERS["4-5"](REBOUND_CASE) == (True, "")

    def test_rebound_outlier_count_is_verified(self) -> None:
        case = json.loads(json.dumps(REBOUND_CASE))
        case["expected"]["outliers"] = [1, 2]
        ok, message = ENGINE_CASE_CHECKERS["4-5"](case)
        assert not ok and "outlier" in message

    def test_rebound_estimate_must_match_the_declared_correlation(self) -> None:
        case = json.loads(json.dumps(REBOUND_CASE))
        case["expected"]["estimated_strength_mpa"] = 35.0
        ok, message = ENGINE_CASE_CHECKERS["4-5"](case)
        assert not ok and "estimated_strength" in message

    def test_simple_case_without_a_mapping_is_reported(self) -> None:
        case = {"test_id": "1-2",
                "inputs": {"W1_g": 1050, "W2_g": 1000},
                "expected": {"not_a_known_output": 5.0},
                "tolerance": 0.1}
        ok, message = ENGINE_CASE_CHECKERS["1-2"](case)
        assert not ok and "no engine mapping" in message

    def test_raising_checker_becomes_a_failure_not_a_crash(self, tmp_path: Path) -> None:
        """A QA tier must degrade gracefully: ``failed``, never a traceback."""
        write_case(tmp_path, "4-1.json", {
            "test_id": "4-1",
            "inputs": {"geometry": "cube", "diameter_mm": 150,
                       "length_mm": 150, "load_kn": 715},
            "expected": {"reported_strength_mpa": 40.5},
            "tolerance": 0.01,
        })
        report = run_golden_suite(tmp_path)
        assert report.failed == 1 and report.passed == 0
        assert "raised" in report.failures[0]
        assert "OutOfScopeError" in report.failures[0]


# ─── Pending / warning notes ──────────────────────────────────────────────


class TestPendingNotes:
    def test_pending_test_is_named(self) -> None:
        note = _pending_note(Path("4-3.json"), {"test_id": "4-3"})
        assert "4-3.json" in note and "no engine implementation yet" in note

    def test_historical_erratum_is_named(self) -> None:
        legacy_id = LEGACY_ERRATA_IDS[0]
        note = _pending_note(Path("1-4c.json"), {"test_id": legacy_id})
        assert "erratum" in note and "errata.yaml" in note

    def test_implemented_but_unchecked_is_called_out(self) -> None:
        """The message that would have exposed the 4.0.0 coverage gap."""
        note = _pending_note(Path("1-2.json"), {"test_id": "1-2"})
        assert "IS implemented" in note and "ENGINE_CASE_CHECKERS" in note

    def test_missing_test_id_is_handled(self) -> None:
        assert "?" in _pending_note(Path("x.json"), {})


# ─── Suite-level behaviour ────────────────────────────────────────────────


class TestSuiteBehaviour:
    def test_missing_corpus_is_a_failure_not_a_silent_pass(self, tmp_path: Path) -> None:
        """Regression: an installed package has no ``validation/`` directory,
        the glob matched nothing, and the tier reported PASS having checked
        zero science."""
        report = run_golden_suite(tmp_path / "absent")
        assert report.failed == 1 and report.passed == 0
        assert report.status.value == "fail"
        assert "verified nothing" in report.failures[0]
        assert "does not exist" in report.failures[0]

    def test_empty_corpus_directory_is_a_failure(self, tmp_path: Path) -> None:
        empty = tmp_path / "golden_cases"
        empty.mkdir()
        report = run_golden_suite(empty)
        assert report.failed == 1 and report.passed == 0
        assert "is empty" in report.failures[0]

    def test_non_strict_mode_warns_instead_of_failing(self, tmp_path: Path) -> None:
        """A plain build must still render when the corpus is not installed."""
        report = run_golden_suite(tmp_path / "absent", strict=False)
        assert report.failed == 0 and report.warnings == 1
        assert report.status.value == "warn"
        assert report.notes and "verified nothing" in report.notes[0]

    def test_corpus_availability_probe(self, tmp_path: Path) -> None:
        from concrete_lab.qa.golden import DEFAULT_GOLDEN_DIR, corpus_available

        assert corpus_available() is True          # the repository corpus
        assert corpus_available(DEFAULT_GOLDEN_DIR) is True
        assert corpus_available(tmp_path) is False  # an empty directory

    def test_corpus_of_non_json_files_is_a_failure(self, tmp_path: Path) -> None:
        (tmp_path / "notes.txt").write_text("not a golden case", encoding="utf-8")
        report = run_golden_suite(tmp_path)
        assert report.failed == 1

    def test_unreadable_json_fails_the_tier_without_crashing(self, tmp_path: Path) -> None:
        """Regression: a corrupt case file used to abort the whole build."""
        (tmp_path / "c39_broken.json").write_text("{oops", encoding="utf-8")
        report = run_golden_suite(tmp_path)
        assert report.failed == 1 and report.passed == 0
        assert "unreadable golden case" in report.failures[0]

    def test_non_object_json_fails_the_tier(self, tmp_path: Path) -> None:
        (tmp_path / "c39_list.json").write_text("[1, 2, 3]", encoding="utf-8")
        report = run_golden_suite(tmp_path)
        assert report.failed == 1
        assert "must be a JSON object" in report.failures[0]

    def test_one_bad_file_does_not_hide_the_good_ones(self, tmp_path: Path) -> None:
        (tmp_path / "c39_broken.json").write_text("{oops", encoding="utf-8")
        write_case(tmp_path, "3-1.json", {
            "test_id": "3-1", "inputs": {"h_mm": 200},
            "expected": {"slump_mm": 100}, "tolerance": 1,
        })
        report = run_golden_suite(tmp_path)
        assert report.passed == 1 and report.failed == 1

    def test_corporate_corpus_is_all_green(self) -> None:
        report = run_golden_suite()
        assert report.failed == 0, report.failures
        assert report.source == "golden-suite"
        assert report.passed >= 17

    def test_every_checker_is_reachable_from_the_corpus(self) -> None:
        """No checker may exist only in theory: each needs a real case file."""
        corpus = {json.loads(path.read_text(encoding="utf-8"))["test_id"]
                  for path in golden.DEFAULT_GOLDEN_DIR.glob("*.json")}
        unreachable = sorted(set(ENGINE_CASE_CHECKERS) - corpus)
        assert not unreachable, f"checkers with no golden case: {unreachable}"
