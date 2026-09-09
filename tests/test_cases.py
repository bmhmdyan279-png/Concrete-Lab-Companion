"""Tests for user-supplied case files (``concrete_lab.cases``).

Until 4.1 the engine could only compute its own demo dataset, which made the
published workbook a sample report rather than a tool.  These tests cover the
validation that stands between a technician's JSON file and the domain
engine: every rejection must be a *sentence a human can act on*, never a
traceback, and nothing invalid may reach the science layer.
"""

import json
from pathlib import Path
from typing import Any, Dict

import pytest

from concrete_lab.cases import CaseFileError, implemented_ids, load_case_file
from concrete_lab.domain.engine import CALCULATORS, run_batch
from concrete_lab.specs.base import all_tests

VALID_CASE: Dict[str, Any] = {
    "test_id": "1-2",
    "inputs": {"wet_mass_g": 1050, "dry_mass_g": 1000},
}


def write(tmp_path: Path, payload: Any, name: str = "cases.json") -> Path:
    """Write a case file and return its path."""
    path = tmp_path / name
    path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    return path


# ─── Happy paths ──────────────────────────────────────────────────────────


class TestLoading:
    def test_minimal_file_loads(self, tmp_path: Path) -> None:
        path = write(tmp_path, {"cases": [VALID_CASE]})
        assert load_case_file(path) == (("1-2", {"wet_mass_g": 1050, "dry_mass_g": 1000}),)

    def test_order_is_preserved(self, tmp_path: Path) -> None:
        cases = [
            {"test_id": "4-1", "inputs": {"diameter_mm": 150, "length_mm": 300, "load_kn": 715}},
            {"test_id": "3-1", "inputs": {"measured_height_mm": 200}},
            {"test_id": "1-2", "inputs": {"wet_mass_g": 1050, "dry_mass_g": 1000}},
        ]
        loaded = load_case_file(write(tmp_path, {"cases": cases}))
        assert [test_id for test_id, _ in loaded] == ["4-1", "3-1", "1-2"]

    def test_project_block_is_optional(self, tmp_path: Path) -> None:
        assert load_case_file(write(tmp_path, {"cases": [VALID_CASE]}))

    def test_whitespace_around_ids_is_tolerated(self, tmp_path: Path) -> None:
        loaded = load_case_file(write(tmp_path, {"cases": [dict(VALID_CASE, test_id="  1-2 ")]}))
        assert loaded[0][0] == "1-2"

    def test_retained_mass_form_is_accepted(self, tmp_path: Path) -> None:
        """The 1-1 file may carry weighed masses instead of percentages."""
        case = {
            "test_id": "1-1",
            "inputs": {
                "total_mass_g": 1000, "pan_g": 10,
                "retained": {"9.5": 25, "4.75": 95, "2.36": 180, "1.18": 200,
                             "0.6": 170, "0.3": 150, "0.15": 120, "0.075": 50},
            },
        }
        results = run_batch(load_case_file(write(tmp_path, {"cases": [case]}))).results
        fineness = next(v for v in results[0].values if v.key == "fineness_modulus")
        assert fineness.quantity is not None
        assert fineness.quantity.value == pytest.approx(3.38)

    def test_shipped_example_file_is_valid_and_computes(self) -> None:
        """``examples/sample_input.json`` is documentation: it must actually run."""
        repo_root = Path(__file__).resolve().parents[1]
        outcome = run_batch(load_case_file(repo_root / "examples" / "sample_input.json"))
        assert outcome.ok, outcome.errors
        assert len(outcome.results) == len(implemented_ids())


# ─── Rejections ───────────────────────────────────────────────────────────


class TestRejections:
    def test_missing_file(self, tmp_path: Path) -> None:
        with pytest.raises(CaseFileError, match="not found"):
            load_case_file(tmp_path / "absent.json")

    def test_directory_is_not_a_file(self, tmp_path: Path) -> None:
        with pytest.raises(CaseFileError, match="not found"):
            load_case_file(tmp_path)

    def test_invalid_json(self, tmp_path: Path) -> None:
        path = tmp_path / "broken.json"
        path.write_text("{not json", encoding="utf-8")
        with pytest.raises(CaseFileError, match="not valid JSON"):
            load_case_file(path)

    def test_top_level_must_be_an_object(self, tmp_path: Path) -> None:
        with pytest.raises(CaseFileError, match="JSON object at the top level"):
            load_case_file(write(tmp_path, [VALID_CASE]))

    @pytest.mark.parametrize("payload", [{}, {"cases": "nope"}, {"cases": None}])
    def test_cases_key_is_required_and_must_be_a_list(self, tmp_path: Path, payload: Any) -> None:
        with pytest.raises(CaseFileError, match="must define a 'cases' list"):
            load_case_file(write(tmp_path, payload))

    def test_empty_case_list_is_refused(self, tmp_path: Path) -> None:
        with pytest.raises(CaseFileError, match="nothing to compute"):
            load_case_file(write(tmp_path, {"cases": []}))

    @pytest.mark.parametrize("case,match", [
        ("not-an-object", "must be an object"),
        ({"inputs": {}}, "non-empty string 'test_id'"),
        ({"test_id": 42, "inputs": {}}, "non-empty string 'test_id'"),
        ({"test_id": "   ", "inputs": {}}, "non-empty string 'test_id'"),
        ({"test_id": "1-2"}, "needs an 'inputs' object"),
        ({"test_id": "1-2", "inputs": []}, "needs an 'inputs' object"),
    ])
    def test_malformed_cases(self, tmp_path: Path, case: Any, match: str) -> None:
        with pytest.raises(CaseFileError, match=match):
            load_case_file(write(tmp_path, {"cases": [case]}))

    def test_unimplemented_test_names_the_available_ones(self, tmp_path: Path) -> None:
        """A pending test must be refused, never silently dropped or faked."""
        pending = next(spec.id for spec in all_tests() if not spec.implemented)
        with pytest.raises(CaseFileError) as info:
            load_case_file(write(tmp_path, {"cases": [{"test_id": pending, "inputs": {}}]}))
        message = str(info.value)
        assert pending in message and "not implemented" in message
        for available in implemented_ids():
            assert available in message, "the error must list what *is* available"

    def test_unknown_test_is_refused(self, tmp_path: Path) -> None:
        with pytest.raises(CaseFileError, match="not implemented"):
            load_case_file(write(tmp_path, {"cases": [{"test_id": "9-9", "inputs": {}}]}))

    def test_duplicate_test_is_refused(self, tmp_path: Path) -> None:
        """Each computed test renders to its own worksheet, so ids must be unique."""
        with pytest.raises(CaseFileError, match="more than once"):
            load_case_file(write(tmp_path, {"cases": [VALID_CASE, dict(VALID_CASE)]}))

    def test_every_problem_is_reported_at_once(self, tmp_path: Path) -> None:
        """A bench user should not have to fix one error, re-run, fix the next."""
        payload = {"cases": [
            {"test_id": "9-9", "inputs": {}},
            {"inputs": {}},
            dict(VALID_CASE),
            dict(VALID_CASE),
        ]}
        with pytest.raises(CaseFileError) as info:
            load_case_file(write(tmp_path, payload))
        message = str(info.value)
        assert "3 problem(s)" in message
        assert "cases[0]" in message and "cases[1]" in message and "cases[3]" in message

    def test_project_block_must_be_an_object(self, tmp_path: Path) -> None:
        with pytest.raises(CaseFileError, match="'project' must be an object"):
            load_case_file(write(tmp_path, {"project": "nope", "cases": [VALID_CASE]}))


# ─── Engine batch behaviour ───────────────────────────────────────────────


class TestBatchOutcome:
    def test_all_valid_cases_compute(self) -> None:
        outcome = run_batch((("1-2", {"wet_mass_g": 1050, "dry_mass_g": 1000}),))
        assert outcome.ok and len(outcome.results) == 1

    def test_every_rejection_is_collected_not_raised(self) -> None:
        """Domain errors derive from ValueError by contract — one net catches all."""
        outcome = run_batch((
            ("1-2", {"wet_mass_g": 900, "dry_mass_g": 1000}),   # wet < dry: EngineError
            ("3-1", {"measured_height_mm": 400}),               # > cone height: EngineError
            ("4-1", {"geometry": "cube", "diameter_mm": 150,    # OutOfScopeError
                     "length_mm": 300, "load_kn": 715}),
            ("1-2", {"wet_mass_g": 1050, "dry_mass_g": 1000}),  # the one good case
        ))
        assert not outcome.ok
        assert len(outcome.errors) == 3
        assert len(outcome.results) == 1
        assert all(error.split(":", 1)[0] in {"1-2", "3-1", "4-1"} for error in outcome.errors)

    def test_error_messages_name_the_test(self) -> None:
        outcome = run_batch((("3-1", {"measured_height_mm": 400}),))
        assert outcome.errors[0].startswith("3-1:")

    def test_implemented_ids_match_the_calculators(self) -> None:
        assert implemented_ids() == tuple(sorted(CALCULATORS))
