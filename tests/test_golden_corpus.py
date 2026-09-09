"""Golden-corpus tests: the *data* in ``validation/golden_cases`` is code.

Renamed from ``test_formulas.py`` in 4.1.0 — the old name described the v3
product (Excel formulas), and half of the old ``test_golden.py`` was dead
code pinned to a ``v2.1.0`` workbook that no longer exists (two permanent
skips).  These tests replace both.

What is enforced here:

* every golden file is structurally complete and machine-checkable;
* every ``test_id`` resolves to something real — a registered test, a
  standards-level case, or a *declared* historical erratum;
* every test the engine implements has at least one golden case that is
  actually executed against it (no silent degradation to "structure only");
* no golden file smuggles an Excel formula back into a product whose
  headline contract is "the workbook contains zero formula cells".
"""

import json
from pathlib import Path
from typing import Any, Iterator, Tuple

import pytest
import yaml

from concrete_lab.domain.engine import CALCULATORS
from concrete_lab.qa.golden import (
    DEFAULT_GOLDEN_DIR,
    ENGINE_CASE_CHECKERS,
    LEGACY_ERRATA_IDS,
    STANDARDS_CHECKERS,
    run_golden_suite,
)
from concrete_lab.specs.base import TEST_REGISTRY, implemented_tests

REPO_ROOT = Path(__file__).resolve().parents[1]
ERRATA_FILE = REPO_ROOT / "validation" / "errata.yaml"

#: Keys every golden case must carry.
REQUIRED_KEYS: Tuple[str, ...] = ("test_id", "test_name", "standard", "inputs", "expected", "tolerance")


def golden_files() -> Tuple[Path, ...]:
    """Every golden case file, in a stable order."""
    return tuple(sorted(DEFAULT_GOLDEN_DIR.glob("*.json")))


def load(path: Path) -> dict:
    """Load one golden case."""
    return json.loads(path.read_text(encoding="utf-8"))


def iter_strings(node: Any) -> Iterator[str]:
    """Yield every string value nested in a JSON structure."""
    if isinstance(node, str):
        yield node
    elif isinstance(node, dict):
        for key, value in node.items():
            yield str(key)
            yield from iter_strings(value)
    elif isinstance(node, list):
        for item in node:
            yield from iter_strings(item)


def is_standards_case(path: Path) -> bool:
    """``True`` when the file is dispatched by filename prefix (c39_*, …)."""
    return any(path.name.startswith(prefix) for prefix in STANDARDS_CHECKERS)


def engine_level_files() -> Tuple[Path, ...]:
    """Golden files identified by ``test_id`` rather than by filename prefix.

    Kept separate so the identifier tests can be parametrised over exactly
    the files they apply to — a ``pytest.skip`` per standards case would
    only inflate the skip count this suite is meant to drive to zero.
    """
    return tuple(path for path in golden_files() if not is_standards_case(path))


# ─── Corpus structure ─────────────────────────────────────────────────────


class TestCorpusStructure:
    @pytest.mark.golden
    @pytest.mark.parametrize("path", golden_files(), ids=lambda p: p.name)
    def test_required_keys_present(self, path: Path) -> None:
        case = load(path)
        missing = [key for key in REQUIRED_KEYS if key not in case]
        assert not missing, f"{path.name} is missing {missing}"

    @pytest.mark.golden
    @pytest.mark.parametrize("path", golden_files(), ids=lambda p: p.name)
    def test_containers_have_the_right_types(self, path: Path) -> None:
        case = load(path)
        assert isinstance(case["test_id"], str) and case["test_id"].strip()
        assert isinstance(case["test_name"], str) and case["test_name"].strip()
        assert isinstance(case["inputs"], dict), "inputs must be an object"
        assert isinstance(case["expected"], dict), "expected must be an object"
        assert case["expected"], f"{path.name} expects nothing — it cannot fail"
        assert isinstance(case["tolerance"], (int, float))
        assert case["tolerance"] >= 0, "tolerance must not be negative"

    @pytest.mark.golden
    def test_corpus_is_large_enough_to_mean_something(self) -> None:
        assert len(golden_files()) >= 20, f"only {len(golden_files())} golden cases"

    @pytest.mark.golden
    @pytest.mark.parametrize("path", golden_files(), ids=lambda p: p.name)
    def test_no_excel_formula_smuggled_into_the_corpus(self, path: Path) -> None:
        """The v4 contract is a formula-free product; the data must agree."""
        offenders = [s for s in iter_strings(load(path)) if s.startswith("=")]
        assert not offenders, f"{path.name} contains Excel formula string(s): {offenders}"


# ─── Identifiers resolve to something real ────────────────────────────────


class TestIdentifierResolution:
    @pytest.mark.golden
    @pytest.mark.parametrize("path", engine_level_files(), ids=lambda p: p.name)
    def test_test_id_is_known(self, path: Path) -> None:
        """No orphan ids: registered test, standards case, or declared erratum."""
        test_id = load(path)["test_id"]
        assert test_id in TEST_REGISTRY or test_id in LEGACY_ERRATA_IDS, (
            f"{path.name}: test_id {test_id!r} is neither a registered test nor a "
            f"declared legacy erratum (see qa.golden.LEGACY_ERRATA_IDS)"
        )

    @pytest.mark.golden
    def test_declared_errata_ids_exist_in_errata_yaml(self) -> None:
        errata = yaml.safe_load(ERRATA_FILE.read_text(encoding="utf-8")) or {}
        recorded = {str(item.get("id")) for item in errata.get("errata", [])}
        for legacy_id in LEGACY_ERRATA_IDS:
            assert legacy_id in recorded, (
                f"LEGACY_ERRATA_IDS declares {legacy_id!r} but validation/errata.yaml "
                f"has no such entry"
            )

    @pytest.mark.golden
    def test_legacy_errata_files_are_the_only_unregistered_ids(self) -> None:
        unregistered = [
            path.name for path in engine_level_files()
            if load(path)["test_id"] not in TEST_REGISTRY
        ]
        assert unregistered == ["1-4c.json"], (
            f"unexpected orphan golden files: {unregistered}"
        )


# ─── Engine coverage is enforced, not hoped for ───────────────────────────


class TestEngineCoverage:
    @pytest.mark.golden
    def test_every_calculator_has_a_golden_checker(self) -> None:
        """The rule that would have caught the 4.0.0 silent-coverage gap."""
        assert set(ENGINE_CASE_CHECKERS) == set(CALCULATORS), (
            f"checker/calculator mismatch — missing: "
            f"{sorted(set(CALCULATORS) - set(ENGINE_CASE_CHECKERS))}, "
            f"stale: {sorted(set(ENGINE_CASE_CHECKERS) - set(CALCULATORS))}"
        )

    @pytest.mark.golden
    def test_every_implemented_test_has_a_golden_file(self) -> None:
        corpus_ids = {load(path)["test_id"] for path in golden_files()}
        missing = sorted(spec.id for spec in implemented_tests() if spec.id not in corpus_ids)
        assert not missing, f"implemented tests without a golden case: {missing}"

    @pytest.mark.golden
    def test_no_implemented_test_is_left_as_a_warning(self) -> None:
        """A warning is honest only for tests the engine genuinely lacks."""
        report = run_golden_suite()
        unverified = [note for note in report.notes if "IS implemented" in note]
        assert not unverified, f"implemented tests skipped by the golden suite: {unverified}"
        assert len(report.notes) == report.warnings, "every warning must be explained by a note"

    @pytest.mark.golden
    def test_pending_warnings_match_the_unimplemented_catalogue(self) -> None:
        report = run_golden_suite()
        corpus_ids = {load(path)["test_id"] for path in golden_files()}
        pending = {spec.id for spec in TEST_REGISTRY.values() if not spec.implemented}
        warned_files = {note.split(":", 1)[0] for note in report.notes}
        expected_files = {f"{test_id}.json" for test_id in pending if test_id in corpus_ids}
        expected_files.add("1-4c.json")  # declared historical erratum
        assert warned_files == expected_files, (
            f"golden warnings do not match the pending catalogue: "
            f"{sorted(warned_files ^ expected_files)}"
        )


# ─── The suite itself ─────────────────────────────────────────────────────


class TestGoldenSuiteOutcome:
    @pytest.mark.golden
    def test_suite_passes_with_no_failures(self) -> None:
        report = run_golden_suite()
        assert report.failed == 0, report.failures
        assert report.passed >= 17, f"only {report.passed} golden cases executed"

    @pytest.mark.golden
    def test_a_corrupted_case_is_reported_not_crashed(self, tmp_path: Path) -> None:
        """A QA tier must degrade to a failure message, never to a traceback."""
        broken = tmp_path / "c39_broken.json"
        broken.write_text(
            json.dumps({
                "test_id": "C39-X", "test_name": "broken", "standard": "ASTM C39",
                "inputs": {"diameter_mm": 0, "length_mm": 300, "load_kn": 715},
                "expected": {"ld_ratio": 2.0, "correction_factor": 1.0,
                             "gross_strength_mpa": 1.0, "reported_strength_mpa": 1.0},
                "tolerance": 0.01,
            }),
            encoding="utf-8",
        )
        report = run_golden_suite(tmp_path)
        assert report.failed == 1
        assert report.passed == 0
        assert "raised" in report.failures[0]
