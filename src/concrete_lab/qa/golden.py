"""Golden-suite QA tier: execute golden cases against the engine.

Unlike the pytest tier, this runner is in-process and dependency-light,
so it can be embedded in every build (QA sheet + manifest) without
spawning pytest.

Three dispatch families:

* **engine cases** — one checker per test the domain engine implements
  (:data:`ENGINE_CASE_CHECKERS`).  These run the full product path
  ``spec → ruleset → ResultValue``, not just the ruleset;
* **standards cases** (``c39_*``, ``c805_*``, ``isiri302_*``) — verified
  directly against the matching ruleset;
* **pending cases** — golden files for tests the engine does not
  implement yet.  They are counted as warnings *and named* in
  :attr:`QAReport.notes`, so a warning is always actionable.

Coverage is enforced, not hoped for: a test id present in
:data:`concrete_lab.domain.engine.CALCULATORS` **must** have a checker
here (see ``tests/test_consistency.py``).  Without that rule the suite
silently degrades to "structure-only" the moment someone adds a
calculator — exactly the gap that let three implemented tests go
unverified in 4.0.0.
"""

from __future__ import annotations

import json
import math
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

from concrete_lab.domain.engine import CALCULATORS, Inputs, TestResult, calculate
from concrete_lab.domain.statuses import ValidationStatus
from concrete_lab.qa.engine import QAReport
from concrete_lab.standards.astm import c39, c805
from concrete_lab.standards.isiri import isiri_302

#: Default location of the golden case corpus.
DEFAULT_GOLDEN_DIR: Path = Path(__file__).resolve().parents[3] / "validation" / "golden_cases"

#: Legacy golden files whose inputs are a flat numeric mapping, declared as
#: ``test_id → (legacy key → semantic engine key)``.
LEGACY_ENGINE_CASES: Dict[str, Dict[str, str]] = {
    "1-2": {"W1_g": "wet_mass_g", "W2_g": "dry_mass_g"},
    "2-1": {"Ma_g": "apparatus_mass_g", "Mt_g": "total_mass_g", "V_cm3": "volume_cm3"},
    "3-1": {"h_mm": "measured_height_mm"},
    "4-2": {"d_mm": "diameter_mm", "L_mm": "length_mm", "P_N": "load_n"},
}

#: Legacy expected-output key → engine ResultValue key.
LEGACY_EXPECTED_KEYS: Dict[str, str] = {
    "moisture_percent": "moisture_percent",
    "density_kg_m3": "density",
    "slump_mm": "slump",
    "tensile_strength_mpa": "tensile_strength",
}

#: Golden files kept as historical evidence of a documented book erratum.
#: Their ``test_id`` is not (and must not be) a registered test, so the
#: corpus tests require this marker plus a matching entry in
#: ``validation/errata.yaml``.
LEGACY_ERRATA_IDS: Tuple[str, ...] = ("1-4ج",)


def _close(value: float, expected: float, tolerance: float) -> bool:
    """Tolerance check that treats zero tolerance as near-exact."""
    if tolerance <= 0:
        tolerance = 1e-9
    return math.isclose(value, expected, abs_tol=tolerance, rel_tol=0.0)


def _value_of(result: TestResult, key: str) -> Optional[float]:
    """Return the numeric value of one :class:`ResultValue`, if any."""
    match = next((v for v in result.values if v.key == key), None)
    if match is None or match.quantity is None:
        return None
    return match.quantity.value


def _check_expected(result: TestResult, expected_key: str, expected_value: Any,
                    tolerance: float, test_id: str) -> Optional[str]:
    """Verify one expected entry; return an error message or ``None``."""
    if expected_key == "overall_status":
        if result.status.value != expected_value:
            return f"{test_id}: overall status {result.status.value!r} != {expected_value!r}"
        return None
    if expected_key == "warnings_min":
        if len(result.warnings) < int(expected_value):
            return f"{test_id}: expected >= {expected_value} warning(s), got {len(result.warnings)}"
        return None

    actual = _value_of(result, expected_key)
    if actual is None:
        return f"{test_id}: engine produced no numeric value {expected_key!r}"
    if not _close(actual, float(expected_value), tolerance):
        return f"{test_id}: {expected_key} = {actual:.6g}, expected {expected_value} ± {tolerance}"
    return None


# ─── Engine case checkers (one per implemented test) ──────────────────────


def _check_simple_case(case: Dict[str, Any]) -> Tuple[bool, str]:
    """Run one flat-mapping legacy golden case through the domain engine."""
    test_id = case["test_id"]
    mapping = LEGACY_ENGINE_CASES[test_id]
    inputs: Inputs = {mapping[key]: value for key, value in case["inputs"].items() if key in mapping}
    tolerance = float(case.get("tolerance", 0.01))

    result = calculate(test_id, inputs)
    for expected_key, expected_value in case["expected"].items():
        value_key = LEGACY_EXPECTED_KEYS.get(expected_key)
        if value_key is None:
            return False, f"{test_id}: no engine mapping for expected key {expected_key!r}"
        error = _check_expected(result, value_key, expected_value, tolerance, test_id)
        if error:
            return False, error
    return True, ""


def _check_grading_case(case: Dict[str, Any]) -> Tuple[bool, str]:
    """Verify test 1-1 (ISIRI 302 envelope + ASTM C136 fineness modulus).

    The inputs are what a laboratory actually records — masses retained on
    each sieve — so the mass → percent-passing conversion is exercised too.
    """
    test_id = case["test_id"]
    tolerance = float(case.get("tolerance", 0.01))
    result = calculate(test_id, dict(case["inputs"]))

    for expected_key, expected_value in case["expected"].items():
        if expected_key == "percent_passing":
            for sieve, percent in dict(expected_value).items():
                actual = _value_of(result, f"sieve_{float(sieve):g}")
                if actual is None:
                    return False, f"{test_id}: no engine value for sieve {sieve} mm"
                if not _close(actual, float(percent), tolerance):
                    return False, (
                        f"{test_id}: passing on sieve {sieve} mm = {actual:.6g}, "
                        f"expected {percent} ± {tolerance}"
                    )
            continue
        error = _check_expected(result, expected_key, expected_value, tolerance, test_id)
        if error:
            return False, error
    return True, ""


def _check_compressive_case(case: Dict[str, Any]) -> Tuple[bool, str]:
    """Verify test 4-1 through the engine (spec → C39-26 ruleset → values).

    Expected keys follow the corpus-wide ``*_mpa`` naming; they are mapped
    onto the engine's semantic :class:`ResultValue` keys here so that the
    JSON stays readable and the engine stays free of display concerns.
    """
    return _check_engine_case(case, expected_aliases=_COMPRESSIVE_EXPECTED_KEYS)


def _check_rebound_case(case: Dict[str, Any]) -> Tuple[bool, str]:
    """Verify test 4-5 through the engine (C805-25 ruleset, ESTIMATION scope)."""
    test_id = case["test_id"]
    tolerance = float(case.get("tolerance", 0.01))
    result = calculate(test_id, dict(case["inputs"]))

    if result.scope.value != "estimation":
        return False, f"{test_id}: scope must stay ESTIMATION, got {result.scope.value!r}"
    if not result.disclaimer:
        return False, f"{test_id}: an estimate without the C805 disclaimer is a safety defect"

    for expected_key, expected_value in case["expected"].items():
        if expected_key == "outliers":
            actual = _value_of(result, "outliers") or 0.0
            if int(actual) != len(expected_value):
                return False, f"{test_id}: {int(actual)} outlier(s), expected {expected_value}"
            continue
        value_key = _REBOUND_EXPECTED_KEYS.get(expected_key, expected_key)
        error = _check_expected(result, value_key, expected_value, tolerance, test_id)
        if error:
            return False, error
    return True, ""


def _check_engine_case(
    case: Dict[str, Any],
    input_aliases: Optional[Dict[str, str]] = None,
    expected_aliases: Optional[Dict[str, str]] = None,
) -> Tuple[bool, str]:
    """Generic engine checker: pass inputs through, compare expected values."""
    test_id = case["test_id"]
    tolerance = float(case.get("tolerance", 0.01))
    input_aliases = input_aliases or {}
    expected_aliases = expected_aliases or {}

    inputs = {input_aliases.get(key, key): value for key, value in case["inputs"].items()}
    result = calculate(test_id, inputs)

    for expected_key, expected_value in case["expected"].items():
        value_key = expected_aliases.get(expected_key, expected_key)
        error = _check_expected(result, value_key, expected_value, tolerance, test_id)
        if error:
            return False, error
    return True, ""


#: Corpus key → engine :class:`ResultValue` key for the compressive test.
_COMPRESSIVE_EXPECTED_KEYS: Dict[str, str] = {
    "gross_strength_mpa": "gross_strength",
    "reported_strength_mpa": "reported_strength",
    "area_mm2": "area",
}

#: Corpus key → engine :class:`ResultValue` key for the rebound test.
_REBOUND_EXPECTED_KEYS: Dict[str, str] = {
    "mean": "mean_rebound",
    "retained_mean": "retained_mean",
    "estimated_strength_mpa": "estimated_strength",
}


#: ``test_id`` → checker for every test the engine implements.  Adding a
#: calculator without adding a checker here fails ``tests/test_consistency.py``.
ENGINE_CASE_CHECKERS: Dict[str, Callable[[Dict[str, Any]], Tuple[bool, str]]] = {
    "1-1": _check_grading_case,
    "1-2": _check_simple_case,
    "2-1": _check_simple_case,
    "3-1": _check_simple_case,
    "4-1": _check_compressive_case,
    "4-2": _check_simple_case,
    "4-5": _check_rebound_case,
}


# ─── Standards case checkers ─────────────────────────────────────────────


def _check_c39_case(case: Dict[str, Any]) -> Tuple[bool, str]:
    """Verify one C39 golden case against the C39-26 ruleset."""
    inputs, expected, tolerance = case["inputs"], case["expected"], float(case["tolerance"])
    result = c39.calculate(inputs["diameter_mm"], inputs["length_mm"], inputs["load_kn"])
    if not _close(result.ld_ratio, expected["ld_ratio"], 1e-9):
        return False, f"C39: L/D {result.ld_ratio} != {expected['ld_ratio']}"
    if not _close(result.correction_factor, expected["correction_factor"], 1e-6):
        return False, f"C39: factor {result.correction_factor} != {expected['correction_factor']}"
    if not _close(result.gross_strength.value, expected["gross_strength_mpa"], tolerance):
        return False, f"C39: gross {result.gross_strength.value:.4f} != {expected['gross_strength_mpa']}"
    if not _close(result.reported_strength.value, expected["reported_strength_mpa"], 1e-9):
        return False, f"C39: reported {result.reported_strength.value} != {expected['reported_strength_mpa']}"
    return True, ""


def _check_c805_case(case: Dict[str, Any]) -> Tuple[bool, str]:
    """Verify one C805 golden case against the C805-25 ruleset."""
    inputs, expected, tolerance = case["inputs"], case["expected"], float(case["tolerance"])
    summary = c805.mean_rebound_number(tuple(inputs["readings"]))
    if not _close(summary.mean, expected["mean"], tolerance):
        return False, f"C805: mean {summary.mean} != {expected['mean']}"
    if list(summary.outliers) != expected["outliers"]:
        return False, f"C805: outliers {summary.outliers} != {expected['outliers']}"
    if not _close(summary.retained_mean, expected["retained_mean"], tolerance):
        return False, f"C805: retained {summary.retained_mean} != {expected['retained_mean']}"
    if "estimate" in inputs:
        spec = inputs["estimate"]
        correlation = c805.LinearCorrelation(spec["correlation_id"], spec["slope"], spec["intercept"])
        estimate = c805.estimate_strength(summary.retained_mean, correlation, correlation.correlation_id)
        if not _close(estimate.estimated_strength.value, expected["estimated_strength_mpa"], tolerance):
            return False, f"C805: estimate {estimate.estimated_strength.value} != {expected['estimated_strength_mpa']}"
    return True, ""


def _check_isiri302_case(case: Dict[str, Any]) -> Tuple[bool, str]:
    """Verify one ISIRI 302 golden case against the grading ruleset."""
    curve = {float(size): pct for size, pct in case["inputs"]["curve"].items()}
    expected = case["expected"]
    evaluation = isiri_302.DEFAULT_RULESET.evaluate(curve)
    if evaluation.overall_status.value != expected["overall_status"]:
        return False, f"ISIRI302: overall {evaluation.overall_status.value} != {expected['overall_status']}"
    if "statuses" in expected:
        got = {str(ev.sieve_mm): ev.status.value for ev in evaluation.evaluations}
        if got != expected["statuses"]:
            return False, f"ISIRI302: per-sieve statuses differ: {got}"
    minimum = expected.get("physical_issue_count_min")
    if minimum is not None and len(evaluation.physical_issues) < int(minimum):
        return False, (
            f"ISIRI302: expected >= {minimum} physical issue(s), "
            f"got {evaluation.physical_issues}"
        )
    return True, ""


#: Filename prefix → checker for standards-owned golden cases.
STANDARDS_CHECKERS: Dict[str, Callable[[Dict[str, Any]], Tuple[bool, str]]] = {
    "c39_": _check_c39_case,
    "c805_": _check_c805_case,
    "isiri302_": _check_isiri302_case,
}


def _checker_for(path: Path, case: Dict[str, Any]) -> Optional[Callable[[Dict[str, Any]], Tuple[bool, str]]]:
    """Select the checker for one golden file, or ``None`` when pending."""
    for prefix, checker in STANDARDS_CHECKERS.items():
        if path.name.startswith(prefix):
            return checker
    return ENGINE_CASE_CHECKERS.get(str(case.get("test_id", "")))


def _pending_note(path: Path, case: Dict[str, Any]) -> str:
    """Explain *why* a golden file was not executed — warnings must be actionable."""
    test_id = str(case.get("test_id", "?"))
    if test_id in LEGACY_ERRATA_IDS:
        return (f"{path.name}: historical erratum {test_id!r} "
                f"(see validation/errata.yaml) — kept as evidence, not an engine case")
    if test_id in CALCULATORS:
        return (f"{path.name}: test {test_id} IS implemented but has no golden checker — "
                f"add one to ENGINE_CASE_CHECKERS")
    return (f"{path.name}: test {test_id} has no engine implementation yet — "
            f"golden case kept as a structural record of the intended contract")


def corpus_available(golden_dir: Path = DEFAULT_GOLDEN_DIR) -> bool:
    """Return ``True`` when at least one golden case can be read.

    The corpus lives in the repository, not inside the installed package, so
    a plain ``pip install`` has no ``validation/`` directory.  Callers use
    this to distinguish "nothing to verify here" from "verification failed".
    """
    return bool(sorted(Path(golden_dir).glob("*.json")))


def run_golden_suite(golden_dir: Path = DEFAULT_GOLDEN_DIR, strict: bool = True) -> QAReport:
    """Execute every golden case the engine understands.

    Args:
        golden_dir: Directory of golden case JSON files.
        strict: What to do when no golden case can be found at all.  A tier
            that verified nothing must never report success, so by default
            this is a **failure**.  ``--validate`` keeps that behaviour.
            A plain build passes ``strict=False``: an installed package
            legitimately has no corpus, and refusing to render a workbook
            because of it would be worse than reporting the gap loudly as a
            warning (which the manifest and the QA sheet then carry).

    Returns:
        A :class:`QAReport`.  Unsupported cases count as warnings and are
        individually named in :attr:`QAReport.notes`; an implemented test
        without a checker is reported as a failure, never a warning.
    """
    golden_dir = Path(golden_dir)
    case_files = sorted(golden_dir.glob("*.json"))

    if not case_files:
        reason = "directory does not exist" if not golden_dir.is_dir() else "directory is empty"
        message = (
            f"no golden cases at {golden_dir} ({reason}) — this tier verified nothing; "
            f"run from a repository checkout to verify the science"
        )
        if strict:
            return QAReport(source="golden-suite", passed=0, failed=1, failures=(message,))
        return QAReport(source="golden-suite", passed=0, failed=0, warnings=1, notes=(message,))

    passed = failed = warnings = 0
    failures: List[str] = []
    notes: List[str] = []

    for path in case_files:
        try:
            with open(path, encoding="utf-8") as handle:
                case = json.load(handle)
        except (OSError, ValueError, UnicodeDecodeError) as exc:
            # An unreadable case file is a failure to report, never a reason
            # to abort the build: this tier promises it cannot crash.
            failed += 1
            failures.append(f"{path.name}: unreadable golden case ({exc})")
            continue

        if not isinstance(case, dict):
            failed += 1
            failures.append(f"{path.name}: a golden case must be a JSON object, got "
                            f"{type(case).__name__}")
            continue

        checker = _checker_for(path, case)
        if checker is None:
            note = _pending_note(path, case)
            if str(case.get("test_id")) in CALCULATORS:
                # Implemented but unverified: a coverage hole, not a pending test.
                failed += 1
                failures.append(note)
            else:
                warnings += 1
                notes.append(note)
            continue

        try:
            ok, message = checker(case)
        except Exception as exc:
            ok, message = False, f"{path.name}: raised {exc!r}"
        if ok:
            passed += 1
        else:
            failed += 1
            failures.append(message)

    return QAReport(
        source="golden-suite",
        passed=passed,
        failed=failed,
        warnings=warnings,
        failures=tuple(failures),
        notes=tuple(notes),
    )


#: Statuses a golden case may expect, so corpus tests can validate files
#: without importing the renderer.
VALID_STATUS_VALUES: Tuple[str, ...] = tuple(status.value for status in ValidationStatus)
