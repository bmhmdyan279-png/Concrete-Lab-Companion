"""Golden-suite QA tier: execute golden cases against the engine.

Unlike the pytest tier, this runner is in-process and dependency-light,
so it can be embedded in every build (QA sheet + manifest) without
spawning pytest.  Three dispatch families:

* legacy engine cases (``1-2``, ``2-1``, ``3-1``, ``4-2``) — mapped to
  semantic engine inputs and verified within tolerance;
* standards cases (``c39_*``, ``c805_*``, ``isiri302_*``) — verified
  against the matching ruleset;
* everything else — legacy structure-only cases, counted as warnings.
"""

from __future__ import annotations

import json
import math
from pathlib import Path
from typing import Any, Callable, Dict, List, Tuple

from concrete_lab.domain.engine import calculate
from concrete_lab.qa.engine import QAReport
from concrete_lab.standards.astm import c39, c805
from concrete_lab.standards.isiri import isiri_302

#: Default location of the golden case corpus.
DEFAULT_GOLDEN_DIR: Path = Path(__file__).resolve().parents[3] / "validation" / "golden_cases"

#: Legacy golden files whose inputs the engine can consume, mapped as
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


def _close(value: float, expected: float, tolerance: float) -> bool:
    """Tolerance check that treats zero tolerance as near-exact."""
    if tolerance <= 0:
        tolerance = 1e-9
    return math.isclose(value, expected, abs_tol=tolerance, rel_tol=0.0)


def _check_legacy_engine_case(case: Dict[str, Any]) -> Tuple[bool, str]:
    """Run one legacy engine golden case through the domain engine."""
    test_id = case["test_id"]
    mapping = LEGACY_ENGINE_CASES[test_id]
    inputs = {mapping[key]: value for key, value in case["inputs"].items() if key in mapping}
    tolerance = float(case.get("tolerance", 0.01))

    result = calculate(test_id, inputs)
    for expected_key, expected_value in case["expected"].items():
        value_key = LEGACY_EXPECTED_KEYS.get(expected_key)
        if value_key is None:
            return False, f"{test_id}: no engine mapping for expected key {expected_key!r}"
        match = next((v for v in result.values if v.key == value_key), None)
        if match is None or match.quantity is None:
            return False, f"{test_id}: engine produced no value {value_key!r}"
        if not _close(match.quantity.value, float(expected_value), tolerance):
            return False, (
                f"{test_id}: {value_key} = {match.quantity.value:.6g}, "
                f"expected {expected_value} ± {tolerance}"
            )
    return True, ""


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
    return True, ""


#: Filename prefix → checker for standards-owned golden cases.
_STANDARDS_CHECKERS: Dict[str, Callable[[Dict[str, Any]], Tuple[bool, str]]] = {
    "c39_": _check_c39_case,
    "c805_": _check_c805_case,
    "isiri302_": _check_isiri302_case,
}


def run_golden_suite(golden_dir: Path = DEFAULT_GOLDEN_DIR) -> QAReport:
    """Execute every golden case the engine understands.

    Args:
        golden_dir: Directory of golden case JSON files.

    Returns:
        A :class:`QAReport`; unsupported legacy cases count as warnings
        (structure-only files kept for history), never as failures.
    """
    passed = failed = warnings = 0
    failures: List[str] = []

    for path in sorted(Path(golden_dir).glob("*.json")):
        with open(path, "r", encoding="utf-8") as handle:
            case = json.load(handle)

        checker = next((c for prefix, c in _STANDARDS_CHECKERS.items() if path.name.startswith(prefix)), None)
        if checker is None and case.get("test_id") in LEGACY_ENGINE_CASES:
            checker = _check_legacy_engine_case

        if checker is None:
            warnings += 1  # structure-only legacy case
            continue

        try:
            ok, message = checker(case)
        except Exception as exc:  # noqa: BLE001 — a QA tier must not crash the build
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
    )
