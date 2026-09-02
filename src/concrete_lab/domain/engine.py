"""Domain calculation engine — the single source of scientific truth.

Review feedback drove this layer: Excel formulas used to mix the
scientific model, its implementation and its presentation.  Here the
engine computes everything in Python (testable, golden-verifiable) and
the Excel renderer only *displays* the results.

Supported tests:

* ``1-1`` grading — ISIRI 302 ruleset (envelope + physical rules);
* ``1-2`` moisture — ASTM C566 dry-basis formula;
* ``2-1`` fresh density — ASTM C138 mass/volume;
* ``3-1`` slump — ASTM C143 cone geometry;
* ``4-1`` compressive strength — ASTM C39 ruleset (cylinders only,
  interpolated L/D correction);
* ``4-2`` splitting tensile — ASTM C496 formula;
* ``4-5`` rebound number — ASTM C805 ruleset (ESTIMATION scope,
  mandatory disclaimer, project-specific correlation).
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, Mapping, Optional, Tuple

from concrete_lab import constants
from concrete_lab.domain.quantities import Quantity
from concrete_lab.domain.statuses import ValidationStatus
from concrete_lab.specs.base import TEST_REGISTRY, TestSpec
from concrete_lab.standards import registry as standards_registry
from concrete_lab.standards.astm import c39, c805
from concrete_lab.standards.base import Scope, StandardSpec
from concrete_lab.standards.isiri import isiri_302

#: Inputs mapping type: semantic input key → value.
Inputs = Mapping[str, Any]


class EngineError(ValueError):
    """Raised for invalid engine inputs (missing keys, wrong types)."""


@dataclass(frozen=True)
class ResultValue:
    """One computed/displayed value inside a test result.

    Attributes:
        key: Semantic identifier (e.g. ``"reported_strength"``).
        label: Human-readable label for the renderer.
        quantity: Magnitude + unit; ``None`` for pure-status rows.
        status: Validation status of this individual value.
        note: Optional note shown next to the value.
    """

    key: str
    label: str
    quantity: Optional[Quantity] = None
    status: ValidationStatus = ValidationStatus.PENDING
    note: str = ""


@dataclass(frozen=True)
class TestResult:
    """Complete outcome of one laboratory test computation.

    Attributes:
        test_id: Id of the governing :class:`TestSpec`.
        title: Test title (from the spec).
        standard: Metadata of the governing standard.
        scope: Usage scope; estimation results carry a badge.
        status: Aggregate validation status of the test.
        values: Computed values in display order.
        warnings: Human-readable plausibility warnings.
        disclaimer: Mandatory usage disclaimer (e.g. ASTM C805).
    """

    test_id: str
    title: str
    standard: StandardSpec
    scope: Scope
    status: ValidationStatus
    values: Tuple[ResultValue, ...]
    warnings: Tuple[str, ...] = field(default_factory=tuple)
    disclaimer: str = ""


# ─── Input helpers ────────────────────────────────────────────────────────


def _number(inputs: Inputs, key: str) -> float:
    """Extract a finite numeric input or raise :class:`EngineError`."""
    if key not in inputs:
        raise EngineError(f"missing input {key!r}")
    value = inputs[key]
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value):
        raise EngineError(f"input {key!r} must be a finite number, got {value!r}")
    return float(value)


def _sequence(inputs: Inputs, key: str) -> Tuple[float, ...]:
    """Extract a numeric sequence input or raise :class:`EngineError`."""
    if key not in inputs:
        raise EngineError(f"missing input {key!r}")
    raw = inputs[key]
    try:
        return tuple(float(v) for v in raw)
    except TypeError as exc:
        raise EngineError(f"input {key!r} must be a numeric sequence") from exc


def _plausibility(value: float, window: Tuple[float, float], label: str) -> Tuple[ValidationStatus, Tuple[str, ...]]:
    """Classify a value against a physical plausibility window."""
    low, high = window
    if low <= value <= high:
        return ValidationStatus.PASS, ()
    return ValidationStatus.WARN, (
        f"{label} = {value:.4g} is outside the plausible range [{low:g}, {high:g}] — check inputs",
    )


def _spec(test_id: str) -> TestSpec:
    """Fetch the registered spec for ``test_id`` or raise."""
    try:
        return TEST_REGISTRY[test_id]
    except KeyError:
        raise EngineError(f"unknown test id {test_id!r}") from None


def _result(
    test_id: str,
    values: Tuple[ResultValue, ...],
    status: ValidationStatus,
    warnings: Tuple[str, ...] = (),
    disclaimer: str = "",
) -> TestResult:
    """Assemble a :class:`TestResult` from its registered spec."""
    spec = _spec(test_id)
    return TestResult(
        test_id=test_id,
        title=spec.title,
        standard=standards_registry.get(spec.standard_code),
        scope=spec.scope,
        status=status,
        values=values,
        warnings=warnings,
        disclaimer=disclaimer,
    )


# ─── Calculators (one per implemented test) ───────────────────────────────


def _calc_moisture(inputs: Inputs) -> TestResult:
    """ASTM C566 dry-basis moisture: (W_wet − W_dry) / W_dry × 100."""
    wet = _number(inputs, "wet_mass_g")
    dry = _number(inputs, "dry_mass_g")
    if dry <= 0:
        raise EngineError("dry_mass_g must be positive")
    if wet < dry:
        raise EngineError("wet_mass_g cannot be below dry_mass_g (physical rule)")
    moisture = (wet - dry) / dry * 100.0
    status, warnings = _plausibility(moisture, constants.PLAUSIBLE_MOISTURE_RANGE, "Moisture content")
    return _result(
        "1-2",
        (ResultValue("moisture_percent", "رطوبت سنگدانه", Quantity(moisture, "%"), status),),
        status,
        warnings,
    )


def _calc_fresh_density(inputs: Inputs) -> TestResult:
    """ASTM C138 fresh density: (M_total − M_apparatus) / V."""
    apparatus = _number(inputs, "apparatus_mass_g")
    total = _number(inputs, "total_mass_g")
    volume = _number(inputs, "volume_cm3")
    if volume <= 0:
        raise EngineError("volume_cm3 must be positive")
    if total <= apparatus:
        raise EngineError("total_mass_g must exceed apparatus_mass_g (physical rule)")
    density = (total - apparatus) / volume * 1000.0  # g/cm³ → kg/m³
    status, warnings = _plausibility(density, constants.PLAUSIBLE_DENSITY_RANGE_KG_M3, "Fresh density")
    return _result(
        "2-1",
        (ResultValue("density", "چگالی بتن تازه", Quantity(density, "kg/m³"), status),),
        status,
        warnings,
    )


def _calc_slump(inputs: Inputs) -> TestResult:
    """ASTM C143 slump: 300 mm cone height − measured height."""
    measured = _number(inputs, "measured_height_mm")
    if not 0.0 <= measured <= constants.SLUMP_CONE_HEIGHT_MM:
        raise EngineError(
            f"measured_height_mm must lie within [0, {constants.SLUMP_CONE_HEIGHT_MM:g}], got {measured!r}"
        )
    slump = constants.SLUMP_CONE_HEIGHT_MM - measured
    return _result(
        "3-1",
        (ResultValue("slump", "اسلامپ", Quantity(slump, "mm"), ValidationStatus.PASS),),
        ValidationStatus.PASS,
    )


def _calc_tensile(inputs: Inputs) -> TestResult:
    """ASTM C496 splitting tensile strength: 2P / (π·d·L)."""
    diameter = _number(inputs, "diameter_mm")
    length = _number(inputs, "length_mm")
    load = _number(inputs, "load_n")
    if diameter <= 0 or length <= 0 or load <= 0:
        raise EngineError("diameter_mm, length_mm and load_n must all be positive")
    tensile = (2.0 * load) / (math.pi * diameter * length)
    status, warnings = _plausibility(tensile, constants.PLAUSIBLE_TENSILE_RANGE_MPA, "Tensile strength")
    return _result(
        "4-2",
        (ResultValue("tensile_strength", "مقاومت کششی برزیلی", Quantity(tensile, "MPa"), status),),
        status,
        warnings,
    )


def _calc_compressive(inputs: Inputs) -> TestResult:
    """ASTM C39 compressive strength via the C39-26 ruleset."""
    geometry_name = str(inputs.get("geometry", "cylinder"))
    try:
        geometry = c39.SpecimenGeometry(geometry_name)
    except ValueError:
        raise EngineError(f"unknown geometry {geometry_name!r}") from None

    computed = c39.calculate(
        diameter_mm=_number(inputs, "diameter_mm"),
        length_mm=_number(inputs, "length_mm"),
        load_kn=_number(inputs, "load_kn"),
        geometry=geometry,
    )
    values = (
        ResultValue("area", "سطح مقطع", computed.area, ValidationStatus.PASS),
        ResultValue("ld_ratio", "نسبت L/D", Quantity(computed.ld_ratio, "–"), ValidationStatus.PASS),
        ResultValue("correction_factor", "ضریب اصلاح L/D", Quantity(computed.correction_factor, "–"),
                    ValidationStatus.PASS),
        ResultValue("gross_strength", "مقاومت خام (P/A)", computed.gross_strength, ValidationStatus.PASS),
        ResultValue("reported_strength", "مقاومت گزارش‌شده", computed.reported_strength, ValidationStatus.PASS),
    )
    status = ValidationStatus.WARN if computed.warnings else ValidationStatus.PASS
    return _result("4-1", values, status, computed.warnings)


def _calc_rebound(inputs: Inputs) -> TestResult:
    """ASTM C805 rebound number via the C805-25 ruleset (ESTIMATION)."""
    readings = _sequence(inputs, "readings")
    summary = c805.mean_rebound_number(readings)

    values = [
        ResultValue("mean_rebound", "میانگین عدد بازگشت", Quantity(summary.mean, "–"), ValidationStatus.PASS),
        ResultValue("retained_mean", "میانگین بدون پرتی", Quantity(summary.retained_mean, "–"),
                    ValidationStatus.PASS),
    ]
    if summary.outliers:
        values.append(ResultValue(
            "outliers", "قرائت‌های پرت (بیش از ۷ واحد)", Quantity(float(len(summary.outliers)), "–"),
            ValidationStatus.WARN, note="در محل اندازه‌گیری تکرار شود",
        ))

    correlation_id = inputs.get("correlation_id")
    if correlation_id:
        correlation = c805.LinearCorrelation(
            correlation_id=str(correlation_id),
            slope=_number(inputs, "correlation_slope"),
            intercept=_number(inputs, "correlation_intercept"),
        )
        estimate = c805.estimate_strength(summary.retained_mean, correlation, correlation.correlation_id)
        values.append(ResultValue(
            "estimated_strength", "مقاومت تخمینی (نیازمند همبستگی پروژه)",
            estimate.estimated_strength, ValidationStatus.WARN,
            note=f"correlation: {correlation.correlation_id}",
        ))

    status = ValidationStatus.WARN if summary.outliers or correlation_id else ValidationStatus.PASS
    return _result("4-5", tuple(values), status, disclaimer=c805.DISCLAIMER)


def _calc_grading(inputs: Inputs) -> TestResult:
    """ISIRI 302 grading check via the full ruleset."""
    if "curve" not in inputs:
        raise EngineError("missing input 'curve' (sieve mm → percent passing)")
    curve = {float(size): float(pct) for size, pct in dict(inputs["curve"]).items()}

    evaluation = isiri_302.DEFAULT_RULESET.evaluate(curve)
    values = [
        ResultValue(
            key=f"sieve_{ev.sieve_mm}",
            label=f"الک {ev.sieve_mm:g} میلی‌متر",
            quantity=Quantity(ev.percent_passing, "%"),
            status=ev.status,
            note=f"محدوده [{ev.limit.min_passing:g}, {ev.limit.max_passing:g}]",
        )
        for ev in evaluation.evaluations
    ]
    for missing in evaluation.missing_sieves:
        values.append(ResultValue(
            key=f"sieve_{missing}_missing", label=f"الک {missing:g} میلی‌متر",
            status=ValidationStatus.ERROR, note="داده وارد نشده است",
        ))
    values.append(ResultValue(
        "grading_overall", "نتیجه دانه‌بندی", None, evaluation.overall_status,
        note="بر اساس پاکت دانه‌بندی ISIRI 302",
    ))
    warnings = tuple(evaluation.physical_issues)
    return _result("1-1", tuple(values), evaluation.overall_status, warnings)


#: Dispatch table: test id → calculator.
CALCULATORS: Dict[str, Callable[[Inputs], TestResult]] = {
    "1-1": _calc_grading,
    "1-2": _calc_moisture,
    "2-1": _calc_fresh_density,
    "3-1": _calc_slump,
    "4-1": _calc_compressive,
    "4-2": _calc_tensile,
    "4-5": _calc_rebound,
}


def calculate(test_id: str, inputs: Inputs) -> TestResult:
    """Run the domain engine for one test.

    Args:
        test_id: Registered test identifier (e.g. ``"4-1"``).
        inputs: Semantic input mapping for the test's calculator.

    Returns:
        The computed :class:`TestResult`.

    Raises:
        EngineError: For unknown tests or malformed inputs.
        concrete_lab.standards.astm.c39.OutOfScopeError: When C39
            inputs are outside the standard's scope (e.g. cube).
    """
    try:
        calculator = CALCULATORS[test_id]
    except KeyError:
        raise EngineError(f"test {test_id!r} has no engine implementation (pending)") from None
    return calculator(inputs)


#: Demo inputs used by ``--demo`` builds and by the QA golden dispatch.
DEMO_CASES: Tuple[Tuple[str, Inputs], ...] = (
    ("1-1", {"curve": {9.5: 97.5, 4.75: 88.0, 2.36: 70.0, 1.18: 50.0,
                       0.6: 33.0, 0.3: 18.0, 0.15: 6.0, 0.075: 1.0}}),
    ("1-2", {"wet_mass_g": 1050, "dry_mass_g": 1000}),
    ("2-1", {"apparatus_mass_g": 5000, "total_mass_g": 25000, "volume_cm3": 10000}),
    ("3-1", {"measured_height_mm": 200}),
    ("4-1", {"diameter_mm": 150, "length_mm": 300, "load_kn": 715}),
    ("4-2", {"diameter_mm": 150, "length_mm": 300, "load_n": 358000}),
    ("4-5", {"readings": [42, 44, 41, 43, 45, 40, 43, 42, 44, 43],
             "correlation_id": "demo-linear-correlation",
             "correlation_slope": 0.5,
             "correlation_intercept": -5.0}),
)


def run_demo_cases() -> Tuple[TestResult, ...]:
    """Compute every demo case; the backbone of the ``--demo`` workbook.

    Returns:
        Results in :data:`DEMO_CASES` order.
    """
    return tuple(calculate(test_id, inputs) for test_id, inputs in DEMO_CASES)
