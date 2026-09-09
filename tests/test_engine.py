"""Unit tests for the domain calculation engine and golden dispatch."""

import pytest

from concrete_lab.domain.engine import (
    CALCULATORS,
    DEMO_CASES,
    EngineError,
    calculate,
    run_demo_cases,
)
from concrete_lab.domain.statuses import ValidationStatus
from concrete_lab.qa.golden import run_golden_suite
from concrete_lab.standards.astm import c39
from concrete_lab.standards.base import Scope

# ─── Simple calculators ───────────────────────────────────────────────────


class TestMoisture:
    def test_golden_value(self) -> None:
        result = calculate("1-2", {"wet_mass_g": 1050, "dry_mass_g": 1000})
        assert result.values[0].quantity.value == pytest.approx(5.0)
        assert result.status is ValidationStatus.PASS

    def test_dry_zero_rejected(self) -> None:
        with pytest.raises(EngineError):
            calculate("1-2", {"wet_mass_g": 100, "dry_mass_g": 0})

    def test_wet_below_dry_is_physical_error(self) -> None:
        with pytest.raises(EngineError, match="physical rule"):
            calculate("1-2", {"wet_mass_g": 900, "dry_mass_g": 1000})

    def test_implausible_moisture_warns(self) -> None:
        result = calculate("1-2", {"wet_mass_g": 1300, "dry_mass_g": 1000})
        assert result.status is ValidationStatus.WARN
        assert result.warnings


class TestFreshDensity:
    def test_golden_value(self) -> None:
        result = calculate("2-1", {"apparatus_mass_g": 5000, "total_mass_g": 25000,
                                   "volume_cm3": 10000})
        assert result.values[0].quantity.value == pytest.approx(2000.0)

    def test_mass_order_physical_rule(self) -> None:
        with pytest.raises(EngineError, match="physical rule"):
            calculate("2-1", {"apparatus_mass_g": 25000, "total_mass_g": 5000,
                              "volume_cm3": 10000})


class TestSlump:
    def test_golden_value(self) -> None:
        result = calculate("3-1", {"measured_height_mm": 200})
        assert result.values[0].quantity.value == pytest.approx(100.0)
        assert result.values[0].quantity.unit == "mm"

    def test_measured_above_cone_height_rejected(self) -> None:
        with pytest.raises(EngineError):
            calculate("3-1", {"measured_height_mm": 350})


class TestTensile:
    def test_golden_value(self) -> None:
        result = calculate("4-2", {"diameter_mm": 150, "length_mm": 300, "load_n": 358000})
        assert result.values[0].quantity.value == pytest.approx(5.0647, abs=1e-3)


# ─── Standards-backed calculators ─────────────────────────────────────────


class TestCompressive:
    INPUTS = {"diameter_mm": 150, "length_mm": 300, "load_kn": 715}

    def test_reported_strength_matches_c39_golden(self) -> None:
        result = calculate("4-1", self.INPUTS)
        reported = next(v for v in result.values if v.key == "reported_strength")
        assert reported.quantity.value == pytest.approx(40.5, abs=1e-9)
        assert result.status is ValidationStatus.PASS

    def test_cube_geometry_is_rejected(self) -> None:
        with pytest.raises(c39.OutOfScopeError):
            calculate("4-1", {**self.INPUTS, "geometry": "cube"})

    def test_unknown_geometry_rejected(self) -> None:
        with pytest.raises(EngineError):
            calculate("4-1", {**self.INPUTS, "geometry": "prism"})


class TestRebound:
    READINGS = [42, 44, 41, 43, 45, 40, 43, 42, 44, 43]

    def test_disclaimer_always_attached(self) -> None:
        result = calculate("4-5", {"readings": self.READINGS})
        assert result.scope is Scope.ESTIMATION
        assert "NOT FOR ACCEPTANCE" in result.disclaimer

    def test_no_correlation_no_strength_estimate(self) -> None:
        result = calculate("4-5", {"readings": self.READINGS})
        assert all(v.key != "estimated_strength" for v in result.values)

    def test_correlation_produces_flagged_estimate(self) -> None:
        result = calculate("4-5", {"readings": self.READINGS,
                                   "correlation_id": "demo", "correlation_slope": 0.5,
                                   "correlation_intercept": -5.0})
        estimate = next(v for v in result.values if v.key == "estimated_strength")
        assert estimate.quantity.value == pytest.approx(16.35)
        assert estimate.status is ValidationStatus.WARN  # estimation, never PASS

    def test_outliers_are_flagged(self) -> None:
        readings = [*self.READINGS[:-1], 25]
        result = calculate("4-5", {"readings": readings})
        assert any(v.key == "outliers" for v in result.values)


class TestGrading:
    PASS_CURVE = {9.5: 97.5, 4.75: 88.0, 2.36: 70.0, 1.18: 50.0,
                  0.6: 33.0, 0.3: 18.0, 0.15: 6.0, 0.075: 1.0}

    def test_clean_curve_passes(self) -> None:
        result = calculate("1-1", {"curve": self.PASS_CURVE})
        assert result.status is ValidationStatus.PASS

    def test_missing_sieve_reports_error(self) -> None:
        incomplete = {k: v for k, v in self.PASS_CURVE.items() if k != 0.075}
        result = calculate("1-1", {"curve": incomplete})
        assert result.status is ValidationStatus.ERROR

    def test_missing_curve_key_rejected(self) -> None:
        with pytest.raises(EngineError):
            calculate("1-1", {})


# ─── Engine plumbing ──────────────────────────────────────────────────────


class TestEnginePlumbing:
    def test_unknown_test_raises(self) -> None:
        with pytest.raises(EngineError, match="pending"):
            calculate("9-9", {})

    def test_missing_input_raises(self) -> None:
        with pytest.raises(EngineError, match="missing input"):
            calculate("1-2", {"wet_mass_g": 100})

    def test_non_numeric_input_raises(self) -> None:
        with pytest.raises(EngineError):
            calculate("3-1", {"measured_height_mm": "tall"})

    def test_demo_cases_cover_implemented_tests(self) -> None:
        assert {test_id for test_id, _ in DEMO_CASES} == set(CALCULATORS)

    def test_run_demo_cases_produces_one_result_per_case(self) -> None:
        results = run_demo_cases()
        assert len(results) == len(DEMO_CASES)
        assert all(r.values for r in results)


# ─── Golden suite (in-process QA tier) ────────────────────────────────────


class TestGoldenSuite:
    def test_every_executable_golden_case_passes(self) -> None:
        from concrete_lab.qa.golden import DEFAULT_GOLDEN_DIR, ENGINE_CASE_CHECKERS

        report = run_golden_suite()
        assert report.failed == 0, report.failures

        # Derived, never hand-counted: standards families + engine cases.
        standards = sum(
            1 for path in DEFAULT_GOLDEN_DIR.glob("*.json")
            if path.name.startswith(("c39_", "c805_", "isiri302_"))
        )
        assert report.passed == standards + len(ENGINE_CASE_CHECKERS)
        assert report.passed + report.warnings == len(list(DEFAULT_GOLDEN_DIR.glob("*.json")))
        # pending tests are warnings, and every warning is explained
        assert report.warnings > 0 and len(report.notes) == report.warnings
