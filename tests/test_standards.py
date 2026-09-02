"""Golden-case and unit tests for the refactored standard rulesets.

Covers the three standards refactored per reviewer feedback:

* ``concrete_lab.standards.astm.c39`` — cylinders only, interpolated
  L/D correction;
* ``concrete_lab.standards.astm.c805`` — ESTIMATION scope, mandatory
  project-specific correlation, strong disclaimer;
* ``concrete_lab.standards.isiri.isiri_302`` — full grading ruleset
  (envelope + physical rules), not a bare dictionary.

Golden cases live in ``validation/golden_cases/`` and follow the same
schema as the existing workbook golden cases.
"""

import json
from pathlib import Path
from typing import Any, Dict, List

import pytest

from concrete_lab.domain.statuses import ValidationStatus
from concrete_lab.standards.astm import c39, c805
from concrete_lab.standards.base import Scope
from concrete_lab.standards.isiri import isiri_302

GOLDEN_DIR = Path(__file__).resolve().parent.parent / "validation" / "golden_cases"


def load_golden_cases(prefix: str) -> List[Dict[str, Any]]:
    """Load golden case files whose name starts with ``prefix``."""
    cases: List[Dict[str, Any]] = []
    for path in sorted(GOLDEN_DIR.glob(f"{prefix}*.json")):
        with open(path, "r", encoding="utf-8") as fp:
            data = json.load(fp)
            data["file"] = path.name
            cases.append(data)
    assert cases, f"no golden cases found with prefix {prefix!r}"
    return cases


# ─── ASTM C39 golden cases ────────────────────────────────────────────────


class TestC39Golden:
    @pytest.mark.parametrize("case", load_golden_cases("c39_"), ids=lambda c: c["file"])
    def test_c39_golden_case(self, case: Dict[str, Any]) -> None:
        """The engine must reproduce the golden worksheet values."""
        inputs = case["inputs"]
        expected = case["expected"]
        tolerance = case["tolerance"]

        result = c39.calculate(
            diameter_mm=inputs["diameter_mm"],
            length_mm=inputs["length_mm"],
            load_kn=inputs["load_kn"],
        )

        assert result.ld_ratio == pytest.approx(expected["ld_ratio"], abs=1e-9)
        assert result.correction_factor == pytest.approx(expected["correction_factor"], abs=1e-6)
        assert result.gross_strength.value == pytest.approx(expected["gross_strength_mpa"], abs=tolerance)
        assert result.reported_strength.value == pytest.approx(expected["reported_strength_mpa"], abs=1e-9)
        assert result.reported_strength.unit == "MPa"


class TestC39Unit:
    @pytest.mark.parametrize(
        ("ld", "factor"),
        [(1.0, 0.87), (1.25, 0.93), (1.5, 0.96), (1.75, 0.98), (2.0, 1.0)],
    )
    def test_tabulated_points_exact(self, ld: float, factor: float) -> None:
        assert c39.ld_correction_factor(ld) == pytest.approx(factor, abs=1e-12)

    @pytest.mark.parametrize(
        ("ld", "factor"),
        [(1.125, 0.90), (1.375, 0.945), (1.6, 0.968), (1.875, 0.99)],
    )
    def test_linear_interpolation_between_points(self, ld: float, factor: float) -> None:
        assert c39.ld_correction_factor(ld) == pytest.approx(factor, abs=1e-12)

    def test_ratio_above_two_is_uncorrected(self) -> None:
        assert c39.ld_correction_factor(2.01) == 1.0
        assert c39.ld_correction_factor(3.0) == 1.0

    def test_ratio_below_table_is_out_of_scope(self) -> None:
        with pytest.raises(c39.OutOfScopeError):
            c39.ld_correction_factor(0.99)

    def test_non_finite_ratio_rejected(self) -> None:
        with pytest.raises(ValueError):
            c39.ld_correction_factor(float("nan"))

    def test_cube_geometry_is_rejected_not_corrected(self) -> None:
        """The legacy 0.95 cube hack must never come back."""
        with pytest.raises(c39.OutOfScopeError, match="Cube specimens are outside the scope"):
            c39.calculate(150, 150, 500, geometry=c39.SpecimenGeometry.CUBE)

    @pytest.mark.parametrize("kwargs", [
        {"diameter_mm": 0}, {"length_mm": -1}, {"load_kn": 0}, {"load_kn": float("inf")},
    ])
    def test_invalid_inputs_rejected(self, kwargs: Dict[str, float]) -> None:
        base = {"diameter_mm": 150, "length_mm": 300, "load_kn": 500}
        base.update(kwargs)
        with pytest.raises(ValueError):
            c39.calculate(**base)

    def test_reported_strength_rounded_to_0p1_mpa(self) -> None:
        result = c39.calculate(150, 300, 715)
        assert result.reported_strength.value == pytest.approx(40.5, abs=1e-9)

    def test_plausible_strength_has_no_warnings(self) -> None:
        assert c39.calculate(150, 300, 715).warnings == ()

    def test_implausible_strength_is_flagged(self) -> None:
        result = c39.calculate(100, 200, 2000)  # ~255 MPa is suspicious
        assert len(result.warnings) == 1
        assert "plausible" in result.warnings[0]

    def test_spec_metadata_matches_active_edition(self) -> None:
        assert c39.SPEC.edition == "2026"
        assert c39.SPEC.is_active
        assert c39.SCOPE is Scope.STANDARD
        assert c39.RULESET_VERSION.startswith("C39-2026")


# ─── ASTM C805 golden cases ───────────────────────────────────────────────


class TestC805Golden:
    @pytest.mark.parametrize("case", load_golden_cases("c805_"), ids=lambda c: c["file"])
    def test_c805_golden_case(self, case: Dict[str, Any]) -> None:
        """Rebound averaging and the correlation estimate match golden values."""
        inputs = case["inputs"]
        expected = case["expected"]
        tolerance = case["tolerance"]

        summary = c805.mean_rebound_number(tuple(inputs["readings"]))
        assert summary.mean == pytest.approx(expected["mean"], abs=tolerance)
        assert list(summary.outliers) == expected["outliers"]
        assert summary.retained_mean == pytest.approx(expected["retained_mean"], abs=tolerance)

        correlation = c805.LinearCorrelation(
            correlation_id=inputs["estimate"]["correlation_id"],
            slope=inputs["estimate"]["slope"],
            intercept=inputs["estimate"]["intercept"],
        )
        estimate = c805.estimate_strength(summary.retained_mean, correlation, correlation.correlation_id)
        assert estimate.estimated_strength.value == pytest.approx(
            expected["estimated_strength_mpa"], abs=tolerance
        )
        assert estimate.estimated_strength.unit == "MPa"
        assert estimate.scope is Scope.ESTIMATION
        assert estimate.disclaimer == c805.DISCLAIMER


class TestC805Unit:
    READINGS = (42, 44, 41, 43, 45, 40, 43, 42, 44, 43)

    def test_fewer_than_ten_readings_rejected(self) -> None:
        with pytest.raises(c805.ReboundError, match="at least 10"):
            c805.mean_rebound_number(self.READINGS[:9])

    @pytest.mark.parametrize("bad", [-1.0, 101.0, float("nan")])
    def test_reading_out_of_scale_rejected(self, bad: float) -> None:
        with pytest.raises(c805.ReboundError):
            c805.mean_rebound_number(self.READINGS[:-1] + (bad,))

    def test_outlier_rule_flags_only_deviations_over_seven_units(self) -> None:
        summary = c805.mean_rebound_number(self.READINGS[:-1] + (25,))
        assert summary.outliers == (25,)
        assert summary.retained_mean > summary.mean  # outlier pulled mean down

    def test_estimate_requires_traceable_correlation_id(self) -> None:
        with pytest.raises(c805.ReboundError, match="correlation_id"):
            c805.estimate_strength(42.0, lambda r: r * 0.5, correlation_id="   ")

    @pytest.mark.parametrize("bad", [-1.0, 101.0, float("nan")])
    def test_estimate_rejects_out_of_range_rebound(self, bad: float) -> None:
        with pytest.raises(c805.ReboundError):
            c805.estimate_strength(bad, lambda r: 20.0, correlation_id="ok")

    def test_estimate_rejects_implausible_correlation_output(self) -> None:
        with pytest.raises(c805.ReboundError, match="implausible"):
            c805.estimate_strength(42.0, lambda r: -3.0, correlation_id="bad-corr")

    def test_disclaimers_forbid_acceptance_use(self) -> None:
        assert "NOT FOR ACCEPTANCE" in c805.DISCLAIMER
        assert "ASTM C39" in c805.DISCLAIMER  # defers to cylinder specimens
        assert "قابل استناد نیست" in c805.DISCLAIMER_FA
        assert c805.BADGE_TEXT == "ESTIMATED — NOT FOR ACCEPTANCE"

    def test_spec_metadata_matches_active_edition(self) -> None:
        assert c805.SPEC.edition == "2025"
        assert c805.SPEC.is_active
        assert c805.SCOPE is Scope.ESTIMATION


# ─── ISIRI 302 golden cases ───────────────────────────────────────────────


class TestISIRI302Golden:
    @pytest.mark.parametrize("case", load_golden_cases("isiri302_"), ids=lambda c: c["file"])
    def test_isiri302_golden_case(self, case: Dict[str, Any]) -> None:
        """The ruleset must classify golden curves exactly as documented."""
        curve = {float(size): pct for size, pct in case["inputs"]["curve"].items()}
        expected = case["expected"]

        evaluation = isiri_302.DEFAULT_RULESET.evaluate(curve)
        assert evaluation.overall_status.value == expected["overall_status"]

        if "statuses" in expected:
            got = {str(ev.sieve_mm): ev.status.value for ev in evaluation.evaluations}
            assert got == expected["statuses"]
        if "physical_issue_count_min" in expected:
            assert len(evaluation.physical_issues) >= expected["physical_issue_count_min"]


class TestISIRI302Unit:
    CURVE_KEYS = (9.5, 4.75, 2.36, 1.18, 0.6, 0.3, 0.15, 0.075)
    PASS_CURVE = dict(zip(CURVE_KEYS, (97.5, 88.0, 70.0, 50.0, 33.0, 18.0, 6.0, 1.0)))

    def test_missing_sieve_is_reported_as_error(self) -> None:
        incomplete = {k: v for k, v in self.PASS_CURVE.items() if k != 0.075}
        evaluation = isiri_302.DEFAULT_RULESET.evaluate(incomplete)
        assert evaluation.missing_sieves == (0.075,)
        assert evaluation.overall_status is ValidationStatus.ERROR

    def test_unknown_sieve_sizes_are_ignored(self) -> None:
        curve = dict(self.PASS_CURVE)
        curve[19.0] = 100.0  # not part of this envelope
        evaluation = isiri_302.DEFAULT_RULESET.evaluate(curve)
        assert evaluation.overall_status is ValidationStatus.PASS

    def test_narrow_band_gets_no_warning_margin(self) -> None:
        """The 0.075 mm band [0, 2] is too narrow for a warning zone."""
        limit = isiri_302.DEFAULT_RULESET.limit_for(0.075)
        assert limit is not None
        assert limit.status_for(0.1) is ValidationStatus.PASS
        assert limit.status_for(1.9) is ValidationStatus.PASS
        assert limit.status_for(2.5) is ValidationStatus.FAIL

    def test_wide_band_warns_near_boundaries(self) -> None:
        limit = isiri_302.DEFAULT_RULESET.limit_for(2.36)  # [60, 80]
        assert limit is not None
        assert limit.status_for(61.0) is ValidationStatus.WARN
        assert limit.status_for(79.5) is ValidationStatus.WARN
        assert limit.status_for(70.0) is ValidationStatus.PASS
        assert limit.status_for(59.9) is ValidationStatus.FAIL

    def test_physical_range_violation_is_error(self) -> None:
        curve = dict(self.PASS_CURVE)
        curve[2.36] = 105.0
        evaluation = isiri_302.DEFAULT_RULESET.evaluate(curve)
        assert evaluation.overall_status is ValidationStatus.ERROR
        assert any("physical range" in issue for issue in evaluation.physical_issues)

    def test_limit_row_validation(self) -> None:
        with pytest.raises(ValueError):
            isiri_302.SieveLimit(sieve_mm=4.75, min_passing=80.0, max_passing=70.0)
        with pytest.raises(ValueError):
            isiri_302.SieveLimit(sieve_mm=4.75, min_passing=-1.0, max_passing=90.0)
        with pytest.raises(ValueError):
            isiri_302.SieveLimit(sieve_mm=0.0, min_passing=0.0, max_passing=100.0)

    def test_ruleset_matches_legacy_isiri_limits_envelope(self) -> None:
        """Regression guard: the envelope must stay identical to build.py's ISIRI_LIMITS."""
        legacy = {
            9.5: (100, 95), 4.75: (95, 80), 2.36: (80, 60), 1.18: (60, 40),
            0.600: (40, 25), 0.300: (25, 10), 0.150: (10, 2), 0.075: (2, 0),
        }
        for limit in isiri_302.GRADING_LIMITS:
            upper, lower = legacy[limit.sieve_mm]
            assert limit.max_passing == upper
            assert limit.min_passing == lower

    def test_limit_for_unknown_sieve_returns_none(self) -> None:
        assert isiri_302.DEFAULT_RULESET.limit_for(12.5) is None

    def test_spec_metadata(self) -> None:
        assert isiri_302.SPEC.edition == "1394"
        assert isiri_302.SPEC.is_active
        assert isiri_302.SCOPE is Scope.STANDARD
