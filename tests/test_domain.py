"""Unit tests for the core domain classes introduced in the v4 refactor.

Covers:
    * ``concrete_lab.domain.quantities.Quantity``
    * ``concrete_lab.domain.statuses.ValidationStatus``
    * ``concrete_lab.domain.results.ResultSpec``
    * ``concrete_lab.standards.base.StandardSpec``
"""

import dataclasses

import pytest

from concrete_lab.domain.quantities import Quantity
from concrete_lab.domain.results import ResultSpec
from concrete_lab.domain.statuses import ValidationStatus
from concrete_lab.standards.base import StandardSpec


# ─── Quantity ─────────────────────────────────────────────────────────────


class TestQuantity:
    def test_valid_construction(self) -> None:
        q = Quantity(500.0, "g")
        assert q.value == 500.0
        assert q.unit == "g"

    def test_int_value_is_coerced_to_float(self) -> None:
        q = Quantity(3, "mm")
        assert q.value == 3.0
        assert isinstance(q.value, float)

    def test_unit_is_stripped(self) -> None:
        assert Quantity(1.0, "  MPa  ").unit == "MPa"

    @pytest.mark.parametrize("bad", ["text", None, [1], object()])
    def test_non_numeric_value_raises_type_error(self, bad: object) -> None:
        with pytest.raises(TypeError):
            Quantity(bad, "g")  # type: ignore[arg-type]

    def test_bool_value_is_rejected(self) -> None:
        with pytest.raises(TypeError):
            Quantity(True, "g")  # type: ignore[arg-type]

    @pytest.mark.parametrize("bad", [float("nan"), float("inf"), float("-inf")])
    def test_non_finite_value_raises_value_error(self, bad: float) -> None:
        with pytest.raises(ValueError):
            Quantity(bad, "g")

    @pytest.mark.parametrize("bad", ["", "   ", "\t"])
    def test_empty_unit_raises_value_error(self, bad: str) -> None:
        with pytest.raises(ValueError):
            Quantity(1.0, bad)

    def test_non_string_unit_raises_type_error(self) -> None:
        with pytest.raises(TypeError):
            Quantity(1.0, 42)  # type: ignore[arg-type]

    def test_is_immutable(self) -> None:
        q = Quantity(1.0, "g")
        with pytest.raises(dataclasses.FrozenInstanceError):
            q.value = 2.0  # type: ignore[misc]

    def test_equality_and_hash(self) -> None:
        a = Quantity(500, "g")
        b = Quantity(500.0, "g")
        assert a == b
        assert hash(a) == hash(b)
        assert Quantity(500, "g") != Quantity(500, "kg")

    @pytest.mark.parametrize(
        ("value", "precision", "expected"),
        [
            (12.3456, 2, "12.35 g"),
            (12.3456, 0, "12 g"),
            (12.3456, 3, "12.346 g"),
            (0.0, 2, "0.00 g"),
        ],
    )
    def test_format(self, value: float, precision: int, expected: str) -> None:
        assert Quantity(value, "g").format(precision) == expected

    def test_format_negative_precision_raises(self) -> None:
        with pytest.raises(ValueError):
            Quantity(1.0, "g").format(-1)

    def test_str_defaults_to_two_decimals(self) -> None:
        assert str(Quantity(500, "g")) == "500.00 g"

    def test_dimensionless(self) -> None:
        q = Quantity.dimensionless(2.8)
        assert q.is_dimensionless
        assert q.format() == "2.80"  # no trailing unit
        assert not Quantity(2.8, "%").is_dimensionless

    def test_is_zero(self) -> None:
        assert Quantity(0.0, "g").is_zero
        assert not Quantity(0.001, "g").is_zero


# ─── ValidationStatus ─────────────────────────────────────────────────────


class TestValidationStatus:
    def test_has_exactly_six_members(self) -> None:
        assert len(ValidationStatus) == 6

    def test_member_names(self) -> None:
        assert {s.name for s in ValidationStatus} == {
            "PENDING", "PASS", "WARN", "FAIL", "SKIPPED", "ERROR",
        }

    def test_values_are_stable_lowercase_strings(self) -> None:
        assert ValidationStatus.PASS.value == "pass"
        assert ValidationStatus.FAIL.value == "fail"

    def test_round_trip_from_value(self) -> None:
        for status in ValidationStatus:
            assert ValidationStatus(status.value) is status

    def test_invalid_value_raises(self) -> None:
        with pytest.raises(ValueError):
            ValidationStatus("not-a-status")

    def test_only_pass_is_success(self) -> None:
        assert ValidationStatus.PASS.is_success
        for status in ValidationStatus:
            if status is not ValidationStatus.PASS:
                assert not status.is_success

    def test_only_pending_is_not_terminal(self) -> None:
        assert not ValidationStatus.PENDING.is_terminal
        for status in ValidationStatus:
            if status is not ValidationStatus.PENDING:
                assert status.is_terminal

    @pytest.mark.parametrize(
        ("status", "symbol"),
        [
            (ValidationStatus.PENDING, "—"),
            (ValidationStatus.PASS, "✅"),
            (ValidationStatus.WARN, "⚠️"),
            (ValidationStatus.FAIL, "❌"),
        ],
    )
    def test_legacy_symbols_are_preserved(
        self, status: ValidationStatus, symbol: str
    ) -> None:
        assert status.symbol == symbol

    def test_every_status_has_a_symbol(self) -> None:
        for status in ValidationStatus:
            assert isinstance(status.symbol, str)
            assert status.symbol


# ─── ResultSpec ───────────────────────────────────────────────────────────


class TestResultSpec:
    def test_defaults(self) -> None:
        spec = ResultSpec(key="moisture", label="رطوبت", formula="=(B6-B7)/B7*100")
        assert spec.unit == ""
        assert spec.precision == 2
        assert spec.critical is False
        assert spec.tooltip == ""

    @pytest.mark.parametrize("precision", [0, 2, 3])
    def test_num_format(self, precision: int) -> None:
        spec = ResultSpec(key="k", label="l", formula="=1", precision=precision)
        expected = "0" if precision == 0 else "0." + "0" * precision
        assert spec.num_format() == expected

    @pytest.mark.parametrize("bad_key", ["", "   ", "not an id", "1abc", "a-b"])
    def test_invalid_key_raises(self, bad_key: str) -> None:
        with pytest.raises(ValueError):
            ResultSpec(key=bad_key, label="l", formula="=1")

    @pytest.mark.parametrize("field", ["label", "formula"])
    def test_empty_required_field_raises(self, field: str) -> None:
        kwargs = {"key": "k", "label": "l", "formula": "=1", field: "  "}
        with pytest.raises(ValueError):
            ResultSpec(**kwargs)

    def test_negative_precision_raises(self) -> None:
        with pytest.raises(ValueError):
            ResultSpec(key="k", label="l", formula="=1", precision=-1)

    @pytest.mark.parametrize("bad", [2.0, "2", True])
    def test_non_int_precision_raises_type_error(self, bad: object) -> None:
        with pytest.raises(TypeError):
            ResultSpec(key="k", label="l", formula="=1", precision=bad)  # type: ignore[arg-type]

    def test_non_bool_critical_raises_type_error(self) -> None:
        with pytest.raises(TypeError):
            ResultSpec(key="k", label="l", formula="=1", critical="yes")  # type: ignore[arg-type]

    def test_is_immutable(self) -> None:
        spec = ResultSpec(key="k", label="l", formula="=1")
        with pytest.raises(dataclasses.FrozenInstanceError):
            spec.precision = 4  # type: ignore[misc]

    def test_is_dimensionless(self) -> None:
        plain = ResultSpec(key="fm", label="FM", formula="=1")
        assert plain.is_dimensionless
        with_unit = ResultSpec(key="s", label="S", formula="=1", unit="MPa")
        assert not with_unit.is_dimensionless

    def test_str_representation(self) -> None:
        assert str(ResultSpec(key="k", label="l", formula="=1", unit="MPa")) == "k [MPa]"
        assert str(ResultSpec(key="k", label="l", formula="=1")) == "k"


# ─── StandardSpec ─────────────────────────────────────────────────────────


class TestStandardSpec:
    def test_valid_construction(self) -> None:
        spec = StandardSpec(
            code="C136", name="ASTM C136/C136M",
            title="Sieve Analysis", edition="2024",
        )
        assert spec.code == "C136"
        assert spec.status == "active"  # default

    def test_is_active_property(self) -> None:
        base = dict(code="C187", name="ASTM C187",
                    title="Normal Consistency", edition="2016")
        assert StandardSpec(status="active", **base).is_active
        assert not StandardSpec(status="withdrawn", **base).is_active

    @pytest.mark.parametrize("field", ["code", "name", "title", "edition"])
    def test_empty_required_field_raises(self, field: str) -> None:
        kwargs = {"code": "C1", "name": "n", "title": "t", "edition": "2024"}
        kwargs[field] = "   "
        with pytest.raises(ValueError):
            StandardSpec(**kwargs)  # type: ignore[arg-type]

    def test_invalid_status_raises(self) -> None:
        with pytest.raises(ValueError):
            StandardSpec(code="C1", name="n", title="t",
                         edition="2024", status="draft")

    def test_citation(self) -> None:
        spec = StandardSpec(code="C136", name="ASTM C136/C136M",
                            title="Sieve Analysis", edition="2024")
        assert spec.citation() == "ASTM C136/C136M (2024)"
        assert str(spec) == "C136: ASTM C136/C136M (2024)"

    def test_is_immutable(self) -> None:
        spec = StandardSpec(code="C1", name="n", title="t", edition="2024")
        with pytest.raises(dataclasses.FrozenInstanceError):
            spec.edition = "2025"  # type: ignore[misc]

    def test_from_dict_legacy_format(self) -> None:
        legacy = {"name": "ASTM C29/C29M", "edition": "2023",
                  "status": "active", "title": "Bulk Density of Aggregates"}
        spec = StandardSpec.from_dict("C29", legacy)
        assert spec.code == "C29"
        assert spec.name == "ASTM C29/C29M"
        assert spec.edition == "2023"
        assert spec.is_active

    def test_from_dict_defaults_status_to_active(self) -> None:
        legacy = {"name": "ASTM C29", "edition": "2023",
                  "title": "Bulk Density"}
        assert StandardSpec.from_dict("C29", legacy).status == "active"

    def test_from_dict_missing_key_raises_value_error(self) -> None:
        with pytest.raises(ValueError, match="missing key"):
            StandardSpec.from_dict("C29", {"name": "n", "edition": "2023"})
