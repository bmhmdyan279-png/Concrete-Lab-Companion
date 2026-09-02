"""Declarative specifications of computed lab results.

A :class:`ResultSpec` describes *what* a computed result is — its key,
human-readable label, the formula that produces it, its unit and display
precision — without carrying any runtime value.  It is the domain-level
successor of the legacy ``OutputField`` dataclass in ``build.py``, with
the sheet geometry (``row``/``col``) moved out of the domain and a
proper ``precision`` replacing the free-form ``num_fmt`` string.
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class ResultSpec:
    """Immutable specification of a single computed result.

    Attributes:
        key: Stable machine identifier of the result (e.g.
            ``"fineness_modulus"``, ``"compressive_strength"``).  Must
            be non-empty and a valid Python identifier so it can be used
            as a keyword/registry key.
        label: Human-readable label rendered next to the value in the
            workbook (may be in Persian).  Must be non-empty.
        formula: The formula (Excel syntax) that computes the value.
            Must be non-empty.
        unit: Physical unit of the value (``"g"``, ``"MPa"``, ...).
            Empty string for dimensionless results.
        precision: Number of decimal places used for display and for
            the generated Excel number format.  Must be ``>= 0``.
        critical: When ``True``, failing this result fails the whole
            test (used by dashboard rollups).
        tooltip: Optional explanatory text shown as a cell comment.

    Raises:
        TypeError: If ``precision`` is not an ``int`` (``bool`` is
            rejected) or ``critical`` is not a ``bool``.
        ValueError: If ``key`` is empty or not an identifier, if
            ``label``/``formula`` are empty/whitespace-only, or if
            ``precision`` is negative.

    Example:
        >>> spec = ResultSpec(key="moisture", label="رطوبت",
        ...                   formula="=(B6-B7)/B7*100", unit="%")
        >>> spec.num_format()
        '0.00'
    """

    key: str
    label: str
    formula: str
    unit: str = ""
    precision: int = 2
    critical: bool = False
    tooltip: str = ""

    def __post_init__(self) -> None:
        """Validate the specification fields."""
        if not isinstance(self.key, str) or not self.key.strip():
            raise ValueError("ResultSpec.key must be a non-empty string")
        if not self.key.strip().isidentifier():
            raise ValueError(
                f"ResultSpec.key must be a valid identifier, got {self.key!r}"
            )
        if not isinstance(self.label, str) or not self.label.strip():
            raise ValueError("ResultSpec.label must be a non-empty string")
        if not isinstance(self.formula, str) or not self.formula.strip():
            raise ValueError("ResultSpec.formula must be a non-empty string")
        if isinstance(self.precision, bool) or not isinstance(self.precision, int):
            raise TypeError(
                f"ResultSpec.precision must be an int, got {type(self.precision).__name__!r}"
            )
        if self.precision < 0:
            raise ValueError(f"ResultSpec.precision must be >= 0, got {self.precision}")
        if not isinstance(self.critical, bool):
            raise TypeError(
                f"ResultSpec.critical must be a bool, got {type(self.critical).__name__!r}"
            )

    # -- Helpers -------------------------------------------------------------

    def num_format(self) -> str:
        """Return the Excel number format matching :attr:`precision`.

        Returns:
            ``"0"`` for zero decimals, otherwise ``"0.00..."`` with
            exactly :attr:`precision`` digits after the point.
        """
        if self.precision == 0:
            return "0"
        return "0." + "0" * self.precision

    @property
    def is_dimensionless(self) -> bool:
        """Return ``True`` when the result carries no physical unit."""
        return self.unit == ""

    def __str__(self) -> str:
        """Return a short debugging representation ``key [unit]``."""
        return f"{self.key} [{self.unit}]" if self.unit else self.key
