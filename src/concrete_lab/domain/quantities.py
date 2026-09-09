"""Physical quantities with explicit units.

A :class:`Quantity` couples a numeric measurement with the unit it was
measured in.  Lab results are meaningless without their unit, so this
pair is the atomic value object of the domain layer: every input read
from a sheet and every computed output should travel as a ``Quantity``.

The class is immutable and hashable, which makes quantities safe to use
as dictionary keys and cache entries.
"""

from __future__ import annotations

import math
from dataclasses import dataclass
from typing import ClassVar, Union

#: Accepted input types for the numeric part of a quantity.
Number = Union[int, float]


@dataclass(frozen=True)
class Quantity:
    """An immutable numeric measurement with a mandatory unit.

    Attributes:
        value: The magnitude of the measurement.  Integers are coerced
            to ``float`` on construction.  ``NaN`` and infinities are
            rejected because they cannot represent a physical reading.
        unit: The measurement unit as a human-readable string, e.g.
            ``"g"``, ``"mm"``, ``"MPa"`` or ``"%"``.  Must be a
            non-empty, non-whitespace string.

    Raises:
        TypeError: If ``value`` is not a real number (``bool`` is
            explicitly rejected) or ``unit`` is not a string.
        ValueError: If ``value`` is ``NaN``/infinite or ``unit`` is
            empty or whitespace-only.

    Example:
        >>> mass = Quantity(500, "g")
        >>> str(mass)
        '500.00 g'
    """

    value: float
    unit: str

    #: Unit marker used by :meth:`dimensionless` for pure ratios.
    DIMENSIONLESS_UNIT: ClassVar[str] = "–"

    def __post_init__(self) -> None:
        """Validate and normalise the raw field values."""
        value = self.value
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            raise TypeError(
                f"Quantity.value must be a real number, got {type(value).__name__!r}"
            )
        value = float(value)
        if math.isnan(value) or math.isinf(value):
            raise ValueError(f"Quantity.value must be finite, got {value!r}")
        object.__setattr__(self, "value", value)

        unit = self.unit
        if not isinstance(unit, str):
            raise TypeError(
                f"Quantity.unit must be a string, got {type(unit).__name__!r}"
            )
        unit = unit.strip()
        if not unit:
            raise ValueError("Quantity.unit must not be empty or whitespace-only")
        object.__setattr__(self, "unit", unit)

    # -- Construction helpers ------------------------------------------------

    @classmethod
    def dimensionless(cls, value: Number) -> Quantity:
        """Create a pure ratio/percentage value without a physical unit.

        Args:
            value: The numeric magnitude (e.g. a fineness modulus or a
                percentage passing).

        Returns:
            A :class:`Quantity` flagged with :data:`DIMENSIONLESS_UNIT`.
        """
        return cls(value, cls.DIMENSIONLESS_UNIT)

    # -- Queries -------------------------------------------------------------

    @property
    def is_dimensionless(self) -> bool:
        """Return ``True`` when this quantity carries no physical unit."""
        return self.unit == self.DIMENSIONLESS_UNIT

    @property
    def is_zero(self) -> bool:
        """Return ``True`` when the magnitude is exactly zero."""
        return self.value == 0.0

    # -- Rendering -----------------------------------------------------------

    def format(self, precision: int = 2) -> str:
        """Render the quantity as a fixed-point string with its unit.

        Args:
            precision: Number of digits after the decimal point
                (must be ``>= 0``).

        Returns:
            A string such as ``"12.50 g"``.  Dimensionless quantities
            render without a trailing unit, e.g. ``"2.80"``.

        Raises:
            ValueError: If ``precision`` is negative.
        """
        if precision < 0:
            raise ValueError(f"precision must be >= 0, got {precision}")
        text = f"{self.value:.{precision}f}"
        if self.is_dimensionless:
            return text
        return f"{text} {self.unit}"

    def __str__(self) -> str:
        """Return the default (2-decimal) rendering of the quantity."""
        return self.format()
