"""ASTM C39/C39M-26 — Compressive strength of cylindrical concrete specimens.

Refactored ruleset addressing review feedback:

* **Scope is cylindrical specimens only.**  The legacy workbook let the
  user pick a cube and silently applied a 0.95 convenience factor,
  which misrepresents this standard.  Here, any geometry other than a
  cylinder raises :class:`OutOfScopeError` with an explicit message
  pointing to the proper cube method.
* **L/D correction uses the published table with linear interpolation**
  instead of the coarse ``IF(L/D<1.8, 0.96, ...)`` step formula.
* **Scientific logic lives in Python**, not in Excel formula strings;
  the workbook renderer will consume :func:`calculate` results.

All dimensions are millimetres, loads are kilonewtons, strengths are
megapascals.  Reported strength is rounded to the nearest 0.1 MPa.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from enum import Enum, unique
from typing import Tuple

from concrete_lab.domain.quantities import Quantity
from concrete_lab.standards.base import Scope, StandardSpec

#: Metadata for this ruleset (active edition at the time of writing).
SPEC: StandardSpec = StandardSpec(
    code="C39",
    name="ASTM C39/C39M",
    title="Compressive Strength of Cylindrical Concrete Specimens",
    edition="2026",
    status="active",
)

#: This is a normative calculation method.
SCOPE: Scope = Scope.STANDARD

#: Version tag so every derived result can be traced to a ruleset revision.
RULESET_VERSION: str = "C39-2026-v1"

#: Correction factors for specimens with L/D less than 2.00, as published
#: in the standard.  Intermediate values are obtained by linear
#: interpolation; at L/D >= 2.00 the factor is 1.00.
LD_CORRECTION_TABLE: Tuple[Tuple[float, float], ...] = (
    (1.00, 0.87),
    (1.25, 0.93),
    (1.50, 0.96),
    (1.75, 0.98),
    (2.00, 1.00),
)

#: Lowest L/D ratio covered by the correction table.
MIN_LD_RATIO: float = 1.00

#: Reporting precision required for SI results (nearest 0.1 MPa).
REPORT_STEP_MPA: float = 0.1

#: Engineering plausibility window for concrete compressive strength.
#: Values outside it are technically computable but flagged as suspicious.
PLAUSIBLE_STRENGTH_RANGE_MPA: Tuple[float, float] = (1.0, 150.0)


class OutOfScopeError(ValueError):
    """Raised when an input lies outside the scope of ASTM C39/C39M."""


@unique
class SpecimenGeometry(Enum):
    """Specimen geometries recognised by the form.

    Only :attr:`CYLINDER` is inside the scope of ASTM C39/C39M; the
    enum still lists ``CUBE`` so that callers can *detect* the wrong
    choice and receive a clear rejection instead of a silent factor.
    """

    CYLINDER = "cylinder"
    CUBE = "cube"


def ld_correction_factor(ld_ratio: float) -> float:
    """Return the L/D correction factor with linear interpolation.

    Args:
        ld_ratio: Length-to-diameter ratio of the specimen.

    Returns:
        The correction factor from :data:`LD_CORRECTION_TABLE`,
        linearly interpolated for intermediate ratios.  Ratios above
        2.00 are uncorrected (factor ``1.00``).

    Raises:
        OutOfScopeError: If ``ld_ratio`` is below
            :data:`MIN_LD_RATIO` (outside the published table).
        ValueError: If ``ld_ratio`` is not finite.

    Example:
        >>> ld_correction_factor(1.6)
        0.968
    """
    if not math.isfinite(ld_ratio):
        raise ValueError(f"L/D ratio must be finite, got {ld_ratio!r}")
    if ld_ratio < MIN_LD_RATIO:
        raise OutOfScopeError(
            f"L/D ratio {ld_ratio:.3f} is below {MIN_LD_RATIO:.2f}, outside the "
            f"correction table of {SPEC.name}"
        )
    if ld_ratio >= LD_CORRECTION_TABLE[-1][0]:
        return 1.0

    for (lo_ld, lo_f), (hi_ld, hi_f) in zip(LD_CORRECTION_TABLE, LD_CORRECTION_TABLE[1:]):
        if lo_ld <= ld_ratio <= hi_ld:
            t = (ld_ratio - lo_ld) / (hi_ld - lo_ld)
            return lo_f + t * (hi_f - lo_f)
    raise AssertionError("unreachable: table must bracket every ratio >= 1.0")  # pragma: no cover


def _round_to_step(value: float, step: float) -> float:
    """Round ``value`` to the nearest multiple of ``step`` (half-even)."""
    return round(round(value / step) * step, 10)


@dataclass(frozen=True)
class C39Result:
    """Immutable outcome of one ASTM C39 compressive-strength calculation.

    Attributes:
        area: Average cross-sectional area (mm²).
        ld_ratio: Measured length-to-diameter ratio.
        correction_factor: Interpolated L/D correction factor.
        gross_strength: Uncorrected strength P/A (MPa).
        reported_strength: Corrected strength rounded to 0.1 MPa.
        warnings: Human-readable plausibility warnings (may be empty).
    """

    area: Quantity
    ld_ratio: float
    correction_factor: float
    gross_strength: Quantity
    reported_strength: Quantity
    warnings: Tuple[str, ...] = field(default_factory=tuple)


def calculate(
    diameter_mm: float,
    length_mm: float,
    load_kn: float,
    geometry: SpecimenGeometry = SpecimenGeometry.CYLINDER,
) -> C39Result:
    """Compute the reported compressive strength per ASTM C39/C39M.

    Args:
        diameter_mm: Specimen diameter in millimetres (> 0).
        length_mm: Specimen length in millimetres (> 0).
        load_kn: Maximum load in kilonewtons (> 0).
        geometry: Specimen geometry; only ``CYLINDER`` is accepted.

    Returns:
        A :class:`C39Result` with area, L/D ratio, interpolated
        correction factor, gross strength and the reported strength
        rounded to the nearest 0.1 MPa.

    Raises:
        OutOfScopeError: If ``geometry`` is not a cylinder (cube
            specimens belong to a different method and must never be
            folded into C39 with an empirical factor), or if L/D is
            below :data:`MIN_LD_RATIO`.
        ValueError: If any dimension/load is not finite or not positive.

    Example:
        >>> result = calculate(diameter_mm=150, length_mm=300, load_kn=715)
        >>> result.reported_strength.format(1)
        '40.5 MPa'
    """
    if geometry is SpecimenGeometry.CUBE:
        raise OutOfScopeError(
            "Cube specimens are outside the scope of ASTM C39/C39M "
            "(Compressive Strength of Cylindrical Concrete Specimens). "
            "این هندسه در دامنه این استاندارد نیست؛ از روش مخصوص نمونه مکعبی استفاده کنید."
        )
    if geometry is not SpecimenGeometry.CYLINDER:  # defensive: future members
        raise OutOfScopeError(f"Geometry {geometry!r} is outside the scope of {SPEC.name}")

    for name, value in (("diameter_mm", diameter_mm), ("length_mm", length_mm), ("load_kn", load_kn)):
        if not isinstance(value, (int, float)) or isinstance(value, bool) or not math.isfinite(value):
            raise ValueError(f"{name} must be a finite number, got {value!r}")
        if value <= 0:
            raise ValueError(f"{name} must be positive, got {value!r}")

    area_mm2 = math.pi * diameter_mm**2 / 4.0
    ld_ratio = length_mm / diameter_mm
    factor = ld_correction_factor(ld_ratio)
    gross_mpa = (load_kn * 1000.0) / area_mm2
    reported_mpa = _round_to_step(gross_mpa * factor, REPORT_STEP_MPA)

    warnings: Tuple[str, ...] = ()
    lo, hi = PLAUSIBLE_STRENGTH_RANGE_MPA
    if not lo <= reported_mpa <= hi:
        warnings = (
            (f"Reported strength {reported_mpa:.1f} MPa is outside the plausible "
             f"range [{lo:.0f}, {hi:.0f}] MPa — check inputs"),
        )

    return C39Result(
        area=Quantity(area_mm2, "mm²"),
        ld_ratio=ld_ratio,
        correction_factor=factor,
        gross_strength=Quantity(gross_mpa, "MPa"),
        reported_strength=Quantity(reported_mpa, "MPa"),
        warnings=warnings,
    )
