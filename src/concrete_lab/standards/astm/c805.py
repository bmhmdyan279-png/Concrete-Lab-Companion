"""ASTM C805/C805M-25 — Assessing the rebound number of hardened concrete.

Refactored ruleset addressing review feedback:

* The module declares ``SCOPE = Scope.ESTIMATION`` and ships a **strong
  disclaimer** attached to every strength estimate.  ASTM C805 states
  explicitly that the rebound method must not be used as the sole basis
  for accepting or rejecting concrete, and that strength estimation
  requires a project-specific correlation.
* Consequently this module deliberately contains **no built-in
  universal rebound-number → strength formula**.  Any estimate must be
  produced from an explicit caller-supplied correlation, whose identity
  is recorded on the result for traceability.
* The rebound value is treated first as a **quality/consistency
  indicator** (mean of repeated readings with the standard's outlier
  rule) and only optionally — and with the disclaimer — as a strength
  estimate.

Units: rebound numbers are dimensionless hammer readings (0–100 scale);
strengths are megapascals.
"""

from __future__ import annotations

import math
from dataclasses import dataclass
from typing import Callable, Tuple

from concrete_lab.domain.quantities import Quantity
from concrete_lab.standards.base import Scope, StandardSpec

#: Metadata for this ruleset (active edition at the time of writing).
SPEC: StandardSpec = StandardSpec(
    code="C805",
    name="ASTM C805/C805M",
    title="Assessing the Rebound Number of Hardened Concrete",
    edition="2025",
    status="active",
)

#: Results of this module are indicative estimates, never normative.
SCOPE: Scope = Scope.ESTIMATION

#: Version tag so every derived result can be traced to a ruleset revision.
RULESET_VERSION: str = "C805-2025-v1"

#: Badge text renderers should display next to any derived value.
BADGE_TEXT: str = "ESTIMATED — NOT FOR ACCEPTANCE"

#: C805 procedure: take at least this many readings per test area.
MIN_READINGS_PER_AREA: int = 10

#: A single reading differing from the area mean by more than this many
#: rebound units is discarded (and, on site, re-measured).
OUTLIER_THRESHOLD_UNITS: float = 7.0

#: Strong disclaimer attached to every strength estimate (English).
DISCLAIMER: str = (
    "⚠️ ESTIMATED VALUE — NOT FOR ACCEPTANCE OR REJECTION OF CONCRETE. "
    "ASTM C805/C805M measures the rebound number of hardened concrete; it does not "
    "measure strength directly. Any strength derived from a rebound number depends on a "
    "project-specific correlation and shall not be used as the sole basis for acceptance "
    "or rejection. Rebound numbers are affected by surface moisture, surface condition "
    "and finish, carbonation, hammer orientation and temperature. Confirm all acceptance "
    "decisions with standard-cured test specimens (e.g. ASTM C39 cylinders)."
)

#: Persian translation of :data:`DISCLAIMER` for the workbook UI.
DISCLAIMER_FA: str = (
    "⚠️ مقدار تخمینی — برای پذیرش یا رد بتن قابل استناد نیست. "
    "روش چکش بازگشتی (ASTM C805/C805M) عدد بازگشت را اندازه می‌گیرد، نه مقاومت را. "
    "هر مقاومت استخراج‌شده از عدد بازگشت وابسته به همبستگی اختصاصی همان پروژه است و "
    "نباید به‌تنهایی مبنای پذیرش یا رد قرار گیرد. عدد بازگشت تحت تأثیر رطوبت سطح، "
    "وضعیت و پرداخت سطح، کربناسیون، جهت‌گیری چکش و دما است. تصمیم‌های پذیرش را با "
    "نمونه‌های استاندارد عمل‌آوری‌شده (مانند استوانه‌های ASTM C39) تأیید کنید."
)


class ReboundError(ValueError):
    """Raised for invalid rebound data or correlations."""


@dataclass(frozen=True)
class ReboundSummary:
    """Quality/consistency summary of one test area.

    Attributes:
        readings: All readings taken in the area.
        mean: Arithmetic mean of all readings.
        outliers: Readings deviating from ``mean`` by more than
            :data:`OUTLIER_THRESHOLD_UNITS`; on site these would be
            discarded and re-measured.
        retained_mean: Mean of the readings after excluding outliers
            (equals ``mean`` when there are no outliers).
    """

    readings: Tuple[float, ...]
    mean: float
    outliers: Tuple[float, ...]
    retained_mean: float


@dataclass(frozen=True)
class LinearCorrelation:
    """A *project-specific* linear rebound → strength correlation.

    The coefficients must be established for the concrete and curing
    conditions of the actual project per ASTM C805; they are never
    universal.

    Attributes:
        correlation_id: Traceable identifier of the correlation study
            (e.g. ``"PRJ-1403-cubes-vs-rebound"``).
        slope: Strength gain per rebound unit (MPa/unit).
        intercept: Strength offset (MPa).
    """

    correlation_id: str
    slope: float
    intercept: float

    def __call__(self, rebound_number: float) -> float:
        """Evaluate the correlation for one rebound number."""
        return self.slope * rebound_number + self.intercept


@dataclass(frozen=True)
class StrengthEstimate:
    """An indicative strength estimate with its mandatory disclaimer.

    Attributes:
        rebound_number: The (corrected) rebound number used.
        correlation_id: Identity of the correlation that produced the
            estimate — empty correlations are rejected upstream.
        estimated_strength: Indicative strength (MPa).
        scope: Always :attr:`Scope.ESTIMATION`.
        disclaimer: The full disclaimer text that must accompany the
            value wherever it is rendered.
    """

    rebound_number: float
    correlation_id: str
    estimated_strength: Quantity
    scope: Scope = Scope.ESTIMATION
    disclaimer: str = DISCLAIMER


def mean_rebound_number(readings: Tuple[float, ...]) -> ReboundSummary:
    """Average repeated readings using the standard's outlier rule.

    The mean of all readings is computed first; every reading that
    differs from that mean by more than :data:`OUTLIER_THRESHOLD_UNITS`
    is flagged as an outlier, and the retained mean (excluding outliers)
    is reported.  Software cannot re-measure discarded points, so the
    caller should repeat them on site.

    Args:
        readings: Rebound readings for one test area.  At least
            :data:`MIN_READINGS_PER_AREA` values, each within 0–100.

    Returns:
        A :class:`ReboundSummary` with mean, outliers and retained mean.

    Raises:
        ReboundError: If fewer than :data:`MIN_READINGS_PER_AREA`
            readings are supplied or any reading is outside 0–100.
    """
    if len(readings) < MIN_READINGS_PER_AREA:
        raise ReboundError(
            f"{SPEC.name} requires at least {MIN_READINGS_PER_AREA} readings per test "
            f"area, got {len(readings)}"
        )
    for reading in readings:
        if not math.isfinite(reading) or not 0.0 <= reading <= 100.0:
            raise ReboundError(f"Rebound reading must lie within [0, 100], got {reading!r}")

    mean = sum(readings) / len(readings)
    outliers = tuple(r for r in readings if abs(r - mean) > OUTLIER_THRESHOLD_UNITS)
    retained = tuple(r for r in readings if abs(r - mean) <= OUTLIER_THRESHOLD_UNITS)
    retained_mean = sum(retained) / len(retained) if retained else mean

    return ReboundSummary(
        readings=tuple(readings),
        mean=mean,
        outliers=outliers,
        retained_mean=retained_mean,
    )


def estimate_strength(
    rebound_number: float,
    correlation: Callable[[float], float],
    correlation_id: str,
) -> StrengthEstimate:
    """Derive an *indicative* strength from a rebound number.

    This function refuses to run without an explicit caller-supplied
    correlation: there is deliberately no default formula, because a
    universal rebound → strength conversion does not exist.

    Args:
        rebound_number: Rebound number to evaluate (0–100), typically
            :attr:`ReboundSummary.retained_mean`.
        correlation: Project-specific callable mapping a rebound number
            to a strength in MPa (see :class:`LinearCorrelation`).
        correlation_id: Non-empty, traceable identifier of the
            correlation study used.

    Returns:
        A :class:`StrengthEstimate` that always carries
        :attr:`Scope.ESTIMATION` and the full disclaimer.

    Raises:
        ReboundError: If the rebound number is out of range, the
            correlation id is empty, or the correlation yields a
            non-finite or negative strength.
    """
    if not math.isfinite(rebound_number) or not 0.0 <= rebound_number <= 100.0:
        raise ReboundError(f"Rebound number must lie within [0, 100], got {rebound_number!r}")
    if not isinstance(correlation_id, str) or not correlation_id.strip():
        raise ReboundError("correlation_id must identify the correlation study used")

    estimated = float(correlation(rebound_number))
    if not math.isfinite(estimated) or estimated < 0.0:
        raise ReboundError(
            f"Correlation {correlation_id!r} produced an implausible strength "
            f"{estimated!r} MPa for rebound {rebound_number!r}"
        )

    return StrengthEstimate(
        rebound_number=rebound_number,
        correlation_id=correlation_id,
        estimated_strength=Quantity(estimated, "MPa"),
    )
