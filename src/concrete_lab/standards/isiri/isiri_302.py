"""ISIRI 302 — Characteristics of concrete aggregates: grading ruleset.

Refactored ruleset addressing review feedback: ISIRI 302 (ویژگی‌های
سنگدانه‌های بتن, 3rd edition approved 1394) is modelled as a real
ruleset instead of a bare ``{sieve: (upper, lower)}`` dictionary:

* Each sieve limit is a typed :class:`SieveLimit` with its own
  evaluation logic and a configurable *warning margin* near the
  envelope boundaries.
* :class:`GradingRuleSet` evaluates a full grading curve, separating
  **standard rules** (the grading envelope) from **physical rules**
  (passing percentages must be within 0–100 and non-increasing as the
  sieve size decreases) and reporting missing sieves explicitly.
* Results use :class:`ValidationStatus` so that FAIL (outside the
  envelope), WARN (inside but near a boundary), ERROR (incomplete or
  physically impossible data) are never conflated.

Two derived quantities live here as well, because they are part of the
same science and must not be re-implemented by callers or spreadsheets:

* :func:`curve_from_retained` — the laboratory measures *masses retained
  on each sieve*; the ruleset evaluates *cumulative percent passing*.
  The conversion (including the ASTM C136 mass-balance rule) is done once,
  here, under test.
* :func:`fineness_modulus` — the ASTM C136 fineness-modulus index over the
  fine-aggregate sieve series.

Percent passing values are percentages (0–100); sieve sizes are mm.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import Dict, List, Mapping, Optional, Tuple

from concrete_lab.domain.statuses import ValidationStatus
from concrete_lab.standards.base import Scope, StandardSpec

#: Metadata for this ruleset (3rd edition, approved 1394).
SPEC: StandardSpec = StandardSpec(
    code="ISIRI302",
    name="ISIRI 302",
    title="ویژگی‌های سنگدانه‌های بتن (Characteristics of Concrete Aggregates)",
    edition="1394",
    status="active",
)

#: Grading envelope checks are normative for aggregate acceptance.
SCOPE: Scope = Scope.STANDARD

#: Version tag so every derived result can be traced to a ruleset revision.
RULESET_VERSION: str = "ISIRI-302-1394-v1"

#: Values inside an envelope band but closer than this many percentage
#: points to a boundary are reported as WARN (near-limit caution).
WARNING_MARGIN_POINTS: float = 2.0

#: Sieve series used for the **fineness modulus of fine aggregate**
#: (ASTM C136/C136M: 3/8 in., No. 4, No. 8, No. 16, No. 30, No. 50,
#: No. 100).  The No. 200 sieve (0.075 mm) and the pan are deliberately
#: excluded: including the pan would add a constant 100 % to the sum and
#: inflate the index by exactly 1.00.
FM_SIEVE_SERIES_MM: Tuple[float, ...] = (9.5, 4.75, 2.36, 1.18, 0.600, 0.300, 0.150)

#: Reporting precision of the fineness modulus (ASTM C136: nearest 0.01).
FM_REPORT_STEP: float = 0.01

#: ASTM C136 mass-balance rule: the sum of the masses retained on all
#: sieves plus the pan must agree with the original sample mass within
#: this relative tolerance, otherwise the analysis is discarded.
MASS_BALANCE_TOLERANCE: float = 0.003

#: Reporting precision of cumulative percent passing / percent retained
#: (ASTM C136: nearest 0.1 %).  Rounding here is a *standard requirement*,
#: not cosmetics: it also keeps binary floating-point noise such as
#: ``32.99999999999999`` out of the evaluated curve and out of the
#: rendered worksheet.
PERCENT_REPORT_STEP: float = 0.1


class GradingError(ValueError):
    """Raised for invalid grading data (incomplete curve, bad masses)."""


@dataclass(frozen=True)
class SieveLimit:
    """One row of the ISIRI 302 grading envelope.

    Attributes:
        sieve_mm: Sieve opening in millimetres.
        min_passing: Lower bound of percent passing (inclusive).
        max_passing: Upper bound of percent passing (inclusive).

    Raises:
        ValueError: If bounds are outside 0–100, unordered, or the
            sieve size is not positive.
    """

    sieve_mm: float
    min_passing: float
    max_passing: float

    def __post_init__(self) -> None:
        """Validate the limit row."""
        if self.sieve_mm <= 0 or not math.isfinite(self.sieve_mm):
            raise ValueError(f"sieve_mm must be positive, got {self.sieve_mm!r}")
        if not 0.0 <= self.min_passing <= self.max_passing <= 100.0:
            raise ValueError(
                f"limits must satisfy 0 <= min <= max <= 100, got "
                f"[{self.min_passing}, {self.max_passing}]"
            )

    @property
    def band_width(self) -> float:
        """Width of the acceptance band in percentage points."""
        return self.max_passing - self.min_passing

    def contains(self, percent_passing: float) -> bool:
        """Return ``True`` when the value lies inside the envelope band."""
        return self.min_passing <= percent_passing <= self.max_passing

    def status_for(self, percent_passing: float, margin: float = WARNING_MARGIN_POINTS) -> ValidationStatus:
        """Classify one percent-passing value against this limit.

        Args:
            percent_passing: Measured cumulative percent passing.
            margin: Near-boundary warning margin (percentage points).
                A warning band is only applied when the acceptance band
                is wide enough (wider than ``2 * margin``) so that
                narrow bands keep a meaningful PASS region.

        Returns:
            * :attr:`ValidationStatus.FAIL` when outside the band;
            * :attr:`ValidationStatus.WARN` when inside but within
              ``margin`` points of a boundary (band wide enough);
            * :attr:`ValidationStatus.PASS` otherwise.
        """
        if not self.contains(percent_passing):
            return ValidationStatus.FAIL
        if self.band_width > 2 * margin:
            near_lower = percent_passing < self.min_passing + margin
            near_upper = percent_passing > self.max_passing - margin
            if near_lower or near_upper:
                return ValidationStatus.WARN
        return ValidationStatus.PASS


#: The ISIRI 302 grading envelope for fine aggregate, ordered from the
#: largest sieve down (percent passing decreases with sieve size).
GRADING_LIMITS: Tuple[SieveLimit, ...] = (
    SieveLimit(sieve_mm=9.5, min_passing=95.0, max_passing=100.0),
    SieveLimit(sieve_mm=4.75, min_passing=80.0, max_passing=95.0),
    SieveLimit(sieve_mm=2.36, min_passing=60.0, max_passing=80.0),
    SieveLimit(sieve_mm=1.18, min_passing=40.0, max_passing=60.0),
    SieveLimit(sieve_mm=0.600, min_passing=25.0, max_passing=40.0),
    SieveLimit(sieve_mm=0.300, min_passing=10.0, max_passing=25.0),
    SieveLimit(sieve_mm=0.150, min_passing=2.0, max_passing=10.0),
    SieveLimit(sieve_mm=0.075, min_passing=0.0, max_passing=2.0),
)


@dataclass(frozen=True)
class SieveEvaluation:
    """Result of checking one sieve against the envelope.

    Attributes:
        sieve_mm: Sieve opening in millimetres.
        percent_passing: Measured cumulative percent passing.
        limit: The envelope row that was applied.
        status: PASS / WARN / FAIL classification.
    """

    sieve_mm: float
    percent_passing: float
    limit: SieveLimit
    status: ValidationStatus


@dataclass(frozen=True)
class GradingEvaluation:
    """Full outcome of evaluating a grading curve against the ruleset.

    Attributes:
        evaluations: Per-sieve envelope checks (largest sieve first).
        missing_sieves: Sieves required by the ruleset but absent from
            the submitted curve.
        physical_issues: Violations of physical sanity rules (values
            outside 0–100 or a non-monotonic curve).
    """

    evaluations: Tuple[SieveEvaluation, ...]
    missing_sieves: Tuple[float, ...] = field(default_factory=tuple)
    physical_issues: Tuple[str, ...] = field(default_factory=tuple)

    @property
    def overall_status(self) -> ValidationStatus:
        """Aggregate status of the whole evaluation.

        * ``ERROR`` when the curve is incomplete (missing sieves) or
          physically impossible;
        * ``FAIL`` when any sieve lies outside the envelope;
        * ``WARN`` when all sieves pass but at least one is near a
          boundary;
        * ``PASS`` otherwise.
        """
        if self.missing_sieves or self.physical_issues:
            return ValidationStatus.ERROR
        statuses = {ev.status for ev in self.evaluations}
        if ValidationStatus.FAIL in statuses:
            return ValidationStatus.FAIL
        if ValidationStatus.WARN in statuses:
            return ValidationStatus.WARN
        return ValidationStatus.PASS


@dataclass(frozen=True)
class GradingRuleSet:
    """Executable ISIRI 302 grading ruleset.

    Attributes:
        standard: The standard metadata this ruleset implements.
        ruleset_version: Traceable ruleset revision tag.
        limits: Envelope rows, largest sieve first.
        warning_margin: Near-boundary warning margin in percentage points.
    """

    standard: StandardSpec = SPEC
    ruleset_version: str = RULESET_VERSION
    limits: Tuple[SieveLimit, ...] = GRADING_LIMITS
    warning_margin: float = WARNING_MARGIN_POINTS

    # -- Queries -------------------------------------------------------------

    @property
    def sieve_sizes(self) -> Tuple[float, ...]:
        """All sieve openings required by the ruleset (mm, descending)."""
        return tuple(limit.sieve_mm for limit in self.limits)

    def limit_for(self, sieve_mm: float) -> Optional[SieveLimit]:
        """Return the envelope row for a sieve size, if the ruleset has one."""
        for limit in self.limits:
            if math.isclose(limit.sieve_mm, sieve_mm, abs_tol=1e-9):
                return limit
        return None

    # -- Evaluation ----------------------------------------------------------

    def evaluate(self, curve: Mapping[float, float]) -> GradingEvaluation:
        """Evaluate a grading curve against the envelope and physical rules.

        Args:
            curve: Mapping of sieve opening (mm) to cumulative percent
                passing.  Keys are matched to the ruleset's sieves with
                floating-point tolerance.

        Returns:
            A :class:`GradingEvaluation`; see
            :attr:`GradingEvaluation.overall_status` for the status
            precedence (ERROR > FAIL > WARN > PASS).
        """
        matched: Dict[float, float] = {}
        for size, percent in curve.items():
            limit = self.limit_for(size)
            if limit is not None:
                matched[limit.sieve_mm] = percent

        evaluations = []
        missing = []
        for limit in self.limits:
            if limit.sieve_mm not in matched:
                missing.append(limit.sieve_mm)
                continue
            percent = matched[limit.sieve_mm]
            evaluations.append(
                SieveEvaluation(
                    sieve_mm=limit.sieve_mm,
                    percent_passing=percent,
                    limit=limit,
                    status=limit.status_for(percent, self.warning_margin),
                )
            )

        return GradingEvaluation(
            evaluations=tuple(evaluations),
            missing_sieves=tuple(missing),
            physical_issues=self._physical_issues(evaluations),
        )

    def _physical_issues(self, evaluations: Tuple[SieveEvaluation, ...]) -> Tuple[str, ...]:
        """Check physical sanity of the evaluated (sieve, passing) pairs.

        Rules: every percent passing must lie in 0–100, and the curve
        must be non-increasing as the sieve opening decreases (material
        cannot pass a smaller sieve more than a larger one).
        """
        issues = []
        for ev in evaluations:
            if not math.isfinite(ev.percent_passing) or not 0.0 <= ev.percent_passing <= 100.0:
                issues.append(
                    f"percent passing {ev.percent_passing!r} on sieve {ev.sieve_mm} mm "
                    f"is outside the physical range [0, 100]"
                )
        for upper, lower in zip(evaluations, evaluations[1:]):
            if lower.percent_passing > upper.percent_passing:
                issues.append(
                    f"curve increases from {upper.percent_passing:g}% at {upper.sieve_mm} mm "
                    f"to {lower.percent_passing:g}% at {lower.sieve_mm} mm (must be non-increasing)"
                )
        return tuple(issues)


#: Shared default ruleset instance (the envelope formerly known as the
#: ``ISIRI_LIMITS`` dictionary in ``build.py``).
DEFAULT_RULESET: GradingRuleSet = GradingRuleSet()


# ─── Derived grading quantities ───────────────────────────────────────────


def _round_to_step(value: float, step: float) -> float:
    """Round ``value`` to the nearest multiple of ``step`` (half-even)."""
    return round(round(value / step) * step, 10)


def fineness_modulus(
    curve: Mapping[float, float],
    series: Tuple[float, ...] = FM_SIEVE_SERIES_MM,
) -> float:
    """Compute the ASTM C136/C136M fineness modulus of a grading curve.

    The fineness modulus is the sum of the **cumulative percentages
    retained** on the specified sieve series, divided by 100.  It is an
    index of average particle size, not an acceptance criterion: ASTM
    C136 defines no pass/fail limit for it.

    Args:
        curve: Mapping of sieve opening (mm) to cumulative percent
            passing.  Keys are matched to ``series`` with floating-point
            tolerance.
        series: Sieves to include in the sum.  Defaults to the
            fine-aggregate series :data:`FM_SIEVE_SERIES_MM`.

    Returns:
        The fineness modulus, rounded to :data:`FM_REPORT_STEP`.

    Raises:
        GradingError: If any sieve of ``series`` is missing from the
            curve (an incomplete analysis cannot yield a meaningful
            index) or if a percent passing is not finite.

    Example:
        >>> curve = {9.5: 97.5, 4.75: 88.0, 2.36: 70.0, 1.18: 50.0,
        ...          0.6: 33.0, 0.3: 18.0, 0.15: 6.0, 0.075: 1.0}
        >>> fineness_modulus(curve)
        3.38
    """
    total = 0.0
    missing: List[float] = []
    for size in series:
        percent_passing = next(
            (value for key, value in curve.items() if math.isclose(float(key), size, abs_tol=1e-9)),
            None,
        )
        if percent_passing is None:
            missing.append(size)
            continue
        if not math.isfinite(float(percent_passing)):
            raise GradingError(f"percent passing on sieve {size} mm is not finite: {percent_passing!r}")
        total += 100.0 - float(percent_passing)

    if missing:
        raise GradingError(
            f"fineness modulus needs the full series {series}; missing sieves (mm): {missing}"
        )
    return _round_to_step(total / 100.0, FM_REPORT_STEP)


def curve_from_retained(
    retained_by_sieve: Mapping[float, float],
    total_mass: Optional[float] = None,
    pan_mass: float = 0.0,
) -> Dict[float, float]:
    """Convert weighed retained masses into a cumulative percent-passing curve.

    A laboratory records the **mass retained on each sieve**; every
    grading rule in this module works on **cumulative percent passing**.
    Keeping that conversion here (rather than in a spreadsheet column or
    a QA adapter) means it is unit-tested and traceable like any other
    piece of science.

    Args:
        retained_by_sieve: Sieve opening (mm) → mass retained on that
            sieve (any consistent mass unit; grams by convention).
        total_mass: Mass of the original sample.  When supplied, the
            ASTM C136 mass-balance rule is enforced (see
            :data:`MASS_BALANCE_TOLERANCE`); when omitted it defaults to
            ``sum(retained) + pan_mass`` and no balance check is possible.
        pan_mass: Mass collected in the pan.  The pan is not a sieve, so
            it enters the balance but not the curve.

    Returns:
        Sieve opening (mm) → cumulative percent passing rounded to
        :data:`PERCENT_REPORT_STEP`, ordered from the largest sieve down
        (material that never reached a sieve passes 100 %).

    Raises:
        GradingError: If the retained masses are empty, any mass is
            negative or not finite, the effective total is not positive,
            or the mass balance disagrees beyond tolerance.

    Example:
        >>> curve_from_retained({4.75: 25.0, 2.36: 95.0}, total_mass=500.0,
        ...                     pan_mass=380.0)
        {4.75: 95.0, 2.36: 76.0}
    """
    if not retained_by_sieve:
        raise GradingError("retained_by_sieve must contain at least one sieve")

    masses: Dict[float, float] = {}
    for size, mass in retained_by_sieve.items():
        size_f = float(size)
        mass_f = float(mass)
        if size_f <= 0 or not math.isfinite(size_f):
            raise GradingError(f"sieve size must be a positive finite number, got {size!r}")
        if mass_f < 0 or not math.isfinite(mass_f):
            raise GradingError(f"retained mass on sieve {size_f:g} mm must be >= 0, got {mass!r}")
        if size_f in masses:
            raise GradingError(f"duplicate sieve size {size_f:g} mm")
        masses[size_f] = mass_f

    pan = float(pan_mass)
    if pan < 0 or not math.isfinite(pan):
        raise GradingError(f"pan_mass must be >= 0, got {pan_mass!r}")

    summed = sum(masses.values()) + pan
    if total_mass is None:
        effective_total = summed
    else:
        effective_total = float(total_mass)
        if effective_total <= 0 or not math.isfinite(effective_total):
            raise GradingError(f"total_mass must be positive, got {total_mass!r}")
        drift = abs(summed - effective_total) / effective_total
        if drift > MASS_BALANCE_TOLERANCE:
            raise GradingError(
                f"mass balance fails ASTM C136: sieves + pan sum to {summed:.4g} but the "
                f"sample mass is {effective_total:.4g} (drift {drift * 100:.2f} % > "
                f"{MASS_BALANCE_TOLERANCE * 100:.1f} %) — re-weigh or check for material loss"
            )

    if effective_total <= 0:
        raise GradingError("total mass must be positive to express a grading curve in percent")

    curve: Dict[float, float] = {}
    cumulative_retained = 0.0
    for size in sorted(masses, reverse=True):
        cumulative_retained += masses[size]
        percent_passing = (1.0 - cumulative_retained / effective_total) * 100.0
        curve[size] = _round_to_step(percent_passing, PERCENT_REPORT_STEP)
    return curve
