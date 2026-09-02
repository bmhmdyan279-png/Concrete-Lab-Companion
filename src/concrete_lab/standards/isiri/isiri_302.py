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

Percent passing values are percentages (0–100); sieve sizes are mm.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import Dict, Mapping, Optional, Tuple

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
