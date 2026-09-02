"""QA engine: run the real pytest suite and report PASS / FAIL / WARN.

The legacy ``--validate`` only printed the registry size; this module
executes pytest and turns the outcome into a structured
:class:`QAReport` that the CLI, the manifest and the QA sheet consume.
"""

from __future__ import annotations

import time
from dataclasses import dataclass, field
from typing import List, Optional, Sequence, Tuple

import pytest

from concrete_lab.domain.statuses import ValidationStatus


@dataclass(frozen=True)
class QAReport:
    """Structured outcome of one QA tier.

    Attributes:
        source: Name of the tier (e.g. ``"pytest-suite"``).
        passed: Number of successful checks/tests.
        failed: Number of failed checks/tests.
        warnings: Number of skipped/xfail checks — not failures, but
            they deserve attention.
        duration_seconds: Wall-clock time of the tier.
        failures: Short human-readable descriptions of each failure.
    """

    source: str
    passed: int
    failed: int
    warnings: int = 0
    duration_seconds: float = 0.0
    failures: Tuple[str, ...] = field(default_factory=tuple)

    @property
    def total(self) -> int:
        """Total number of evaluated checks/tests."""
        return self.passed + self.failed + self.warnings

    @property
    def status(self) -> ValidationStatus:
        """Aggregate status: FAIL > WARN > PASS precedence."""
        if self.failed:
            return ValidationStatus.FAIL
        if self.warnings:
            return ValidationStatus.WARN
        return ValidationStatus.PASS

    def to_dict(self) -> dict:
        """Serialise the report for the manifest / QA sheet."""
        return {
            "source": self.source,
            "passed": self.passed,
            "failed": self.failed,
            "warnings": self.warnings,
            "total": self.total,
            "duration_seconds": round(self.duration_seconds, 3),
            "status": self.status.value,
            "failures": list(self.failures),
        }

    def merged_with(self, other: "QAReport") -> "QAReport":
        """Combine two reports into one (failures concatenated)."""
        return QAReport(
            source=f"{self.source}+{other.source}",
            passed=self.passed + other.passed,
            failed=self.failed + other.failed,
            warnings=self.warnings + other.warnings,
            duration_seconds=self.duration_seconds + other.duration_seconds,
            failures=self.failures + other.failures,
        )


class _Collector:
    """Pytest plugin that tallies outcomes into counters."""

    def __init__(self) -> None:
        self.passed = 0
        self.failed = 0
        self.skipped = 0
        self.xfailed = 0
        self.failures: List[str] = []

    # pytest hook: called for every test phase report
    def pytest_runtest_logreport(self, report: "pytest.TestReport") -> None:
        if report.when == "call":
            if report.passed:
                self.passed += 1
            elif report.failed:
                self.failed += 1
                self.failures.append(_describe(report))
            elif report.skipped:
                if hasattr(report, "wasxfail"):
                    self.xfailed += 1
                else:
                    self.skipped += 1
        elif report.when == "setup":
            if report.skipped:  # pytest.mark.skip reports during setup
                if hasattr(report, "wasxfail"):
                    self.xfailed += 1
                else:
                    self.skipped += 1
            elif report.failed:
                self.failed += 1
                self.failures.append(_describe(report))
        elif report.when == "teardown" and report.failed:
            self.failed += 1
            self.failures.append(_describe(report))


def _describe(report: "pytest.TestReport") -> str:
    """One-line failure description for reports and the manifest."""
    longrepr = report.longrepr
    message = "unknown failure"
    crash = getattr(longrepr, "reprcrash", None)
    if crash is not None and getattr(crash, "message", None):
        message = str(crash.message)
    elif isinstance(longrepr, tuple) and len(longrepr) >= 3:
        message = str(longrepr[2])
    elif longrepr is not None:
        lines = [line for line in str(longrepr).splitlines() if line.strip()]
        if lines:
            message = lines[-1]
    return f"{report.nodeid}: {message[:200]}"


class QAEngine:
    """Runs QA tiers and produces :class:`QAReport` objects."""

    def run_pytest(self, pytest_args: Optional[Sequence[str]] = None) -> QAReport:
        """Execute pytest in-process and classify the outcome.

        Args:
            pytest_args: Extra arguments (test paths, ``-k`` filters…).
                Defaults to the repository ``pytest.ini`` selection.

        Returns:
            A :class:`QAReport` where skipped/xfail tests count as
            warnings; any failure flips the status to FAIL.
        """
        collector = _Collector()
        args = ["-q", "--tb=line", "-p", "no:cacheprovider", *(pytest_args or ())]
        started = time.perf_counter()
        pytest.main(args, plugins=[collector])  # exit code handled via collector
        duration = time.perf_counter() - started

        return QAReport(
            source="pytest-suite",
            passed=collector.passed,
            failed=collector.failed,
            warnings=collector.skipped + collector.xfailed,
            duration_seconds=duration,
            failures=tuple(collector.failures),
        )
