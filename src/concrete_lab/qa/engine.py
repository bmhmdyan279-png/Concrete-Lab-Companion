"""QA engine: run the real pytest suite and report PASS / FAIL / WARN.

The legacy ``--validate`` only printed the registry size; this module
executes pytest and turns the outcome into a structured
:class:`QAReport` that the CLI, the manifest and the QA sheet consume.

pytest runs **in-process** on purpose: the CLI is a short-lived process,
and every automated caller (``build.py --validate`` in CI, the end-to-end
tests) invokes it as a subprocess, so the host interpreter is never left
with pytest's global state.  What in-process execution *cannot* hide is a
crash, so :meth:`QAEngine.run_pytest` reads pytest's exit code as well as
the per-test hook: a collection error, an internal error or an interrupt
is reported as a failure instead of silently counting zero tests and
declaring PASS.
"""

from __future__ import annotations

import time
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any, Dict, List, Optional, Sequence, Tuple

from concrete_lab.domain.statuses import ValidationStatus

if TYPE_CHECKING:  # pragma: no cover - annotation-only, never imported at runtime
    import pytest


class QAError(RuntimeError):
    """Raised when a QA tier cannot run because its tooling is missing."""


def _import_pytest() -> Any:
    """Import pytest lazily.

    pytest is a **development** dependency, not a runtime one.  Importing it
    at module level made ``pip install concrete-lab-companion`` produce a
    package whose ``concrete_lab.cli`` could not even be imported without
    the dev extras installed — building a workbook does not need a test
    runner.  Only the ``--validate`` tier does, so the import happens here
    and the failure is reported as an actionable message.

    Raises:
        QAError: If pytest is not installed.
    """
    try:
        import pytest
    except ModuleNotFoundError as exc:
        raise QAError(
            "the pytest QA tier needs the development dependencies — "
            "install them with `pip install -r requirements-dev.txt` "
            "or `pip install concrete-lab-companion[dev]`"
        ) from exc
    return pytest


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
        notes: Short human-readable descriptions of each warning.  A
            warning count with no explanation is not actionable, so
            every tier that can warn says *what* it skipped and why.
    """

    source: str
    passed: int
    failed: int
    warnings: int = 0
    duration_seconds: float = 0.0
    failures: Tuple[str, ...] = field(default_factory=tuple)
    notes: Tuple[str, ...] = field(default_factory=tuple)

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
            "notes": list(self.notes),
        }

    def merged_with(self, other: QAReport) -> QAReport:
        """Combine two reports into one (failures and notes concatenated)."""
        return QAReport(
            source=f"{self.source}+{other.source}",
            passed=self.passed + other.passed,
            failed=self.failed + other.failed,
            warnings=self.warnings + other.warnings,
            duration_seconds=self.duration_seconds + other.duration_seconds,
            failures=self.failures + other.failures,
            notes=self.notes + other.notes,
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
    def pytest_runtest_logreport(self, report: pytest.TestReport) -> None:
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


def _describe(report: pytest.TestReport) -> str:
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


#: pytest exit codes that mean "the run itself broke", as opposed to
#: "tests ran and some failed".  ``ExitCode`` is an IntEnum, so these are
#: compared against plain ints to stay independent of pytest internals.
PYTEST_OK: int = 0
PYTEST_TESTS_FAILED: int = 1
PYTEST_INTERRUPTED: int = 2
PYTEST_USAGE_ERROR: int = 3
PYTEST_INTERNAL_ERROR: int = 4
PYTEST_NO_TESTS_COLLECTED: int = 5

#: Human-readable description of each non-test-run exit code.
PYTEST_EXIT_MEANING: Dict[int, str] = {
    PYTEST_INTERRUPTED: "interrupted before completion",
    PYTEST_USAGE_ERROR: "usage error (bad arguments or configuration)",
    PYTEST_INTERNAL_ERROR: "aborted before the suite ran (internal or usage error)",
    PYTEST_NO_TESTS_COLLECTED: "no tests were collected",
}


class QAEngine:
    """Runs QA tiers and produces :class:`QAReport` objects."""

    def run_pytest(self, pytest_args: Optional[Sequence[str]] = None) -> QAReport:
        """Execute pytest in-process and classify the outcome.

        Args:
            pytest_args: Extra arguments (test paths, ``-k`` filters…).
                Defaults to the configured ``testpaths`` selection.

        Returns:
            A :class:`QAReport` where skipped/xfail tests count as
            warnings; any test failure flips the status to FAIL.  A run
            that never executed tests (crash, bad arguments, empty
            collection) is reported as a failure with pytest's exit code,
            never as an empty PASS.
        """
        pytest = _import_pytest()
        collector = _Collector()
        # ``--import-mode=importlib`` keeps repeated in-process runs from
        # colliding in ``sys.modules`` when two temporary test files share a
        # basename (pytest's default prepend mode raises "import file
        # mismatch" and aborts collection, which would be reported as an
        # internal error rather than as the test outcome).
        args = ["-q", "--tb=line", "-p", "no:cacheprovider",
                "--import-mode=importlib", *(pytest_args or ())]
        started = time.perf_counter()
        exit_code = int(pytest.main(args, plugins=[collector]))
        duration = time.perf_counter() - started

        failures = list(collector.failures)
        notes: List[str] = []
        failed = collector.failed

        if exit_code not in (PYTEST_OK, PYTEST_TESTS_FAILED):
            meaning = PYTEST_EXIT_MEANING.get(exit_code, "did not complete normally")
            failures.append(f"pytest exit code {exit_code}: {meaning}")
            failed += 1
        elif exit_code == PYTEST_OK and collector.passed == 0 and not collector.skipped:
            # Exit code 0 with nothing counted means the selection matched no
            # test; a QA tier that checked nothing must not claim success.
            notes.append("pytest reported success without running any test")

        return QAReport(
            source="pytest-suite",
            passed=collector.passed,
            failed=failed,
            warnings=collector.skipped + collector.xfailed,
            duration_seconds=duration,
            failures=tuple(failures),
            notes=tuple(notes),
        )
