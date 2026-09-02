"""Validation lifecycle statuses for lab results.

Every computed result in a concrete lab test moves through a small
state machine: it starts *pending*, may end up passing, warning,
failing, being skipped, or ending in an error.  :class:`ValidationStatus`
is the single source of truth for those six states so that styles,
dashboard rollups and reports never disagree about what a status means.

The enum supersedes the legacy ``TestStatus`` from ``build.py`` (which
only had ``NOT_READY``/``PASS``/``WARN``/``FAIL``) while keeping its
emoji symbols for visual continuity in the generated workbook.
"""

from __future__ import annotations

from enum import Enum, unique


@unique
class ValidationStatus(Enum):
    """The six possible outcomes of validating a lab result.

    Attributes (members):
        PENDING: The result has not been evaluated yet (inputs missing
            or the check has not run).  Replaces the legacy
            ``TestStatus.NOT_READY``.
        PASS: The result satisfies the standard's acceptance limits.
        WARN: The result is acceptable but lies in a warning band
            (e.g. close to a limit, or a soft guideline is violated).
        FAIL: The result violates the standard's acceptance limits.
        SKIPPED: The check does not apply to this test/specimen and was
            intentionally not evaluated.
        ERROR: The result could not be computed at all (invalid or
            inconsistent inputs, broken formula, ...).

    Note:
        Member *values* are stable lowercase strings intended for
        persistence (manifests, sheet cells); member *names* are for
        Python code.
    """

    PENDING = "pending"
    PASS = "pass"
    WARN = "warn"
    FAIL = "fail"
    SKIPPED = "skipped"
    ERROR = "error"

    # -- Properties ----------------------------------------------------------

    @property
    def is_success(self) -> bool:
        """Return ``True`` only for :attr:`PASS`.

        ``WARN`` is deliberately *not* a success: downstream QA rollups
        must opt in explicitly if they want to treat warnings as
        acceptable.
        """
        return self is ValidationStatus.PASS

    @property
    def is_terminal(self) -> bool:
        """Return ``True`` when the status is a final evaluation outcome.

        ``PENDING`` is the only non-terminal state; every other member
        represents a concluded evaluation (including ``SKIPPED`` and
        ``ERROR``).
        """
        return self is not ValidationStatus.PENDING

    @property
    def symbol(self) -> str:
        """Return the emoji/typographic symbol shown in Excel cells.

        The mapping for ``PENDING``/``PASS``/``WARN``/``FAIL`` matches
        the legacy ``TestStatus`` symbols so regenerated workbooks look
        identical to the v3 output.
        """
        return _SYMBOLS[self]


#: Display symbols per status, used by :attr:`ValidationStatus.symbol`.
_SYMBOLS = {
    ValidationStatus.PENDING: "—",
    ValidationStatus.PASS: "✅",
    ValidationStatus.WARN: "⚠️",
    ValidationStatus.FAIL: "❌",
    ValidationStatus.SKIPPED: "⏭️",
    ValidationStatus.ERROR: "❗",
}
