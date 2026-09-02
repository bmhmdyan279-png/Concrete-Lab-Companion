"""Declarative descriptions of laboratory standards.

The legacy ``build.py`` keeps every referenced standard in a plain
dictionary (``STANDARDS = {"C136": {"name": ..., "edition": ...}}``).
:class:`StandardSpec` replaces that loosely-typed structure with a
validated, immutable dataclass so that typos in an edition or an
unknown status fail at construction time instead of surfacing as a
broken cell in the generated workbook.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, ClassVar, FrozenSet, Mapping


@dataclass(frozen=True)
class StandardSpec:
    """Immutable specification of a single laboratory standard.

    Attributes:
        code: Short registry key used throughout the code base, e.g.
            ``"C136"``, ``"D2419"``, ``"ISIRI302"``.  Must be non-empty.
        name: Official designation, e.g. ``"ASTM C136/C136M"``.  Must
            be non-empty.
        title: Human-readable title of the method, e.g.
            ``"Sieve Analysis"``.  Must be non-empty.
        edition: Revision year as a string (``"2024"``).  Kept as a
            string because some standards embed letters in their
            designation.  Must be non-empty.
        status: Publication status; one of :data:`VALID_STATUSES`.

    Raises:
        ValueError: If any text field is empty/whitespace-only or
            ``status`` is not one of :data:`VALID_STATUSES`.

    Example:
        >>> spec = StandardSpec(code="C136", name="ASTM C136/C136M",
        ...                     title="Sieve Analysis", edition="2024")
        >>> spec.is_active
        True
        >>> spec.citation()
        'ASTM C136/C136M (2024)'
    """

    code: str
    name: str
    title: str
    edition: str
    status: str = "active"

    #: Statuses accepted for :attr:`status`.
    VALID_STATUSES: ClassVar[FrozenSet[str]] = frozenset(
        {"active", "withdrawn", "superseded"}
    )

    def __post_init__(self) -> None:
        """Validate every field of the specification."""
        for field_name in ("code", "name", "title", "edition"):
            value = getattr(self, field_name)
            if not isinstance(value, str) or not value.strip():
                raise ValueError(
                    f"StandardSpec.{field_name} must be a non-empty string"
                )
        if self.status not in self.VALID_STATUSES:
            raise ValueError(
                f"StandardSpec.status must be one of {sorted(self.VALID_STATUSES)}, "
                f"got {self.status!r}"
            )

    # -- Construction helpers ------------------------------------------------

    @classmethod
    def from_dict(cls, code: str, data: Mapping[str, Any]) -> "StandardSpec":
        """Build a spec from a legacy ``STANDARDS`` dictionary entry.

        This is the migration bridge from the old ``build.py`` format::

            {"C136": {"name": ..., "title": ..., "edition": ..., "status": ...}}

        Args:
            code: Registry key of the standard (the dict key).
            data: Mapping with ``name``, ``title``, ``edition`` and
                optionally ``status`` (defaults to ``"active"``).

        Returns:
            A validated :class:`StandardSpec`.

        Raises:
            ValueError: If a required key is missing or any field is
                invalid (see :meth:`__post_init__`).
        """
        try:
            return cls(
                code=code,
                name=str(data["name"]),
                title=str(data["title"]),
                edition=str(data["edition"]),
                status=str(data.get("status", "active")),
            )
        except KeyError as exc:
            raise ValueError(
                f"legacy standard entry {code!r} is missing key {exc.args[0]!r}"
            ) from exc

    # -- Queries -------------------------------------------------------------

    @property
    def is_active(self) -> bool:
        """Return ``True`` when the standard is still in force."""
        return self.status == "active"

    def citation(self) -> str:
        """Return the display citation, e.g. ``"ASTM C136/C136M (2024)"``."""
        return f"{self.name} ({self.edition})"

    def __str__(self) -> str:
        """Return the short form ``code: citation``."""
        return f"{self.code}: {self.citation()}"
