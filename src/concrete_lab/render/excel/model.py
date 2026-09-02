"""Pure-data model of a rendered workbook.

These dataclasses are the contract between the domain layer (what was
computed) and the renderer (how it is displayed).  They carry no
behaviour — a renderer that only receives a :class:`WorkbookModel`
cannot possibly re-implement scientific logic.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Optional, Tuple


@dataclass(frozen=True)
class CellModel:
    """One rendered cell.

    Attributes:
        value: Display value (text or number — never a formula).
        style: Name of the named style applied to the cell.
        number_format: Optional Excel number format (e.g. ``"0.00"``).
    """

    value: Any
    style: str = "label"
    number_format: Optional[str] = None


@dataclass(frozen=True)
class RowModel:
    """One rendered row (tuple of cells, left-to-right storage)."""

    cells: Tuple[CellModel, ...]


@dataclass(frozen=True)
class SheetModel:
    """One rendered worksheet.

    Attributes:
        title: Worksheet title (Excel's 31-char limit is the caller's
            responsibility).
        rows: Rows in display order.
        tab_color: Optional tab colour (hex without ``#``).
        rtl: Right-to-left layout (Persian sheets).
        column_widths: Optional ``(column, width)`` pairs; defaults
            applied by the renderer for the rest.
    """

    title: str
    rows: Tuple[RowModel, ...] = field(default_factory=tuple)
    tab_color: Optional[str] = None
    rtl: bool = True
    column_widths: Tuple[Tuple[int, float], ...] = field(default_factory=tuple)


@dataclass(frozen=True)
class WorkbookModel:
    """A complete workbook as pure data."""

    sheets: Tuple[SheetModel, ...]

    @property
    def sheet_titles(self) -> Tuple[str, ...]:
        """Titles in workbook order."""
        return tuple(sheet.title for sheet in self.sheets)
