"""Named constants for the Concrete Lab Companion.

Every magic number/string used by the renderers and engine is declared
here with a descriptive name (review feedback: no anonymous literals
buried in logic).
"""

from __future__ import annotations

from typing import Dict, Final

# ─── Workbook layout ──────────────────────────────────────────────────────

#: Default zoom level for every worksheet.
DEFAULT_ZOOM: Final[int] = 80

#: Default worksheet column widths (1-based column index → width in
#: character units).  Single source of truth for the renderer; a sheet
#: may override individual columns via ``SheetModel.column_widths``.
DEFAULT_COLUMN_WIDTHS: Final[Dict[int, float]] = {1: 36.0, 2: 18.0, 3: 12.0, 4: 14.0, 5: 48.0}

#: Worksheet titles that are internal references (never protected, hidden
#: from user-facing lists).
INTERNAL_SHEET_PREFIX: Final[str] = "_"

# ─── Plausibility windows (physical validation) ───────────────────────────

#: Physically plausible moisture content of aggregates (percent).
PLAUSIBLE_MOISTURE_RANGE: Final[tuple] = (0.0, 20.0)

#: Physically plausible fresh-concrete density (kg/m³).
PLAUSIBLE_DENSITY_RANGE_KG_M3: Final[tuple] = (1600.0, 2800.0)

#: Physically plausible splitting tensile strength (MPa).
PLAUSIBLE_TENSILE_RANGE_MPA: Final[tuple] = (0.5, 15.0)

#: Maximum slump measurable with the standard 300 mm cone (mm).
SLUMP_CONE_HEIGHT_MM: Final[float] = 300.0

#: Data-error window for the fine-aggregate fineness modulus.  ASTM C136
#: sets **no** acceptance limit for the index, so this window only flags
#: values that indicate broken input data (a sieve row dropped from the
#: curve, masses in the wrong unit, …).
FM_DATA_ERROR_WINDOW: Final[tuple] = (1.0, 5.0)

#: Commonly recommended fineness modulus of fine aggregate for concrete
#: (ACI 211.1 practice).  Displayed as an *informational note* — never as
#: a pass/fail criterion, because ASTM C136 defines no limit.
RECOMMENDED_FM_RANGE_FINE: Final[tuple] = (2.3, 3.1)

# ─── Colours (mirrors config.yaml; config wins when loaded) ───────────────

COLOR_HEADER_FILL: Final[str] = "1F4E79"
COLOR_HEADER_FONT: Final[str] = "FFFFFF"
COLOR_PASS_FILL: Final[str] = "C6EFCE"
COLOR_PASS_FONT: Final[str] = "006100"
COLOR_WARN_FILL: Final[str] = "FCE4D6"
COLOR_WARN_FONT: Final[str] = "C00000"
COLOR_FAIL_FILL: Final[str] = "FFC7CE"
COLOR_FAIL_FONT: Final[str] = "9C0006"
COLOR_VALUE_FILL: Final[str] = "F2F2F2"
COLOR_BADGE_FILL: Final[str] = "FFF2CC"

# ─── Fonts ────────────────────────────────────────────────────────────────

FONT_PRIMARY: Final[str] = "Tahoma"
FONT_NUMBERS: Final[str] = "Calibri"

# ─── Protection ───────────────────────────────────────────────────────────

#: Environment variable that supplies the worksheet protection password.
PASSWORD_ENV_VAR: Final[str] = "WORKBOOK_PASSWORD"

#: Sheet protection is a convenience guard, never a security boundary.
PROTECTION_DISCLAIMER: Final[str] = (
    "Worksheet protection prevents accidental modification. "
    "It is not a security boundary."
)
