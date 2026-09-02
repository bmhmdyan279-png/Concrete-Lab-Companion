"""Worksheet protection manager.

Applies openpyxl sheet protection consistently and makes the security
posture explicit: protection prevents accidental edits only — it is
documented (and reported in the manifest) as *not* a security boundary.
"""

from __future__ import annotations

from logging import getLogger
from typing import Optional

from openpyxl import Workbook

from concrete_lab.constants import INTERNAL_SHEET_PREFIX, PROTECTION_DISCLAIMER

logger = getLogger(__name__)


class ProtectionManager:
    """Applies or skips worksheet protection, logging the outcome."""

    @staticmethod
    def apply(workbook: Workbook, password: Optional[str]) -> bool:
        """Protect all user-facing sheets when a password is available.

        Args:
            workbook: The rendered workbook.
            password: Protection password; ``None`` disables protection.

        Returns:
            ``True`` when protection was applied, ``False`` otherwise.
        """
        if not password:
            logger.warning("Protection: DISABLED (no password configured)")
            logger.warning(PROTECTION_DISCLAIMER)
            return False

        for worksheet in workbook.worksheets:
            if worksheet.title.startswith(INTERNAL_SHEET_PREFIX):
                continue
            protection = worksheet.protection
            protection.sheet = True
            protection.password = password
            protection.selectLockedCells = False
            protection.selectUnlockedCells = True
            protection.formatCells = False
            protection.formatColumns = False
            protection.formatRows = False
            protection.insertColumns = False
            protection.insertRows = False
            protection.deleteColumns = False
            protection.deleteRows = False

        workbook.security.lockStructure = True
        workbook.security.workbookPassword = password
        logger.info("Protection: ENABLED (%d sheets)", len(workbook.worksheets))
        logger.info(PROTECTION_DISCLAIMER)
        return True
