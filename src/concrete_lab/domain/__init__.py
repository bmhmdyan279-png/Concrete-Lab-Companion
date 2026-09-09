"""Domain layer: immutable value objects shared by the whole package.

Only three concepts live here:

* :class:`Quantity` — a magnitude bound to its unit;
* :class:`ValidationStatus` — the canonical six-state outcome enum;
* the calculation engine (:mod:`concrete_lab.domain.engine`) with its
  :class:`~concrete_lab.domain.engine.ResultValue` /
  :class:`~concrete_lab.domain.engine.TestResult` records.

There is deliberately **no** result *specification* type carrying an
Excel formula string: since v4 the workbook is display-only and the
science lives in Python, so a formula-bearing domain object would
contradict the architecture (the former ``ResultSpec`` was unused dead
code and was removed in 4.1.0).
"""

from concrete_lab.domain.quantities import Quantity
from concrete_lab.domain.statuses import ValidationStatus

__all__ = ["Quantity", "ValidationStatus"]
