"""Domain layer: immutable value objects shared by the whole package."""

from concrete_lab.domain.quantities import Quantity
from concrete_lab.domain.results import ResultSpec
from concrete_lab.domain.statuses import ValidationStatus

__all__ = ["Quantity", "ResultSpec", "ValidationStatus"]
