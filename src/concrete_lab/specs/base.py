"""Data-driven test specifications (the product catalogue).

Each laboratory test the product offers is declared once, next to its
chapter, as a :class:`TestSpec` (review feedback: decentralised
registration instead of one 500-line ``define_tests()`` function).

A spec describes *what the test is* — identity, standard, scope,
implementation state.  The actual science lives in the standards
rulesets and the domain engine; the Excel renderer displays results.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, Tuple

from concrete_lab.standards.base import Scope


@dataclass(frozen=True)
class TestSpec:
    """Declarative specification of one laboratory test.

    Attributes:
        id: Stable semantic identifier (e.g. ``"4-1"``).
        title: Human-readable title (Persian in this product).
        standard_code: Registry key of the governing standard.
        sheet_name: Worksheet title in the rendered workbook.
        scope: How results may be used (see :class:`Scope`); renderer
            shows an ``ESTIMATED — NOT FOR ACCEPTANCE`` badge for
            estimation scope.
        implemented: ``True`` when the domain engine can compute this
            test; pending tests are listed but never faked.
        tab_color: Worksheet tab colour (hex, no ``#``).
        notes: Free-form engineering notes shown on the guide sheet.
    """

    id: str
    title: str
    standard_code: str
    sheet_name: str
    scope: Scope = Scope.STANDARD
    implemented: bool = False
    tab_color: str = "1F4E79"
    notes: str = ""


#: Global catalogue, keyed by test id.
TEST_REGISTRY: Dict[str, TestSpec] = {}


def register_test(spec: TestSpec) -> TestSpec:
    """Register a test spec; duplicates are a programming error.

    Args:
        spec: The specification to register.

    Returns:
        The same spec.

    Raises:
        ValueError: If the id was already registered.
    """
    if spec.id in TEST_REGISTRY:
        raise ValueError(f"test {spec.id!r} is already registered")
    TEST_REGISTRY[spec.id] = spec
    return spec


def all_tests() -> Tuple[TestSpec, ...]:
    """Return every registered test in sheet order."""
    return tuple(sorted(TEST_REGISTRY.values(), key=lambda s: s.sheet_name))


def implemented_tests() -> Tuple[TestSpec, ...]:
    """Return only the tests the domain engine can compute."""
    return tuple(t for t in all_tests() if t.implemented)
