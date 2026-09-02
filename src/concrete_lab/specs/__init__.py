"""Test specification catalogue.

Importing this package registers every chapter's test specs into
:mod:`concrete_lab.specs.base` — nothing is registered until imported.
"""

import importlib

# Chapter modules register their specs as a side effect of import.
for _chapter in ("aggregates", "cement", "fresh_concrete", "hardened_concrete"):
    importlib.import_module(f"{__name__}.{_chapter}")

from concrete_lab.specs.base import TEST_REGISTRY, TestSpec, all_tests, implemented_tests

__all__ = ["TEST_REGISTRY", "TestSpec", "all_tests", "implemented_tests"]
