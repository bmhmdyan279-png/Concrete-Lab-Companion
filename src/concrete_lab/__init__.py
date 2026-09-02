"""Concrete Lab Companion — calculation & QA engine (v4 architecture).

Version 4 separates the layers that the legacy monolith mixed together:

* ``domain``    — value objects and the calculation engine (the science);
* ``standards`` — typed rulesets per laboratory standard;
* ``qa``        — real QA: pytest runner, golden suite, structural checks;
* ``render``    — display-only Excel renderer (never computes);
* ``report``    — build manifest with traceability metadata.
"""

__version__ = "4.0.0"
