"""Registry of laboratory standards with audited editions.

The legacy ``build.py`` kept standard metadata in a bare dictionary
with stale editions.  This registry:

* stores :class:`StandardSpec` objects (validated, immutable);
* records the **ruleset module** that implements each standard, where
  one exists (the standard/science boundary demanded by review);
* reflects the 2026 standards audit: ``C39`` is edition **2026** (the
  legacy dict said 2022) and ``C805`` is the active **-25** revision.
"""

from __future__ import annotations

from typing import Dict, Optional, Tuple

from concrete_lab.standards.astm import c39, c805
from concrete_lab.standards.base import StandardSpec
from concrete_lab.standards.isiri import isiri_302

#: All standards referenced anywhere in the product, keyed by short code.
#: Editions marked "(audited 2026)" were verified against the active
#: ASTM/ISIRI catalogue during the v4 standards audit.
_REGISTRY: Dict[str, StandardSpec] = {}

#: Standards with an executable ruleset module, keyed by short code.
RULESETS: Dict[str, object] = {
    "C39": c39,
    "C805": c805,
    "ISIRI302": isiri_302,
}


def register(spec: StandardSpec) -> StandardSpec:
    """Register a standard; duplicate codes are a programming error.

    Args:
        spec: The standard metadata to register.

    Returns:
        The same spec (so the call can be used declaratively).

    Raises:
        ValueError: If the code was already registered.
    """
    if spec.code in _REGISTRY:
        raise ValueError(f"standard {spec.code!r} is already registered")
    _REGISTRY[spec.code] = spec
    return spec


def get(code: str) -> StandardSpec:
    """Return the registered spec for ``code``.

    Raises:
        KeyError: If the code is unknown.
    """
    try:
        return _REGISTRY[code]
    except KeyError:
        raise KeyError(f"unknown standard {code!r}") from None


def ruleset_version(code: str) -> Optional[str]:
    """Return the ruleset version tag for a standard, if implemented."""
    module = RULESETS.get(code)
    return getattr(module, "RULESET_VERSION", None)


def all_standards() -> Tuple[StandardSpec, ...]:
    """Return all registered standards in registration order."""
    return tuple(_REGISTRY.values())


# ─── Registration (audited 2026) ──────────────────────────────────────────

register(StandardSpec("C29", "ASTM C29/C29M", "Bulk Density of Aggregates", "2023"))
register(StandardSpec("C39", "ASTM C39/C39M", "Compressive Strength of Cylindrical Concrete Specimens", "2026"))
register(StandardSpec("C78", "ASTM C78/C78M", "Flexural Strength (Third-Point)", "2022"))
register(StandardSpec("C127", "ASTM C127", "Density of Coarse Aggregate", "2023"))
register(StandardSpec("C128", "ASTM C128", "Density of Fine Aggregate", "2022"))
register(StandardSpec("C136", "ASTM C136/C136M", "Sieve Analysis", "2024"))
register(StandardSpec("C138", "ASTM C138/C138M", "Density of Fresh Concrete", "2024"))
register(StandardSpec("C143", "ASTM C143/C143M", "Slump of Concrete", "2024"))
register(StandardSpec("C187", "ASTM C187", "Normal Consistency (historical)", "2016", status="withdrawn"))
register(StandardSpec("C191", "ASTM C191", "Setting Time by Vicat", "2024"))
register(StandardSpec("C232", "ASTM C232/C232M", "Bleeding of Concrete", "2023"))
register(StandardSpec("C293", "ASTM C293/C293M", "Flexural Strength (Center-Point)", "2023"))
register(StandardSpec("C496", "ASTM C496/C496M", "Splitting Tensile Strength", "2023"))
register(StandardSpec("C566", "ASTM C566", "Moisture Content of Aggregates", "2023"))
register(StandardSpec("C597", "ASTM C597", "Pulse Velocity Through Concrete", "2023"))
register(StandardSpec("C805", "ASTM C805/C805M", "Assessing the Rebound Number of Hardened Concrete", "2025"))
register(StandardSpec("D2419", "ASTM D2419", "Sand Equivalent", "2022"))
register(StandardSpec("D4791", "ASTM D4791", "Flat & Elongated Particles", "2023"))
register(StandardSpec("EN196-1", "EN 196-1", "Mortar Strength", "2023"))
register(StandardSpec("ISIRI302", "ISIRI 302", "ویژگی‌های سنگدانه‌های بتن (Characteristics of Concrete Aggregates)", "1394"))
