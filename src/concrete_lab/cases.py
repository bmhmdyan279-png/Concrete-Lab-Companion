"""Loading laboratory data from a user-supplied case file.

The v4 engine always computed *something* — but until 4.1 it could only
compute its own built-in demo dataset, which made the published workbook
a sample report rather than a tool.  This module is the real user-input
door: a small JSON file describing the measurements a technician actually
took, validated before it ever reaches the domain engine.

File format::

    {
      "project": {"name": "پل سرخه‌رود", "operator": "…"},   // optional
      "cases": [
        {"test_id": "4-1",
         "inputs": {"diameter_mm": 150, "length_mm": 300, "load_kn": 715}},
        {"test_id": "1-1",
         "inputs": {"retained": {"9.5": 25, "4.75": 95}, "total_mass_g": 1000,
                    "pan_g": 10}}
      ]
    }

Rules enforced here (before any calculation):

* ``cases`` must be a non-empty list of objects;
* every case needs a string ``test_id`` and an object ``inputs``;
* a ``test_id`` may appear only once — each computed test renders to its
  own worksheet, so duplicates would collide on the sheet name;
* every ``test_id`` must be implemented by the engine (pending tests are
  reported with the list of ids that *are* available, never silently
  dropped and never faked).
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, List, Mapping, Optional, Tuple

from concrete_lab.domain.engine import CALCULATORS, Inputs

#: Key holding the list of cases in a case file.
CASES_KEY: str = "cases"

#: Key holding optional project metadata in a case file.
PROJECT_KEY: str = "project"


class CaseFileError(ValueError):
    """Raised when a user-supplied case file is missing or malformed."""


def implemented_ids() -> Tuple[str, ...]:
    """Return the test ids the engine can currently compute (sorted)."""
    return tuple(sorted(CALCULATORS))


def load_case_file(path: Path) -> Tuple[Tuple[str, Inputs], ...]:
    """Read and validate a case file.

    Args:
        path: Location of the JSON case file.

    Returns:
        ``(test_id, inputs)`` pairs in file order.

    Raises:
        CaseFileError: If the file cannot be read, is not valid JSON, or
            violates any structural rule documented in this module.
    """
    path = Path(path)
    if not path.is_file():
        raise CaseFileError(f"case file not found: {path}")

    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise CaseFileError(f"{path} is not valid JSON: {exc}") from exc
    except OSError as exc:
        raise CaseFileError(f"{path} could not be read: {exc}") from exc

    if not isinstance(raw, dict):
        raise CaseFileError(f"{path} must contain a JSON object at the top level")

    project = raw.get(PROJECT_KEY)
    if project is not None and not isinstance(project, Mapping):
        raise CaseFileError(f"{PROJECT_KEY!r} must be an object when present")

    cases = raw.get(CASES_KEY)
    if not isinstance(cases, list):
        raise CaseFileError(
            f"{path} must define a {CASES_KEY!r} list; available test ids: "
            f"{', '.join(implemented_ids())}"
        )
    if not cases:
        raise CaseFileError(f"{CASES_KEY!r} is empty — nothing to compute")

    available = set(CALCULATORS)
    loaded: List[Tuple[str, Inputs]] = []
    seen: Dict[str, int] = {}
    problems: List[str] = []

    for index, case in enumerate(cases):
        where = f"{CASES_KEY}[{index}]"
        if not isinstance(case, Mapping):
            problems.append(f"{where} must be an object")
            continue

        test_id = case.get("test_id")
        if not isinstance(test_id, str) or not test_id.strip():
            problems.append(f"{where} needs a non-empty string 'test_id'")
            continue
        test_id = test_id.strip()

        inputs = case.get("inputs")
        if not isinstance(inputs, Mapping):
            problems.append(f"{where} ({test_id}) needs an 'inputs' object")
            continue

        if test_id not in available:
            problems.append(
                f"{where}: test {test_id!r} is not implemented by the engine; "
                f"available: {', '.join(sorted(available))}"
            )
            continue
        if test_id in seen:
            problems.append(
                f"{where}: test {test_id!r} appears more than once (first at "
                f"{CASES_KEY}[{seen[test_id]}]); each test renders to its own worksheet"
            )
            continue

        seen[test_id] = index
        loaded.append((test_id, dict(inputs)))

    if problems:
        raise CaseFileError(
            f"{path} has {len(problems)} problem(s):\n  - " + "\n  - ".join(problems)
        )
    return tuple(loaded)


def project_metadata(path: Path) -> Optional[Dict[str, Any]]:
    """Return the optional ``project`` block of a case file, if present.

    Kept separate from :func:`load_case_file` so that a build can record
    *where the data came from* even when only the cases are needed.
    """
    try:
        raw = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    if not isinstance(raw, dict):
        return None
    project = raw.get(PROJECT_KEY)
    return dict(project) if isinstance(project, Mapping) else None
