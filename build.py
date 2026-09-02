#!/usr/bin/env python3
"""Concrete Lab Companion — build orchestrator (v4).

This file used to be a 1700-line monolith mixing standards metadata,
scientific formulas, Excel layout, protection, hashing and CLI.  In the
v4 architecture it is intentionally tiny: it only bootstraps the
package and delegates to the CLI.

Layers (all under ``src/concrete_lab``):

* ``domain``    — value objects + the calculation engine;
* ``standards`` — typed rulesets (C39-26, C805-25, ISIRI 302, …);
* ``qa``        — real QA: pytest runner, golden suite, structural;
* ``render``    — display-only Excel renderer;
* ``report``    — build manifest with traceability.

Usage:
    python build.py                 # build + embedded QA + manifest
    python build.py --validate      # run real QA (golden + pytest)
    python build.py --no-protect    # build without sheet protection
"""

import sys
from pathlib import Path

_SRC = Path(__file__).resolve().parent / "src"
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

from concrete_lab.cli import main  # noqa: E402

if __name__ == "__main__":
    sys.exit(main())
