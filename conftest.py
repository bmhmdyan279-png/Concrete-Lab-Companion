"""Root pytest configuration.

The package lives in a ``src`` layout that is not installed as a
distribution during tests, so the ``src`` directory is prepended to
``sys.path`` here to make ``import concrete_lab`` work.
"""

import sys
from pathlib import Path

_SRC = Path(__file__).resolve().parent / "src"
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))
