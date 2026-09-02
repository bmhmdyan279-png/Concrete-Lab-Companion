"""QA layer: real verification, not registry inspection.

Three tiers, addressing review feedback that the legacy ``--validate``
flag did not actually validate anything:

* ``engine``     — runs the real pytest suite and reports PASS/FAIL/WARN;
* ``golden``     — executes golden cases in-process against the engine;
* ``structural`` — verifies built artifacts (hash, manifest, sheets,
  the renderer's no-formula guarantee).
"""
