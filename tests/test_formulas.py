"""
Golden Tests for Excel Formulas
Phase 3 of Obsessive QA - Tests extracted from validation/golden_cases
"""
import pytest
import json
import glob
from pathlib import Path

GOLDEN_DIR = Path("validation/golden_cases")

def load_golden_cases():
    """Load all golden test cases"""
    cases = []
    for f in sorted(GOLDEN_DIR.glob("*.json")):
        with open(f, "r", encoding="utf-8") as fp:
            data = json.load(fp)
            data["file"] = f.name
            cases.append(data)
    return cases

@pytest.mark.golden
@pytest.mark.parametrize("case", load_golden_cases(), ids=lambda c: c["file"])
def test_golden_case_structure(case):
    """Verify each golden case has required fields"""
    assert "test_name" in case, "Missing test_name"
    assert "standard" in case, "Missing standard"
    assert "inputs" in case, "Missing inputs"
    assert "expected" in case, "Missing expected"
    assert "tolerance" in case, "Missing tolerance"

@pytest.mark.golden
def test_golden_cases_count():
    """Verify we have sufficient test coverage"""
    cases = load_golden_cases()
    assert len(cases) >= 20, f"Only {len(cases)} golden cases, need at least 20"
