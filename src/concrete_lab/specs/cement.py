"""Chapter 2 — Cement & fresh-density tests (آزمایش‌های سیمان و بتن تازه): 2-1 … 2-4."""

from concrete_lab.specs.base import TestSpec, register_test

register_test(TestSpec(
    id="2-1", title="چگالی بتن تازه", standard_code="C138",
    sheet_name="10_آزمایش_2-1", implemented=True, tab_color="2E7D32",
    notes="Density = (M_total − M_apparatus) / V.",
))
register_test(TestSpec(
    id="2-2", title="روانایی نرمال (تاریخی)", standard_code="C187",
    sheet_name="11_آزمایش_2-2", tab_color="2E7D32",
    notes="Governing standard is withdrawn; kept for historical reports.",
))
register_test(TestSpec(
    id="2-3", title="زمان گیرش", standard_code="C191",
    sheet_name="12_آزمایش_2-3", tab_color="2E7D32",
))
register_test(TestSpec(
    id="2-4", title="مقاومت ملات", standard_code="EN196-1",
    sheet_name="13_آزمایش_2-4", tab_color="2E7D32",
))
