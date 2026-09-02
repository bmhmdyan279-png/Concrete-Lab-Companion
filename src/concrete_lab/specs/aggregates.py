"""Chapter 1 — Aggregate tests (آزمایش‌های سنگدانه): 1-1 … 1-8."""

from concrete_lab.specs.base import TestSpec, register_test

register_test(TestSpec(
    id="1-1", title="دانه‌بندی سنگدانه", standard_code="ISIRI302",
    sheet_name="02_آزمایش_1-1", implemented=True, tab_color="FF6F00",
    notes="Grading envelope check via the ISIRI 302 ruleset (replaces the "
          "legacy FM-only view; legacy FM golden case encoded a buggy formula).",
))
register_test(TestSpec(
    id="1-2", title="رطوبت سنگدانه", standard_code="C566",
    sheet_name="03_آزمایش_1-2", implemented=True, tab_color="FF6F00",
    notes="Dry-basis moisture: (W_wet − W_dry) / W_dry × 100.",
))
register_test(TestSpec(
    id="1-3", title="وزن مخصوص درشت‌دانه", standard_code="C127",
    sheet_name="04_آزمایش_1-3", tab_color="FF6F00",
))
register_test(TestSpec(
    id="1-4", title="وزن مخصوص ریزدانه", standard_code="C128",
    sheet_name="05_آزمایش_1-4", tab_color="FF6F00",
))
register_test(TestSpec(
    id="1-5", title="وزن واحد سنگدانه", standard_code="C29",
    sheet_name="06_آزمایش_1-5", tab_color="FF6F00",
))
register_test(TestSpec(
    id="1-6", title="معادل ماسه‌ای", standard_code="D2419",
    sheet_name="07_آزمایش_1-6", tab_color="FF6F00",
))
register_test(TestSpec(
    id="1-7", title="شاخص‌های شکلی", standard_code="D4791",
    sheet_name="08_آزمایش_1-7", tab_color="FF6F00",
))
register_test(TestSpec(
    id="1-8", title="جذب آب", standard_code="C127",
    sheet_name="09_آزمایش_1-8", tab_color="FF6F00",
))
