"""Chapter 3 — Fresh concrete tests (آزمایش‌های بتن تازه): 3-1 … 3-3."""

from concrete_lab.specs.base import TestSpec, register_test

register_test(TestSpec(
    id="3-1", title="اسلامپ", standard_code="C143",
    sheet_name="14_آزمایش_3-1", implemented=True, tab_color="1565C0",
    notes="Slump = 300 mm cone height − measured specimen height.",
))
register_test(TestSpec(
    id="3-2", title="آب‌اندازی", standard_code="C232",
    sheet_name="15_آزمایش_3-2", tab_color="1565C0",
))
register_test(TestSpec(
    id="3-3", title="وزن واحد بتن", standard_code="C29",
    sheet_name="16_آزمایش_3-3", tab_color="1565C0",
))
