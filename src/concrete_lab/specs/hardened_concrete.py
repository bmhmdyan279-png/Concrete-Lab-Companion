"""Chapter 4 — Hardened concrete tests (آزمایش‌های بتن سخت‌شده): 4-1 … 4-5."""

from concrete_lab.specs.base import Scope, TestSpec, register_test

register_test(TestSpec(
    id="4-1", title="مقاومت فشاری", standard_code="C39",
    sheet_name="17_آزمایش_4-1", implemented=True, tab_color="B71C1C",
    notes="Cylinders only; interpolated L/D correction (C39-26 ruleset). "
          "Cube specimens are rejected, never silently converted.",
))
register_test(TestSpec(
    id="4-2", title="مقاومت کششی برزیلی", standard_code="C496",
    sheet_name="18_آزمایش_4-2", implemented=True, tab_color="B71C1C",
    notes="Splitting tensile strength: 2P / (π·d·L).",
))
register_test(TestSpec(
    id="4-3", title="مقاومت خمشی", standard_code="C78",
    sheet_name="19_آزمایش_4-3", tab_color="B71C1C",
))
register_test(TestSpec(
    id="4-4", title="فراصوت (UPV)", standard_code="C597",
    sheet_name="20_آزمایش_4-4", tab_color="B71C1C",
))
register_test(TestSpec(
    id="4-5", title="چکش اشمیت", standard_code="C805",
    sheet_name="21_آزمایش_4-5", implemented=True, tab_color="B71C1C",
    scope=Scope.ESTIMATION,
    notes="Rebound number only; strength needs a project-specific "
          "correlation. NOT FOR ACCEPTANCE (C805-25 ruleset).",
))
