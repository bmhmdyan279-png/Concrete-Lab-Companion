"""Assemble the pure-data workbook model from domain results.

This module decides *what to show* (labels, order, badges) but never
recomputes science: every number it places comes from a
:class:`TestResult` produced by the domain engine or from a QA report.
"""

from __future__ import annotations

from typing import Dict, Iterable, List, Mapping, Optional, Tuple

from concrete_lab import __version__, constants
from concrete_lab.config import AppConfig
from concrete_lab.domain.engine import TestResult
from concrete_lab.domain.quantities import Quantity
from concrete_lab.domain.statuses import ValidationStatus
from concrete_lab.qa.engine import QAReport
from concrete_lab.render.excel.model import CellModel, RowModel, SheetModel, WorkbookModel
from concrete_lab.specs.base import TEST_REGISTRY, all_tests, implemented_tests
from concrete_lab.standards import registry as standards_registry
from concrete_lab.standards.base import Scope
from concrete_lab.utils import utc_now_iso

#: Worksheet layout: default column widths (column → width).
DEFAULT_WIDTHS: Tuple[Tuple[int, float], ...] = ((1, 36.0), (2, 18.0), (3, 12.0), (4, 14.0), (5, 48.0))

#: Style name used to draw each validation status.
_STATUS_STYLE: Dict[ValidationStatus, str] = {
    ValidationStatus.PASS: "pass",
    ValidationStatus.WARN: "warn",
    ValidationStatus.FAIL: "fail",
    ValidationStatus.ERROR: "fail",
    ValidationStatus.SKIPPED: "pending",
    ValidationStatus.PENDING: "pending",
}


def _row(*cells: CellModel) -> RowModel:
    """Shorthand row constructor."""
    return RowModel(cells=tuple(cells))


def _quantity_cells(quantity: Optional[Quantity]) -> Tuple[CellModel, CellModel]:
    """Value + unit cells for a quantity (with a sane number format)."""
    if quantity is None:
        return CellModel("—", "value"), CellModel("", "value")
    number_format = "0.00" if isinstance(quantity.value, float) else None
    return CellModel(quantity.value, "value", number_format), CellModel(quantity.unit, "value")


def _status_cell(status: ValidationStatus) -> CellModel:
    """Renderer-side status display: canonical enum in, symbol out."""
    return CellModel(status.symbol, _STATUS_STYLE[status])


# ─── Individual sheets ────────────────────────────────────────────────────


def _guide_sheet(config: AppConfig) -> SheetModel:
    """00 — guide: scope legend and usage rules."""
    rows = [
        _row(CellModel(config.name, "title")),
        _row(CellModel(f"نسخه {__version__} — موتور محاسباتی جدا از نمایش اکسل", "subtitle")),
        _row(CellModel("", "label")),
        _row(CellModel("محیط‌های محاسبه", "header"), CellModel("شرح", "header")),
        _row(CellModel("STANDARD", "value"), CellModel("نتیجه نرماتیو؛ مطابق قواعد استاندارد", "label")),
        _row(CellModel("ESTIMATION", "badge"),
             CellModel("تخمین؛ هرگز مبنای پذیرش یا رد بتن قرار ندهید", "label")),
        _row(CellModel("", "label")),
        _row(CellModel("معماری نسخه ۴", "header"), CellModel("توضیح", "header")),
        _row(CellModel("Domain Engine", "value"), CellModel("محاسبه علم آزمایش در پایتون (قابل تست و راستی‌آزمایی)", "label")),
        _row(CellModel("Excel Renderer", "value"), CellModel("فقط نمایش نتایج؛ هیچ فرمول محاسباتی در فایل نیست", "label")),
        _row(CellModel("QA Engine", "value"), CellModel("اجرای واقعی تست‌ها و موارد طلایی هنگام ساخت و اعتبارسنجی", "label")),
        _row(CellModel("", "label")),
        _row(CellModel(constants.PROTECTION_DISCLAIMER, "note")),
        _row(CellModel(f"ساخت: {utc_now_iso()}", "note")),
    ]
    return SheetModel(title="00_راهنما", rows=tuple(rows))


def _info_sheet(config: AppConfig) -> SheetModel:
    """01 — test/project information from configuration."""
    rows = [
        _row(CellModel("اطلاعات محصول و آزمون", "title")),
        _row(CellModel("", "label")),
        _row(CellModel("نام محصول", "header"), CellModel("مقدار", "header")),
        _row(CellModel("نام", "label"), CellModel(config.name, "value")),
        _row(CellModel("نسخه", "label"), CellModel(__version__, "value")),
        _row(CellModel("زبان", "label"), CellModel(config.language, "value")),
        _row(CellModel("مرجع اصلی استانداردها", "label"), CellModel(config.standards_primary, "value")),
        _row(CellModel("مرجع ثانویه", "label"), CellModel(config.standards_secondary, "value")),
        _row(CellModel("تعداد کل آزمایش‌ها", "label"), CellModel(len(all_tests()), "value")),
        _row(CellModel("آزمایش‌های پیاده‌سازی‌شده", "label"), CellModel(len(implemented_tests()), "value")),
        _row(CellModel("تاریخ ساخت", "label"), CellModel(utc_now_iso(), "value")),
    ]
    return SheetModel(title="01_اطلاعات_آزمون", rows=tuple(rows))


def _test_sheet(result: TestResult) -> SheetModel:
    """One worksheet per computed test (values only, no formulas)."""
    spec = TEST_REGISTRY[result.test_id]
    standard = result.standard
    rows: List[RowModel] = [
        _row(CellModel(f"آزمایش {result.test_id}: {result.title}", "title")),
        _row(CellModel(f"استاندارد: {standard.citation()}", "subtitle")),
    ]
    if result.scope is Scope.ESTIMATION:
        rows.append(_row(CellModel("ESTIMATED — NOT FOR ACCEPTANCE", "badge"),
                         CellModel("تخمینی — برای پذیرش/رد قابل استناد نیست", "badge")))
    ruleset = standards_registry.ruleset_version(standard.code)
    if ruleset:
        rows.append(_row(CellModel(f"Ruleset: {ruleset}", "note")))
    rows.append(_row(CellModel("", "label")))
    rows.append(_row(
        CellModel("شرح", "header"), CellModel("مقدار", "header"),
        CellModel("واحد", "header"), CellModel("وضعیت", "header"),
        CellModel("توضیح", "header"),
    ))
    for value in result.values:
        quantity_cell, unit_cell = _quantity_cells(value.quantity)
        note = value.note
        rows.append(_row(
            CellModel(value.label, "label"), quantity_cell, unit_cell,
            _status_cell(value.status), CellModel(note, "note"),
        ))
    rows.append(_row(CellModel("", "label")))
    rows.append(_row(
        CellModel("وضعیت کلی آزمایش", "label"), CellModel("", "value"), CellModel("", "value"),
        _status_cell(result.status), CellModel("", "note"),
    ))
    for warning in result.warnings:
        rows.append(_row(CellModel("هشدار", "warn"), CellModel(warning, "note")))
    if result.disclaimer:
        rows.append(_row(CellModel("سلب مسئولیت", "badge"), CellModel(result.disclaimer, "note")))
    return SheetModel(title=spec.sheet_name, rows=tuple(rows), tab_color=spec.tab_color)


def _primary_value(result: TestResult) -> Tuple[str, str]:
    """Pick the headline value for report/dashboard rows."""
    candidates = ("reported_strength", "tensile_strength", "slump", "density",
                  "moisture_percent", "retained_mean", "grading_overall")
    for key in candidates:
        match = next((v for v in result.values if v.key == key), None)
        if match is not None and match.quantity is not None:
            return f"{match.quantity.value:.2f} {match.quantity.unit}", match.quantity.unit
    overall = result.values[-1] if result.values else None
    if overall is not None:
        return overall.status.symbol, ""
    return "—", ""


def _report_sheet(results: Tuple[TestResult, ...]) -> SheetModel:
    """22 — summary report of all computed tests."""
    rows: List[RowModel] = [
        _row(CellModel("گزارش نتایج", "title")),
        _row(CellModel("", "label")),
        _row(CellModel("شناسه", "header"), CellModel("آزمایش", "header"),
             CellModel("استاندارد", "header"), CellModel("نتیجه اصلی", "header"),
             CellModel("وضعیت", "header")),
    ]
    for result in results:
        value_text, _unit = _primary_value(result)
        rows.append(_row(
            CellModel(result.test_id, "value"), CellModel(result.title, "label"),
            CellModel(result.standard.citation(), "note"), CellModel(value_text, "value"),
            _status_cell(result.status),
        ))
    pending = [spec for spec in all_tests() if not spec.implemented]
    for spec in pending:
        rows.append(_row(
            CellModel(spec.id, "value"), CellModel(spec.title, "label"),
            CellModel("در انتظار پیاده‌سازی موتور", "note"), CellModel("—", "pending"),
            _status_cell(ValidationStatus.SKIPPED),
        ))
    return SheetModel(title="22_گزارش", rows=tuple(rows))


def _dashboard_sheet(results: Tuple[TestResult, ...], qa_reports: Mapping[str, QAReport]) -> SheetModel:
    """23 — lab state dashboard (counters, quality, key results)."""
    passed = sum(r.status is ValidationStatus.PASS for r in results)
    warned = sum(r.status is ValidationStatus.WARN for r in results)
    failed = sum(r.status in (ValidationStatus.FAIL, ValidationStatus.ERROR) for r in results)

    rows: List[RowModel] = [
        _row(CellModel("داشبورد آزمایشگاه", "title")),
        _row(CellModel("", "label")),
        _row(CellModel("نمای کلی", "header"), CellModel("تعداد", "header")),
        _row(CellModel("کل آزمایش‌های محصول", "label"), CellModel(len(all_tests()), "value")),
        _row(CellModel("پیاده‌سازی‌شده در موتور", "label"), CellModel(len(implemented_tests()), "value")),
        _row(CellModel("در انتظار", "label"), CellModel(len(all_tests()) - len(implemented_tests()), "value")),
        _row(CellModel("", "label")),
        _row(CellModel("کیفیت نتایج این ساخت", "header"), CellModel("تعداد", "header")),
        _row(CellModel("PASS", "pass"), CellModel(passed, "value")),
        _row(CellModel("WARN", "warn"), CellModel(warned, "value")),
        _row(CellModel("FAIL", "fail"), CellModel(failed, "value")),
        _row(CellModel("", "label")),
        _row(CellModel("سطح‌های QA", "header"), CellModel("وضعیت", "header"),
             CellModel("پاس", "header"), CellModel("شکست", "header")),
    ]
    for name, report in qa_reports.items():
        rows.append(_row(
            CellModel(name, "label"), _status_cell(report.status),
            CellModel(report.passed, "value"), CellModel(report.failed, "value"),
        ))
    rows.append(_row(CellModel("", "label")))
    rows.append(_row(CellModel("نتایج کلیدی", "header"), CellModel("مقدار", "header"),
                     CellModel("وضعیت", "header")))
    for result in results:
        value_text, _unit = _primary_value(result)
        rows.append(_row(CellModel(result.title, "label"), CellModel(value_text, "value"),
                         _status_cell(result.status)))
    return SheetModel(title="23_داشبورد", rows=tuple(rows))


def _qa_sheet(qa_reports: Mapping[str, QAReport]) -> SheetModel:
    """24 — QA outcomes actually executed for this build."""
    rows: List[RowModel] = [
        _row(CellModel("نتایج QA این ساخت", "title")),
        _row(CellModel("این اعداد حاصل اجرای واقعی تست‌هاست، نه فرمول‌های اکسل.", "subtitle")),
        _row(CellModel("", "label")),
        _row(CellModel("سطح", "header"), CellModel("پاس", "header"), CellModel("شکست", "header"),
             CellModel("هشدار", "header"), CellModel("وضعیت", "header")),
    ]
    for name, report in qa_reports.items():
        rows.append(_row(
            CellModel(name, "label"), CellModel(report.passed, "value"),
            CellModel(report.failed, "value"), CellModel(report.warnings, "value"),
            _status_cell(report.status),
        ))
    failures = [failure for report in qa_reports.values() for failure in report.failures]
    if failures:
        rows.append(_row(CellModel("", "label")))
        rows.append(_row(CellModel("موارد شکست", "header")))
        for failure in failures:
            rows.append(_row(CellModel(failure, "fail")))
    return SheetModel(title="24_QA_Test", rows=tuple(rows))


def _warnings_sheet(results: Tuple[TestResult, ...]) -> SheetModel:
    """25 — every warning and disclaimer collected from results."""
    rows: List[RowModel] = [
        _row(CellModel("خطاها و هشدارها", "title")),
        _row(CellModel("", "label")),
        _row(CellModel("آزمایش", "header"), CellModel("مورد", "header")),
    ]
    for result in results:
        for warning in result.warnings:
            rows.append(_row(CellModel(result.test_id, "value"), CellModel(warning, "warn")))
        if result.disclaimer:
            rows.append(_row(CellModel(result.test_id, "value"),
                             CellModel(result.disclaimer, "note")))
    if len(rows) == 3:
        rows.append(_row(CellModel("—", "note"), CellModel("بدون هشدار", "note")))
    return SheetModel(title="25_خطاها_هشدارها", rows=tuple(rows))


def _standards_sheet() -> SheetModel:
    """_Standards — the audited standards catalogue with ruleset tags."""
    rows: List[RowModel] = [
        _row(CellModel("فهرست استانداردها (حسابرسی ۲۰۲۶)", "title")),
        _row(CellModel("", "label")),
        _row(CellModel("کد", "header"), CellModel("نام", "header"), CellModel("ویرایش", "header"),
             CellModel("وضعیت", "header"), CellModel("Ruleset", "header")),
    ]
    for spec in standards_registry.all_standards():
        ruleset = standards_registry.ruleset_version(spec.code) or "—"
        rows.append(_row(
            CellModel(spec.code, "value"), CellModel(spec.citation(), "label"),
            CellModel(spec.edition, "value"), CellModel(spec.status, "value"),
            CellModel(ruleset, "note"),
        ))
    return SheetModel(title="_Standards", rows=tuple(rows), rtl=False)


# ─── Entry point ──────────────────────────────────────────────────────────


def assemble_workbook_model(
    results: Iterable[TestResult],
    config: AppConfig,
    qa_reports: Mapping[str, QAReport],
) -> WorkbookModel:
    """Build the complete display model for one build.

    Args:
        results: Domain engine results to display.
        config: Validated application configuration.
        qa_reports: QA tiers to display on the QA sheet/dashboard.

    Returns:
        A :class:`WorkbookModel` ready for the renderer.
    """
    results = tuple(results)
    sheets: List[SheetModel] = [
        _guide_sheet(config),
        _info_sheet(config),
    ]
    sheets.extend(_test_sheet(result) for result in results)
    sheets.append(_report_sheet(results))
    sheets.append(_dashboard_sheet(results, qa_reports))
    sheets.append(_qa_sheet(qa_reports))
    sheets.append(_warnings_sheet(results))
    sheets.append(_standards_sheet())
    return WorkbookModel(sheets=tuple(sheets))
