#!/usr/bin/env python3
"""
Concrete Lab Companion — Canonical Build Script
================================================
تنها نقطه ساخت Workbook.
build_workbook.py و build_project.py حذف شده‌اند.

Usage:
    python build.py
"""

import yaml
import hashlib
import os
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from datetime import datetime
from pathlib import Path

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("ERROR: openpyxl not installed. Run: pip install -r requirements.txt")
    sys.exit(1)


def load_config() -> dict:
    """Load central configuration from config.yaml."""
    config_path = Path(__file__).parent / "config.yaml"
    if not config_path.exists():
        print("ERROR: config.yaml not found")
        sys.exit(1)
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def load_errata() -> list:
    """Load errata from single source of truth."""
    errata_path = Path(__file__).parent / "validation" / "errata.yaml"
    if not errata_path.exists():
        print("WARNING: validation/errata.yaml not found")
        return []
    with open(errata_path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)
    return data.get("errata", [])


def create_styles(cfg: dict) -> dict:
    """Create reusable styles from config."""
    colors = cfg["colors"]
    return {
        "header_fill": PatternFill(start_color=colors["header_fill"],
                                   end_color=colors["header_fill"],
                                   fill_type="solid"),
        "header_font": Font(name=cfg["fonts"]["primary_fa"],
                            color=colors["header_font"],
                            bold=True, size=12),
        "input_fill": PatternFill(start_color=colors["input_fill"],
                                  end_color=colors["input_fill"],
                                  fill_type="solid"),
        "calc_fill": PatternFill(start_color=colors["calc_fill"],
                                 end_color=colors["calc_fill"],
                                 fill_type="solid"),
        "pass_fill": PatternFill(start_color=colors["pass_fill"],
                                 end_color=colors["pass_fill"],
                                 fill_type="solid"),
        "fail_fill": PatternFill(start_color=colors["fail_fill"],
                                 end_color=colors["fail_fill"],
                                 fill_type="solid"),
        "title_font": Font(name=cfg["fonts"]["primary_fa"],
                           bold=True, size=16),
        "normal_font": Font(name=cfg["fonts"]["primary_fa"], size=11),
        "thin_border": Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        ),
    }


def build_guide_sheet(wb: Workbook, cfg: dict, styles: dict) -> None:
    """Build 00_راهنما sheet."""
    ws = wb.active
    ws.title = "00_راهنما"
    ws.sheet_properties.tabColor = "1565C0"

    ws["A1"] = f"🧪 {cfg['project']['name']}"
    ws["A1"].font = styles["title_font"]

    ws["A3"] = f"نسخه: {cfg['project']['version']}"
    ws["A4"] = f"تاریخ ساخت: {cfg['project']['build_date']}"
    ws["A5"] = f"زبان: {cfg['project']['language']}"

    ws["A7"] = "راهنمای رنگ‌بندی:"
    ws["A7"].font = Font(bold=True)

    guide = [
        ("A8", "🟡 زرد", "سلول ورودی — مقدار را وارد کنید", "input_fill"),
        ("A9", "⬜ خاکستری", "سلول محاسبه — خودکار", "calc_fill"),
        ("A10", "🟢 سبز", "نتیجه قبولی", "pass_fill"),
        ("A11", "🔴 قرمز", "خطا / نامطلوب", "fail_fill"),
    ]
    for cell, label, desc, fill_key in guide:
        ws[cell] = label
        ws[cell].fill = styles[fill_key]
        ws[cell].font = styles["normal_font"]
        ws[f"B{cell[1:]}"] = desc

    ws["A13"] = f"رمز شیت‌ها: {cfg['password']['sheet_protection']}"
    ws["A14"] = cfg["password"]["note"]
    ws["A14"].font = Font(italic=True, size=9, color="999999")

    # Metadata
    ws["A16"] = "Build Metadata:"
    ws["A16"].font = Font(bold=True)
    ws["A17"] = f"Generator: build.py v1.0"
    ws["A18"] = f"Build date: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    ws["A19"] = f"Config: config.yaml"


def build_errata_sheet(wb: Workbook, cfg: dict, styles: dict, errata: list) -> None:
    """Build 06_خطاها_هشدارها from errata.yaml."""
    ws = wb.create_sheet("06_خطاها_هشدارها")
    ws.sheet_properties.tabColor = "C62828"

    headers = ["شناسه", "نوع", "شرح", "مقدار کتاب", "مقدار مرجع", "استاندارد", "بند", "وضعیت"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = styles["header_fill"]
        cell.font = styles["header_font"]
        cell.border = styles["thin_border"]

    for row_idx, err in enumerate(errata, 2):
        values = [
            err.get("id", ""),
            err.get("type", ""),
            err.get("title", ""),
            err.get("book_value", ""),
            err.get("reference_value", ""),
            err.get("standard", ""),
            err.get("clause", ""),
            err.get("status", ""),
        ]
        for col, val in enumerate(values, 1):
            cell = ws.cell(row=row_idx, column=col, value=val)
            cell.font = styles["normal_font"]
            cell.border = styles["thin_border"]
            if err.get("status") == "confirmed":
                cell.fill = styles["pass_fill"]
            elif err.get("status") == "pending_validation":
                cell.fill = styles["input_fill"]

    # Auto-width
    for col in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 20


def build_dashboard_sheet(wb: Workbook, cfg: dict, styles: dict) -> None:
    """Build 04_داشبورد — honest, formula-driven where possible."""
    ws = wb.create_sheet("04_داشبورد")
    ws.sheet_properties.tabColor = "2E7D32"

    ws["A1"] = "📊 داشبورد پروژه"
    ws["A1"].font = styles["title_font"]

    ws["A3"] = "⚠️ توجه: مقادیر زیر فعلاً ایستا هستند."
    ws["A3"].font = Font(italic=True, color="C62828", size=10)

    ws["A5"] = "تعداد شیت‌های مورد انتظار:"
    ws["B5"] = cfg["sheets"]["total_expected"]
    ws["A6"] = "تعداد شیت‌های پیاده‌سازی‌شده:"
    ws["B6"] = cfg["sheets"]["implemented"]
    ws["A7"] = "پیشرفت:"
    ws["B7"] = f"=B6/B5"
    ws["B7"].number_format = "0%"

    ws["A9"] = "وضعیت فرمول‌ها:"
    ws["B9"] = "~70% فعال (نه ۱۰۰٪)"
    ws["B9"].font = Font(color="FF6F00", bold=True)

    ws["A11"] = "سازگاری:"
    ws["B11"] = "Desktop: تأییدشده | Web/Mobile: در انتظار تست"


def build_validation_sheet(wb: Workbook, cfg: dict, styles: dict, errata: list) -> None:
    """Build _Validation_Data from errata.yaml."""
    ws = wb.create_sheet("_Validation_Data")
    ws.sheet_properties.tabColor = "FF6F00"
    ws.sheet_state = "hidden"

    headers = ["ID", "Type", "Book Value", "Reference Value", "Standard", "Clause", "Status"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = styles["header_fill"]
        cell.font = styles["header_font"]

    for row_idx, err in enumerate(errata, 2):
        values = [
            err.get("id"), err.get("type"), err.get("book_value"),
            err.get("reference_value"), err.get("standard"),
            err.get("clause"), err.get("status"),
        ]
        for col, val in enumerate(values, 1):
            ws.cell(row=row_idx, column=col, value=val)


def protect_sheets(wb: Workbook, cfg: dict) -> None:
    """Apply sheet protection (accidental edit prevention, NOT security)."""
    password = cfg["password"]["sheet_protection"]
    for ws in wb.worksheets:
        if ws.title.startswith("_"):
            ws.protection.sheet = True
            ws.protection.password = password


def compute_sha256(filepath: str) -> str:
    """Compute SHA-256 hash of a file."""
    sha256 = hashlib.sha256()
    with open(filepath, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            sha256.update(chunk)
    return sha256.hexdigest()


def main() -> None:
    print("═" * 50)
    print("  Concrete Lab Companion — Build Script v1.0")
    print("═" * 50)

    cfg = load_config()
    errata = load_errata()
    styles = create_styles(cfg)

    print(f"\n📋 Config loaded: {cfg['project']['name']} v{cfg['project']['version']}")
    print(f"📋 Errata loaded: {len(errata)} items")

    wb = Workbook()

    print("\n🔨 Building sheets...")
    build_guide_sheet(wb, cfg, styles)
    print("  ✅ 00_راهنما")

    build_dashboard_sheet(wb, cfg, styles)
    print("  ✅ 04_داشبورد")

    build_errata_sheet(wb, cfg, styles, errata)
    print("  ✅ 06_خطاها_هشدارها")

    build_validation_sheet(wb, cfg, styles, errata)
    print("  ✅ _Validation_Data (hidden)")

    # TODO: Add remaining 20 test sheets here

    protect_sheets(wb, cfg)
    print("  🔒 Sheet protection applied")

    # Save
    output_dir = Path("output")
    output_dir.mkdir(exist_ok=True)

    filename = f"{cfg['project']['product_file']}_v{cfg['project']['version']}.xlsx"
    output_path = output_dir / filename

    wb.save(str(output_path))
    print(f"\n💾 Saved: {output_path}")

    # Hash
    file_hash = compute_sha256(str(output_path))
    hash_path = output_path.with_suffix(".xlsx.sha256")
    with open(hash_path, "w") as f:
        f.write(f"{file_hash}  {filename}\n")
    print(f"🔐 SHA-256: {file_hash}")
    print(f"💾 Hash file: {hash_path}")

    print("\n✅ Build complete!")
    print("═" * 50)


if __name__ == "__main__":
    main()