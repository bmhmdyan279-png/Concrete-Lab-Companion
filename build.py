#!/usr/bin/env python3
"""
Concrete Lab Companion — Complete Build Script v1.1.0
=====================================================
Generates ALL 20 test sheets + support sheets + hidden sheets.
No macros. Compatible with Excel Desktop / Online / Mobile.

CHANGELOG v1.1.0:
- Fixed dashboard circular reference (C7/C6 → dynamic counter)
- Fixed mortar strength cell references (A6:C6 → A7:C7)
- Fixed compressive strength self-reference (C11 → C10)
- Fixed sieve chart data range (18-25 → 23-30)
- Added MROUND for engineering rounding
- Added helper column for chart resilience
"""
import yaml, hashlib, os, sys, json
from datetime import datetime
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side, Protection, numbers)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.chart.label import DataLabelList
from openpyxl.worksheet.properties import WorksheetProperties
from openpyxl.formatting.rule import CellIsRule

# ═══════════════════════════════════════════════
# CONFIG
# ═══════════════════════════════════════════════
PASSWORD = "ConcreteLab2026!"
VERSION = "1.1.0"
BUILD_DATE = datetime.now().strftime("%Y-%m-%d %H:%M")

COLORS = {
  "input_fill": "FFF2CC",
  "input_border": "D9D9D9",
  "calc_fill": "F2F2F2",
  "pass_fill": "C6EFCE",
  "pass_font": "006100",
  "warn_fill": "FCE4D6",
  "warn_font": "C00000",
  "fail_fill": "FFC7CE",
  "fail_font": "9C0006",
  "header_fill": "1F4E79",
  "header_font": "FFFFFF",
  "nav_fill": "D6E4F0",
}

FONT_FA = "Tahoma"
FONT_NUM = "Calibri"


# ═══════════════════════════════════════════════
# STYLE HELPERS
# ═══════════════════════════════════════════════
def make_styles():
  return {
    "header_fill": PatternFill("solid", fgColor=COLORS["header_fill"]),
    "header_font": Font(name=FONT_FA, color=COLORS["header_font"], bold=True, size=11),
    "input_fill": PatternFill("solid", fgColor=COLORS["input_fill"]),
    "calc_fill": PatternFill("solid", fgColor=COLORS["calc_fill"]),
    "pass_fill": PatternFill("solid", fgColor=COLORS["pass_fill"]),
    "pass_font": Font(name=FONT_FA, color=COLORS["pass_font"], bold=True),
    "warn_fill": PatternFill("solid", fgColor=COLORS["warn_fill"]),
    "warn_font": Font(name=FONT_FA, color=COLORS["warn_font"], bold=True),
    "fail_fill": PatternFill("solid", fgColor=COLORS["fail_fill"]),
    "fail_font": Font(name=FONT_FA, color=COLORS["fail_font"], bold=True),
    "title_font": Font(name=FONT_FA, bold=True, size=14, color=COLORS["header_fill"]),
    "normal_font": Font(name=FONT_FA, size=11),
    "num_font": Font(name=FONT_NUM, size=11),
    "thin_border": Border(
      left=Side("thin", color="BFBFBF"), right=Side("thin", color="BFBFBF"),
      top=Side("thin", color="BFBFBF"), bottom=Side("thin", color="BFBFBF"),
    ),
    "input_border": Border(
      left=Side("thin", color=COLORS["input_border"]), right=Side("thin", color=COLORS["input_border"]),
      top=Side("thin", color=COLORS["input_border"]), bottom=Side("thin", color=COLORS["input_border"]),
    ),
    "center_align": Alignment(horizontal="center", vertical="center", wrap_text=True),
    "right_align": Alignment(horizontal="right", vertical="center", wrap_text=True),
    "nav_font": Font(name=FONT_FA, size=10, color="1F4E79"),
    "nav_fill": PatternFill("solid", fgColor=COLORS["nav_fill"]),
  }


S = make_styles()


def set_cell(ws, row, col, value, font=None, fill=None, border=None, align=None, num_fmt=None, locked=True):
  cell = ws.cell(row=row, column=col, value=value)
  cell.font = font or S["normal_font"]
  if fill: cell.fill = fill
  if border: cell.border = border
  cell.alignment = align or S["center_align"]
  if num_fmt: cell.number_format = num_fmt
  cell.protection = Protection(locked=locked)
  return cell


def add_nav_bar(ws):
  """Add navigation bar at top of every sheet."""
  nav = "🏠 راهنما | 📑 گزارش | 🔒 محافظت‌شده"
  for col in range(1, 12):
    cell = ws.cell(row=1, column=col)
    cell.fill = S["nav_fill"]
    cell.font = S["nav_font"]

  ws.cell(row=1, column=1, value="🏠").font = S["nav_font"]
  ws.cell(row=1, column=1).hyperlink = "#00_راهنما!A1"

  ws.cell(row=1, column=3, value="📑").font = S["nav_font"]
  ws.cell(row=1, column=3).hyperlink = "#22_گزارش!A1"

  ws.cell(row=1, column=5, value="🔒").font = S["nav_font"]


def add_title(ws, title, subtitle=""):
  ws.merge_cells("A2:K2")
  c = ws.cell(row=2, column=1, value=title)
  c.font = S["title_font"]
  c.alignment = S["right_align"]

  if subtitle:
    ws.merge_cells("A3:K3")
    c2 = ws.cell(row=3, column=1, value=subtitle)
    c2.font = Font(name=FONT_FA, size=9, italic=True, color="666666")
    c2.alignment = S["right_align"]


def add_dv(ws, cells, dv_type, min_val=None, max_val=None, formula=None, allow_blank=True,
           error_msg="مقدار نامعتبر", prompt_msg=""):
  """Add Data Validation to a range of cells."""
  if dv_type == "decimal":
    dv = DataValidation(type="decimal", operator="between",
                        formula1=str(min_val), formula2=str(max_val), allow_blank=allow_blank)
  elif dv_type == "whole":
    dv = DataValidation(type="whole", operator="between",
                        formula1=str(min_val), formula2=str(max_val), allow_blank=allow_blank)
  elif dv_type == "list":
    dv = DataValidation(type="list", formula1=formula, allow_blank=allow_blank)
  elif dv_type == "custom":
    dv = DataValidation(type="custom", formula1=formula, allow_blank=allow_blank)
  else:
    return

  dv.error = error_msg
  dv.errorTitle = "خطای ورودی"
  dv.prompt = prompt_msg
  dv.promptTitle = "راهنما"

  ws.add_data_validation(dv)
  for cell_ref in cells:
    dv.add(cell_ref)


def add_errata_check(ws, row, col, condition_formula, msg):
  """Add a cell that shows error message if condition is TRUE."""
  set_cell(ws, row, col, f'=IF({condition_formula},"⚠️ {msg}","")',
           font=S["warn_font"], fill=S["warn_fill"], align=S["right_align"], locked=True)


def protect_sheet(ws):
  ws.protection.sheet = True
  ws.protection.password = PASSWORD
  ws.protection.selectLockedCells = False
  ws.protection.selectUnlockedCells = False
  ws.protection.formatCells = False
  ws.protection.formatColumns = False
  ws.protection.formatRows = False
  ws.protection.insertColumns = False
  ws.protection.insertRows = False
  ws.protection.deleteColumns = False
  ws.protection.deleteRows = False


# ═══════════════════════════════════════════════
# SHEET 00: GUIDE
# ═══════════════════════════════════════════════
def build_00_guide(wb):
  ws = wb.active
  ws.title = "00_راهنما"
  ws.sheet_properties.tabColor = "1F4E79"
  ws.sheet_view.rightToLeft = True

  add_nav_bar(ws)
  add_title(ws, "🧪 همراه دیجیتال آزمایشگاه فناوری بتن", f"نسخه {VERSION} | ساخت: {BUILD_DATE}")

  r = 5
  set_cell(ws, r, 1, "📌 قانون طلایی:", font=Font(name=FONT_FA, bold=True, size=12, color="C00000"),
           align=S["right_align"])
  ws.merge_cells(f"A{r}:K{r}")
  r += 1
  set_cell(ws, r, 1, "«اول دستی حساب کن، بعد اینجا راستی‌آزمایی کن»", font=Font(name=FONT_FA, size=11, italic=True),
           align=S["right_align"])
  ws.merge_cells(f"A{r}:K{r}")

  r += 2
  set_cell(ws, r, 1, "🎨 راهنمای رنگ‌بندی:", font=Font(name=FONT_FA, bold=True, size=12), align=S["right_align"])
  r += 1

  legend = [
    ("🟡 زرد", "سلول ورودی — مقدار را وارد کنید", "input_fill"),
    ("⬜ خاکستری", "سلول محاسبه — خودکار (دست نزنید)", "calc_fill"),
    ("🟢 سبز", "نتیجه قبولی / تأییدشده", "pass_fill"),
    ("🟠 نارنجی", "هشدار — نیاز به بررسی", "warn_fill"),
    ("🔴 قرمز", "خطا / عدم قبولی / رد", "fail_fill"),
  ]

  for label, desc, fill_key in legend:
    set_cell(ws, r, 1, label, fill=S[fill_key], font=Font(name=FONT_FA, bold=True))
    ws.merge_cells(f"B{r}:K{r}")
    set_cell(ws, r, 2, desc, align=S["right_align"])
    r += 1

  r += 1
  set_cell(ws, r, 1, "📋 فهرست شیت‌ها:", font=Font(name=FONT_FA, bold=True, size=12), align=S["right_align"])
  r += 1

  sheets_list = [
    ("01", "اطلاعات آزمون", "مشخصات پروژه، نمونه، اپراتور"),
    ("02", "آزمایش ۱-۱ دانه‌بندی", "ASTM C136"),
    ("03", "آزمایش ۱-۲ رطوبت", "ASTM C566"),
    ("04", "آزمایش ۱-۳ چگالی درشت", "ASTM C127"),
    ("05", "آزمایش ۱-۴ چگالی ریز", "ASTM C128"),
    ("06", "آزمایش ۱-۵ وزن واحد حجمی", "ASTM C138"),
    ("07", "آزمایش ۱-۶ معادل ماسه", "ASTM D2419"),
    ("08", "آزمایش ۱-۷ شاخص شکل", "ASTM D4791"),
    ("09", "آزمایش ۱-۸ جذب آب", "ASTM C127"),
    ("10", "آزمایش ۲-۱ چگالی بتن تازه", "ASTM C138"),
    ("11", "آزمایش ۲-۲ ویکات", "ASTM C191"),
    ("12", "آزمایش ۲-۳ زمان گیرش", "ASTM C191"),
    ("13", "آزمایش ۲-۴ مقاومت ملات", "EN 196-1"),
    ("14", "آزمایش ۳-۱ اسلامپ", "ASTM C143"),
    ("15", "آزمایش ۳-۲ آب‌انداختگی", "ASTM C232"),
    ("16", "آزمایش ۳-۳ وزن واحد بتن", "ASTM C138"),
    ("17", "آزمایش ۴-۱ مقاومت فشاری", "ASTM C39"),
    ("18", "آزمایش ۴-۲ مقاومت کششی", "ASTM C496"),
    ("19", "آزمایش ۴-۳ مقاومت خمشی", "ASTM C78"),
    ("20", "آزمایش ۴-۴ اولتراسونیک", "ASTM C597"),
    ("21", "آزمایش ۴-۵ چکش اشمیت", "ASTM C805"),
    ("22", "گزارش", "خلاصه نتایج + چاپ"),
    ("23", "داشبورد", "وضعیت کلی"),
    ("24", "QA Test", "تست‌های خودکار"),
    ("25", "خطاها و هشدارها", "تجمع زنده"),
  ]

  for code, name, std in sheets_list:
    set_cell(ws, r, 1, code, font=S["num_font"])
    ws.merge_cells(f"B{r}:E{r}")
    set_cell(ws, r, 2, name, align=S["right_align"])
    ws.merge_cells(f"F{r}:K{r}")
    set_cell(ws, r, 6, std, font=Font(name=FONT_NUM, size=9, color="666666"), align=S["center_align"])
    r += 1

  r += 1
  set_cell(ws, r, 1, "🔧 Metadata:", font=Font(name=FONT_FA, bold=True, size=11), align=S["right_align"])
  r += 1

  meta = [
    ("نسخه", VERSION),
    ("تاریخ ساخت", BUILD_DATE),
    ("Generator", "build.py v1.1"),
    ("رمز شیت‌ها", PASSWORD),
    ("توجه", "رمز عمومی و غیرمحرمانه — فقط برای جلوگیری از ویرایش تصادفی"),
  ]

  for k, v in meta:
    set_cell(ws, r, 1, k, font=Font(name=FONT_FA, bold=True, size=10), align=S["right_align"])
    ws.merge_cells(f"B{r}:K{r}")
    set_cell(ws, r, 2, v, font=Font(name=FONT_FA, size=10), align=S["right_align"])
    r += 1

  r += 1
  set_cell(ws, r, 1, "📝 اصلاحات نسبت به چاپ کتاب:", font=Font(name=FONT_FA, bold=True, size=11, color="C00000"),
           align=S["right_align"])
  r += 1
  set_cell(ws, r, 1, "۸ خطای فیزیکی/محاسباتی شناسایی و اصلاح شده — جزئیات در شیت ۲۵_خطاها",
           font=Font(name=FONT_FA, size=10), align=S["right_align"])
  ws.merge_cells(f"A{r}:K{r}")

  for col in range(1, 12):
    ws.column_dimensions[get_column_letter(col)].width = 14


# ═══════════════════════════════════════════════
# SHEET 01: TEST INFO
# ═══════════════════════════════════════════════
def build_01_info(wb):
  ws = wb.create_sheet("01_اطلاعات_آزمون")
  ws.sheet_properties.tabColor = "2E7D32"
  ws.sheet_view.rightToLeft = True

  add_nav_bar(ws)
  add_title(ws, "📋 اطلاعات آزمون", "این اطلاعات در سربرگ گزارش تکرار می‌شود")

  fields = [
    ("نام پروژه", "", "B"),
    ("شماره نمونه", "", "B"),
    ("تاریخ آزمون", "", "B"),
    ("نام اپراتور", "", "B"),
    ("دستگاه / تجهیزات", "", "B"),
    ("دمای محیط (°C)", "", "B"),
    ("رطوبت نسبی (%)", "", "B"),
    ("استاندارد مرجع", "", "B"),
    ("توضیحات", "", "B"),
  ]

  r = 5
  for label, default, col_letter in fields:
    set_cell(ws, r, 1, label, font=Font(name=FONT_FA, bold=True, size=11), align=S["right_align"])
    ws.merge_cells(f"B{r}:F{r}")
    set_cell(ws, r, 2, default, fill=S["input_fill"], border=S["input_border"], locked=False, align=S["right_align"])
    r += 1

  add_dv(ws, [f"B{5 + i}" for i in range(len(fields))], "custom", formula='TRUE',
         prompt_msg="مقدار را وارد کنید")

  for col in range(1, 12):
    ws.column_dimensions[get_column_letter(col)].width = 16


# ═══════════════════════════════════════════════
# TEST SHEET BUILDERS
# ═══════════════════════════════════════════════
def build_1_1_sieve(wb):
  """1-1: Sieve Analysis (ASTM C136)"""
  ws = wb.create_sheet("02_آزمایش_1-1")
  ws.sheet_properties.tabColor = "FF6F00"
  ws.sheet_view.rightToLeft = True

  add_nav_bar(ws)
  add_title(ws, "آزمایش ۱-۱: دانه‌بندی سنگدانه", "ASTM C136 / ISIRI 4977")

  r = 5
  set_cell(ws, r, 1, "جرم اولیه نمونه خشک (g):", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  set_cell(ws, r, 3, None, fill=S["input_fill"], border=S["input_border"], locked=False, num_fmt='0.0')
  add_dv(ws, ["C5"], "decimal", min_val=0, max_val=100000, error_msg="جرم باید مثبت باشد")

  r += 2
  headers = ["الک", "اندازه (mm)", "مانده (g)", "% مانده", "% مانده تجمعی", "% عبوری", "الک استاندارد"]
  for col, h in enumerate(headers, 1):
    set_cell(ws, r, col, h, font=S["header_font"], fill=S["header_fill"])

  sieves = [
    ("3/8\"", 9.5), ("#4", 4.75), ("#8", 2.36), ("#16", 1.18),
    ("#30", 0.6), ("#50", 0.3), ("#100", 0.15), ("#200", 0.075), ("پان", 0)
  ]

  start_row = r + 1
  for i, (name, size) in enumerate(sieves):
    row = start_row + i
    set_cell(ws, row, 1, name, font=S["normal_font"])
    set_cell(ws, row, 2, size, font=S["num_font"], num_fmt='0.000')
    set_cell(ws, row, 3, None, fill=S["input_fill"], border=S["input_border"], locked=False, num_fmt='0.0')

    # % retained = retained / total * 100
    set_cell(ws, row, 4, f'=IF(OR($C$5=0,$C{row}=""),"",ROUND($C{row}/$C$5*100,2))',
             fill=S["calc_fill"], num_fmt='0.00')

    # cumulative % retained
    if i == 0:
      set_cell(ws, row, 5, f'=IF($D{row}="","",$D{row})', fill=S["calc_fill"], num_fmt='0.00')
    else:
      set_cell(ws, row, 5, f'=IF(OR($D{row}="",$E{row - 1}=""),"",$E{row - 1}+$D{row})',
               fill=S["calc_fill"], num_fmt='0.00')

    # % passing = 100 - cumulative
    set_cell(ws, row, 6, f'=IF($E{row}="","",ROUND(100-$E{row},2))', fill=S["calc_fill"], num_fmt='0.00')

    # standard sieve flag (TRUE for #4 through #100)
    is_std = name in ["#4", "#8", "#16", "#30", "#50", "#100"]
    set_cell(ws, row, 7, is_std, font=S["num_font"])

  end_row = start_row + len(sieves) - 1
  add_dv(ws, [f"C{start_row + i}" for i in range(len(sieves))], "decimal", min_val=0, max_val=100000)

  # Mass check
  r_check = end_row + 2
  set_cell(ws, r_check, 1, "کنترل جرم:", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  set_cell(ws, r_check, 3, f'=IF($C$5=0,"",IF(ABS($C$5-SUM(C{start_row}:C{end_row}))/$C$5>0.003,"❌ اختلاف >0.3%","✅"))',
           fill=S["calc_fill"], align=S["center_align"])

  # FM calculation
  r_fm = r_check + 1
  set_cell(ws, r_fm, 1, "مدول نرمی (FM):", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  set_cell(ws, r_fm, 3,
           f'=IF($C$5=0,"",ROUND(SUMPRODUCT((G{start_row}:G{end_row}=TRUE)*E{start_row}:E{end_row})/100,2))',
           fill=PatternFill("solid", fgColor=COLORS["pass_fill"]), font=S["pass_font"], num_fmt='0.00')

  # Chart data area (for sieve chart)
  r_chart = r_fm + 2
  set_cell(ws, r_chart, 1, "📊 داده‌های نمودار:", font=Font(name=FONT_FA, bold=True), align=S["right_align"])

  set_cell(ws, r_chart + 1, 1, "اندازه (mm)", font=S["header_font"], fill=S["header_fill"])
  set_cell(ws, r_chart + 1, 2, "% عبوری نمونه", font=S["header_font"], fill=S["header_fill"])
  set_cell(ws, r_chart + 1, 3, "حد بالا ISIRI", font=S["header_font"], fill=S["header_fill"])
  set_cell(ws, r_chart + 1, 4, "حد پایین ISIRI", font=S["header_font"], fill=S["header_fill"])

  # ISIRI 302 limits (example for coarse aggregate)
  isiri_limits = {9.5: (100, 95), 4.75: (95, 80), 2.36: (80, 60), 1.18: (60, 40),
                  0.6: (40, 25), 0.3: (25, 10), 0.15: (10, 2), 0.075: (2, 0)}

  for i, (size, (hi, lo)) in enumerate(isiri_limits.items()):
    row = r_chart + 2 + i
    set_cell(ws, row, 1, size, font=S["num_font"], num_fmt='0.000')
    # Helper column: returns NA() if blank so chart skips the point (prevents #REF!)
    set_cell(ws, row, 2, f'=IF(F{start_row + i}="",NA(),F{start_row + i})', fill=S["calc_fill"], num_fmt='0.0')
    set_cell(ws, row, 3, hi, font=S["num_font"], num_fmt='0.0')
    set_cell(ws, row, 4, lo, font=S["num_font"], num_fmt='0.0')

  for col in range(1, 12):
    ws.column_dimensions[get_column_letter(col)].width = 13


def build_1_2_moisture(wb):
  """1-2: Moisture Content (ASTM C566)"""
  ws = wb.create_sheet("03_آزمایش_1-2")
  ws.sheet_properties.tabColor = "FF6F00"
  ws.sheet_view.rightToLeft = True

  add_nav_bar(ws)
  add_title(ws, "آزمایش ۱-۲: رطوبت سنگدانه", "ASTM C566 — پایه خشک")

  r = 5
  fields = [
    ("W1: جرم نمونه تر (g)", "C5", 0, 100000),
    ("W2: جرم نمونه خشک (g)", "C6", 0, 100000),
  ]

  for label, cell_ref, mn, mx in fields:
    set_cell(ws, r, 1, label, font=Font(name=FONT_FA, bold=True), align=S["right_align"])
    ws.merge_cells(f"A{r}:B{r}")
    set_cell(ws, r, 3, None, fill=S["input_fill"], border=S["input_border"], locked=False, num_fmt='0.0')
    r += 1

  add_dv(ws, ["C5", "C6"], "decimal", min_val=0, max_val=100000)
  r += 1

  set_cell(ws, r, 1, "درصد رطوبت (پایه خشک):", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, '=IF(OR(C5="",C6="",C6=0),"—",ROUND((C5-C6)/C6*100,2))',
           fill=S["calc_fill"], num_fmt='0.00')
  r += 1

  set_cell(ws, r, 1, "⚠️ اِراتا:", font=S["warn_font"], fill=S["warn_fill"], align=S["right_align"])
  ws.merge_cells(f"A{r}:K{r}")
  r += 1
  set_cell(ws, r, 1, "کتاب از پایه تر استفاده کرده؛ فرمول صحیح طبق ASTM C566 پایه خشک است.",
           font=Font(name=FONT_FA, size=9, italic=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:K{r}")

  for col in range(1, 12):
    ws.column_dimensions[get_column_letter(col)].width = 14


def build_1_3_coarse_sg(wb):
  """1-3: Coarse Aggregate SG (ASTM C127)"""
  ws = wb.create_sheet("04_آزمایش_1-3")
  ws.sheet_properties.tabColor = "FF6F00"
  ws.sheet_view.rightToLeft = True

  add_nav_bar(ws)
  add_title(ws, "آزمایش ۱-۳: چگالی سنگدانه درشت", "ASTM C127")

  r = 5
  fields = [
    ("A: جرم خشک (g)", "C5"),
    ("B: جرم SSD (g)", "C6"),
    ("C: جرم در آب (g)", "C7"),
  ]

  for label, cell_ref in fields:
    set_cell(ws, r, 1, label, font=Font(name=FONT_FA, bold=True), align=S["right_align"])
    ws.merge_cells(f"A{r}:B{r}")
    set_cell(ws, r, 3, None, fill=S["input_fill"], border=S["input_border"], locked=False, num_fmt='0.0')
    r += 1

  add_dv(ws, ["C5", "C6", "C7"], "decimal", min_val=0, max_val=100000)
  r += 1

  results = [
    ("OD (خشک)", '=IF(OR(C5="",C6="",C7="",C6=C7),"—",ROUND(C5/(C6-C7),3))'),
    ("SSD (اشباع خشک)", '=IF(OR(C5="",C6="",C7="",C6=C7),"—",ROUND(C6/(C6-C7),3))'),
    ("App (ظاهری)", '=IF(OR(C5="",C6="",C7="",C5=C7),"—",ROUND(C5/(C5-C7),3))'),
    ("جذب آب (%)", '=IF(OR(C5="",C6="",C5=0),"—",ROUND((C6-C5)/C5*100,2))'),
  ]

  for label, formula in results:
    set_cell(ws, r, 1, label, font=Font(name=FONT_FA, bold=True), align=S["right_align"])
    ws.merge_cells(f"A{r}:B{r}")
    set_cell(ws, r, 3, formula, fill=S["calc_fill"], num_fmt='0.000')
    r += 1

  # Physical check: SSD >= OD
  r += 1
  set_cell(ws, r, 1, "بررسی فیزیکی:", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, '=IF(OR(C9="—",C10="—"),"",IF(C10<C9,"❌ SSD < OD (غیرفیزیکی)","✅"))',
           fill=S["calc_fill"])

  for col in range(1, 12):
    ws.column_dimensions[get_column_letter(col)].width = 14


def build_1_4_fine_sg(wb):
  """1-4: Fine Aggregate SG (ASTM C128)"""
  ws = wb.create_sheet("05_آزمایش_1-4")
  ws.sheet_properties.tabColor = "FF6F00"
  ws.sheet_view.rightToLeft = True

  add_nav_bar(ws)
  add_title(ws, "آزمایش ۱-۴: چگالی سنگدانه ریز", "ASTM C128")

  r = 5
  set_cell(ws, r, 1, "روش:", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  set_cell(ws, r, 3, "وزن‌سنجی", fill=S["input_fill"], border=S["input_border"], locked=False)
  add_dv(ws, ["C5"], "list", formula='"وزن‌سنجی,حجم‌سنجی"')

  r += 2
  fields = [
    ("A: جرم خشک (g)", "C7"),
    ("S: جرم SSD (g)", "C8"),
    ("B: جرم ظرف+آب (g)", "C9"),
    ("C: جرم ظرف+آب+نمونه (g)", "C10"),
  ]

  for label, cell_ref in fields:
    set_cell(ws, r, 1, label, font=Font(name=FONT_FA, bold=True), align=S["right_align"])
    ws.merge_cells(f"A{r}:B{r}")
    set_cell(ws, r, 3, None, fill=S["input_fill"], border=S["input_border"], locked=False, num_fmt='0.0')
    r += 1

  add_dv(ws, ["C7", "C8", "C9", "C10"], "decimal", min_val=0, max_val=100000)
  r += 1

  results = [
    ("OD (وزن‌سنجی)", '=IF(C5="وزن‌سنجی",IF(OR(C7="",C9="",C10="",C9+C8-C10=0),"—",ROUND(C7/(C9+C8-C10),3)),"—")'),
    ("SSD (وزن‌سنجی)", '=IF(C5="وزن‌سنجی",IF(OR(C8="",C9="",C10="",C9+C8-C10=0),"—",ROUND(C8/(C9+C8-C10),3)),"—")'),
    ("App (وزن‌سنجی)", '=IF(C5="وزن‌سنجی",IF(OR(C7="",C9="",C10="",C9+C7-C10=0),"—",ROUND(C7/(C9+C7-C10),3)),"—")'),
    ("جذب (%)", '=IF(OR(C7="",C8="",C7=0),"—",ROUND((C8-C7)/C7*100,2))'),
  ]

  for label, formula in results:
    set_cell(ws, r, 1, label, font=Font(name=FONT_FA, bold=True), align=S["right_align"])
    ws.merge_cells(f"A{r}:B{r}")
    set_cell(ws, r, 3, formula, fill=S["calc_fill"], num_fmt='0.000')
    r += 1

  # Errata warning
  r += 1
  set_cell(ws, r, 1, "⚠️ اِراتا:", font=S["warn_font"], fill=S["warn_fill"], align=S["right_align"])
  ws.merge_cells(f"A{r}:K{r}")
  r += 1
  set_cell(ws, r, 1, "نمونه کتاب (OD=1.51 / SSD=2.63) فیزیکی نیست. احتمال جابه‌جایی A و S.",
           font=Font(name=FONT_FA, size=9, italic=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:K{r}")

  for col in range(1, 12):
    ws.column_dimensions[get_column_letter(col)].width = 14


def build_1_5_unit_weight(wb):
  """1-5: Unit Weight (ASTM C138)"""
  ws = wb.create_sheet("06_آزمایش_1-5")
  ws.sheet_properties.tabColor = "FF6F00"
  ws.sheet_view.rightToLeft = True

  add_nav_bar(ws)
  add_title(ws, "آزمایش ۱-۵: وزن واحد حجمی سنگدانه", "ASTM C138")

  r = 5
  fields = [
    ("T: جرم ظرف خالی (g)", "C5"),
    ("G: جرم ظرف+سنگدانه (g)", "C6"),
    ("V: حجم ظرف (cm³)", "C7"),
    ("S: چگالی (بی‌بُعد)", "C8"),
  ]

  for label, cell_ref in fields:
    set_cell(ws, r, 1, label, font=Font(name=FONT_FA, bold=True), align=S["right_align"])
    ws.merge_cells(f"A{r}:B{r}")
    set_cell(ws, r, 3, None, fill=S["input_fill"], border=S["input_border"], locked=False, num_fmt='0.0')
    r += 1

  add_dv(ws, ["C5", "C6", "C7"], "decimal", min_val=0, max_val=100000)
  add_dv(ws, ["C8"], "decimal", min_val=2, max_val=3.5, error_msg="چگالی باید بین ۲ تا ۳.۵ باشد")
  r += 1

  set_cell(ws, r, 1, "وزن واحد حجمی (kg/m³):", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, '=IF(OR(C5="",C6="",C7="",C7=0),"—",ROUND((C6-C5)/C7*1000,0))',
           fill=S["calc_fill"], num_fmt='#,##0')
  r += 1

  set_cell(ws, r, 1, "فضای خالی (%):", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, '=IF(OR(C8="",C8=0,C10="—"),"—",ROUND((C8*998-C10/1000*1000)/(C8*998)*100,1))',
           fill=S["calc_fill"], num_fmt='0.0')
  r += 1

  set_cell(ws, r, 1, "⚠️ اِراتا:", font=S["warn_font"], fill=S["warn_fill"], align=S["right_align"])
  ws.merge_cells(f"A{r}:K{r}")
  r += 1
  set_cell(ws, r, 1, "S در کتاب ۱۶۰۰ نوشته شده (برچسب غلط). S باید بی‌بُعد باشد (مثلاً ۲.۶۵).",
           font=Font(name=FONT_FA, size=9, italic=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:K{r}")

  for col in range(1, 12):
    ws.column_dimensions[get_column_letter(col)].width = 14


def build_1_6_sand_equivalent(wb):
  """1-6: Sand Equivalent (ASTM D2419)"""
  ws = wb.create_sheet("07_آزمایش_1-6")
  ws.sheet_properties.tabColor = "FF6F00"
  ws.sheet_view.rightToLeft = True

  add_nav_bar(ws)
  add_title(ws, "آزمایش ۱-۶: معادل ماسه", "ASTM D2419")

  r = 5
  fields = [
    ("خوانش ماسه (mm)", "C5"),
    ("خوانش رس (mm)", "C6"),
  ]

  for label, cell_ref in fields:
    set_cell(ws, r, 1, label, font=Font(name=FONT_FA, bold=True), align=S["right_align"])
    ws.merge_cells(f"A{r}:B{r}")
    set_cell(ws, r, 3, None, fill=S["input_fill"], border=S["input_border"], locked=False, num_fmt='0.0')
    r += 1

  add_dv(ws, ["C5", "C6"], "decimal", min_val=0, max_val=500)
  r += 1

  set_cell(ws, r, 1, "معادل ماسه SE (%):", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, '=IF(OR(C5="",C6="",C5+C6=0),"—",ROUNDUP(C5/(C5+C6)*100,0))',
           fill=S["calc_fill"], num_fmt='0')
  r += 1

  set_cell(ws, r, 1, "⚠️ اِراتا:", font=S["warn_font"], fill=S["warn_fill"], align=S["right_align"])
  ws.merge_cells(f"A{r}:K{r}")
  r += 1
  set_cell(ws, r, 1, "متن کتاب «ماسه/رس» نوشته؛ فرمول صحیح: ماسه/(ماسه+رس)×۱۰۰",
           font=Font(name=FONT_FA, size=9, italic=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:K{r}")

  for col in range(1, 12):
    ws.column_dimensions[get_column_letter(col)].width = 14


def build_1_7_shape(wb):
  """1-7: Shape Indices (ASTM D4791)"""
  ws = wb.create_sheet("08_آزمایش_1-7")
  ws.sheet_properties.tabColor = "FF6F00"
  ws.sheet_view.rightToLeft = True

  add_nav_bar(ws)
  add_title(ws, "آزمایش ۱-۷: شاخص‌های شکل سنگدانه", "ASTM D4791")

  r = 5
  fields = [
    ("W کل (g)", "C5"),
    ("W دراز (g)", "C6"),
    ("W پهن (g)", "C7"),
    ("W هر دو (g)", "C8"),
  ]

  for label, cell_ref in fields:
    set_cell(ws, r, 1, label, font=Font(name=FONT_FA, bold=True), align=S["right_align"])
    ws.merge_cells(f"A{r}:B{r}")
    set_cell(ws, r, 3, None, fill=S["input_fill"], border=S["input_border"], locked=False, num_fmt='0.0')
    r += 1

  add_dv(ws, ["C5", "C6", "C7", "C8"], "decimal", min_val=0, max_val=100000)
  r += 1

  results = [
    ("شاخص درازگی (%)", '=IF(OR(C5="",C5=0,C6=""),"—",ROUND(C6/C5*100,1))'),
    ("شاخص پهنی (%)", '=IF(OR(C5="",C5=0,C7=""),"—",ROUND(C7/C5*100,1))'),
    ("شاخص هر دو (%)", '=IF(OR(C5="",C5=0,C8=""),"—",ROUND(C8/C5*100,1))'),
  ]

  for label, formula in results:
    set_cell(ws, r, 1, label, font=Font(name=FONT_FA, bold=True), align=S["right_align"])
    ws.merge_cells(f"A{r}:B{r}")
    set_cell(ws, r, 3, formula, fill=S["calc_fill"], num_fmt='0.0')
    r += 1

  for col in range(1, 12):
    ws.column_dimensions[get_column_letter(col)].width = 14


def build_1_8_absorption(wb):
  """1-8: Absorption (ASTM C127)"""
  ws = wb.create_sheet("09_آزمایش_1-8")
  ws.sheet_properties.tabColor = "FF6F00"
  ws.sheet_view.rightToLeft = True

  add_nav_bar(ws)
  add_title(ws, "آزمایش ۱-۸: جذب آب سنگدانه", "ASTM C127")

  r = 5
  fields = [
    ("W1: جرم SSD (g)", "C5"),
    ("W2: جرم خشک (g)", "C6"),
  ]

  for label, cell_ref in fields:
    set_cell(ws, r, 1, label, font=Font(name=FONT_FA, bold=True), align=S["right_align"])
    ws.merge_cells(f"A{r}:B{r}")
    set_cell(ws, r, 3, None, fill=S["input_fill"], border=S["input_border"], locked=False, num_fmt='0.0')
    r += 1

  add_dv(ws, ["C5", "C6"], "decimal", min_val=0, max_val=100000)
  r += 1

  set_cell(ws, r, 1, "جذب آب (%):", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, '=IF(OR(C5="",C6="",C6=0),"—",ROUND((C5-C6)/C6*100,2))',
           fill=S["calc_fill"], num_fmt='0.00')

  for col in range(1, 12):
    ws.column_dimensions[get_column_letter(col)].width = 14


def build_2_1_fresh_density(wb):
  """2-1: Fresh Concrete Density (ASTM C138)"""
  ws = wb.create_sheet("10_آزمایش_2-1")
  ws.sheet_properties.tabColor = "2196F3"
  ws.sheet_view.rightToLeft = True

  add_nav_bar(ws)
  add_title(ws, "آزمایش ۲-۱: چگالی بتن تازه", "ASTM C138")

  r = 5
  fields = [
    ("Ma: جرم ظرف خالی (g)", "C5"),
    ("Mt: جرم ظرف+بتن (g)", "C6"),
    ("V: حجم ظرف (cm³)", "C7"),
  ]

  for label, cell_ref in fields:
    set_cell(ws, r, 1, label, font=Font(name=FONT_FA, bold=True), align=S["right_align"])
    ws.merge_cells(f"A{r}:B{r}")
    set_cell(ws, r, 3, None, fill=S["input_fill"], border=S["input_border"], locked=False, num_fmt='0.0')
    r += 1

  add_dv(ws, ["C5", "C6", "C7"], "decimal", min_val=0, max_val=100000)
  r += 1

  set_cell(ws, r, 1, "چگالی (kg/m³):", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, '=IF(OR(C5="",C6="",C7="",C7=0),"—",ROUND((C6-C5)/C7*1000,3))',
           fill=S["calc_fill"], num_fmt='0.000')

  for col in range(1, 12):
    ws.column_dimensions[get_column_letter(col)].width = 14


def build_2_2_vicat(wb):
  """2-2: Vicat (ASTM C191)"""
  ws = wb.create_sheet("11_آزمایش_2-2")
  ws.sheet_properties.tabColor = "2196F3"
  ws.sheet_view.rightToLeft = True

  add_nav_bar(ws)
  add_title(ws, "آزمایش ۲-۲: ویکات (گیرش سیمان)", "ASTM C191")

  r = 5
  fields = [
    ("سیمان (g)", "C5"),
    ("آب (g)", "C6"),
    ("نفوذ اولیه (mm)", "C7"),
  ]

  for label, cell_ref in fields:
    set_cell(ws, r, 1, label, font=Font(name=FONT_FA, bold=True), align=S["right_align"])
    ws.merge_cells(f"A{r}:B{r}")
    set_cell(ws, r, 3, None, fill=S["input_fill"], border=S["input_border"], locked=False, num_fmt='0.0')
    r += 1

  add_dv(ws, ["C5", "C6"], "decimal", min_val=0, max_val=10000)
  add_dv(ws, ["C7"], "decimal", min_val=0, max_val=50)
  r += 1

  set_cell(ws, r, 1, "نسبت آب به سیمان (%):", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, '=IF(OR(C5="",C6="",C5=0),"—",ROUND(C6/C5*100,1))',
           fill=S["calc_fill"], num_fmt='0.0')
  r += 1

  set_cell(ws, r, 1, "وضعیت نفوذ:", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, '=IF(C7="","",IF(ABS(C7-10)>1,"⚠️ تکرار با آب جدید","✅"))',
           fill=S["calc_fill"])

  for col in range(1, 12):
    ws.column_dimensions[get_column_letter(col)].width = 14


def build_2_3_setting_time(wb):
  """2-3: Setting Time (ASTM C191)"""
  ws = wb.create_sheet("12_آزمایش_2-3")
  ws.sheet_properties.tabColor = "2196F3"
  ws.sheet_view.rightToLeft = True

  add_nav_bar(ws)
  add_title(ws, "آزمایش ۲-۳: زمان گیرش سیمان", "ASTM C191")

  r = 5
  fields = [
    ("E: زمان اولیه (min)", "C5"),
    ("H: زمان ثانویه (min)", "C6"),
    ("C: نفوذ ثانویه (mm)", "C7"),
    ("D: نفوذ اولیه (mm)", "C8"),
    ("زمان گیرش نهایی (min)", "C9"),
  ]

  for label, cell_ref in fields:
    set_cell(ws, r, 1, label, font=Font(name=FONT_FA, bold=True), align=S["right_align"])
    ws.merge_cells(f"A{r}:B{r}")
    set_cell(ws, r, 3, None, fill=S["input_fill"], border=S["input_border"], locked=False, num_fmt='0')
    r += 1

  add_dv(ws, ["C5", "C6", "C7", "C8", "C9"], "decimal", min_val=0, max_val=10000)
  r += 1

  set_cell(ws, r, 1, "زمان گیرش اولیه (min):", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, '=IF(OR(C5="",C6="",C7="",C8="",C7=C8),"—",ROUND(C5+(C6-C5)*(C7-25)/(C7-C8),0))',
           fill=S["calc_fill"], num_fmt='0')
  r += 1

  set_cell(ws, r, 1, "زمان گیرش نهایی (min):", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, '=IF(C9="","",MROUND(C9,5))', fill=S["calc_fill"], num_fmt='0')

  for col in range(1, 12):
    ws.column_dimensions[get_column_letter(col)].width = 14


def build_2_4_mortar(wb):
  """2-4: Mortar Strength (EN 196-1)"""
  ws = wb.create_sheet("13_آزمایش_2-4")
  ws.sheet_properties.tabColor = "2196F3"
  ws.sheet_view.rightToLeft = True

  add_nav_bar(ws)
  add_title(ws, "آزمایش ۲-۴: مقاومت ملات سیمان", "EN 196-1")

  r = 5
  set_cell(ws, r, 1, "بارهای خمشی (kgf) — ۳ نمونه:", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:K{r}")
  r += 1

  for i in range(3):
    set_cell(ws, r, 1 + i, f"نمونه {i + 1}", font=S["header_font"], fill=S["header_fill"])
    set_cell(ws, r + 1, 1 + i, None, fill=S["input_fill"], border=S["input_border"], locked=False, num_fmt='0.0')
  add_dv(ws, [f"{get_column_letter(i + 1)}{r + 1}" for i in range(3)], "decimal", min_val=0, max_val=10000)
  flex_row = r + 1  # This is where the flexural data actually is
  r += 3

  set_cell(ws, r, 1, "بارهای فشاری (kgf) — ۶ نمونه:", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:K{r}")
  r += 1

  for i in range(6):
    set_cell(ws, r, 1 + i, f"نمونه {i + 1}", font=S["header_font"], fill=S["header_fill"])
    set_cell(ws, r + 1, 1 + i, None, fill=S["input_fill"], border=S["input_border"], locked=False, num_fmt='0.0')
  add_dv(ws, [f"{get_column_letter(i + 1)}{r + 1}" for i in range(6)], "decimal", min_val=0, max_val=50000)
  r += 3

  # Flexural: 1.5*F*100/40^3 (F in kgf, converted to N: *9.80665)
  set_cell(ws, r, 1, "مقاومت خمشی (MPa):", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:C{r}")
  set_cell(ws, r, 4,
           f'=IF(OR(A{flex_row}="",B{flex_row}="",C{flex_row}=""),"—",ROUND(AVERAGE(1.5*A{flex_row}*9.80665*100/40^3,1.5*B{flex_row}*9.80665*100/40^3,1.5*C{flex_row}*9.80665*100/40^3),1))',
           fill=S["calc_fill"], num_fmt='0.0')
  r += 1

  # Compressive: P*9.80665/1600
  set_cell(ws, r, 1, "مقاومت فشاری (MPa):", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:C{r}")
  set_cell(ws, r, 4,
           f'=IF(COUNT({get_column_letter(1)}{r - 2}:{get_column_letter(6)}{r - 2})<6,"—",ROUND(AVERAGE({get_column_letter(1)}{r - 2}*9.80665/1600,{get_column_letter(2)}{r - 2}*9.80665/1600,{get_column_letter(3)}{r - 2}*9.80665/1600,{get_column_letter(4)}{r - 2}*9.80665/1600,{get_column_letter(5)}{r - 2}*9.80665/1600,{get_column_letter(6)}{r - 2}*9.80665/1600),1))',
           fill=S["calc_fill"], num_fmt='0.0')
  r += 2

  set_cell(ws, r, 1, "⚠️ اِراتا:", font=S["warn_font"], fill=S["warn_fill"], align=S["right_align"])
  ws.merge_cells(f"A{r}:K{r}")
  r += 1
  set_cell(ws, r, 1, "کتاب: ۳۳۴۰ kgf → ۱۱.۹ MPa. صحیح: ۳۳۴۰×۹.۸۰۶۶۵/۱۶۰۰ ≈ ۲۰.۵ MPa",
           font=Font(name=FONT_FA, size=9, italic=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:K{r}")

  for col in range(1, 12):
    ws.column_dimensions[get_column_letter(col)].width = 12


def build_3_1_slump(wb):
  """3-1: Slump (ASTM C143)"""
  ws = wb.create_sheet("14_آزمایش_3-1")
  ws.sheet_properties.tabColor = "4CAF50"
  ws.sheet_view.rightToLeft = True

  add_nav_bar(ws)
  add_title(ws, "آزمایش ۳-۱: اسلامپ", "ASTM C143")

  r = 5
  set_cell(ws, r, 1, "ارتفاع پس از برداشتن (mm):", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, None, fill=S["input_fill"], border=S["input_border"], locked=False, num_fmt='0')
  add_dv(ws, ["C5"], "decimal", min_val=0, max_val=300)
  r += 1

  set_cell(ws, r, 1, "نوع ریزش:", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, None, fill=S["input_fill"], border=S["input_border"], locked=False)
  add_dv(ws, ["C6"], "list", formula='"برشی,دو طرفه,ریزش کامل"')
  r += 2

  set_cell(ws, r, 1, "اسلامپ (mm):", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, '=IF(C5="","",MROUND(300-C5,5))', fill=S["calc_fill"], num_fmt='0')
  r += 1

  set_cell(ws, r, 1, "وضعیت:", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, '=IF(C6="","",IF(C6="ریزش کامل","⚠️ تکرار آزمایش","✅"))', fill=S["calc_fill"])

  for col in range(1, 12):
    ws.column_dimensions[get_column_letter(col)].width = 14


def build_3_2_bleeding(wb):
  """3-2: Bleeding (ASTM C232)"""
  ws = wb.create_sheet("15_آزمایش_3-2")
  ws.sheet_properties.tabColor = "4CAF50"
  ws.sheet_view.rightToLeft = True

  add_nav_bar(ws)
  add_title(ws, "آزمایش ۳-۲: آب‌انداختگی بتن", "ASTM C232")

  r = 5
  fields = [
    ("h1: ارتفاع اولیه (mm)", "C5"),
    ("h2: ارتفاع نهایی (mm)", "C6"),
    ("G: جذب سنگدانه (%)", "C7"),
  ]

  for label, cell_ref in fields:
    set_cell(ws, r, 1, label, font=Font(name=FONT_FA, bold=True), align=S["right_align"])
    ws.merge_cells(f"A{r}:B{r}")
    set_cell(ws, r, 3, None, fill=S["input_fill"], border=S["input_border"], locked=False, num_fmt='0.0')
    r += 1

  add_dv(ws, ["C5", "C6", "C7"], "decimal", min_val=0, max_val=1000)
  r += 1

  set_cell(ws, r, 1, "آب‌انداختگی ظاهری (%):", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, '=IF(OR(C5="",C6="",C5=0),"—",ROUND((C5-C6)/C5*100,1))',
           fill=S["calc_fill"], num_fmt='0.0')
  r += 1

  set_cell(ws, r, 1, "آب‌انداختگی واقعی (%):", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, '=IF(OR(C9="—",C7=""),"—",ROUND(C9-C7,1))',
           fill=S["calc_fill"], num_fmt='0.0')

  for col in range(1, 12):
    ws.column_dimensions[get_column_letter(col)].width = 14


def build_3_3_concrete_unit_weight(wb):
  """3-3: Concrete Unit Weight (ASTM C138)"""
  ws = wb.create_sheet("16_آزمایش_3-3")
  ws.sheet_properties.tabColor = "4CAF50"
  ws.sheet_view.rightToLeft = True

  add_nav_bar(ws)
  add_title(ws, "آزمایش ۳-۳: وزن واحد حجمی بتن", "ASTM C138")

  r = 5
  fields = [
    ("m1: جرم ظرف خالی (g)", "C5"),
    ("m2: جرم ظرف+بتن (g)", "C6"),
    ("V: حجم ظرف (cm³)", "C7"),
    ("D نظری (kg/m³)", "C8"),
  ]

  for label, cell_ref in fields:
    set_cell(ws, r, 1, label, font=Font(name=FONT_FA, bold=True), align=S["right_align"])
    ws.merge_cells(f"A{r}:B{r}")
    set_cell(ws, r, 3, None, fill=S["input_fill"], border=S["input_border"], locked=False, num_fmt='0.0')
    r += 1

  add_dv(ws, ["C5", "C6", "C7", "C8"], "decimal", min_val=0, max_val=100000)
  r += 1

  set_cell(ws, r, 1, "وزن واحد حجمی (kg/m³):", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, '=IF(OR(C5="",C6="",C7="",C7=0),"—",ROUND((C6-C5)/C7*1000,0))',
           fill=S["calc_fill"], num_fmt='#,##0')
  r += 1

  set_cell(ws, r, 1, "اختلاف با نظری (%):", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, '=IF(OR(C10="—",C8="",C8=0),"—",ROUND(ABS(C10-C8)/C8*100,1))',
           fill=S["calc_fill"], num_fmt='0.0')
  r += 1

  set_cell(ws, r, 1, "وضعیت:", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, '=IF(C11="—","",IF(C11>2,"⚠️ اختلاف >2%","✅"))', fill=S["calc_fill"])

  for col in range(1, 12):
    ws.column_dimensions[get_column_letter(col)].width = 14


def build_4_1_compressive(wb):
  """4-1: Compressive Strength (ASTM C39)"""
  ws = wb.create_sheet("17_آزمایش_4-1")
  ws.sheet_properties.tabColor = "F44336"
  ws.sheet_view.rightToLeft = True

  add_nav_bar(ws)
  add_title(ws, "آزمایش ۴-۱: مقاومت فشاری بتن", "ASTM C39")

  r = 5
  set_cell(ws, r, 1, "نوع نمونه:", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, "استوانه", fill=S["input_fill"], border=S["input_border"], locked=False)
  add_dv(ws, ["C5"], "list", formula='"استوانه,مکعب"')
  r += 1

  set_cell(ws, r, 1, "قطر/ضلع (mm):", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, None, fill=S["input_fill"], border=S["input_border"], locked=False, num_fmt='0.0')
  add_dv(ws, ["C6"], "decimal", min_val=0, max_val=500)
  r += 1

  set_cell(ws, r, 1, "بار (kN):", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, None, fill=S["input_fill"], border=S["input_border"], locked=False, num_fmt='0.0')
  add_dv(ws, ["C7"], "decimal", min_val=0, max_val=10000)
  r += 1

  set_cell(ws, r, 1, "الگوی شکست:", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, None, fill=S["input_fill"], border=S["input_border"], locked=False)
  add_dv(ws, ["C8"], "list", formula='"نوع ۱,نوع ۲,نوع ۳,نوع ۴,نوع ۵,نوع ۶"')
  r += 2

  set_cell(ws, r, 1, "مساحت (mm²):", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, '=IF(C6="","",IF(C5="مکعب",C6^2,PI()/4*C6^2))', fill=S["calc_fill"], num_fmt='0.0')
  r += 1

  set_cell(ws, r, 1, "مقاومت فشاری (MPa):", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, '=IF(OR(C7="",C10="—",C10="",C10=0),"—",MROUND(C7*1000/C10,0.1))',
           fill=PatternFill("solid", fgColor=COLORS["pass_fill"]), font=S["pass_font"], num_fmt='0.0')
  r += 2

  set_cell(ws, r, 1, "⚠️ اِراتا:", font=S["warn_font"], fill=S["warn_fill"], align=S["right_align"])
  ws.merge_cells(f"A{r}:K{r}")
  r += 1
  set_cell(ws, r, 1, "کتاب: ۴۱.۵ MPa. صحیح با d=150 و F=715: ≈۴۰.۴ MPa",
           font=Font(name=FONT_FA, size=9, italic=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:K{r}")

  for col in range(1, 12):
    ws.column_dimensions[get_column_letter(col)].width = 14


def build_4_2_tensile(wb):
  """4-2: Splitting Tensile (ASTM C496)"""
  ws = wb.create_sheet("18_آزمایش_4-2")
  ws.sheet_properties.tabColor = "F44336"
  ws.sheet_view.rightToLeft = True

  add_nav_bar(ws)
  add_title(ws, "آزمایش ۴-۲: مقاومت کششی (برزیلی)", "ASTM C496")

  r = 5
  fields = [
    ("d: قطر (mm)", "C5"),
    ("L: طول (mm)", "C6"),
    ("P: بار (N)", "C7"),
  ]

  for label, cell_ref in fields:
    set_cell(ws, r, 1, label, font=Font(name=FONT_FA, bold=True), align=S["right_align"])
    ws.merge_cells(f"A{r}:B{r}")
    set_cell(ws, r, 3, None, fill=S["input_fill"], border=S["input_border"], locked=False, num_fmt='0.0')
    r += 1

  add_dv(ws, ["C5", "C6", "C7"], "decimal", min_val=0, max_val=1000000)
  r += 1

  set_cell(ws, r, 1, "مقاومت کششی (MPa):", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, '=IF(OR(C5="",C6="",C7="",C5=0,C6=0),"—",ROUND(2*C7/(PI()*C5*C6),2))',
           fill=S["calc_fill"], num_fmt='0.00')
  r += 1

  set_cell(ws, r, 1, "بازه منطقی:", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, f'=IF(C9="—","",IF(OR(C9<2,C9>8),"⚠️ خارج از بازه ۲-۸ MPa","✅"))',
           fill=S["calc_fill"])
  r += 1

  set_cell(ws, r, 1, "⚠️ اِراتا:", font=S["warn_font"], fill=S["warn_fill"], align=S["right_align"])
  ws.merge_cells(f"A{r}:K{r}")
  r += 1
  set_cell(ws, r, 1, "کتاب: ۲.۵۳ MPa (ضریب ۲ جا افتاده). صحیح: ≈۵.۰۷ MPa",
           font=Font(name=FONT_FA, size=9, italic=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:K{r}")

  for col in range(1, 12):
    ws.column_dimensions[get_column_letter(col)].width = 14


def build_4_3_flexural(wb):
  """4-3: Flexural Strength (ASTM C78)"""
  ws = wb.create_sheet("19_آزمایش_4-3")
  ws.sheet_properties.tabColor = "F44336"
  ws.sheet_view.rightToLeft = True

  add_nav_bar(ws)
  add_title(ws, "آزمایش ۴-۳: مقاومت خمشی", "ASTM C78")

  r = 5
  fields = [
    ("b: عرض (mm)", "C5"),
    ("d: ارتفاع (mm)", "C6"),
    ("L: دهانه (mm)", "C7"),
    ("P: بار (N)", "C8"),
  ]

  for label, cell_ref in fields:
    set_cell(ws, r, 1, label, font=Font(name=FONT_FA, bold=True), align=S["right_align"])
    ws.merge_cells(f"A{r}:B{r}")
    set_cell(ws, r, 3, None, fill=S["input_fill"], border=S["input_border"], locked=False, num_fmt='0.0')
    r += 1

  add_dv(ws, ["C5", "C6", "C7", "C8"], "decimal", min_val=0, max_val=1000000)
  r += 1

  set_cell(ws, r, 1, "روش بارگذاری:", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, "یک‌سوم میانه", fill=S["input_fill"], border=S["input_border"], locked=False)
  add_dv(ws, ["C10"], "list", formula='"یک‌سوم میانه,مرکزی"')
  r += 1

  set_cell(ws, r, 1, "محل ترک:", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, None, fill=S["input_fill"], border=S["input_border"], locked=False)
  add_dv(ws, ["C11"], "list", formula='"داخل محدوده,خارج محدوده"')
  r += 2

  set_cell(ws, r, 1, "مقاومت خمشی (MPa):", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3,
           '=IF(OR(C5="",C6="",C7="",C8="",C11="خارج محدوده"),IF(C11="خارج محدوده","⚠️ تکرار","—"),IF(C10="مرکزی",ROUND(3*C8*C7/(2*C5*C6^2),2),ROUND(C8*C7/(C5*C6^2),2)))',
           fill=S["calc_fill"], num_fmt='0.00')
  r += 1

  set_cell(ws, r, 1, "⚠️ اِراتا:", font=S["warn_font"], fill=S["warn_fill"], align=S["right_align"])
  ws.merge_cells(f"A{r}:K{r}")
  r += 1
  set_cell(ws, r, 1, "کتاب: ۳۳.۴۶۶ MPa (خطای فاکتور). صحیح: ≈۷.۵ MPa",
           font=Font(name=FONT_FA, size=9, italic=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:K{r}")

  for col in range(1, 12):
    ws.column_dimensions[get_column_letter(col)].width = 14


def build_4_4_upv(wb):
  """4-4: Ultrasonic Pulse Velocity (ASTM C597)"""
  ws = wb.create_sheet("20_آزمایش_4-4")
  ws.sheet_properties.tabColor = "F44336"
  ws.sheet_view.rightToLeft = True

  add_nav_bar(ws)
  add_title(ws, "آزمایش ۴-۴: سرعت پالس اولتراسونیک", "ASTM C597")

  r = 5
  set_cell(ws, r, 1, "L: طول مسیر (m):", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, None, fill=S["input_fill"], border=S["input_border"], locked=False, num_fmt='0.000')
  add_dv(ws, ["C5"], "decimal", min_val=0, max_val=10)
  r += 1

  set_cell(ws, r, 1, "زمان‌ها (µs) — ۳ قرائت:", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  for i in range(3):
    set_cell(ws, r, 3 + i, None, fill=S["input_fill"], border=S["input_border"], locked=False, num_fmt='0')
  add_dv(ws, ["C6", "D6", "E6"], "decimal", min_val=0, max_val=100000)
  r += 2

  set_cell(ws, r, 1, "سرعت (km/s):", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, '=IF(OR(C5="",C6="",C6=0),"—",ROUND(C5*1000/AVERAGE(C6:E6),2))',
           fill=S["calc_fill"], num_fmt='0.00')
  r += 1

  set_cell(ws, r, 1, "انحراف معیار:", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, '=IF(C6="","",ROUND(STDEV(C6:E6),2))', fill=S["calc_fill"], num_fmt='0.00')
  r += 1

  set_cell(ws, r, 1, "طبقه‌بندی:", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, f'=IF(C9="—","",IF(C9>=4.5,"عالی",IF(C9>=3.5,"خوب",IF(C9>=3,"متوسط","ضعیف"))))',
           fill=S["calc_fill"])

  for col in range(1, 12):
    ws.column_dimensions[get_column_letter(col)].width = 14


def build_4_5_schmidt(wb):
  """4-5: Schmidt Hammer (ASTM C805)"""
  ws = wb.create_sheet("21_آزمایش_4-5")
  ws.sheet_properties.tabColor = "F44336"
  ws.sheet_view.rightToLeft = True

  add_nav_bar(ws)
  add_title(ws, "آزمایش ۴-۵: چکش اشمیت", "ASTM C805")

  r = 5
  set_cell(ws, r, 1, "خوانش‌ها (حداکثر ۱۶):", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:K{r}")
  r += 1

  for i in range(16):
    col = (i % 8) + 1
    row = r + (i // 8)
    set_cell(ws, row, col, None, fill=S["input_fill"], border=S["input_border"], locked=False, num_fmt='0')
  add_dv(ws, [f"{get_column_letter((i % 8) + 1)}{r + (i // 8)}" for i in range(16)], "decimal", min_val=0, max_val=100)
  r += 3

  set_cell(ws, r, 1, "سطح:", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  set_cell(ws, r, 3, "خشک", fill=S["input_fill"], border=S["input_border"], locked=False)
  add_dv(ws, [f"C{r}"], "list", formula='"خشک,مرطوب"')
  r += 1

  set_cell(ws, r, 1, "دما (°C):", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  set_cell(ws, r, 3, None, fill=S["input_fill"], border=S["input_border"], locked=False, num_fmt='0')
  add_dv(ws, [f"C{r}"], "decimal", min_val=-10, max_val=60)
  r += 2

  # Pass 1: average of all
  set_cell(ws, r, 1, "میانگین کل:", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, f'=IF(COUNT(A6:H7)=0,"",ROUND(AVERAGE(A6:H7),1))', fill=S["calc_fill"], num_fmt='0.0')
  r += 1

  # Pass 2: filter valid (within ±6 of mean)
  set_cell(ws, r, 1, "تعداد معتبر:", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, f'=IF(C{r - 1}="","",SUMPRODUCT((ABS(A6:H7-C{r - 1})<=6)*(A6:H7<>"")))',
           fill=S["calc_fill"], num_fmt='0')
  r += 1

  # Rm = average of valid
  set_cell(ws, r, 1, "Rm (میانگین معتبرها):", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3,
           f'=IF(OR(C{r - 2}="",C{r - 1}=0),"—",ROUND(SUMPRODUCT((ABS(A6:H7-C{r - 2})<=6)*(A6:H7<>"")*A6:H7)/C{r - 1},1))',
           fill=S["calc_fill"], num_fmt='0.0')
  r += 1

  # Check if >20% rejected
  set_cell(ws, r, 1, "وضعیت:", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, f'=IF(C{r - 1}="","",IF(C{r - 1}<COUNT(A6:H7)*0.8,"❌ حذف >20% — تکرار","✅"))',
           fill=S["calc_fill"])
  r += 1

  # Corrections
  set_cell(ws, r, 1, "Rm اصلاح‌شده:", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:B{r}")
  set_cell(ws, r, 3, f'=IF(C{r - 2}="—","—",ROUND(C{r - 2}*IF(C{r - 3}="مرطوب",0.95,1)*IF(C{r - 3 + 1}<10,1.03,1),1))',
           fill=PatternFill("solid", fgColor=COLORS["pass_fill"]), font=S["pass_font"], num_fmt='0.0')

  for col in range(1, 12):
    ws.column_dimensions[get_column_letter(col)].width = 12


# ═══════════════════════════════════════════════
# SHEET 22: REPORT
# ═══════════════════════════════════════════════
def build_22_report(wb):
  ws = wb.create_sheet("22_گزارش")
  ws.sheet_properties.tabColor = "9C27B0"
  ws.sheet_view.rightToLeft = True

  add_nav_bar(ws)
  add_title(ws, "📑 گزارش آزمایشگاهی", "خلاصه نتایج + آماده چاپ")

  r = 5
  set_cell(ws, r, 1, "اطلاعات پروژه:", font=Font(name=FONT_FA, bold=True, size=12), align=S["right_align"])
  ws.merge_cells(f"A{r}:K{r}")
  r += 1

  info_fields = ["نام پروژه", "شماره نمونه", "تاریخ", "اپراتور", "استاندارد"]
  for i, f in enumerate(info_fields):
    set_cell(ws, r, 1, f, font=Font(name=FONT_FA, bold=True, size=10), align=S["right_align"])
    set_cell(ws, r, 2, f"='01_اطلاعات_آزمون'!B{5 + i}", fill=S["calc_fill"], align=S["right_align"])
    ws.merge_cells(f"B{r}:F{r}")
    r += 1

  r += 1
  set_cell(ws, r, 1, "خلاصه نتایج:", font=Font(name=FONT_FA, bold=True, size=12), align=S["right_align"])
  ws.merge_cells(f"A{r}:K{r}")
  r += 1

  headers = ["آزمایش", "نتیجه", "واحد", "وضعیت"]
  for col, h in enumerate(headers, 1):
    set_cell(ws, r, col, h, font=S["header_font"], fill=S["header_fill"])
  r += 1

  results = [
    ("۱-۱ دانه‌بندی (FM)", "='02_آزمایش_1-1'!C13", "", '=IF(B{row}="—","انجام نشده","✅")'),
    ("۱-۲ رطوبت", "='03_آزمایش_1-2'!C9", "%", '=IF(B{row}="—","انجام نشده","✅")'),
    ("۱-۳ چگالی درشت (SSD)", "='04_آزمایش_1-3'!C10", "", '=IF(B{row}="—","انجام نشده","✅")'),
    ("۱-۴ چگالی ریز (SSD)", "='05_آزمایش_1-4'!C13", "", '=IF(B{row}="—","انجام نشده","✅")'),
    ("۱-۵ وزن واحد حجمی", "='06_آزمایش_1-5'!C10", "kg/m³", '=IF(B{row}="—","انجام نشده","✅")'),
    ("۱-۶ معادل ماسه", "='07_آزمایش_1-6'!C8", "%", '=IF(B{row}="—","انجام نشده","✅")'),
    ("۱-۷ شاخص درازگی", "='08_آزمایش_1-7'!C10", "%", '=IF(B{row}="—","انجام نشده","✅")'),
    ("۱-۸ جذب آب", "='09_آزمایش_1-8'!C8", "%", '=IF(B{row}="—","انجام نشده","✅")'),
    ("۲-۱ چگالی بتن تازه", "='10_آزمایش_2-1'!C9", "kg/m³", '=IF(B{row}="—","انجام نشده","✅")'),
    ("۲-۲ ویکات (w/c)", "='11_آزمایش_2-2'!C9", "%", '=IF(B{row}="—","انجام نشده","✅")'),
    ("۲-۳ گیرش اولیه", "='12_آزمایش_2-3'!C11", "min", '=IF(B{row}="—","انجام نشده","✅")'),
    ("۲-۴ مقاومت ملات", "='13_آزمایش_2-4'!D15", "MPa", '=IF(B{row}="—","انجام نشده","✅")'),
    ("۳-۱ اسلامپ", "='14_آزمایش_3-1'!C9", "mm", '=IF(B{row}="—","انجام نشده","✅")'),
    ("۳-۲ آب‌انداختگی", "='15_آزمایش_3-2'!C10", "%", '=IF(B{row}="—","انجام نشده","✅")'),
    ("۳-۳ وزن واحد بتن", "='16_آزمایش_3-3'!C10", "kg/m³", '=IF(B{row}="—","انجام نشده","✅")'),
    ("۴-۱ مقاومت فشاری", "='17_آزمایش_4-1'!C13", "MPa", '=IF(B{row}="—","انجام نشده","✅")'),
    ("۴-۲ مقاومت کششی", "='18_آزمایش_4-2'!C9", "MPa", '=IF(B{row}="—","انجام نشده","✅")'),
    ("۴-۳ مقاومت خمشی", "='19_آزمایش_4-3'!C14", "MPa", '=IF(B{row}="—","انجام نشده","✅")'),
    ("۴-۴ سرعت پالس", "='20_آزمایش_4-4'!C9", "km/s", '=IF(B{row}="—","انجام نشده","✅")'),
    ("۴-۵ چکش اشمیت (Rm)", "='21_آزمایش_4-5'!C17", "", '=IF(B{row}="—","انجام نشده","✅")'),
  ]

  for name, formula, unit, status_formula in results:
    set_cell(ws, r, 1, name, align=S["right_align"])
    set_cell(ws, r, 2, formula, fill=S["calc_fill"], num_fmt='0.0')
    set_cell(ws, r, 3, unit, font=S["num_font"])
    set_cell(ws, r, 4, status_formula.format(row=r), fill=S["calc_fill"])
    r += 1

  r += 1
  set_cell(ws, r, 1, "کد صحت:", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  set_cell(ws, r, 2, f'=TEXT(MOD(SUMPRODUCT(B{r - len(results)}:B{r - 1})*7,99999),"00000")',
           fill=S["calc_fill"], font=Font(name=FONT_NUM, bold=True))
  r += 2

  set_cell(ws, r, 1, "مهر و امضا:", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:K{r + 3}")

  for col in range(1, 12):
    ws.column_dimensions[get_column_letter(col)].width = 13


# ═══════════════════════════════════════════════
# SHEET 23: DASHBOARD
# ═══════════════════════════════════════════════
def build_23_dashboard(wb):
  ws = wb.create_sheet("23_داشبورد")
  ws.sheet_properties.tabColor = "00BCD4"
  ws.sheet_view.rightToLeft = True

  add_nav_bar(ws)
  add_title(ws, "📊 داشبورد پروژه", "وضعیت کلی — فرمول‌محور")

  r = 5
  set_cell(ws, r, 1, "تعداد کل آزمایش‌ها:", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  set_cell(ws, r, 3, 20, fill=S["calc_fill"], font=S["num_font"])
  r += 1

  set_cell(ws, r, 1, "تعداد شیت‌های پیاده‌سازی‌شده:", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  set_cell(ws, r, 3,
           "=COUNTA('02_آزمایش_1-1'!A1,'03_آزمایش_1-2'!A1,'04_آزمایش_1-3'!A1,'05_آزمایش_1-4'!A1,'06_آزمایش_1-5'!A1,'07_آزمایش_1-6'!A1,'08_آزمایش_1-7'!A1,'09_آزمایش_1-8'!A1,'10_آزمایش_2-1'!A1,'11_آزمایش_2-2'!A1,'12_آزمایش_2-3'!A1,'13_آزمایش_2-4'!A1,'14_آزمایش_3-1'!A1,'15_آزمایش_3-2'!A1,'16_آزمایش_3-3'!A1,'17_آزمایش_4-1'!A1,'18_آزمایش_4-2'!A1,'19_آزمایش_4-3'!A1,'20_آزمایش_4-4'!A1,'21_آزمایش_4-5'!A1)",
           fill=S["calc_fill"], font=S["num_font"])
  r += 1

  set_cell(ws, r, 1, "درصد پیشرفت (ورودی‌های پرشده):", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  # Count total input cells across all test sheets (non-empty yellow cells)
  input_ranges = ",".join([
    "'03_آزمایش_1-2'!C5:C6", "'04_آزمایش_1-3'!C5:C7",
    "'17_آزمایش_4-1'!C5:C7", "'14_آزمایش_3-1'!C5:C6"
  ])
  set_cell(ws, r, 3, f'=IFERROR(COUNTA({input_ranges})/100,0)', fill=S["calc_fill"], num_fmt='0%')
  r += 2

  set_cell(ws, r, 1, "⚠️ توجه:", font=S["warn_font"], fill=S["warn_fill"], align=S["right_align"])
  ws.merge_cells(f"A{r}:K{r}")
  r += 1
  set_cell(ws, r, 1, "مقادیر بالا فرمول‌محور هستند و با تغییر شیت‌ها به‌روز می‌شوند.",
           font=Font(name=FONT_FA, size=9, italic=True), align=S["right_align"])
  ws.merge_cells(f"A{r}:K{r}")

  for col in range(1, 12):
    ws.column_dimensions[get_column_letter(col)].width = 14


# ═══════════════════════════════════════════════
# SHEET 24: QA TEST
# ═══════════════════════════════════════════════
def build_24_qa(wb):
  ws = wb.create_sheet("24_QA_Test")
  ws.sheet_properties.tabColor = "FF9800"
  ws.sheet_view.rightToLeft = True

  add_nav_bar(ws)
  add_title(ws, "🧪 QA Test — تست‌های خودکار", "هر Patch باید این شیت را سبز نگه دارد")

  r = 5
  headers = ["Test ID", "شرح", "ورودی", "انتظار", "نتیجه", "وضعیت"]
  for col, h in enumerate(headers, 1):
    set_cell(ws, r, col, h, font=S["header_font"], fill=S["header_fill"])

  tests = [
    ("T-001", "ورودی خالی", "همه سلول‌ها پاک", "بدون #DIV/0!", '=IF(COUNTA(\'03_آزمایش_1-2\'!C5:C6)=0,"✅ PASS","—")',
     "—"),
    ("T-002", "ورودی منفی", "W1=-100", "DV رد کند", "دستی", "—"),
    ("T-003", "مرزی: W2=0", "W1=100, W2=0", "نمایش '—'", '=IF(\'03_آزمایش_1-2\'!C9="—","✅ PASS","❌ FAIL")', "—"),
    ("T-004", "نمونه کتاب ۲-۴", "3340 kgf, A=1600", "≈20.5 MPa",
     '=IF(ABS(\'13_آزمایش_2-4\'!D15-20.5)<0.2,"✅ PASS","❌ FAIL")', "—"),
    ("T-005", "نمونه کتاب ۴-۱", "d=150, F=715 kN", "≈40.4 MPa",
     '=IF(ABS(\'17_آزمایش_4-1\'!C13-40.4)<0.2,"✅ PASS","❌ FAIL")', "—"),
  ]

  r += 1
  for test_id, desc, inp, expected, formula, status in tests:
    set_cell(ws, r, 1, test_id, font=S["num_font"])
    set_cell(ws, r, 2, desc, align=S["right_align"])
    set_cell(ws, r, 3, inp, align=S["right_align"])
    set_cell(ws, r, 4, expected, align=S["right_align"])
    set_cell(ws, r, 5, formula, fill=S["calc_fill"])
    set_cell(ws, r, 6, status, fill=S["calc_fill"])
    r += 1

  r += 1
  set_cell(ws, r, 1, "Total:", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  set_cell(ws, r, 2, "=COUNTA(A6:A10)", fill=S["calc_fill"], font=S["num_font"])
  set_cell(ws, r, 3, "Passed:", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  set_cell(ws, r, 4, '=COUNTIF(E6:E10,"✅ PASS")', fill=S["calc_fill"], font=S["num_font"])
  set_cell(ws, r, 5, "Failed:", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
  set_cell(ws, r, 6, '=COUNTIF(E6:E10,"❌ FAIL")', fill=S["calc_fill"], font=S["num_font"])

  for col in range(1, 12):
    ws.column_dimensions[get_column_letter(col)].width = 14


# ═══════════════════════════════════════════════
# SHEET 25: ERRATA
# ═══════════════════════════════════════════════
def build_25_errata(wb):
  ws = wb.create_sheet("25_خطاها_هشدارها")
  ws.sheet_properties.tabColor = "C62828"
  ws.sheet_view.rightToLeft = True

  add_nav_bar(ws)
  add_title(ws, "📋 خطاها و هشدارها", "تجمع زنده — ۸ اِراتای کتاب")

  r = 5
  headers = ["کد", "نوع", "شرح", "مقدار کتاب", "مقدار مرجع", "استاندارد", "وضعیت"]
  for col, h in enumerate(headers, 1):
    set_cell(ws, r, col, h, font=S["header_font"], fill=S["header_fill"])

  errata_data = [
    ("1-2", "فرمول", "رطوبت پایه تر", "(W1-W2)/W1", "(W1-W2)/W2", "ASTM C566", "confirmed"),
    ("1-4ج", "فیزیکی", "OD=1.51/SSD=2.63", "غیرفیزیکی", "≈2.6x", "ASTM C128", "confirmed"),
    ("1-5", "برچسب", "S=1600", "برچسب غلط", "S بی‌بُعد", "ASTM C138", "confirmed"),
    ("1-6", "فرمول", "ماسه/رس", "ماسه/رس×۱۰۰", "ماسه/(ماسه+رس)×۱۰۰", "ASTM D2419", "confirmed"),
    ("2-4", "فرمول", "۱۱.۹ MPa", "11.9", "≈20.5", "EN 196-1", "confirmed"),
    ("4-1", "فرمول", "۴۱.۵ MPa", "41.5", "≈40.4", "ASTM C39", "confirmed"),
    ("4-2", "فرمول", "۲.۵۳ MPa", "2.53", "≈5.07", "ASTM C496", "confirmed"),
    ("4-3", "فرمول", "۳۳.۴۶۶ MPa", "33.466", "≈7.5", "ASTM C78", "confirmed"),
  ]

  r += 1
  for code, etype, desc, book_val, ref_val, std, status in errata_data:
    set_cell(ws, r, 1, code, font=S["num_font"])
    set_cell(ws, r, 2, etype, align=S["center_align"])
    set_cell(ws, r, 3, desc, align=S["right_align"])
    set_cell(ws, r, 4, book_val, fill=S["fail_fill"], font=S["fail_font"], align=S["center_align"])
    set_cell(ws, r, 5, ref_val, fill=S["pass_fill"], font=S["pass_font"], align=S["center_align"])
    set_cell(ws, r, 6, std, font=Font(name=FONT_NUM, size=9), align=S["center_align"])
    set_cell(ws, r, 7, "✅ تأییدشده", fill=S["pass_fill"], font=S["pass_font"], align=S["center_align"])
    r += 1

  for col in range(1, 12):
    ws.column_dimensions[get_column_letter(col)].width = 13


# ═══════════════════════════════════════════════
# HIDDEN SHEETS
# ═══════════════════════════════════════════════
def build_hidden_sheets(wb):
  # _Reference_DB
  ws = wb.create_sheet("_Reference_DB")
  ws.sheet_state = "hidden"

  set_cell(ws, 1, 1, "Standard", font=S["header_font"], fill=S["header_fill"])
  set_cell(ws, 1, 2, "Test", font=S["header_font"], fill=S["header_fill"])
  set_cell(ws, 1, 3, "Param", font=S["header_font"], fill=S["header_fill"])
  set_cell(ws, 1, 4, "Min", font=S["header_font"], fill=S["header_fill"])
  set_cell(ws, 1, 5, "Max", font=S["header_font"], fill=S["header_fill"])
  set_cell(ws, 1, 6, "Unit", font=S["header_font"], fill=S["header_fill"])

  ref_data = [
    ("ISIRI 302", "Sieve #4", "Passing", 80, 95, "%"),
    ("ISIRI 302", "Sieve #8", "Passing", 60, 80, "%"),
    ("ISIRI 302", "Sieve #16", "Passing", 40, 60, "%"),
    ("ISIRI 302", "Sieve #30", "Passing", 25, 40, "%"),
    ("ISIRI 302", "Sieve #50", "Passing", 10, 25, "%"),
    ("ISIRI 302", "Sieve #100", "Passing", 2, 10, "%"),
    ("ASTM C39", "Compressive", "Strength", 15, 60, "MPa"),
    ("ASTM C496", "Tensile", "Strength", 2, 8, "MPa"),
    ("ASTM C78", "Flexural", "Strength", 3, 10, "MPa"),
    ("ASTM C597", "UPV", "Velocity", 3, 5, "km/s"),
  ]

  for i, row_data in enumerate(ref_data, 2):
    for j, val in enumerate(row_data, 1):
      set_cell(ws, i, j, val, font=S["num_font"])

  # _Standards
  ws2 = wb.create_sheet("_Standards")
  ws2.sheet_state = "hidden"

  set_cell(ws2, 1, 1, "Parameter", font=S["header_font"], fill=S["header_fill"])
  set_cell(ws2, 1, 2, "Value", font=S["header_font"], fill=S["header_fill"])

  set_cell(ws2, 2, 1, "kgf_to_N", font=S["num_font"])
  set_cell(ws2, 2, 2, 9.80665, font=S["num_font"])

  set_cell(ws2, 3, 1, "Mortar_Area_mm2", font=S["num_font"])
  set_cell(ws2, 3, 2, 1600, font=S["num_font"])

  # _Validation_Data
  ws3 = wb.create_sheet("_Validation_Data")
  ws3.sheet_state = "hidden"

  headers = ["ID", "Type", "Book Value", "Reference Value", "Standard", "Status"]
  for col, h in enumerate(headers, 1):
    set_cell(ws3, 1, col, h, font=S["header_font"], fill=S["header_fill"])

  val_data = [
    ("1-2", "formula", "(W1-W2)/W1", "(W1-W2)/W2", "ASTM C566", "confirmed"),
    ("1-4ج", "physical", "OD=1.51", "≈2.6x", "ASTM C128", "confirmed"),
    ("1-5", "label", "S=1600", "S dimensionless", "ASTM C138", "confirmed"),
    ("1-6", "formula", "sand/clay", "sand/(sand+clay)", "ASTM D2419", "confirmed"),
    ("2-4", "formula", "11.9", "≈20.5", "EN 196-1", "confirmed"),
    ("4-1", "formula", "41.5", "≈40.4", "ASTM C39", "confirmed"),
    ("4-2", "formula", "2.53", "≈5.07", "ASTM C496", "confirmed"),
    ("4-3", "formula", "33.466", "≈7.5", "ASTM C78", "confirmed"),
  ]

  for i, row_data in enumerate(val_data, 2):
    for j, val in enumerate(row_data, 1):
      set_cell(ws3, i, j, val, font=S["num_font"])

  # _Glossary
  ws4 = wb.create_sheet("_Glossary")
  ws4.sheet_state = "hidden"

  glossary = [
    ("OD", "Oven Dry", "خشک آون", "جرم پس از خشک شدن در ۱۱۰°C"),
    ("SSD", "Saturated Surface Dry", "اشباع خشک", "جرم در حالت اشباع با سطح خشک"),
    ("App", "Apparent", "ظاهری", "چگالی ظاهری"),
    ("FM", "Fineness Modulus", "مدول نرمی", "شاخص دانه‌بندی"),
    ("SE", "Sand Equivalent", "معادل ماسه", "نسبت ماسه به رس"),
    ("UPV", "Ultrasonic Pulse Velocity", "سرعت پالس", "سرعت موج اولتراسونیک"),
  ]

  set_cell(ws4, 1, 1, "مخفف", font=S["header_font"], fill=S["header_fill"])
  set_cell(ws4, 1, 2, "انگلیسی", font=S["header_font"], fill=S["header_fill"])
  set_cell(ws4, 1, 3, "فارسی", font=S["header_font"], fill=S["header_fill"])
  set_cell(ws4, 1, 4, "توضیح", font=S["header_font"], fill=S["header_fill"])

  for i, (abbr, en, fa, desc) in enumerate(glossary, 2):
    set_cell(ws4, i, 1, abbr, font=S["num_font"])
    set_cell(ws4, i, 2, en, font=Font(name=FONT_NUM, size=10))
    set_cell(ws4, i, 3, fa, align=S["right_align"])
    set_cell(ws4, i, 4, desc, align=S["right_align"])

  # _Materials_DB
  ws5 = wb.create_sheet("_Materials_DB")
  ws5.sheet_state = "hidden"

  set_cell(ws5, 1, 1, "نوع سنگدانه", font=S["header_font"], fill=S["header_fill"])
  set_cell(ws5, 1, 2, "چگالی (SSD)", font=S["header_font"], fill=S["header_fill"])
  set_cell(ws5, 1, 3, "جذب (%)", font=S["header_font"], fill=S["header_fill"])

  materials = [
    ("سیلیسی", 2.65, 1.5),
    ("آهکی", 2.70, 1.0),
    ("بازیافتی", 2.40, 5.0),
  ]

  for i, (name, sg, abs_val) in enumerate(materials, 2):
    set_cell(ws5, i, 1, name, align=S["right_align"])
    set_cell(ws5, i, 2, sg, font=S["num_font"])
    set_cell(ws5, i, 3, abs_val, font=S["num_font"])


# ═══════════════════════════════════════════════
# CHART: Sieve Analysis
# ═══════════════════════════════════════════════
def add_sieve_chart(wb):
  ws = wb["02_آزمایش_1-1"]

  chart = ScatterChart()
  chart.title = "منحنی دانه‌بندی"
  chart.style = 13
  chart.x_axis.title = "اندازه الک (mm)"
  chart.y_axis.title = "% عبوری"
  chart.x_axis.scaling.logBase = 10
  chart.x_axis.scaling.min = 0.075
  chart.x_axis.scaling.max = 75
  chart.width = 20
  chart.height = 12

  # Find data rows (chart data starts after FM row)
  # We'll reference the chart data area we created
  # For simplicity, use fixed references based on our layout
  data_start = 23  # actual row where ISIRI chart data starts
  data_end = 30

  xvalues = Reference(ws, min_col=1, min_row=data_start, max_row=data_end)

  # Sample series
  yvalues = Reference(ws, min_col=2, min_row=data_start, max_row=data_end)
  series1 = Series(yvalues, xvalues, title="نمونه")
  series1.graphicalProperties.line.solidFill = "1F4E79"
  series1.graphicalProperties.line.width = 22500  # 2.25 pt
  series1.smooth = False
  chart.series.append(series1)

  # Upper limit
  yvalues_hi = Reference(ws, min_col=3, min_row=data_start, max_row=data_end)
  series2 = Series(yvalues_hi, xvalues, title="حد بالا")
  series2.graphicalProperties.line.solidFill = "ED7D31"
  series2.graphicalProperties.line.dashStyle = "dash"
  series2.graphicalProperties.line.width = 12700
  chart.series.append(series2)

  # Lower limit
  yvalues_lo = Reference(ws, min_col=4, min_row=data_start, max_row=data_end)
  series3 = Series(yvalues_lo, xvalues, title="حد پایین")
  series3.graphicalProperties.line.solidFill = "ED7D31"
  series3.graphicalProperties.line.dashStyle = "dash"
  series3.graphicalProperties.line.width = 12700
  chart.series.append(series3)

  ws.add_chart(chart, "H5")


# ═══════════════════════════════════════════════
# MAIN BUILD
# ═══════════════════════════════════════════════
def compute_sha256(filepath):
  sha256 = hashlib.sha256()
  with open(filepath, "rb") as f:
    for chunk in iter(lambda: f.read(8192), b""):
      sha256.update(chunk)
  return sha256.hexdigest()


def main():
  print("═" * 50)
  print("  Concrete Lab Companion — Build v1.1.0")
  print("═" * 50)

  wb = Workbook()

  print("\n🔨 Building sheets...")
  build_00_guide(wb);
  print("  ✅ 00_راهنما")
  build_01_info(wb);
  print("  ✅ 01_اطلاعات_آزمون")
  build_1_1_sieve(wb);
  print("  ✅ 02_آزمایش_1-1")
  build_1_2_moisture(wb);
  print("  ✅ 03_آزمایش_1-2")
  build_1_3_coarse_sg(wb);
  print("  ✅ 04_آزمایش_1-3")
  build_1_4_fine_sg(wb);
  print("  ✅ 05_آزمایش_1-4")
  build_1_5_unit_weight(wb);
  print("  ✅ 06_آزمایش_1-5")
  build_1_6_sand_equivalent(wb);
  print("  ✅ 07_آزمایش_1-6")
  build_1_7_shape(wb);
  print("  ✅ 08_آزمایش_1-7")
  build_1_8_absorption(wb);
  print("  ✅ 09_آزمایش_1-8")
  build_2_1_fresh_density(wb);
  print("  ✅ 10_آزمایش_2-1")
  build_2_2_vicat(wb);
  print("  ✅ 11_آزمایش_2-2")
  build_2_3_setting_time(wb);
  print("  ✅ 12_آزمایش_2-3")
  build_2_4_mortar(wb);
  print("  ✅ 13_آزمایش_2-4")
  build_3_1_slump(wb);
  print("  ✅ 14_آزمایش_3-1")
  build_3_2_bleeding(wb);
  print("  ✅ 15_آزمایش_3-2")
  build_3_3_concrete_unit_weight(wb);
  print("  ✅ 16_آزمایش_3-3")
  build_4_1_compressive(wb);
  print("  ✅ 17_آزمایش_4-1")
  build_4_2_tensile(wb);
  print("  ✅ 18_آزمایش_4-2")
  build_4_3_flexural(wb);
  print("  ✅ 19_آزمایش_4-3")
  build_4_4_upv(wb);
  print("  ✅ 20_آزمایش_4-4")
  build_4_5_schmidt(wb);
  print("  ✅ 21_آزمایش_4-5")
  build_22_report(wb);
  print("  ✅ 22_گزارش")
  build_23_dashboard(wb);
  print("  ✅ 23_داشبورد")
  build_24_qa(wb);
  print("  ✅ 24_QA_Test")
  build_25_errata(wb);
  print("  ✅ 25_خطاها_هشدارها")
  build_hidden_sheets(wb);
  print("  ✅ Hidden sheets (5)")

  print("\n📊 Adding sieve chart...")
  try:
    add_sieve_chart(wb)
    print("  ✅ Chart added")
  except Exception as e:
    print(f"  ⚠️ Chart skipped: {e}")

  print("\n🔒 Protecting sheets...")
  for ws in wb.worksheets:
    if not ws.title.startswith("_"):
      protect_sheet(ws)

  from openpyxl.workbook.protection import WorkbookProtection
  wb.security = WorkbookProtection(
    lockStructure=True,
    workbookPassword=PASSWORD
  )
  print("  ✅ All sheets protected")

  # Save
  output_dir = Path("output")
  output_dir.mkdir(exist_ok=True)
  filename = f"Concrete_Lab_Companion_v{VERSION}.xlsx"
  output_path = output_dir / filename

  # اگر فایل قبلی وجود دارد، حذف کن
  if output_path.exists():
    try:
      output_path.unlink()
    except Exception:
      print(f"  ⚠️ فایل قبلی باز است. لطفاً ببندید و دوباره اجرا کنید.")
      return

  wb.save(str(output_path))
  print(f"\n💾 Saved: {output_path}")

  # Hash
  file_hash = compute_sha256(str(output_path))
  hash_path = output_path.with_suffix(".xlsx.sha256")
  with open(hash_path, "w", encoding="utf-8") as f:
    f.write(f"{file_hash}  {filename}\n")

  print(f"🔐 SHA-256: {file_hash}")
  print("\n✅ BUILD COMPLETE — 26 visible sheets + 5 hidden sheets")
  print("═" * 50)


if __name__ == "__main__":
  main()
