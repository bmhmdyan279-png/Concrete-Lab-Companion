#!/usr/bin/env python3
"""
Concrete Lab Companion — GENERATION v3.0.0
==========================================
Professional Excel workbook for concrete testing.
Fully data‑driven, maintainable, standards‑compliant, and QA‑verified.
"""

import argparse
import hashlib
import json
import logging
import os
import sys
from dataclasses import dataclass, field
from datetime import datetime, timezone
from enum import Enum
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Union

from openpyxl import Workbook
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import (
    Alignment, Border, Font, PatternFill, Protection, Side
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.hyperlink import Hyperlink

# ─── Logging ────────────────────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

# ─── Constants ──────────────────────────────────────────────────────────────
VERSION = "3.0.0"
BUILD_DATE = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
PASSWORD = os.getenv("WORKBOOK_PASSWORD")  # No default; set or use --no-protect

# Standard specifications (ASTM / ISIRI) — with status and edition
STANDARDS = {
    "C29":   {"name": "ASTM C29/C29M", "edition": "2023", "status": "active", "title": "Bulk Density of Aggregates"},
    "C39":   {"name": "ASTM C39/C39M", "edition": "2022", "status": "active", "title": "Compressive Strength"},
    "C78":   {"name": "ASTM C78/C78M", "edition": "2022", "status": "active", "title": "Flexural Strength (Third-Point)"},
    "C127":  {"name": "ASTM C127", "edition": "2023", "status": "active", "title": "Density of Coarse Aggregate"},
    "C128":  {"name": "ASTM C128", "edition": "2022", "status": "active", "title": "Density of Fine Aggregate"},
    "C136":  {"name": "ASTM C136/C136M", "edition": "2024", "status": "active", "title": "Sieve Analysis"},
    "C138":  {"name": "ASTM C138/C138M", "edition": "2024", "status": "active", "title": "Density of Fresh Concrete"},
    "C143":  {"name": "ASTM C143/C143M", "edition": "2024", "status": "active", "title": "Slump of Concrete"},
    "C187":  {"name": "ASTM C187", "edition": "2016", "status": "withdrawn", "title": "Normal Consistency (historical)"},
    "C191":  {"name": "ASTM C191", "edition": "2024", "status": "active", "title": "Setting Time by Vicat"},
    "C232":  {"name": "ASTM C232/C232M", "edition": "2023", "status": "active", "title": "Bleeding of Concrete"},
    "C293":  {"name": "ASTM C293/C293M", "edition": "2023", "status": "active", "title": "Flexural Strength (Center-Point)"},
    "C496":  {"name": "ASTM C496/C496M", "edition": "2023", "status": "active", "title": "Splitting Tensile Strength"},
    "C566":  {"name": "ASTM C566", "edition": "2023", "status": "active", "title": "Moisture Content of Aggregates"},
    "C597":  {"name": "ASTM C597", "edition": "2023", "status": "active", "title": "Pulse Velocity Through Concrete"},
    "C805":  {"name": "ASTM C805/C805M", "edition": "2025", "status": "active", "title": "Rebound Hammer"},
    "D2419": {"name": "ASTM D2419", "edition": "2022", "status": "active", "title": "Sand Equivalent"},
    "D4791": {"name": "ASTM D4791", "edition": "2023", "status": "active", "title": "Flat & Elongated Particles"},
    "EN196-1": {"name": "EN 196-1", "edition": "2023", "status": "active", "title": "Mortar Strength"},
    "ISIRI302": {"name": "ISIRI 302", "edition": "2020", "status": "active", "title": "Aggregate Grading Limits"},
}
# ISIRI 302 limits for sieve analysis (sieve size mm → upper%, lower%)
ISIRI_LIMITS = {
    9.5:   (100, 95),
    4.75:  (95, 80),
    2.36:  (80, 60),
    1.18:  (60, 40),
    0.600: (40, 25),
    0.300: (25, 10),
    0.150: (10, 2),
    0.075: (2, 0),
}
SIEVE_SIZES = [9.5, 4.75, 2.36, 1.18, 0.600, 0.300, 0.150, 0.075]

# ─── Style System ──────────────────────────────────────────────────────────
class StyleManager:
    """Centralised style definitions with fallback fonts."""
    COLORS = {
        "input_fill": "FFF2CC", "input_border": "D9D9D9", "calc_fill": "F2F2F2",
        "pass_fill": "C6EFCE", "pass_font": "006100", "warn_fill": "FCE4D6",
        "warn_font": "C00000", "fail_fill": "FFC7CE", "fail_font": "9C0006",
        "header_fill": "1F4E79", "header_font": "FFFFFF", "nav_fill": "D6E4F0",
        "neutral_fill": "FFFFFF", "neutral_font": "000000",
    }
    FONT_FAMILY = "Tahoma, B Nazanin, Arial, sans-serif"
    FONT_NUM = "Calibri, Arial, sans-serif"

    def __init__(self):
        self.styles = {}
        self._build_styles()

    def _build_styles(self):
        c = self.COLORS
        self.styles.update({
            "header": {
                "fill": PatternFill("solid", fgColor=c["header_fill"]),
                "font": Font(name=self.FONT_FAMILY, color=c["header_font"], bold=True, size=11),
                "alignment": Alignment(horizontal="center", vertical="center", wrap_text=True),
            },
            "input": {
                "fill": PatternFill("solid", fgColor=c["input_fill"]),
                "border": Border(
                    left=Side("thin", c["input_border"]), right=Side("thin", c["input_border"]),
                    top=Side("thin", c["input_border"]), bottom=Side("thin", c["input_border"])
                ),
                "alignment": Alignment(horizontal="right", vertical="center", wrap_text=True),
            },
            "calc": {
                "fill": PatternFill("solid", fgColor=c["calc_fill"]),
                "alignment": Alignment(horizontal="right", vertical="center", wrap_text=True),
            },
            "neutral": {
                "fill": PatternFill("solid", fgColor=c["neutral_fill"]),
                "font": Font(name=self.FONT_FAMILY, color=c["neutral_font"]),
                "alignment": Alignment(horizontal="right", vertical="center", wrap_text=True),
            },
            "pass": {
                "fill": PatternFill("solid", fgColor=c["pass_fill"]),
                "font": Font(name=self.FONT_FAMILY, color=c["pass_font"], bold=True),
                "alignment": Alignment(horizontal="right", vertical="center", wrap_text=True),
            },
            "warn": {
                "fill": PatternFill("solid", fgColor=c["warn_fill"]),
                "font": Font(name=self.FONT_FAMILY, color=c["warn_font"], bold=True),
                "alignment": Alignment(horizontal="right", vertical="center", wrap_text=True),
            },
            "fail": {
                "fill": PatternFill("solid", fgColor=c["fail_fill"]),
                "font": Font(name=self.FONT_FAMILY, color=c["fail_font"], bold=True),
                "alignment": Alignment(horizontal="right", vertical="center", wrap_text=True),
            },
            "title": {
                "font": Font(name=self.FONT_FAMILY, bold=True, size=14, color=c["header_fill"]),
                "alignment": Alignment(horizontal="right", vertical="center", wrap_text=True),
            },
            "subtitle": {
                "font": Font(name=self.FONT_FAMILY, size=9, italic=True, color="666666"),
                "alignment": Alignment(horizontal="right", vertical="center", wrap_text=True),
            },
            "nav": {
                "fill": PatternFill("solid", fgColor=c["nav_fill"]),
                "font": Font(name=self.FONT_FAMILY, size=10, color=c["header_fill"]),
                "alignment": Alignment(horizontal="center", vertical="center", wrap_text=True),
            },
            "normal": {
                "font": Font(name=self.FONT_FAMILY, size=11),
                "alignment": Alignment(horizontal="right", vertical="center", wrap_text=True),
            },
            "num": {
                "font": Font(name=self.FONT_NUM, size=11),
                "alignment": Alignment(horizontal="right", vertical="center", wrap_text=True),
            },
            "thin_border": Border(
                left=Side("thin", "BFBFBF"), right=Side("thin", "BFBFBF"),
                top=Side("thin", "BFBFBF"), bottom=Side("thin", "BFBFBF")
            ),
        })

    def get(self, key: str) -> Dict:
        return self.styles.get(key, {})

    def apply(self, cell, style_key: str):
        style = self.get(style_key)
        if not style:
            return
        if "font" in style:
            cell.font = style["font"]
        if "fill" in style:
            cell.fill = style["fill"]
        if "border" in style:
            cell.border = style["border"]
        if "alignment" in style:
            cell.alignment = style["alignment"]

STYLES = StyleManager()

# ─── Status Enum ────────────────────────────────────────────────────────────
class TestStatus(Enum):
    NOT_READY = "—"
    PASS = "✅"
    WARN = "⚠️"
    FAIL = "❌"

# ─── Data Models ──────────────────────────────────────────────────────────
@dataclass
class InputField:
    key: str
    label: str
    row: int
    col: int = 3
    unit: str = ""
    validation: Dict[str, Any] = field(default_factory=dict)
    tooltip: str = ""
    instruction: str = ""

@dataclass
class OutputField:
    key: str
    label: str
    row: int
    col: int
    formula: str
    unit: str = ""
    num_fmt: str = "0.00"
    style: str = "neutral"
    tooltip: str = ""

@dataclass
class Check:
    label: str
    formula: str
    row: int
    col: int = 3
    style_pass: str = "pass"
    style_fail: str = "fail"
    severity: str = "error"

@dataclass
class ChartConfig:
    type: str = "scatter"
    title: str = ""
    x_col: int = 1
    y_col: int = 2
    data_start: int = 22
    data_end: int = 29
    x_axis_title: str = "اندازه الک (mm)"
    y_axis_title: str = "% عبوری"
    log_scale: bool = True
    series: List[Dict[str, Any]] = field(default_factory=list)

@dataclass
class TestSpec:
    id: str
    title: str
    standard: str
    sheet_name: str
    edition: str = ""
    tab_color: str = "1F4E79"
    subtitle: str = ""
    inputs: List[InputField] = field(default_factory=list)
    outputs: List[OutputField] = field(default_factory=list)
    checks: List[Check] = field(default_factory=list)
    chart: Optional[ChartConfig] = None
    summary_key: str = ""
    layout: str = "simple"
    custom_builder: Optional[Callable] = None

# ─── Test Registry ─────────────────────────────────────────────────────────
TEST_REGISTRY: Dict[str, TestSpec] = {}

def register_test(spec: TestSpec) -> TestSpec:
    TEST_REGISTRY[spec.id] = spec
    return spec

# ─── Row Tracker ─────────────────────────────────────────────────────────
class RowTracker:
    def __init__(self, start: int = 5):
        self.current = start

    def next(self, step: int = 1) -> int:
        self.current += step
        return self.current

# ─── Utility functions ────────────────────────────────────────────────────
def cell_ref(row: int, col: int = 3) -> str:
    return f"{get_column_letter(col)}{row}"

def safe_formula(formula: str) -> str:
    return f"=IFERROR({formula[1:] if formula.startswith('=') else formula}, \"—\")"

# ─── Define all tests ────────────────────────────────────────────────────
def build_sieve_table(ws, spec: TestSpec, wb: Workbook):
    """Custom builder for sieve analysis (test 1-1)."""
    STYLES.apply(ws.cell(row=7, column=1), "header")
    ws.merge_cells(start_row=7, start_column=1, end_row=7, end_column=11)
    ws.cell(row=7, column=1, value="جدول دانه‌بندی").alignment = Alignment(horizontal="center")
    headers = ["اندازه الک", "مانده (g)", "درصد مانده", "درصد عبوری", "حد بالا", "حد پایین", "در محدوده؟"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=8, column=col, value=h)
        STYLES.apply(cell, "header")
    base_row = 9
    for i, size in enumerate(SIEVE_SIZES):
        r = base_row + i
        ws.cell(row=r, column=1, value=size).number_format = "0.000"
        ws.cell(row=r, column=3)  # residue input
        STYLES.apply(ws.cell(row=r, column=3), "input")
        ws.cell(row=r, column=4, value=f"=IF($C$5=0,\"—\",C{r}/$C$5*100)").number_format = "0.00"
        ws.cell(row=r, column=5, value=f"=IF(ISNUMBER(D{r}),100-SUM($D$9:D{r}),\"—\")").number_format = "0.00"
        upper, lower = ISIRI_LIMITS.get(size, (100, 0))
        ws.cell(row=r, column=6, value=upper).number_format = "0.0"
        ws.cell(row=r, column=7, value=lower).number_format = "0.0"
        ws.cell(row=r, column=8, value=f"=IF(ISNUMBER(E{r}),AND(E{r}<=F{r},E{r}>=G{r}),\"—\")")
    r_total = base_row + len(SIEVE_SIZES)
    ws.cell(row=r_total, column=2, value="جمع مانده‌ها:").alignment = Alignment(horizontal="right")
    ws.cell(row=r_total, column=3, value=f"=SUM(C{base_row}:C{r_total-1})").number_format = "0.00"
    STYLES.apply(ws.cell(row=r_total, column=3), "calc")
    fm_row = r_total + 1
    ws.cell(row=fm_row, column=2, value="مدول نرمی (FM):").alignment = Alignment(horizontal="right")
    ws.cell(row=fm_row, column=3,
            value="=IF(OR($C$5=0,C9=\"\"),\"—\",ROUND(SUMPRODUCT((H9:H16=TRUE)*E9:E16)/100,2))").number_format = "0.00"
    STYLES.apply(ws.cell(row=fm_row, column=3), "calc")
    chk_row = r_total + 2
    ws.cell(row=chk_row, column=2, value="کنترل جرم:").alignment = Alignment(horizontal="right")
    ws.cell(row=chk_row, column=3,
            value=f"=IF($C$5=0,\"—\",IF(ABS($C$5-SUM(C{base_row}:C{r_total-1}))/$C$5>0.003,\"❌ اختلاف >0.3%\",\"✅\"))")
    STYLES.apply(ws.cell(row=chk_row, column=3), "calc")
    chart_start = r_total + 4
    ws.cell(row=chart_start-1, column=1, value="داده‌های نمودار").font = Font(bold=True)
    for col, h in enumerate(["اندازه الک", "% عبوری", "حد بالا", "حد پایین"], 1):
        ws.cell(row=chart_start-1, column=col, value=h).font = Font(bold=True)
    for i, size in enumerate(SIEVE_SIZES):
        r = chart_start + i
        ws.cell(row=r, column=1, value=size).number_format = "0.000"
        ws.cell(row=r, column=2, value=f"=E{base_row+i}")
        ws.cell(row=r, column=3, value=f"=F{base_row+i}")
        ws.cell(row=r, column=4, value=f"=G{base_row+i}")
        for col in (2,3,4):
            STYLES.apply(ws.cell(row=r, column=col), "calc")
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
    xvalues = Reference(ws, min_col=1, min_row=chart_start, max_row=chart_start+len(SIEVE_SIZES)-1)
    yvalues = Reference(ws, min_col=2, min_row=chart_start, max_row=chart_start+len(SIEVE_SIZES)-1)
    series = Series(yvalues, xvalues, title="نمونه")
    series.graphicalProperties.line.solidFill = "1F4E79"
    series.graphicalProperties.line.width = 22500
    chart.series.append(series)
    upper_ref = Reference(ws, min_col=3, min_row=chart_start, max_row=chart_start+len(SIEVE_SIZES)-1)
    lower_ref = Reference(ws, min_col=4, min_row=chart_start, max_row=chart_start+len(SIEVE_SIZES)-1)
    for ref, label, color in [(upper_ref, "حد بالا", "ED7D31"), (lower_ref, "حد پایین", "ED7D31")]:
        s = Series(ref, xvalues, title=label)
        s.graphicalProperties.line.solidFill = color
        s.graphicalProperties.line.dashStyle = "dash"
        chart.series.append(s)
    ws.add_chart(chart, f"I{chart_start-2}")
    spec.outputs = [OutputField(key="fm", label="FM", row=fm_row, col=3,
                                formula=f"=C{fm_row}", style="neutral", num_fmt="0.00")]
    spec.checks = [Check(label="کنترل جرم", formula=f"=C{chk_row}", row=chk_row, col=3)]
    return {"fm_row": fm_row, "chk_row": chk_row}

def build_matrix_input(ws, spec: TestSpec, wb: Workbook):
    """Custom builder for matrix layouts (2-4, 4-5)."""
    by_row = {}
    for inp in spec.inputs:
        by_row.setdefault(inp.row, []).append(inp)
    for row, inps in by_row.items():
        if len(inps) == 1 and inps[0].col == 3:
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
            ws.cell(row=row, column=1, value=inps[0].label)
            STYLES.apply(ws.cell(row=row, column=1), "normal")
            ws.cell(row=row, column=3, value=None)
            STYLES.apply(ws.cell(row=row, column=3), "input")
            if inps[0].unit:
                ws.cell(row=row, column=4, value=f"[{inps[0].unit}]")
                STYLES.apply(ws.cell(row=row, column=4), "num")
        else:
            for inp in inps:
                ws.cell(row=row-1, column=inp.col, value=inp.label)
                STYLES.apply(ws.cell(row=row-1, column=inp.col), "header")
                ws.cell(row=row, column=inp.col, value=None)
                STYLES.apply(ws.cell(row=row, column=inp.col), "input")
                if inp.unit:
                    ws.cell(row=row-1, column=inp.col+1, value=f"[{inp.unit}]")
                    STYLES.apply(ws.cell(row=row-1, column=inp.col+1), "num")

# ─── Register all tests ──────────────────────────────────────────────────
def define_tests():
    # 1-1 Sieve Analysis
    register_test(TestSpec(
        id="1-1",
        title="دانه‌بندی سنگدانه",
        standard=STANDARDS["C136"]["name"],
        edition=STANDARDS["C136"]["edition"],
        tab_color="FF6F00",
        sheet_name="02_آزمایش_1-1",
        subtitle="ISIRI 302",
        summary_key="دانه‌بندی (FM)",
        layout="table",
        custom_builder=build_sieve_table,
        inputs=[
            InputField("initial_mass", "جرم اولیه نمونه خشک (g):", 5, 3, "g",
                       validation={"type":"decimal","min":0,"max":100000},
                       tooltip="جرم اولیه نمونه خشک را به گرم وارد کنید.",
                       instruction="نمونه را تا وزن ثابت در آون خشک کرده و توزین کنید.")
        ],
    ))

    # 1-2 Moisture
    register_test(TestSpec(
        id="1-2",
        title="رطوبت سنگدانه",
        standard=STANDARDS["C566"]["name"],
        edition=STANDARDS["C566"]["edition"],
        tab_color="FF6F00",
        sheet_name="03_آزمایش_1-2",
        summary_key="رطوبت",
        inputs=[
            InputField("W1", "جرم نمونه تر (g):", 5, 3, "g",
                       validation={"type":"decimal","min":0,"max":100000},
                       tooltip="جرم نمونه در حالت طبیعی (تر)"),
            InputField("W2", "جرم نمونه خشک (g):", 6, 3, "g",
                       validation={"type":"decimal","min":0,"max":100000},
                       tooltip="جرم نمونه پس از خشک شدن در آون"),
        ],
        outputs=[
            OutputField("moisture", "درصد رطوبت (پایه خشک):", 8, 3,
                       formula="=IF(OR(C5=\"\",C6=\"\",C6=0),\"—\",ROUND((C5-C6)/C6*100,2))",
                       num_fmt="0.00", style="neutral", unit="%",
                       tooltip="رطوبت بر اساس وزن خشک محاسبه می‌شود."),
        ]
    ))

    # 1-3 Coarse aggregate density
    register_test(TestSpec(
        id="1-3",
        title="چگالی سنگدانه درشت",
        standard=STANDARDS["C127"]["name"],
        edition=STANDARDS["C127"]["edition"],
        tab_color="FF6F00",
        sheet_name="04_آزمایش_1-3",
        summary_key="چگالی درشت (SSD)",
        inputs=[
            InputField("A", "جرم خشک (g):", 5, 3, "g",
                       validation={"type":"decimal","min":0,"max":100000},
                       tooltip="جرم نمونه خشک شده در آون"),
            InputField("B", "جرم SSD (g):", 6, 3, "g",
                       validation={"type":"decimal","min":0,"max":100000},
                       tooltip="جرم نمونه در حالت اشباع با سطح خشک"),
            InputField("C", "جرم در آب (g):", 7, 3, "g",
                       validation={"type":"decimal","min":0,"max":100000},
                       tooltip="جرم نمونه در آب (وزن هیدرواستاتیک)"),
        ],
        outputs=[
            OutputField("OD", "چگالی خشک (OD):", 9, 3,
                       formula="=IF(OR(C5=\"\",C6=\"\",C7=\"\",C6=C7),\"—\",ROUND(C5/(C6-C7),3))",
                       num_fmt="0.000", style="neutral", unit="g/cm³"),
            OutputField("SSD", "چگالی SSD:", 10, 3,
                       formula="=IF(OR(C5=\"\",C6=\"\",C7=\"\",C6=C7),\"—\",ROUND(C6/(C6-C7),3))",
                       num_fmt="0.000", style="neutral", unit="g/cm³"),
            OutputField("App", "چگالی ظاهری:", 11, 3,
                       formula="=IF(OR(C5=\"\",C6=\"\",C7=\"\",C5=C7),\"—\",ROUND(C5/(C5-C7),3))",
                       num_fmt="0.000", style="neutral", unit="g/cm³"),
            OutputField("Abs", "جذب آب (%):", 12, 3,
                       formula="=IF(OR(C5=\"\",C6=\"\",C5=0),\"—\",ROUND((C6-C5)/C5*100,2))",
                       num_fmt="0.01", style="neutral", unit="%"),
        ],
        checks=[
            Check("بررسی فیزیکی (SSD ≥ OD)", "=IF(OR(C9=\"—\",C10=\"—\"),\"—\",IF(C10<C9-0.01,\"❌ غیرفیزیکی\",\"✅\"))", 14, 3, severity="error"),
            Check("بررسی فیزیکی (App ≥ SSD)", "=IF(OR(C10=\"—\",C11=\"—\"),\"—\",IF(C11<C10-0.01,\"❌ غیرفیزیکی\",\"✅\"))", 15, 3, severity="error"),
        ]
    ))

    # 1-4 Fine aggregate density
    register_test(TestSpec(
        id="1-4",
        title="چگالی سنگدانه ریز",
        standard=STANDARDS["C128"]["name"],
        edition=STANDARDS["C128"]["edition"],
        tab_color="FF6F00",
        sheet_name="05_آزمایش_1-4",
        summary_key="چگالی ریز (SSD)",
        inputs=[
            InputField("A", "جرم خشک (g):", 7, 3, "g",
                       validation={"type":"decimal","min":0,"max":100000},
                       tooltip="جرم نمونه خشک شده در آون"),
            InputField("S", "جرم SSD (g):", 8, 3, "g",
                       validation={"type":"decimal","min":0,"max":100000},
                       tooltip="جرم نمونه در حالت اشباع با سطح خشک"),
            InputField("B", "جرم ظرف+آب (g):", 9, 3, "g",
                       validation={"type":"decimal","min":0,"max":100000},
                       tooltip="جرم پیکنومتر + آب"),
            InputField("C", "جرم ظرف+آب+نمونه (g):", 10, 3, "g",
                       validation={"type":"decimal","min":0,"max":100000},
                       tooltip="جرم پیکنومتر + آب + نمونه"),
        ],
        outputs=[
            OutputField("OD", "چگالی خشک (OD):", 12, 3,
                       formula="=IF(OR(C7=\"\",C9=\"\",C10=\"\",C9+C8-C10=0),\"—\",ROUND(C7/(C9+C8-C10),3))",
                       num_fmt="0.000", style="neutral", unit="g/cm³"),
            OutputField("SSD", "چگالی SSD:", 13, 3,
                       formula="=IF(OR(C8=\"\",C9=\"\",C10=\"\",C9+C8-C10=0),\"—\",ROUND(C8/(C9+C8-C10),3))",
                       num_fmt="0.000", style="neutral", unit="g/cm³"),
            OutputField("App", "چگالی ظاهری:", 14, 3,
                       formula="=IF(OR(C7=\"\",C9=\"\",C10=\"\",C9+C7-C10=0),\"—\",ROUND(C7/(C9+C7-C10),3))",
                       num_fmt="0.000", style="neutral", unit="g/cm³"),
            OutputField("Abs", "جذب آب (%):", 15, 3,
                       formula="=IF(OR(C7=\"\",C8=\"\",C7=0),\"—\",ROUND((C8-C7)/C7*100,2))",
                       num_fmt="0.01", style="neutral", unit="%"),
        ],
        checks=[
            Check("بررسی فیزیکی (SSD ≥ OD)", "=IF(OR(C12=\"—\",C13=\"—\"),\"—\",IF(C13<C12-0.01,\"❌ غیرفیزیکی\",\"✅\"))", 17, 3, severity="error"),
            Check("بررسی فیزیکی (App ≥ SSD)", "=IF(OR(C13=\"—\",C14=\"—\"),\"—\",IF(C14<C13-0.01,\"❌ غیرفیزیکی\",\"✅\"))", 18, 3, severity="error"),
        ]
    ))

    # 1-5 Unit weight of aggregate
    register_test(TestSpec(
        id="1-5",
        title="وزن واحد حجمی سنگدانه",
        standard=STANDARDS["C29"]["name"],
        edition=STANDARDS["C29"]["edition"],
        tab_color="FF6F00",
        sheet_name="06_آزمایش_1-5",
        summary_key="وزن واحد حجمی",
        inputs=[
            InputField("T", "جرم ظرف خالی (g):", 5, 3, "g",
                       validation={"type":"decimal","min":0,"max":100000},
                       tooltip="جرم ظرف خالی"),
            InputField("G", "جرم ظرف+سنگدانه (g):", 6, 3, "g",
                       validation={"type":"decimal","min":0,"max":100000},
                       tooltip="جرم ظرف همراه با سنگدانه"),
            InputField("V", "حجم ظرف (cm³):", 7, 3, "cm³",
                       validation={"type":"decimal","min":0,"max":100000},
                       tooltip="حجم کالیبره شده ظرف"),
            InputField("S", "چگالی (بی‌بعد):", 8, 3, "",
                       validation={"type":"decimal","min":2,"max":3.5},
                       tooltip="چگالی ظاهری یا SSD سنگدانه"),
        ],
        outputs=[
            OutputField("UW", "وزن واحد حجمی (kg/m³):", 10, 3,
                       formula="=IF(OR(C5=\"\",C6=\"\",C7=\"\",C7=0),\"—\",ROUND((C6-C5)/C7*1000,0))",
                       num_fmt="#,##0", style="neutral", unit="kg/m³"),
            OutputField("Voids", "فضای خالی (%):", 11, 3,
                       formula="=IF(OR(C8=\"\",C8=0,C10=\"—\"),\"—\",ROUND((C8*1000-C10)/(C8*1000)*100,1))",
                       num_fmt="0.0", style="neutral", unit="%"),
        ],
        checks=[]
    ))

    # 1-6 Sand Equivalent (ASTM D2419) — corrected to standard formula
    register_test(TestSpec(
        id="1-6",
        title="معادل ماسه",
        standard=STANDARDS["D2419"]["name"],
        edition=STANDARDS["D2419"]["edition"],
        tab_color="FF6F00",
        sheet_name="07_آزمایش_1-6",
        summary_key="معادل ماسه",
        inputs=[
            InputField("Sand", "خوانش ماسه (mm):", 5, 3, "mm",
                       validation={"type":"decimal","min":0,"max":500},
                       tooltip="ارتفاع لایه ماسه"),
            InputField("Clay", "خوانش رس (mm):", 6, 3, "mm",
                       validation={"type":"decimal","min":0,"max":500},
                       tooltip="ارتفاع لایه رس"),
        ],
        outputs=[
            OutputField("SE", "معادل ماسه SE (%):", 8, 3,
                       formula="=IF(OR(C5=\"\",C6=\"\",C6=0),\"—\",ROUND(C5/C6*100,1))",
                       num_fmt="0.0", style="neutral", unit="%",
                       tooltip="SE = (خوانش ماسه / خوانش رس) × 100 — طبق ASTM D2419"),
        ],
        checks=[]
    ))

    # 1-7 Shape indices
    register_test(TestSpec(
        id="1-7",
        title="شاخص‌های شکل سنگدانه",
        standard=STANDARDS["D4791"]["name"],
        edition=STANDARDS["D4791"]["edition"],
        tab_color="FF6F00",
        sheet_name="08_آزمایش_1-7",
        summary_key="شاخص درازگی",
        inputs=[
            InputField("Wtot", "جرم کل (g):", 5, 3, "g",
                       validation={"type":"decimal","min":0,"max":100000},
                       tooltip="جرم کل نمونه"),
            InputField("Wlong", "جرم دراز (g):", 6, 3, "g",
                       validation={"type":"decimal","min":0,"max":100000},
                       tooltip="جرم ذرات دراز (نسبت طول به عرض > 3)"),
            InputField("Wflat", "جرم پهن (g):", 7, 3, "g",
                       validation={"type":"decimal","min":0,"max":100000},
                       tooltip="جرم ذرات پهن (نسبت عرض به ضخامت > 3)"),
            InputField("Wboth", "جرم هر دو (g):", 8, 3, "g",
                       validation={"type":"decimal","min":0,"max":100000},
                       tooltip="جرم ذرات دارای هر دو ویژگی"),
        ],
        outputs=[
            OutputField("LI", "شاخص درازگی (%):", 10, 3,
                       formula="=IF(OR(C5=\"\",C5=0,C6=\"\"),\"—\",ROUND(C6/C5*100,1))",
                       num_fmt="0.0", style="neutral", unit="%"),
            OutputField("FI", "شاخص پهنی (%):", 11, 3,
                       formula="=IF(OR(C5=\"\",C5=0,C7=\"\"),\"—\",ROUND(C7/C5*100,1))",
                       num_fmt="0.0", style="neutral", unit="%"),
            OutputField("BI", "شاخص هر دو (%):", 12, 3,
                       formula="=IF(OR(C5=\"\",C5=0,C8=\"\"),\"—\",ROUND(C8/C5*100,1))",
                       num_fmt="0.0", style="neutral", unit="%"),
        ],
        checks=[]
    ))

    # 1-8 Absorption
    register_test(TestSpec(
        id="1-8",
        title="جذب آب سنگدانه",
        standard=STANDARDS["C127"]["name"],
        edition=STANDARDS["C127"]["edition"],
        tab_color="FF6F00",
        sheet_name="09_آزمایش_1-8",
        summary_key="جذب آب",
        inputs=[
            InputField("W1", "جرم SSD (g):", 5, 3, "g",
                       validation={"type":"decimal","min":0,"max":100000},
                       tooltip="جرم نمونه در حالت اشباع با سطح خشک"),
            InputField("W2", "جرم خشک (g):", 6, 3, "g",
                       validation={"type":"decimal","min":0,"max":100000},
                       tooltip="جرم نمونه خشک شده در آون"),
        ],
        outputs=[
            OutputField("Abs", "جذب آب (%):", 8, 3,
                       formula="=IF(OR(C5=\"\",C6=\"\",C6=0),\"—\",ROUND((C5-C6)/C6*100,2))",
                       num_fmt="0.01", style="neutral", unit="%"),
        ],
        checks=[]
    ))

    # 2-1 Fresh concrete density (ASTM C138) with air content
    register_test(TestSpec(
        id="2-1",
        title="چگالی بتن تازه",
        standard=STANDARDS["C138"]["name"],
        edition=STANDARDS["C138"]["edition"],
        tab_color="2196F3",
        sheet_name="10_آزمایش_2-1",
        summary_key="چگالی بتن تازه",
        inputs=[
            InputField("Ma", "جرم ظرف خالی (g):", 5, 3, "g",
                       validation={"type":"decimal","min":0,"max":100000},
                       tooltip="جرم ظرف خالی"),
            InputField("Mt", "جرم ظرف+بتن (g):", 6, 3, "g",
                       validation={"type":"decimal","min":0,"max":100000},
                       tooltip="جرم ظرف همراه با بتن تازه"),
            InputField("V", "حجم ظرف (cm³):", 7, 3, "cm³",
                       validation={"type":"decimal","min":0,"max":100000},
                       tooltip="حجم کالیبره شده ظرف"),
            InputField("TheoD", "چگالی نظری (kg/m³):", 8, 3, "kg/m³",
                       validation={"type":"decimal","min":0,"max":10000},
                       tooltip="چگالی نظری بر اساس طرح اختلاط"),
        ],
        outputs=[
            OutputField("Density", "چگالی (kg/m³):", 10, 3,
                       formula="=IF(OR(C5=\"\",C6=\"\",C7=\"\",C7=0),\"—\",ROUND((C6-C5)/C7*1000,0))",
                       num_fmt="#,##0", style="neutral", unit="kg/m³"),
            OutputField("AirPct", "درصد هوای بتن (%):", 11, 3,
                       formula="=IF(OR(C10=\"—\",C8=\"\",C8=0),\"—\",ROUND((C8-C10)/C8*100,1))",
                       num_fmt="0.0", style="neutral", unit="%",
                       tooltip="هوای بتن = (چگالی نظری - چگالی اندازه‌گیری شده) / چگالی نظری × 100"),
        ],
        checks=[]
    ))

    # 2-2 Normal Consistency (historical)
    register_test(TestSpec(
        id="2-2",
        title="قوام نرمال سیمان (ویکات)",
        standard=STANDARDS["C187"]["name"] + " (historical reference)",
        edition=STANDARDS["C187"]["edition"],
        tab_color="2196F3",
        sheet_name="11_آزمایش_2-2",
        summary_key="قوام نرمال (w/c)",
        inputs=[
            InputField("Cement", "سیمان (g):", 5, 3, "g",
                       validation={"type":"decimal","min":0,"max":10000},
                       tooltip="جرم سیمان"),
            InputField("Water", "آب (g):", 6, 3, "g",
                       validation={"type":"decimal","min":0,"max":10000},
                       tooltip="جرم آب مخلوط"),
            InputField("Penetration", "نفوذ اولیه (mm):", 7, 3, "mm",
                       validation={"type":"decimal","min":0,"max":50},
                       tooltip="نفوذ سوزن ویکات پس از 30 ثانیه"),
        ],
        outputs=[
            OutputField("WC", "نسبت آب به سیمان (%):", 9, 3,
                       formula="=IF(OR(C5=\"\",C6=\"\",C5=0),\"—\",ROUND(C6/C5*100,1))",
                       num_fmt="0.0", style="neutral", unit="%"),
        ],
        checks=[
            Check("وضعیت نفوذ", "=IF(C7=\"\",\"\",IF(ABS(C7-10)>1,\"⚠️ تکرار با آب جدید\",\"✅\"))", 10, 3, severity="warning"),
        ]
    ))

    # 2-3 Setting Time
    register_test(TestSpec(
        id="2-3",
        title="زمان گیرش سیمان",
        standard=STANDARDS["C191"]["name"],
        edition=STANDARDS["C191"]["edition"],
        tab_color="2196F3",
        sheet_name="12_آزمایش_2-3",
        summary_key="گیرش اولیه",
        inputs=[
            InputField("E", "زمان اولیه (min):", 5, 3, "min",
                       validation={"type":"decimal","min":0,"max":10000},
                       tooltip="زمان شروع گیرش (min)"),
            InputField("H", "زمان ثانویه (min):", 6, 3, "min",
                       validation={"type":"decimal","min":0,"max":10000},
                       tooltip="زمان دوم (min)"),
            InputField("C", "نفوذ ثانویه (mm):", 7, 3, "mm",
                       validation={"type":"decimal","min":0,"max":50},
                       tooltip="مقدار نفوذ در زمان دوم"),
            InputField("D", "نفوذ اولیه (mm):", 8, 3, "mm",
                       validation={"type":"decimal","min":0,"max":50},
                       tooltip="مقدار نفوذ در زمان اولیه"),
            InputField("Final", "زمان گیرش نهایی (min):", 9, 3, "min",
                       validation={"type":"decimal","min":0,"max":10000},
                       tooltip="زمان گیرش نهایی بر اساس آزمایش جداگانه"),
        ],
        outputs=[
            OutputField("Initial", "زمان گیرش اولیه (min):", 11, 3,
                       formula="=IF(OR(C5=\"\",C6=\"\",C7=\"\",C8=\"\",C7=C8),\"—\",ROUND(C5+(C6-C5)*(25-C8)/(C7-C8),0))",
                       num_fmt="0", style="neutral", unit="min"),
            OutputField("FinalM", "زمان گیرش نهایی (min):", 12, 3,
                       formula="=IF(C9=\"\",\"\",MROUND(C9,5))", num_fmt="0", style="neutral", unit="min"),
        ],
        checks=[
            Check("ترتیب منطقی (نهایی > اولیه)", "=IF(OR(C11=\"—\",C12=\"—\"),\"—\",IF(C12<=C11,\"❌ خطا\",\"✅\"))", 14, 3, severity="error"),
        ]
    ))

    # 2-4 Mortar strength (matrix layout)
    register_test(TestSpec(
        id="2-4",
        title="مقاومت ملات سیمان",
        standard=STANDARDS["EN196-1"]["name"],
        edition=STANDARDS["EN196-1"]["edition"],
        tab_color="2196F3",
        sheet_name="13_آزمایش_2-4",
        summary_key="مقاومت ملات",
        layout="matrix",
        custom_builder=build_matrix_input,
        inputs=[
            InputField("F1", "بار خمشی ۱ (kgf)", 7, 1, "kgf",
                       validation={"type":"decimal","min":0,"max":10000},
                       tooltip="بار خمشی نمونه ۱"),
            InputField("F2", "بار خمشی ۲ (kgf)", 7, 2, "kgf",
                       validation={"type":"decimal","min":0,"max":10000},
                       tooltip="بار خمشی نمونه ۲"),
            InputField("F3", "بار خمشی ۳ (kgf)", 7, 3, "kgf",
                       validation={"type":"decimal","min":0,"max":10000},
                       tooltip="بار خمشی نمونه ۳"),
            InputField("C1", "بار فشاری ۱ (kgf)", 11, 1, "kgf",
                       validation={"type":"decimal","min":0,"max":50000},
                       tooltip="بار فشاری نمونه ۱"),
            InputField("C2", "بار فشاری ۲ (kgf)", 11, 2, "kgf",
                       validation={"type":"decimal","min":0,"max":50000},
                       tooltip="بار فشاری نمونه ۲"),
            InputField("C3", "بار فشاری ۳ (kgf)", 11, 3, "kgf",
                       validation={"type":"decimal","min":0,"max":50000},
                       tooltip="بار فشاری نمونه ۳"),
            InputField("C4", "بار فشاری ۴ (kgf)", 11, 4, "kgf",
                       validation={"type":"decimal","min":0,"max":50000},
                       tooltip="بار فشاری نمونه ۴"),
            InputField("C5", "بار فشاری ۵ (kgf)", 11, 5, "kgf",
                       validation={"type":"decimal","min":0,"max":50000},
                       tooltip="بار فشاری نمونه ۵"),
            InputField("C6", "بار فشاری ۶ (kgf)", 11, 6, "kgf",
                       validation={"type":"decimal","min":0,"max":50000},
                       tooltip="بار فشاری نمونه ۶"),
        ],
        outputs=[
            OutputField("Flex", "مقاومت خمشی (MPa):", 9, 7,
                       formula="=IF(COUNT(A7:C7)<3,\"—\",ROUND(AVERAGE(1.5*A7*9.80665*100/40^3,1.5*B7*9.80665*100/40^3,1.5*C7*9.80665*100/40^3),1))",
                       num_fmt="0.0", style="neutral", unit="MPa"),
            OutputField("Comp", "مقاومت فشاری (MPa):", 13, 7,
                       formula="=IF(COUNT(A11:F11)<6,\"—\",ROUND(AVERAGE(A11*9.80665/1600,B11*9.80665/1600,C11*9.80665/1600,D11*9.80665/1600,E11*9.80665/1600,F11*9.80665/1600),1))",
                       num_fmt="0.0", style="neutral", unit="MPa"),
        ],
        checks=[
            Check("حذف داده پرت ±10% (فشاری)", "=IF(OR(C13=\"—\",COUNT(A11:F11)<6),\"—\",IF(MAX(A11:F11)>1.1*AVERAGE(A11:F11),\"⚠️ داده پرت\",\"✅\"))", 15, 7, severity="warning"),
        ]
    ))

    # 3-1 Slump — using 300 mm per ASTM C143M (metric)
    register_test(TestSpec(
        id="3-1",
        title="اسلامپ بتن",
        standard=STANDARDS["C143"]["name"],
        edition=STANDARDS["C143"]["edition"],
        tab_color="4CAF50",
        sheet_name="14_آزمایش_3-1",
        summary_key="اسلامپ",
        inputs=[
            InputField("H", "ارتفاع پس از برداشتن (mm):", 5, 3, "mm",
                       validation={"type":"decimal","min":0,"max":300},
                       tooltip="ارتفاع بتن پس از برداشتن قالب (میلی‌متر)"),
            InputField("Type", "نوع ریزش:", 6, 3,
                       validation={"type":"list","formula":"برشی,دو طرفه,ریزش کامل"},
                       tooltip="نوع ریزش بتن پس از برداشتن قالب"),
        ],
        outputs=[
            OutputField("Slump", "اسلامپ (mm):", 8, 3,
                       formula="=IF(C5=\"\",\"\",MROUND(300-C5,5))",
                       num_fmt="0", style="neutral", unit="mm",
                       tooltip="اسلامپ = 300 - ارتفاع باقی‌مانده (طبق ASTM C143M)"),
        ],
        checks=[
            Check("وضعیت", "=IF(C6=\"\",\"\",IF(C6=\"ریزش کامل\",\"⚠️ تکرار آزمایش\",\"✅\"))", 9, 3, severity="warning"),
        ]
    ))

    # 3-2 Bleeding (approximate height-based)
    register_test(TestSpec(
        id="3-2",
        title="آب‌انداختگی بتن (تقریبی)",
        standard=STANDARDS["C232"]["name"] + " (روش تقریبی ارتفاعی)",
        edition=STANDARDS["C232"]["edition"],
        tab_color="4CAF50",
        sheet_name="15_آزمایش_3-2",
        summary_key="آب‌انداختگی",
        inputs=[
            InputField("h1", "ارتفاع اولیه (mm):", 5, 3, "mm",
                       validation={"type":"decimal","min":0,"max":1000},
                       tooltip="ارتفاع بتن تازه در قالب قبل از آب‌انداختگی"),
            InputField("h2", "ارتفاع نهایی (mm):", 6, 3, "mm",
                       validation={"type":"decimal","min":0,"max":1000},
                       tooltip="ارتفاع بتن پس از آب‌انداختگی"),
            InputField("G", "جذب سنگدانه (%):", 7, 3, "%",
                       validation={"type":"decimal","min":0,"max":1000},
                       tooltip="جذب آب سنگدانه (درصد) برای تصحیح"),
        ],
        outputs=[
            OutputField("Bleed", "آب‌انداختگی ظاهری (%):", 9, 3,
                       formula="=IF(OR(C5=\"\",C6=\"\",C5=0),\"—\",ROUND((C5-C6)/C5*100,1))",
                       num_fmt="0.0", style="neutral", unit="%"),
            OutputField("BleedReal", "آب‌انداختگی واقعی (%):", 10, 3,
                       formula="=IF(OR(C9=\"—\",C7=\"\"),\"—\",ROUND(C9-C7,1))",
                       num_fmt="0.0", style="neutral", unit="%"),
        ],
        checks=[]
    ))

    # 3-3 Unit weight of concrete
    register_test(TestSpec(
        id="3-3",
        title="وزن واحد حجمی بتن",
        standard=STANDARDS["C138"]["name"],
        edition=STANDARDS["C138"]["edition"],
        tab_color="4CAF50",
        sheet_name="16_آزمایش_3-3",
        summary_key="وزن واحد بتن",
        inputs=[
            InputField("m1", "جرم ظرف خالی (g):", 5, 3, "g",
                       validation={"type":"decimal","min":0,"max":100000},
                       tooltip="جرم ظرف خالی"),
            InputField("m2", "جرم ظرف+بتن (g):", 6, 3, "g",
                       validation={"type":"decimal","min":0,"max":100000},
                       tooltip="جرم ظرف همراه با بتن"),
            InputField("V", "حجم ظرف (cm³):", 7, 3, "cm³",
                       validation={"type":"decimal","min":0,"max":100000},
                       tooltip="حجم کالیبره شده ظرف"),
            InputField("D_theo", "چگالی نظری (kg/m³):", 8, 3, "kg/m³",
                       validation={"type":"decimal","min":0,"max":10000},
                       tooltip="چگالی نظری بر اساس طرح اختلاط"),
        ],
        outputs=[
            OutputField("UW", "وزن واحد حجمی (kg/m³):", 10, 3,
                       formula="=IF(OR(C5=\"\",C6=\"\",C7=\"\",C7=0),\"—\",ROUND((C6-C5)/C7*1000,0))",
                       num_fmt="#,##0", style="neutral", unit="kg/m³"),
            OutputField("Dev", "اختلاف با نظری (%):", 11, 3,
                       formula="=IF(OR(C10=\"—\",C8=\"\",C8=0),\"—\",ROUND(ABS(C10-C8)/C8*100,1))",
                       num_fmt="0.0", style="neutral", unit="%"),
        ],
        checks=[
            Check("وضعیت", "=IF(C11=\"—\",\"\",IF(C11>2,\"⚠️ اختلاف >2%\",\"✅\"))", 12, 3, severity="warning"),
        ]
    ))

    # 4-1 Compressive strength (ASTM C39) with L/D correction
    register_test(TestSpec(
        id="4-1",
        title="مقاومت فشاری بتن (استوانه)",
        standard=STANDARDS["C39"]["name"],
        edition=STANDARDS["C39"]["edition"],
        tab_color="F44336",
        sheet_name="17_آزمایش_4-1",
        summary_key="مقاومت فشاری",
        inputs=[
            InputField("D", "قطر (mm):", 6, 3, "mm",
                       validation={"type":"decimal","min":0,"max":500},
                       tooltip="قطر نمونه استوانه‌ای (میلی‌متر)"),
            InputField("L", "ارتفاع (mm):", 7, 3, "mm",
                       validation={"type":"decimal","min":0,"max":500},
                       tooltip="ارتفاع نمونه استوانه‌ای (میلی‌متر)"),
            InputField("P", "بار (kN):", 8, 3, "kN",
                       validation={"type":"decimal","min":0,"max":10000},
                       tooltip="بار نهایی (کیلونیوتن)"),
            InputField("Pattern", "الگوی شکست:", 9, 3,
                       validation={"type":"list","formula":"نوع ۱,نوع ۲,نوع ۳,نوع ۴,نوع ۵,نوع ۶"},
                       tooltip="الگوی شکست بر اساس ASTM C39"),
            InputField("Shape", "نوع نمونه:", 10, 3,
                       validation={"type":"list","formula":"استوانه,مکعب"},
                       tooltip="نوع نمونه (استوانه یا مکعب)"),
        ],
        outputs=[
            OutputField("Area", "مساحت (mm²):", 12, 3,
                       formula="=IF(OR(C6=\"\",C6=0),\"—\",PI()/4*C6^2)",
                       num_fmt="0.0", style="neutral", unit="mm²"),
            OutputField("LD", "نسبت L/D:", 13, 3,
                       formula="=IF(OR(C6=\"\",C7=\"\",C6=0),\"—\",ROUND(C7/C6,3))",
                       num_fmt="0.000", style="neutral"),
            OutputField("Strength", "مقاومت فشاری (MPa):", 14, 3,
                       formula="=IF(OR(C8=\"\",C12=\"—\",C10=\"مکعب\"),IF(C10=\"مکعب\",ROUND(C8*1000/(C6^2)*0.95,1),\"—\"),MROUND(C8*1000/C12*IF(C13<1.8,0.96,IF(C13<1.9,0.98,1)),0.1))",
                       num_fmt="0.1", style="neutral", unit="MPa",
                       tooltip="مقاومت فشاری با اعمال ضریب اصلاح L/D بر اساس ASTM C39"),
        ],
        checks=[
            Check("الگوی شکست معتبر", "=IF(C9=\"\",\"\",IF(OR(C9=\"نوع ۱\",C9=\"نوع ۲\",C9=\"نوع ۳\"),\"✅ معتبر\",\"⚠️ نامعتبر\"))", 15, 3, severity="warning"),
            Check("نسبت L/D معتبر (≥1.8)", "=IF(C13=\"—\",\"\",IF(C13<1.8,\"⚠️ L/D<1.8\",\"✅\"))", 16, 3, severity="warning"),
        ]
    ))

    # 4-2 Splitting tensile
    register_test(TestSpec(
        id="4-2",
        title="مقاومت کششی (برزیلی)",
        standard=STANDARDS["C496"]["name"],
        edition=STANDARDS["C496"]["edition"],
        tab_color="F44336",
        sheet_name="18_آزمایش_4-2",
        summary_key="مقاومت کششی",
        inputs=[
            InputField("d", "قطر (mm):", 5, 3, "mm",
                       validation={"type":"decimal","min":0,"max":1000},
                       tooltip="قطر نمونه استوانه‌ای"),
            InputField("L", "طول (mm):", 6, 3, "mm",
                       validation={"type":"decimal","min":0,"max":1000},
                       tooltip="طول نمونه استوانه‌ای"),
            InputField("P", "بار (N):", 7, 3, "N",
                       validation={"type":"decimal","min":0,"max":1000000},
                       tooltip="بار نهایی (نیوتن)"),
        ],
        outputs=[
            OutputField("Fct", "مقاومت کششی (MPa):", 9, 3,
                       formula="=IF(OR(C5=\"\",C6=\"\",C7=\"\",C5=0,C6=0),\"—\",ROUND(2*C7/(PI()*C5*C6),2))",
                       num_fmt="0.00", style="neutral", unit="MPa"),
        ],
        checks=[
            Check("بازه منطقی (۲-۸ MPa)", "=IF(C9=\"—\",\"\",IF(OR(C9<2,C9>8),\"⚠️ خارج از بازه\",\"✅\"))", 10, 3, severity="warning"),
        ]
    ))

    # 4-3 Flexural strength (C78 + C293)
    register_test(TestSpec(
        id="4-3",
        title="مقاومت خمشی",
        standard=STANDARDS["C78"]["name"] + " / " + STANDARDS["C293"]["name"],
        edition="2022/2023",
        tab_color="F44336",
        sheet_name="19_آزمایش_4-3",
        summary_key="مقاومت خمشی",
        inputs=[
            InputField("b", "عرض (mm):", 5, 3, "mm",
                       validation={"type":"decimal","min":0,"max":1000},
                       tooltip="عرض مقطع تیر"),
            InputField("d", "ارتفاع (mm):", 6, 3, "mm",
                       validation={"type":"decimal","min":0,"max":1000},
                       tooltip="ارتفاع مقطع تیر"),
            InputField("L", "دهانه (mm):", 7, 3, "mm",
                       validation={"type":"decimal","min":0,"max":1000},
                       tooltip="دهانه تکیه‌گاهی"),
            InputField("P", "بار (N):", 8, 3, "N",
                       validation={"type":"decimal","min":0,"max":1000000},
                       tooltip="بار نهایی (نیوتن)"),
            InputField("Method", "روش بارگذاری:", 10, 3,
                       validation={"type":"list","formula":"یک‌سوم میانه,مرکزی"},
                       tooltip="روش بارگذاری: یک‌سوم میانه (C78) یا مرکزی (C293)"),
            InputField("Crack", "محل ترک:", 11, 3,
                       validation={"type":"list","formula":"داخل محدوده,خارج محدوده,خارج 5%"},
                       tooltip="محل ترک نسبت به دهانه میانی"),
        ],
        outputs=[
            OutputField("Flex", "مقاومت خمشی (MPa):", 13, 3,
                       formula="=IF(OR(C5=\"\",C6=\"\",C7=\"\",C8=\"\"),\"—\",IF(C11=\"خارج محدوده\",\"⚠️ باطل\",IF(C11=\"خارج 5%\",ROUND(3*C8*C7/(2*C5*C6^2),2)*0.85,IF(C10=\"مرکزی\",ROUND(3*C8*C7/(2*C5*C6^2),2),ROUND(C8*C7/(C5*C6^2),2)))))",
                       num_fmt="0.00", style="neutral", unit="MPa"),
        ],
        checks=[]
    ))

    # 4-4 UPV (ASTM C597) — educational classification
    register_test(TestSpec(
        id="4-4",
        title="سرعت پالس اولتراسونیک",
        standard=STANDARDS["C597"]["name"],
        edition=STANDARDS["C597"]["edition"],
        tab_color="F44336",
        sheet_name="20_آزمایش_4-4",
        summary_key="سرعت پالس",
        layout="matrix",
        custom_builder=build_matrix_input,
        inputs=[
            InputField("L", "طول مسیر (m):", 5, 3, "m",
                       validation={"type":"decimal","min":0,"max":10},
                       tooltip="طول مسیر پالس (متر)"),
            InputField("T1", "زمان ۱ (µs)", 7, 1, "µs",
                       validation={"type":"decimal","min":0,"max":100000},
                       tooltip="زمان عبور پالس - اندازه‌گیری ۱"),
            InputField("T2", "زمان ۲ (µs)", 7, 2, "µs",
                       validation={"type":"decimal","min":0,"max":100000},
                       tooltip="زمان عبور پالس - اندازه‌گیری ۲"),
            InputField("T3", "زمان ۳ (µs)", 7, 3, "µs",
                       validation={"type":"decimal","min":0,"max":100000},
                       tooltip="زمان عبور پالس - اندازه‌گیری ۳"),
        ],
        outputs=[
            OutputField("Velocity", "سرعت (km/s):", 9, 3,
                       formula="=IF(OR(C5=\"\",C6=\"\",C6=0),\"—\",ROUND(C5*1000/AVERAGE(C7:E7),2))",
                       num_fmt="0.00", style="neutral", unit="km/s"),
            OutputField("Stdev", "انحراف معیار:", 10, 3,
                       formula="=IF(C7=\"\",\"\",IFERROR(ROUND(STDEV(C7:E7),2),\"—\"))",
                       num_fmt="0.00", style="neutral", unit="µs"),
        ],
        checks=[
            Check("طبقه‌بندی کیفی (مرجع آموزشی)", "=IF(C9=\"—\",\"\",IF(C9>=4.5,\"عالی\",IF(C9>=3.5,\"خوب\",IF(C9>=3,\"متوسط\",\"ضعیف\"))))", 11, 3, severity="warning"),
        ]
    ))

    # 4-5 Schmidt rebound (ASTM C805) — with angle/surface correction
    register_test(TestSpec(
        id="4-5",
        title="چکش اشمیت",
        standard=STANDARDS["C805"]["name"],
        edition=STANDARDS["C805"]["edition"],
        tab_color="F44336",
        sheet_name="21_آزمایش_4-5",
        summary_key="چکش اشمیت (Rm)",
        layout="matrix",
        custom_builder=build_matrix_input,
        inputs=[
            InputField("R1", "R1", 6, 1, "", validation={"type":"decimal","min":0,"max":100},
                       tooltip="خوانش شماره ۱"),
            InputField("R2", "R2", 6, 2, "", validation={"type":"decimal","min":0,"max":100},
                       tooltip="خوانش شماره ۲"),
            InputField("R3", "R3", 6, 3, "", validation={"type":"decimal","min":0,"max":100},
                       tooltip="خوانش شماره ۳"),
            InputField("R4", "R4", 6, 4, "", validation={"type":"decimal","min":0,"max":100},
                       tooltip="خوانش شماره ۴"),
            InputField("R5", "R5", 6, 5, "", validation={"type":"decimal","min":0,"max":100},
                       tooltip="خوانش شماره ۵"),
            InputField("R6", "R6", 6, 6, "", validation={"type":"decimal","min":0,"max":100},
                       tooltip="خوانش شماره ۶"),
            InputField("R7", "R7", 6, 7, "", validation={"type":"decimal","min":0,"max":100},
                       tooltip="خوانش شماره ۷"),
            InputField("R8", "R8", 6, 8, "", validation={"type":"decimal","min":0,"max":100},
                       tooltip="خوانش شماره ۸"),
            InputField("R9", "R9", 7, 1, "", validation={"type":"decimal","min":0,"max":100},
                       tooltip="خوانش شماره ۹"),
            InputField("R10", "R10", 7, 2, "", validation={"type":"decimal","min":0,"max":100},
                       tooltip="خوانش شماره ۱۰"),
            InputField("R11", "R11", 7, 3, "", validation={"type":"decimal","min":0,"max":100},
                       tooltip="خوانش شماره ۱۱"),
            InputField("R12", "R12", 7, 4, "", validation={"type":"decimal","min":0,"max":100},
                       tooltip="خوانش شماره ۱۲"),
            InputField("R13", "R13", 7, 5, "", validation={"type":"decimal","min":0,"max":100},
                       tooltip="خوانش شماره ۱۳"),
            InputField("R14", "R14", 7, 6, "", validation={"type":"decimal","min":0,"max":100},
                       tooltip="خوانش شماره ۱۴"),
            InputField("R15", "R15", 7, 7, "", validation={"type":"decimal","min":0,"max":100},
                       tooltip="خوانش شماره ۱۵"),
            InputField("R16", "R16", 7, 8, "", validation={"type":"decimal","min":0,"max":100},
                       tooltip="خوانش شماره ۱۶"),
            InputField("Surface", "سطح:", 9, 3,
                       validation={"type":"list","formula":"خشک,مرطوب"},
                       tooltip="وضعیت سطح بتن"),
            InputField("Angle", "زاویه ضربه:", 10, 3,
                       validation={"type":"list","formula":"عمودی به پایین-90,عمودی به بالا+90,افقی 0,45 درجه پایین,45 درجه بالا"},
                       tooltip="زاویه چکش نسبت به افق"),
            InputField("Temp", "دما (°C):", 11, 3, "°C",
                       validation={"type":"decimal","min":-10,"max":60},
                       tooltip="دمای محیط در زمان آزمایش"),
        ],
        outputs=[
            OutputField("Avg", "میانگین کل:", 13, 3,
                       formula="=IF(COUNT(A6:H7)=0,\"\",ROUND(AVERAGE(A6:H7),1))",
                       num_fmt="0.0", style="neutral"),
            OutputField("Nvalid", "تعداد معتبر:", 14, 3,
                       formula="=IF(C13=\"\",\"\",COUNTIFS(A6:H7,\">=\"&C13-6,A6:H7,\"<=\"&C13+6))",
                       num_fmt="0", style="neutral"),
            OutputField("Rm", "میانگین معتبرها:", 15, 3,
                       formula="=IF(OR(C13=\"\",C14=0),\"—\",ROUND(SUMPRODUCT((ABS(IF(ISNUMBER(A6:H7),A6:H7,999)-C13)<=6)*IF(ISNUMBER(A6:H7),A6:H7,0))/C14,1))",
                       num_fmt="0.0", style="neutral"),
            OutputField("Rm_corr", "Rm اصلاح‌شده (مرجع):", 17, 3,
                       formula="=IF(OR(C15=\"—\",C9=\"\",C10=\"\",C11=\"\"),\"—\",ROUND(C15*IF(C9=\"مرطوب\",1.05,1)*IF(C11<10,1.03,1)*IF(C10=\"عمودی به پایین-90\",1.08,IF(C10=\"عمودی به بالا+90\",0.92,IF(C10=\"45 درجه پایین\",1.04,IF(C10=\"45 درجه بالا\",0.96,1)))),1))",
                       num_fmt="0.0", style="neutral",
                       tooltip="Rm اصلاح‌شده برای دما و زاویه (ضرایب مرجع ACI 228.1R). این مقدار تخمینی است و نیاز به کالیبراسیون ویژه پروژه دارد."),
        ],
        checks=[
            Check("وضعیت معتبرها", "=IF(C15=\"—\",\"\",IF(C14<COUNT(A6:H7)*0.8,\"❌ حذف >20% — تکرار\",\"✅\"))", 16, 3, severity="error"),
            Check("Disclaimer", "=IF(C17=\"—\",\"\",\"⚠️ تخمین مقاومت نیازمند کالیبراسیون اختصاصی\" & CHAR(10) & \"مرجع: ACI 228.1R\")", 18, 3, severity="warning"),
        ]
    ))

define_tests()

# ─── Workbook Builder ──────────────────────────────────────────────────
class WorkbookBuilder:
    def __init__(self):
        self.wb = Workbook()
        self.wb.remove(self.wb.active)
        self._named_ranges = {}
        self._result_map = {}

    def build(self, protect: bool = True, demo: bool = False) -> Workbook:
        self._build_guide()
        self._build_info()
        self._build_all_tests()
        self._build_reference_sheets()
        self._build_report()
        self._build_dashboard()
        self._build_qa()
        if protect and PASSWORD:
            self._protect_all()
        self._apply_global_settings()
        self._wb_calculation()
        if demo:
            self._fill_demo_data()
        return self.wb

    # ─── Helper methods ──────────────────────────────────────────────────
    def _cell(self, ws, row, col, value=None, style=None, num_fmt=None,
              locked=True, hyperlink=None, align=None, comment=None):
        cell = ws.cell(row=row, column=col, value=value)
        if style:
            STYLES.apply(cell, style)
        if num_fmt:
            cell.number_format = num_fmt
        cell.protection = Protection(locked=locked)
        if hyperlink:
            if isinstance(hyperlink, str):
                cell.hyperlink = Hyperlink(ref=cell.coordinate, location=hyperlink)
            else:
                cell.hyperlink = hyperlink
        if align == "right":
            cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
        elif align == "center":
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        elif align == "left":
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        if comment:
            cell.comment = Comment(comment, "سیستم")
            cell.comment.width = 200
            cell.comment.height = 80
        return cell

    @staticmethod
    def _merge(ws, r1, c1, r2, c2):
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)

    def _add_nav(self, ws):
        for col in range(1, 12):
            cell = ws.cell(row=1, column=col)
            STYLES.apply(cell, "nav")
        self._cell(ws, 1, 1, "🏠", hyperlink="'00_راهنما'!A1", align="center")
        self._cell(ws, 1, 3, "📋", hyperlink="'01_اطلاعات_آزمون'!A1", align="center")
        self._cell(ws, 1, 5, "📑", hyperlink="'22_گزارش'!A1", align="center")
        self._cell(ws, 1, 7, "📊", hyperlink="'23_داشبورد'!A1", align="center")
        self._cell(ws, 1, 9, "🔒", align="center")

    def _add_title(self, ws, title, subtitle=""):
        self._merge(ws, 2, 1, 2, 11)
        self._cell(ws, 2, 1, title, style="title", align="right")
        if subtitle:
            self._merge(ws, 3, 1, 3, 11)
            self._cell(ws, 3, 1, subtitle, style="subtitle", align="right")

    @staticmethod
    def _add_dv(ws, cells, dv_type, **kwargs):
        if dv_type == "decimal":
            dv = DataValidation(type="decimal", operator="between",
                                formula1=str(kwargs.get("min", 0)),
                                formula2=str(kwargs.get("max", 100000)),
                                allow_blank=kwargs.get("allow_blank", True))
        elif dv_type == "list":
            formula = kwargs.get("formula", "")
            dv = DataValidation(type="list", formula1=formula.replace(";", ","),
                                allow_blank=kwargs.get("allow_blank", True))
        elif dv_type == "custom":
            dv = DataValidation(type="custom", formula1=kwargs.get("formula", ""),
                                allow_blank=kwargs.get("allow_blank", True))
        elif dv_type == "textLength":
            dv = DataValidation(type="textLength", operator="greaterThan",
                                formula1="0",
                                allow_blank=kwargs.get("allow_blank", False))
        else:
            return
        dv.error = kwargs.get("error_msg", "مقدار نامعتبر")
        dv.errorTitle = "خطای ورودی"
        dv.prompt = kwargs.get("prompt", "")
        dv.promptTitle = "راهنما"
        dv.errorStyle = kwargs.get("error_style", "stop")
        ws.add_data_validation(dv)
        for cell_ref in cells:
            dv.add(cell_ref)

    def _apply_input_validation(self, ws, inputs):
        for inp in inputs:
            if inp.validation:
                val_copy = inp.validation.copy()
                dv_type = val_copy.pop("type", None)
                if dv_type:
                    self._add_dv(ws, [f"{get_column_letter(inp.col)}{inp.row}"],
                                 dv_type, **val_copy)
            if inp.tooltip or inp.instruction:
                comment_text = inp.tooltip
                if inp.instruction:
                    comment_text += "\n\n" + inp.instruction
                ws.cell(row=inp.row, column=inp.col).comment = Comment(comment_text, "سیستم")
                ws.cell(row=inp.row, column=inp.col).comment.width = 250
                ws.cell(row=inp.row, column=inp.col).comment.height = 100

    @staticmethod
    def _freeze_panes(ws, row=4, col=1):
        ws.freeze_panes = f"{get_column_letter(col)}{row}"

    @staticmethod
    def _setup_print(ws, orientation="portrait", fit_to_width=1, fit_to_height=1,
                     margins=(0.3, 0.3, 0.5, 0.5), header_text=None, footer_text=None):
      ws.page_setup.orientation = orientation
      ws.page_setup.paperSize = ws.PAPERSIZE_A4
      ws.page_setup.fitToWidth = fit_to_width
      ws.page_setup.fitToHeight = fit_to_height
      ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
      ws.print_options.horizontalCentered = True
      ws.page_margins = PageMargins(
        left=margins[0], right=margins[0],
        top=margins[2], bottom=margins[3]
      )

      # تنظیم هدر و فوتر – پشتیبانی از دو روش
      if hasattr(ws, 'header_footer'):
        # روش جدید (openpyxl >= 2.4)
        if header_text:
          ws.header_footer.header.center.text = header_text
        if footer_text:
          ws.header_footer.footer.center.text = footer_text
        ws.header_footer.footer.center.text += f" &P/&N  نسخه {VERSION}"
      else:
        # روش قدیمی (برای نسخه‌های بسیار قدیمی یا در صورت عدم دسترسی)
        if header_text:
          ws.oddHeader.center.text = header_text
        if footer_text:
          ws.oddFooter.center.text = footer_text
        ws.oddFooter.center.text += f" &P/&N  نسخه {VERSION}"

    @staticmethod
    def _set_widths(ws, widths=None):
        if widths is None:
            widths = [12, 14, 16, 14, 14, 14, 14, 14, 14, 14, 14]
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.sheet_view.zoomScale = 85
        ws.sheet_view.showGridLines = False

    # ─── Build sheets ────────────────────────────────────────────────────
    def _build_guide(self):
        ws = self.wb.create_sheet("00_راهنما")
        ws.sheet_properties.tabColor = "1F4E79"
        ws.sheet_view.rightToLeft = True
        self._add_nav(ws)
        self._add_title(ws, "🧪 همراه دیجیتال آزمایشگاه فناوری بتن",
                        f"نسخه {VERSION} | ساخت: {BUILD_DATE}")
        r = 5
        self._cell(ws, r, 1, "📌 راهنمای سریع", style="title", align="right")
        self._merge(ws, r, 1, r, 11)
        r += 1
        self._cell(ws, r, 1, "سلول‌های زرد = ورودی کاربر | خاکستری = محاسبه | سبز = قبول | نارنجی = هشدار | قرمز = خطا",
                   style="normal", align="right")
        self._merge(ws, r, 1, r, 11)
        r += 2
        self._cell(ws, r, 1, "⚠️ هشدار حقوقی:", style="title", align="right")
        self._merge(ws, r, 1, r, 11)
        r += 1
        self._cell(ws, r, 1,
                   "این نرم‌افزار صرفاً ابزار کمکی است و مسئولیت نهایی تأیید نتایج و ایمنی سازه بر عهده مهندس ناظر است.",
                   style="normal", align="right")
        self._merge(ws, r, 1, r, 11)
        r += 2
        self._cell(ws, r, 1, "📋 فهرست شیت‌ها:", style="title", align="right")
        self._merge(ws, r, 1, r, 11)
        r += 1
        for code, name, std in [
            ("01", "اطلاعات آزمون", ""),
            ("02", "۱-۱ دانه‌بندی", STANDARDS["C136"]["name"]),
            ("03", "۱-۲ رطوبت", STANDARDS["C566"]["name"]),
            ("04", "۱-۳ چگالی درشت", STANDARDS["C127"]["name"]),
            ("05", "۱-۴ چگالی ریز", STANDARDS["C128"]["name"]),
            ("06", "۱-۵ وزن واحد سنگدانه", STANDARDS["C29"]["name"]),
            ("07", "۱-۶ معادل ماسه", STANDARDS["D2419"]["name"]),
            ("08", "۱-۷ شاخص شکل", STANDARDS["D4791"]["name"]),
            ("09", "۱-۸ جذب آب", STANDARDS["C127"]["name"]),
            ("10", "۲-۱ چگالی بتن تازه", STANDARDS["C138"]["name"]),
            ("11", "۲-۲ قوام نرمال", STANDARDS["C187"]["name"] + " (مرجع)"),
            ("12", "۲-۳ زمان گیرش", STANDARDS["C191"]["name"]),
            ("13", "۲-۴ مقاومت ملات", STANDARDS["EN196-1"]["name"]),
            ("14", "۳-۱ اسلامپ", STANDARDS["C143"]["name"]),
            ("15", "۳-۲ آب‌انداختگی", STANDARDS["C232"]["name"] + " (تقریبی)"),
            ("16", "۳-۳ وزن واحد بتن", STANDARDS["C138"]["name"]),
            ("17", "۴-۱ مقاومت فشاری", STANDARDS["C39"]["name"]),
            ("18", "۴-۲ مقاومت کششی", STANDARDS["C496"]["name"]),
            ("19", "۴-۳ مقاومت خمشی", STANDARDS["C78"]["name"] + "/" + STANDARDS["C293"]["name"]),
            ("20", "۴-۴ اولتراسونیک", STANDARDS["C597"]["name"]),
            ("21", "۴-۵ چکش اشمیت", STANDARDS["C805"]["name"]),
            ("22", "گزارش", ""),
            ("23", "داشبورد", ""),
            ("24", "QA Test", ""),
        ]:
            self._cell(ws, r, 1, code, style="num", align="right")
            self._cell(ws, r, 2, name, style="normal", align="right")
            self._cell(ws, r, 6, std, style="num", align="right")
            self._merge(ws, r, 2, r, 5)
            self._merge(ws, r, 6, r, 11)
            r += 1
        self._setup_print(ws, orientation="landscape", footer_text="راهنما")
        self._freeze_panes(ws, row=4)
        self._set_widths(ws)

    def _build_info(self):
        ws = self.wb.create_sheet("01_اطلاعات_آزمون")
        ws.sheet_properties.tabColor = "2E7D32"
        ws.sheet_view.rightToLeft = True
        self._add_nav(ws)
        self._add_title(ws, "📋 اطلاعات آزمون", "این اطلاعات در گزارش تکرار می‌شود")
        fields = ["نام پروژه", "شماره نمونه", "تاریخ آزمون", "نام اپراتور",
                  "دستگاه / تجهیزات", "دمای محیط (°C)", "رطوبت نسبی (%)", "استاندارد مرجع", "توضیحات"]
        r = 5
        for i, label in enumerate(fields):
            self._cell(ws, r+i, 1, label, style="normal", align="right")
            self._merge(ws, r+i, 1, r+i, 2)
            self._cell(ws, r+i, 3, None, style="input", locked=False, align="right")
            self._merge(ws, r+i, 3, r+i, 6)
            self._add_dv(ws, [f"C{r+i}"], "textLength", allow_blank=False, error_msg="تکمیل این فیلد الزامی است")
        self._setup_print(ws, footer_text="اطلاعات آزمون")
        self._freeze_panes(ws, row=4)
        self._set_widths(ws)

    def _build_all_tests(self):
        for test_id, spec in TEST_REGISTRY.items():
            self._build_test_sheet(spec)

    def _build_test_sheet(self, spec: TestSpec):
        ws = self.wb.create_sheet(spec.sheet_name)
        ws.sheet_properties.tabColor = spec.tab_color
        ws.sheet_view.rightToLeft = True
        self._add_nav(ws)
        self._add_title(ws, spec.title, f"{spec.standard} — نسخه {spec.edition}")
        if spec.subtitle:
            self._cell(ws, 4, 1, spec.subtitle, style="subtitle", align="right")
            self._merge(ws, 4, 1, 4, 11)

        if spec.custom_builder:
            spec.custom_builder(ws, spec, self.wb)
        else:
            self._build_simple_layout(ws, spec)

        self._apply_input_validation(ws, spec.inputs)
        self._add_status_cf(ws, spec)

        for out in spec.outputs:
            if spec.summary_key:
                self._result_map[spec.summary_key] = (spec.sheet_name, cell_ref(out.row, out.col))

        self._setup_print(ws, footer_text=spec.title)
        self._freeze_panes(ws, row=4)
        self._set_widths(ws)

    def _build_simple_layout(self, ws, spec):
        for inp in spec.inputs:
            self._cell(ws, inp.row, 1, inp.label, style="normal", align="right")
            self._merge(ws, inp.row, 1, inp.row, 2)
            self._cell(ws, inp.row, inp.col, None, style="input", locked=False, align="right")
            if inp.unit:
                self._cell(ws, inp.row, inp.col+1, f"[{inp.unit}]", style="num", align="right")
        for out in spec.outputs:
            self._cell(ws, out.row, 1, out.label, style="normal", align="right")
            self._merge(ws, out.row, 1, out.row, 2)
            self._cell(ws, out.row, out.col, out.formula, style=out.style,
                       num_fmt=out.num_fmt, locked=True, align="right")
            if out.tooltip:
                ws.cell(row=out.row, column=out.col).comment = Comment(out.tooltip, "سیستم")
        for chk in spec.checks:
            self._cell(ws, chk.row, 1, chk.label, style="normal", align="right")
            self._merge(ws, chk.row, 1, chk.row, 2)
            self._cell(ws, chk.row, chk.col, chk.formula, style="calc", locked=True, align="right")

    @staticmethod
    def _add_status_cf(ws, spec):
        for chk in spec.checks:
            ref = f"{get_column_letter(chk.col)}{chk.row}"
            ws.conditional_formatting.add(ref, FormulaRule(
                formula=[f'ISNUMBER(FIND("✅", {ref}))'],
                fill=PatternFill("solid", fgColor=STYLES.COLORS["pass_fill"]),
                font=Font(color=STYLES.COLORS["pass_font"])
            ))
            ws.conditional_formatting.add(ref, FormulaRule(
                formula=[f'ISNUMBER(FIND("❌", {ref}))'],
                fill=PatternFill("solid", fgColor=STYLES.COLORS["fail_fill"]),
                font=Font(color=STYLES.COLORS["fail_font"])
            ))
            ws.conditional_formatting.add(ref, FormulaRule(
                formula=[f'ISNUMBER(FIND("⚠️", {ref}))'],
                fill=PatternFill("solid", fgColor=STYLES.COLORS["warn_fill"]),
                font=Font(color=STYLES.COLORS["warn_font"])
            ))

    # ─── Reference Sheets ──────────────────────────────────────────────
    def _build_reference_sheets(self):
        ws = self.wb.create_sheet("_Standards")
        ws.sheet_state = "hidden"
        ws.append(["ID", "Standard", "Edition", "Status", "Title"])
        for key, std in STANDARDS.items():
            ws.append([key, std["name"], std["edition"], std.get("status",""), std["title"]])
        ws = self.wb.create_sheet("_ISIRI_Limits")
        ws.sheet_state = "hidden"
        ws.append(["Sieve (mm)", "Upper (%)", "Lower (%)"])
        for size, (upper, lower) in ISIRI_LIMITS.items():
            ws.append([size, upper, lower])
        ws = self.wb.create_sheet("_Errata")
        ws.sheet_state = "hidden"
        ws.append(["ID", "Sheet", "Type", "Description", "Standard", "Status", "Date"])
        errata = [
            ("1-4j", "1-4", "Physical", "OD=1.51 vs SSD=2.63", "ASTM C128", "✅ Fixed", "2026-08-13"),
            ("2-4", "2-4", "Formula", "3340 kgf → 11.9 MPa", "EN 196-1", "✅ Fixed", "2026-08-13"),
            ("1-6", "1-6", "Formula", "ROUNDUP → ROUND per standard", "ASTM D2419", "✅ Fixed", "2026-08-22"),
            ("3-1", "3-1", "Formula", "304.8 → 300 mm (metric standard)", "ASTM C143M", "✅ Fixed", "2026-08-22"),
            ("4-5", "4-5", "Coefficient", "Wet surface correction reversed", "ACI 228.1R", "✅ Fixed", "2026-08-22"),
            ("4-5", "4-5", "Coefficient", "Added 45° angle corrections", "ACI 228.1R", "✅ Fixed", "2026-08-22"),
            ("1-5", "1-5", "Formula", "Voids referenced C9 → C10", "ASTM C29", "✅ Fixed", "2026-08-22"),
        ]
        for row in errata:
            ws.append(row)
        ws = self.wb.create_sheet("_Validation_Data")
        ws.sheet_state = "hidden"
        ws.append(["Test ID", "Case", "Inputs", "Expected", "Tolerance", "Status"])
        golden = [
            ("1-2", "Normal", "W1=500,W2=480", "4.17", "0.01", "Pending"),
            ("1-3", "Normal", "A=2500,B=2600,C=1500", "2.27", "0.01", "Pending"),
            ("1-6", "Normal", "Sand=40,Clay=10", "400.0", "0.1", "Pending"),
            ("3-1", "Normal", "H=50", "250", "5", "Pending"),
            ("4-1", "Normal", "D=150,L=300,P=715", "40.4", "0.1", "Pending"),
        ]
        for row in golden:
            ws.append(row)
        ws = self.wb.create_sheet("_Glossary")
        ws.sheet_state = "hidden"
        ws.append(["Persian", "English", "Definition"])
        ws = self.wb.create_sheet("_Materials_DB")
        ws.sheet_state = "hidden"
        ws.append(["Material", "Property", "Value", "Unit"])

    # ─── Report ──────────────────────────────────────────────────────────
    def _build_report(self):
        ws = self.wb.create_sheet("22_گزارش")
        ws.sheet_properties.tabColor = "9C27B0"
        ws.sheet_view.rightToLeft = True
        self._add_nav(ws)
        self._add_title(ws, "📑 گزارش آزمایشگاهی", "خلاصه نتایج")
        r = 5
        self._cell(ws, r, 1, "اطلاعات پروژه:", style="title", align="right")
        self._merge(ws, r, 1, r, 11)
        r += 1
        info_labels = ["نام پروژه", "شماره نمونه", "تاریخ", "اپراتور", "استاندارد"]
        for i, label in enumerate(info_labels):
            self._cell(ws, r+i, 1, label, style="normal", align="right")
            self._cell(ws, r+i, 2, f"='01_اطلاعات_آزمون'!C{5+i}", style="calc", align="right")
            self._merge(ws, r+i, 2, r+i, 6)
        r += len(info_labels) + 1
        self._cell(ws, r, 1, "خلاصه نتایج:", style="title", align="right")
        self._merge(ws, r, 1, r, 11)
        r += 1
        headers = ["آزمایش", "نتیجه", "واحد", "وضعیت"]
        for col, h in enumerate(headers, 1):
            self._cell(ws, r, col, h, style="header", align="center")
        r += 1
        for summary_key in sorted(self._result_map.keys()):
            sheet, cell = self._result_map[summary_key]
            unit = ""
            for spec in TEST_REGISTRY.values():
                if spec.summary_key == summary_key:
                    for out in spec.outputs:
                        if out.key in sheet or out.key == summary_key:
                            unit = out.unit
                            break
                    break
            self._cell(ws, r, 1, summary_key, style="normal", align="right")
            self._cell(ws, r, 2, f"='{sheet}'!{cell}", style="calc", num_fmt="0.0", align="right")
            self._cell(ws, r, 3, unit, style="num", align="right")
            status_formula = f'=IF(ISNUMBER(B{r}),"✅",IF(B{r}="—","انجام نشده",B{r}))'
            self._cell(ws, r, 4, status_formula, style="calc", align="right")
            ws.conditional_formatting.add(f"D{r}", FormulaRule(
                formula=[f'ISNUMBER(FIND("✅", D{r}))'],
                fill=PatternFill("solid", fgColor=STYLES.COLORS["pass_fill"]),
                font=Font(color=STYLES.COLORS["pass_font"])
            ))
            ws.conditional_formatting.add(f"D{r}", FormulaRule(
                formula=[f'ISNUMBER(FIND("❌", D{r}))'],
                fill=PatternFill("solid", fgColor=STYLES.COLORS["fail_fill"]),
                font=Font(color=STYLES.COLORS["fail_font"])
            ))
            ws.conditional_formatting.add(f"D{r}", FormulaRule(
                formula=[f'ISNUMBER(FIND("⚠️", D{r}))'],
                fill=PatternFill("solid", fgColor=STYLES.COLORS["warn_fill"]),
                font=Font(color=STYLES.COLORS["warn_font"])
            ))
            r += 1
        r += 2
        self._cell(ws, r, 1, "امضاها:", style="title", align="right")
        self._merge(ws, r, 1, r, 11)
        r += 1
        self._cell(ws, r, 1, "تکنسین:", style="normal", align="right")
        self._cell(ws, r, 3, "_________________", style="normal", align="right")
        self._merge(ws, r, 3, r, 6)
        r += 1
        self._cell(ws, r, 1, "مسئول فنی:", style="normal", align="right")
        self._cell(ws, r, 3, "_________________", style="normal", align="right")
        self._merge(ws, r, 3, r, 6)
        self._setup_print(ws, orientation="landscape", footer_text="گزارش")
        self._freeze_panes(ws, row=4)
        self._set_widths(ws)

    # ─── Dashboard ──────────────────────────────────────────────────────
    def _build_dashboard(self):
        ws = self.wb.create_sheet("23_داشبورد")
        ws.sheet_properties.tabColor = "00BCD4"
        ws.sheet_view.rightToLeft = True
        self._add_nav(ws)
        self._add_title(ws, "📊 داشبورد پروژه", "وضعیت کلی")
        r = 5
        self._cell(ws, r, 1, "تعداد کل آزمایش‌ها:", style="normal", align="right")
        self._cell(ws, r, 3, len(TEST_REGISTRY), style="calc", num_fmt="0", align="right")
        r += 1
        self._cell(ws, r, 1, "شیت‌های پیاده‌سازی‌شده:", style="normal", align="right")
        self._cell(ws, r, 3, len(TEST_REGISTRY), style="calc", num_fmt="0", align="right")
        r += 1
        input_ranges = []
        for spec in TEST_REGISTRY.values():
            for inp in spec.inputs:
                input_ranges.append(f"'{spec.sheet_name}'!{get_column_letter(inp.col)}{inp.row}")
        if input_ranges:
            range_str = ",".join(input_ranges)
            total_cells = len(input_ranges)
            progress_formula = f'=IFERROR(COUNTA({range_str})/{total_cells},0)'
            self._cell(ws, r, 1, "درصد تکمیل داده‌های ورودی:", style="normal", align="right")
            self._cell(ws, r, 3, progress_formula, style="calc", num_fmt="0%", align="right")
        r += 2
        self._cell(ws, r, 1, "نسخه:", style="normal", align="right")
        self._cell(ws, r, 3, VERSION, style="calc", align="right")
        r += 1
        self._cell(ws, r, 1, "تاریخ ساخت:", style="normal", align="right")
        self._cell(ws, r, 3, BUILD_DATE, style="calc", align="right")
        self._setup_print(ws, footer_text="داشبورد")
        self._freeze_panes(ws, row=4)
        self._set_widths(ws)

    # ─── QA Test ──────────────────────────────────────────────────────────
    def _build_qa(self):
        ws = self.wb.create_sheet("24_QA_Test")
        ws.sheet_properties.tabColor = "FF9800"
        ws.sheet_view.rightToLeft = True
        self._add_nav(ws)
        self._add_title(ws, "🧪 QA Test — تست‌های خودکار", "اعتبارسنجی فرمول‌ها")
        r = 5
        headers = ["Test ID", "شرح", "ورودی", "انتظار", "نتیجه", "وضعیت"]
        for col, h in enumerate(headers, 1):
            self._cell(ws, r, col, h, style="header", align="center")
        r += 1
        tests = [
            ("T-001", "ورودی خالی", "همه سلول‌ها پاک", "بدون #DIV/0!",
             '=IF(COUNTA(\'03_آزمایش_1-2\'!C5:C6)=0,"✅ PASS","—")'),
            ("T-002", "مرزی: W2=0", "W1=100, W2=0", "نمایش '—'",
             '=IF(\'03_آزمایش_1-2\'!C8="—","✅ PASS","❌ FAIL")'),
            ("T-003", "نمونه کتاب ۲-۴", "3340 kgf, A=1600", "≈20.5 MPa",
             '=IF(ABS(\'13_آزمایش_2-4\'!D12-20.5)<0.2,"✅ PASS","❌ FAIL")'),
            ("T-004", "نمونه کتاب ۴-۱", "d=150, P=715 kN", "≈40.4 MPa",
             '=IF(ABS(\'17_آزمایش_4-1\'!C14-40.4)<0.2,"✅ PASS","❌ FAIL")'),
            ("T-005", "SE فرمول استاندارد", "Sand=40, Clay=10", "400.0",
             '=IF(ABS(\'07_آزمایش_1-6\'!C8-400)<0.1,"✅ PASS","❌ FAIL")'),
            ("T-006", "اسلامپ متریک", "H=50", "250 mm",
             '=IF(\'14_آزمایش_3-1\'!C8=250,"✅ PASS","❌ FAIL")'),
            ("T-007", "Voids مرجع 1-5", "T=1000,G=2600,V=1000,S=2.6", "≈38.5%",
             '=IF(ABS(\'06_آزمایش_1-5\'!C11-38.5)<0.1,"✅ PASS","❌ FAIL")'),
        ]
        for test_id, desc, inp, expected, formula in tests:
            self._cell(ws, r, 1, test_id, style="num", align="right")
            self._cell(ws, r, 2, desc, style="normal", align="right")
            self._cell(ws, r, 3, inp, style="normal", align="right")
            self._cell(ws, r, 4, expected, style="normal", align="right")
            self._cell(ws, r, 5, formula, style="calc", align="right")
            status_cell = ws.cell(row=r, column=6)
            status_cell.value = f'=IF(ISNUMBER(FIND("PASS",E{r})),"✅",IF(ISNUMBER(FIND("FAIL",E{r})),"❌","—"))'
            status_cell.font = Font(name=STYLES.FONT_FAMILY, size=11)
            status_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            r += 1
        self._setup_print(ws, footer_text="QA Test")
        self._freeze_panes(ws, row=4)
        self._set_widths(ws)

    # ─── Protection & Global Settings ──────────────────────────────────
    def _protect_all(self):
        for ws in self.wb.worksheets:
            if ws.title.startswith("_"):
                continue
            ws.protection.sheet = True
            ws.protection.password = PASSWORD
            ws.protection.selectLockedCells = False
            ws.protection.selectUnlockedCells = True
            ws.protection.formatCells = False
            ws.protection.formatColumns = False
            ws.protection.formatRows = False
            ws.protection.insertColumns = False
            ws.protection.insertRows = False
            ws.protection.deleteColumns = False
            ws.protection.deleteRows = False
        self.wb.security.lockStructure = True
        self.wb.security.workbookPassword = PASSWORD

    def _apply_global_settings(self):
        for ws in self.wb.worksheets:
            ws.sheet_view.zoomScale = 80

    def _wb_calculation(self):
        self.wb.calculation.fullCalcOnLoad = True

    def _fill_demo_data(self):
        pass

# ─── Main ──────────────────────────────────────────────────────────────────
def compute_sha256(filepath: Path) -> str:
    sha256 = hashlib.sha256()
    with open(filepath, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            sha256.update(chunk)
    return sha256.hexdigest()

def main():
    parser = argparse.ArgumentParser(description="Concrete Lab Companion Generator")
    parser.add_argument("--output", default="output", help="Output directory")
    parser.add_argument("--no-protect", action="store_true", help="Disable sheet protection")
    parser.add_argument("--demo", action="store_true", help="Fill demo data")
    parser.add_argument("--validate", action="store_true", help="Validate formulas without building")
    parser.add_argument("--version", action="version", version=f"v{VERSION}")
    args = parser.parse_args()

    logger.info("═" * 50)
    logger.info(f" Concrete Lab Companion Generator v{VERSION}")
    logger.info("═" * 50)

    if args.validate:
        logger.info("🔍 Validation mode: checking registry...")
        logger.info(f"Tests registered: {len(TEST_REGISTRY)}")
        for tid, spec in TEST_REGISTRY.items():
            logger.info(f"  {tid}: {spec.title} — {len(spec.outputs)} outputs, {len(spec.checks)} checks")
        logger.info("✅ Validation complete")
        sys.exit(0)

    builder = WorkbookBuilder()
    wb = builder.build(protect=not args.no_protect, demo=args.demo)

    output_dir = Path(args.output)
    output_dir.mkdir(exist_ok=True)
    filename = f"Concrete_Lab_Companion_v{VERSION}.xlsx"
    output_path = output_dir / filename

    if output_path.exists():
        try:
            output_path.unlink()
        except PermissionError:
            logger.error("فایل قبلی باز است. لطفاً آن را ببندید و دوباره اجرا کنید.")
            sys.exit(1)

    wb.save(str(output_path))
    logger.info(f"💾 Saved: {output_path}")

    sha = compute_sha256(output_path)
    sha_path = output_path.parent / (output_path.name + ".sha256")
    with open(sha_path, "w", encoding="utf-8") as f:
        f.write(f"{sha}  {filename}\n")
    logger.info(f"🔐 SHA-256: {sha}")

    manifest = {
        "version": VERSION,
        "build_date": BUILD_DATE,
        "filename": filename,
        "sha256": sha,
        "tests": len(TEST_REGISTRY),
        "sheets": [ws.title for ws in wb.worksheets if not ws.title.startswith("_")],
    }
    manifest_path = output_path.parent / (output_path.stem + ".json")
    with open(manifest_path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2, ensure_ascii=False)
    logger.info(f"📄 Manifest: {manifest_path}")

    logger.info("✅ BUILD COMPLETE")
    logger.info("═" * 50)

if __name__ == "__main__":
    main()
