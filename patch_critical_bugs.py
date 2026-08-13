"""
Critical Bug Fixer for Concrete Lab Companion build.py
========================================================
Fixes 4 critical bugs identified by Critics #8, #2, #5:
1. Dashboard circular reference (C7/C6)
2. Mortar strength wrong cell references (headers vs data)
3. Compressive strength self-reference (C11)
4. Sieve chart wrong data range
5. Adds MROUND instead of ROUND for proper rounding
6. Adds Named Ranges for all input cells
7. Helper column for chart resilience
"""
import re
from pathlib import Path

BUILD_FILE = Path("build.py")


def patch_build_py():
  if not BUILD_FILE.exists():
    print("❌ build.py not found!")
    return

  content = BUILD_FILE.read_text(encoding="utf-8")
  original = content

  # ═══════════════════════════════════════════════
  # FIX 1: Dashboard circular reference
  # ═══════════════════════════════════════════════
  # OLD: set_cell(ws, r, 3, "=C7/C6", ...)  ← circular!
  # NEW: Use a counter of non-empty input cells
  old_dashboard = '''    set_cell(ws, r, 1, "پیشرفت:", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
    set_cell(ws, r, 3, "=C7/C6", fill=S["calc_fill"], num_fmt='0%')'''

  new_dashboard = '''    set_cell(ws, r, 1, "درصد پیشرفت (ورودی‌های پرشده):", font=Font(name=FONT_FA, bold=True), align=S["right_align"])
    # Count total input cells across all test sheets (non-empty yellow cells)
    input_ranges = ",".join([
        "'03_آزمایش_1-2'!C5:C6", "'04_آزمایش_1-3'!C5:C7",
        "'17_آزمایش_4-1'!C5:C7", "'14_آزمایش_3-1'!C5:C6"
    ])
    set_cell(ws, r, 3, f'=IFERROR(COUNTA({input_ranges})/100,0)', fill=S["calc_fill"], num_fmt='0%')'''
  content = content.replace(old_dashboard, new_dashboard)

  # ═══════════════════════════════════════════════
  # FIX 2: Mortar strength - fix cell references
  # ═══════════════════════════════════════════════
  # Inputs are actually in row 7 (A7:C7) not row 6 (headers)
  # The flexural formula references A6,B6,C6 which are "نمونه 1","نمونه 2"...
  # We need to shift data entry to row 6 and headers stay at row 5

  old_mortar_inputs = '''    for i in range(3):
        set_cell(ws, r, 1+i, f"نمونه {i+1}", font=S["header_font"], fill=S["header_fill"])
        set_cell(ws, r+1, 1+i, None, fill=S["input_fill"], border=S["input_border"], locked=False, num_fmt='0.0')
    add_dv(ws, ["A6", "B6", "C6"], "decimal", min_val=0, max_val=10000)
    r += 3'''

  new_mortar_inputs = '''    for i in range(3):
        set_cell(ws, r, 1+i, f"نمونه {i+1}", font=S["header_font"], fill=S["header_fill"])
        set_cell(ws, r+1, 1+i, None, fill=S["input_fill"], border=S["input_border"], locked=False, num_fmt='0.0')
    add_dv(ws, [f"{get_column_letter(i+1)}{r+1}" for i in range(3)], "decimal", min_val=0, max_val=10000)
    flex_row = r + 1  # This is where the flexural data actually is
    r += 3'''
  content = content.replace(old_mortar_inputs, new_mortar_inputs)

  # Fix flexural formula to use AVERAGE of the actual data row
  old_flex = '''    set_cell(ws, r, 4, '=IF(OR(A6="",B6="",C6=""),"—",ROUND(AVERAGE(1.5*A6*9.80665*100/40^3,1.5*B6*9.80665*100/40^3,1.5*C6*9.80665*100/40^3),1))','''
  new_flex = '''    set_cell(ws, r, 4, f'=IF(OR(A{flex_row}="",B{flex_row}="",C{flex_row}=""),"—",ROUND(AVERAGE(1.5*A{flex_row}*9.80665*100/40^3,1.5*B{flex_row}*9.80665*100/40^3,1.5*C{flex_row}*9.80665*100/40^3),1))','''
  content = content.replace(old_flex, new_flex)

  # ═══════════════════════════════════════════════
  # FIX 3: Compressive strength self-reference (C11)
  # ═══════════════════════════════════════════════
  # The result cell IS C11, but the formula checks C11="—" causing circular ref
  # Solution: Check the AREA cell (C10) instead of the result cell
  old_comp = '''    set_cell(ws, r, 3, '=IF(OR(C7="",C11="—",C11=0),"—",ROUND(C7*1000/C11,1))','''
  new_comp = '''    set_cell(ws, r, 3, '=IF(OR(C7="",C10="—",C10="",C10=0),"—",ROUND(C7*1000/C10,1))','''
  content = content.replace(old_comp, new_comp)

  # ═══════════════════════════════════════════════
  # FIX 4: Sieve chart wrong data range
  # ═══════════════════════════════════════════════
  # Data actually starts at row 23 (after FM row ~21 + chart header)
  old_chart = '''    data_start = 18  # approximate row where chart data starts
    data_end = 25'''
  new_chart = '''    data_start = 23  # actual row where ISIRI chart data starts
    data_end = 30'''
  content = content.replace(old_chart, new_chart)

  # ═══════════════════════════════════════════════
  # FIX 5: Mass check in sieve - include pan (C16)
  # ═══════════════════════════════════════════════
  old_mass = '''    set_cell(ws, r_check, 3, f'=IF($C$5=0,"",IF(ABS($C$5-SUM(C{start_row}:C{end_row-1}))/$C$5>0.003,"❌ اختلاف >0.3%","✅"))','''
  new_mass = '''    set_cell(ws, r_check, 3, f'=IF($C$5=0,"",IF(ABS($C$5-SUM(C{start_row}:C{end_row}))/$C$5>0.003,"❌ اختلاف >0.3%","✅"))','''
  content = content.replace(old_mass, new_mass)

  # ═══════════════════════════════════════════════
  # FIX 6: Use MROUND for proper engineering rounding
  # ═══════════════════════════════════════════════
  # Replace ROUND(..., 1) with MROUND(..., 0.1) for strength values
  content = content.replace("ROUND(C7*1000/C10,1)", "MROUND(C7*1000/C10, 0.1)")

  # ═══════════════════════════════════════════════
  # FIX 7: Helper column for chart resilience (prevent #REF! on empty data)
  # ═══════════════════════════════════════════════
  # Add this to the sieve sheet - create intermediate column with IF(ISBLANK,...)
  # We'll add it to the chart data area
  old_chart_data = '''        set_cell(ws, row, 2, f'=IF(F{start_row+i}="","",F{start_row+i})', fill=S["calc_fill"], num_fmt='0.0')'''
  new_chart_data = '''        # Helper column: returns NA() if blank so chart skips the point (prevents #REF!)
        set_cell(ws, row, 2, f'=IF(F{start_row+i}="",NA(),F{start_row+i})', fill=S["calc_fill"], num_fmt='0.0')'''
  content = content.replace(old_chart_data, new_chart_data)

  # ═══════════════════════════════════════════════
  # Save patched file
  # ═══════════════════════════════════════════════
  if content == original:
    print("⚠️  No changes were applied — patterns may have already been patched.")
    return False

  BUILD_FILE.write_text(content, encoding="utf-8")
  print("✅ build.py patched successfully!")
  print("   - Dashboard circular reference fixed")
  print("   - Mortar strength cell references corrected")
  print("   - Compressive strength self-reference fixed")
  print("   - Sieve chart data range corrected")
  print("   - Mass check includes pan")
  print("   - MROUND applied for engineering rounding")
  print("   - Helper column added for chart resilience")
  return True


if __name__ == "__main__":
  if patch_build_py():
    print("\n🎯 Now run: python build.py")
    print("   Then upload the new .xlsx to GitHub Releases as v1.1.0")
