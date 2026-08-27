import datetime
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import BarChart, LineChart, PieChart, Reference

wb = Workbook()

BUSINESS_NAME = "PREMIUM BAKERY"
FONT_NAME = "Arial"
TITLE_FONT_NAME = "Georgia"

# ---------- Premium bakery palette: espresso brown, gold, cream ----------
COL_ESPRESSO = "3E2723"
COL_BROWN = "6D4C41"
COL_GOLD = "C9A227"
COL_GOLD_LIGHT = "E9D9A0"
COL_CREAM = "FFF8ED"
COL_INCOME_GREEN = "3F6B32"
COL_EXPENSE_MAROON = "8E3B3B"
COL_SLATE = "37474F"

# ---------- Styles (sized up a bit for easy reading) ----------
HEADER_FONT = Font(name=FONT_NAME, size=12, bold=True, color="FFFFFF")
SECTION_FONT = Font(name=TITLE_FONT_NAME, size=13, bold=True, color="FFFFFF")
LABEL_FONT = Font(name=FONT_NAME, size=11, bold=True, color="5C4033")
KPI_VALUE_FONT = Font(name=FONT_NAME, size=20, bold=True, color=COL_INCOME_GREEN)
KPI_VALUE_FONT_RED = Font(name=FONT_NAME, size=20, bold=True, color=COL_EXPENSE_MAROON)
NORMAL_FONT = Font(name=FONT_NAME, size=11, color="000000")
EXAMPLE_FONT = Font(name=FONT_NAME, size=11, italic=True, color="1F5FA8")
NOTE_FONT = Font(name=FONT_NAME, size=10, italic=True, color="8A7A6B")
HOWTO_FONT = Font(name=FONT_NAME, size=13, color="3E2723")
HOWTO_HEAD_FONT = Font(name=TITLE_FONT_NAME, size=15, bold=True, color=COL_INCOME_GREEN)

HEADER_FILL_INCOME = PatternFill("solid", fgColor=COL_INCOME_GREEN)
HEADER_FILL_EXPENSE = PatternFill("solid", fgColor=COL_EXPENSE_MAROON)
HEADER_FILL_DASH = PatternFill("solid", fgColor=COL_ESPRESSO)
HEADER_FILL_SUMMARY = PatternFill("solid", fgColor=COL_BROWN)
HEADER_FILL_CAT = PatternFill("solid", fgColor=COL_SLATE)
KPI_FILL = PatternFill("solid", fgColor=COL_CREAM)
SECTION_FILL = PatternFill("solid", fgColor=COL_BROWN)
TOTAL_FILL = PatternFill("solid", fgColor=COL_GOLD_LIGHT)
EXAMPLE_FILL = PatternFill("solid", fgColor="FFF9E6")
HOWTO_FILL = PatternFill("solid", fgColor=COL_CREAM)
GOLD_FILL = PatternFill("solid", fgColor=COL_GOLD)

thin = Side(style="thin", color="D9D9D9")
BORDER_ALL = Border(left=thin, right=thin, top=thin, bottom=thin)

def gold_accent_bar(ws, row, last_col, height=5):
    """A slim gold accent stripe under a title banner, for a premium touch."""
    ws.row_dimensions[row].height = height
    for c in range(1, last_col + 1):
        ws.cell(row=row, column=c).fill = GOLD_FILL

CURRENCY_FMT = '"R"#,##0;[RED]-"R"#,##0'
PERCENT_FMT = '0.0%'
DATE_FMT = 'dd-mmm-yyyy'

DATA_ROWS_END = 5002  # rows 5..5002 => 4,998 entries per log (~13+ years of daily entries)

def style_header_row(ws, row, last_col, fill, font=HEADER_FONT, height=24):
    ws.row_dimensions[row].height = height
    for c in range(1, last_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER_ALL

def autosize(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

# ---------- Sheet protection helpers ----------
# openpyxl cells default to locked=True. Locking only takes effect once a
# sheet's protection is switched on. No password is set anywhere in this
# workbook — this is meant to prevent accidental edits, not to be a real
# security barrier.

def lock_whole_sheet(ws):
    """Protect a sheet with everything locked (no editable cells)."""
    ws.protection.sheet = True

def protect_with_unlocked_cells(ws, unlocked_addresses):
    """Protect a sheet, but leave specific cell addresses editable."""
    ws.protection.sheet = True
    for addr in unlocked_addresses:
        ws[addr].protection = Protection(locked=False)

def unlock_range(ws, min_row, max_row, min_col, max_col):
    """Mark a rectangular block of cells as editable (used before enabling
    sheet protection, so the data-entry area of Income/Expenses stays open)."""
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            ws.cell(row=r, column=c).protection = Protection(locked=False)

# =====================================================================
# 0. HOW TO USE (plain language, first tab)
# =====================================================================
howto = wb.active
howto.title = "How To Use"
howto.sheet_view.showGridLines = False

howto.merge_cells("A1:F1")
howto["A1"] = f"🧁 {BUSINESS_NAME} — How To Use This Book"
howto["A1"].font = Font(name=TITLE_FONT_NAME, size=20, bold=True, color="FFFFFF")
howto["A1"].fill = HEADER_FILL_DASH
howto["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
for c in range(1, 7):
    howto.cell(row=1, column=c).fill = HEADER_FILL_DASH
howto.row_dimensions[1].height = 34
gold_accent_bar(howto, 2, 6)

steps = [
    ("1. Money coming in?", "Go to the green \"Income\" tab at the bottom. Write down the date, what the order was for, pick a Category from the drop-down list, pick how you were paid, and type the Amount."),
    ("2. Money going out?", "Go to the red \"Expenses\" tab at the bottom. Do the same thing — date, what you bought, category, payment method, and amount."),
    ("3. See how the business is doing?", "Go to the \"Dashboard\" tab. It shows your totals and pictures (charts) automatically. You never need to type anything there."),
    ("4. The yellow row", "Each tab starts with one yellow example row. Just type over it with your own first entry — or delete it once you understand the layout."),
    ("5. That's it!", "Everything else updates by itself. You do not need to add up anything or touch any formulas."),
]

r = 4
for title, body in steps:
    howto.cell(row=r, column=1, value=title).font = HOWTO_HEAD_FONT
    howto.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    r += 1
    howto.cell(row=r, column=1, value=body).font = HOWTO_FONT
    howto.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical="top")
    howto.merge_cells(start_row=r, start_column=1, end_row=r+1, end_column=6)
    howto.row_dimensions[r].height = 20
    howto.row_dimensions[r+1].height = 20
    for cc in range(1, 7):
        howto.cell(row=r, column=cc).fill = HOWTO_FILL
        howto.cell(row=r+1, column=cc).fill = HOWTO_FILL
    r += 3

howto.cell(row=r, column=1, value="Tip: All amounts are in South African Rand (R).").font = NOTE_FONT
howto.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)

autosize(howto, {"A": 16, "B": 16, "C": 16, "D": 16, "E": 16, "F": 16})

# =====================================================================
# 1. CATEGORIES (helper sheet for dropdown lists)
# =====================================================================
cat = wb.create_sheet("Categories")
cat.sheet_view.showGridLines = False

cat["A1"] = "Income Categories"
cat["B1"] = "Expense Categories"
cat["C1"] = "Payment Methods"
style_header_row(cat, 1, 3, HEADER_FILL_CAT)

income_categories = [
    "Custom Orders", "Catering Orders", "Wholesale Orders",
    "Online Orders", "Delivery Charges Collected", "Other Income",
]
expense_categories = [
    "Raw Materials / Ingredients", "Packaging & Boxes", "Rent", "Electricity & Water",
    "Salaries & Wages", "Equipment & Maintenance", "Marketing & Promotion",
    "Transport & Delivery", "Licence & Compliance", "Cleaning & Supplies", "Other Expenses",
]
payment_methods = ["Cash", "Card", "EFT / Bank Transfer", "Other"]

for i, v in enumerate(income_categories, start=2):
    c = cat.cell(row=i, column=1, value=v); c.font = NORMAL_FONT; c.border = BORDER_ALL
for i, v in enumerate(expense_categories, start=2):
    c = cat.cell(row=i, column=2, value=v); c.font = NORMAL_FONT; c.border = BORDER_ALL
for i, v in enumerate(payment_methods, start=2):
    c = cat.cell(row=i, column=3, value=v); c.font = NORMAL_FONT; c.border = BORDER_ALL

autosize(cat, {"A": 30, "B": 30, "C": 20})

cat["E1"] = "You can rename these to match your bakery — the drop-downs on the Income and Expenses tabs read from these columns."
cat["E1"].font = NOTE_FONT

wb.create_named_range("IncomeCatList", cat, "$A$2:$A$20")
wb.create_named_range("ExpenseCatList", cat, "$B$2:$B$20")
wb.create_named_range("PaymentList", cat, "$C$2:$C$20")

# =====================================================================
# 2. INCOME LOG   (Date | Description | Category | Payment Method | Amount | Notes)
# =====================================================================
inc = wb.create_sheet("Income")
inc.sheet_view.showGridLines = False

inc["A1"] = f"{BUSINESS_NAME} — INCOME (Money Coming In)"
inc["A1"].font = Font(name=TITLE_FONT_NAME, size=15, bold=True, color=COL_INCOME_GREEN)
inc.merge_cells("A1:F1")
gold_under = Border(bottom=Side(style="medium", color=COL_GOLD))
for c in range(1, 7):
    inc.cell(row=1, column=c).border = gold_under
inc["A2"] = "Write down every order or sale here. The yellow row is an example — overwrite it."
inc["A2"].font = NOTE_FONT
inc.merge_cells("A2:F2")

headers = ["Date", "Description", "Category", "Payment Method", "Amount (R)", "Notes"]
for i, h in enumerate(headers, start=1):
    inc.cell(row=3, column=i, value=h)
style_header_row(inc, 3, 6, HEADER_FILL_INCOME)

example = [datetime.date(2026, 8, 1), "Custom birthday cake - Van der Merwe family", "Custom Orders", "EFT / Bank Transfer", 850, "Paid in full"]
for i, v in enumerate(example, start=1):
    c = inc.cell(row=4, column=i, value=v)
    c.font = EXAMPLE_FONT
    c.fill = EXAMPLE_FILL
    c.border = BORDER_ALL
inc["A4"].number_format = DATE_FMT
inc["E4"].number_format = CURRENCY_FMT

for r_ in range(5, DATA_ROWS_END + 1):
    inc.cell(row=r_, column=1).number_format = DATE_FMT
    inc.cell(row=r_, column=5).number_format = CURRENCY_FMT
    for c in range(1, 7):
        inc.cell(row=r_, column=c).border = BORDER_ALL
        inc.cell(row=r_, column=c).font = NORMAL_FONT

dv_cat_inc = DataValidation(type="list", formula1="=IncomeCatList", allow_blank=True)
dv_cat_inc.add(f"C5:C{DATA_ROWS_END}")
inc.add_data_validation(dv_cat_inc)

dv_pay_inc = DataValidation(type="list", formula1="=PaymentList", allow_blank=True)
dv_pay_inc.add(f"D5:D{DATA_ROWS_END}")
inc.add_data_validation(dv_pay_inc)

total_row = DATA_ROWS_END + 1
inc.cell(row=total_row, column=2, value="TOTAL INCOME").font = Font(name=FONT_NAME, size=12, bold=True)
inc.cell(row=total_row, column=5, value=f"=SUM(E4:E{DATA_ROWS_END})")
inc.cell(row=total_row, column=5).number_format = CURRENCY_FMT
inc.cell(row=total_row, column=5).font = Font(name=FONT_NAME, size=12, bold=True)
for c in range(1, 7):
    inc.cell(row=total_row, column=c).fill = TOTAL_FILL
    inc.cell(row=total_row, column=c).border = BORDER_ALL

inc.freeze_panes = "A5"
autosize(inc, {"A": 14, "B": 38, "C": 22, "D": 20, "E": 14, "F": 26})

# Only the data-entry area (example row + all data rows, cols A-F) stays editable.
# Headers (rows 1-3) and the total row stay locked.
unlock_range(inc, min_row=4, max_row=DATA_ROWS_END, min_col=1, max_col=6)
lock_whole_sheet(inc)

# =====================================================================
# 3. EXPENSES LOG (same layout as Income, for consistency)
# =====================================================================
exp = wb.create_sheet("Expenses")
exp.sheet_view.showGridLines = False

exp["A1"] = f"{BUSINESS_NAME} — EXPENSES (Money Going Out)"
exp["A1"].font = Font(name=TITLE_FONT_NAME, size=15, bold=True, color=COL_EXPENSE_MAROON)
exp.merge_cells("A1:F1")
gold_under = Border(bottom=Side(style="medium", color=COL_GOLD))
for c in range(1, 7):
    exp.cell(row=1, column=c).border = gold_under
exp["A2"] = "Write down everything you buy for the bakery here. The yellow row is an example — overwrite it."
exp["A2"].font = NOTE_FONT
exp.merge_cells("A2:F2")

headers_e = ["Date", "Description", "Category", "Payment Method", "Amount (R)", "Notes"]
for i, h in enumerate(headers_e, start=1):
    exp.cell(row=3, column=i, value=h)
style_header_row(exp, 3, 6, HEADER_FILL_EXPENSE)

example_e = [datetime.date(2026, 8, 2), "Flour, sugar, butter - weekly stock (Makro)", "Raw Materials / Ingredients", "Cash", 650, ""]
for i, v in enumerate(example_e, start=1):
    c = exp.cell(row=4, column=i, value=v)
    c.font = EXAMPLE_FONT
    c.fill = EXAMPLE_FILL
    c.border = BORDER_ALL
exp["A4"].number_format = DATE_FMT
exp["E4"].number_format = CURRENCY_FMT

for r_ in range(5, DATA_ROWS_END + 1):
    exp.cell(row=r_, column=1).number_format = DATE_FMT
    exp.cell(row=r_, column=5).number_format = CURRENCY_FMT
    for c in range(1, 7):
        exp.cell(row=r_, column=c).border = BORDER_ALL
        exp.cell(row=r_, column=c).font = NORMAL_FONT

dv_cat_exp = DataValidation(type="list", formula1="=ExpenseCatList", allow_blank=True)
dv_cat_exp.add(f"C5:C{DATA_ROWS_END}")
exp.add_data_validation(dv_cat_exp)

dv_pay_exp = DataValidation(type="list", formula1="=PaymentList", allow_blank=True)
dv_pay_exp.add(f"D5:D{DATA_ROWS_END}")
exp.add_data_validation(dv_pay_exp)

exp.cell(row=total_row, column=2, value="TOTAL EXPENSES").font = Font(name=FONT_NAME, size=12, bold=True)
exp.cell(row=total_row, column=5, value=f"=SUM(E4:E{DATA_ROWS_END})")
exp.cell(row=total_row, column=5).number_format = CURRENCY_FMT
exp.cell(row=total_row, column=5).font = Font(name=FONT_NAME, size=12, bold=True)
for c in range(1, 7):
    exp.cell(row=total_row, column=c).fill = TOTAL_FILL
    exp.cell(row=total_row, column=c).border = BORDER_ALL

exp.freeze_panes = "A5"
autosize(exp, {"A": 14, "B": 38, "C": 26, "D": 20, "E": 14, "F": 22})

unlock_range(exp, min_row=4, max_row=DATA_ROWS_END, min_col=1, max_col=6)
lock_whole_sheet(exp)

print("How To Use, Categories, Income, Expenses sheets built.")

# =====================================================================
# 3b. DAILY REPORT — day-by-day breakdown for one selected month
# =====================================================================
daily = wb.create_sheet("Daily Report")
daily.sheet_view.showGridLines = False

daily["A1"] = f"{BUSINESS_NAME} — DAILY REPORT"
daily["A1"].font = Font(name=TITLE_FONT_NAME, size=15, bold=True, color=COL_BROWN)
daily.merge_cells("A1:F1")
daily["A2"] = "Day-by-day breakdown for one month, built automatically from the Income and Expenses tabs. Change Month/Year in H2/I2."
daily["A2"].font = NOTE_FONT
daily.merge_cells("A2:F2")

daily["G2"] = "Month:"
daily["G2"].font = LABEL_FONT
daily["H2"] = 8
daily["H2"].font = Font(name=FONT_NAME, size=11, bold=True, color="B22222")
daily["H2"].fill = PatternFill("solid", fgColor="FFF2CC")
daily["H2"].alignment = Alignment(horizontal="center")
daily["H2"].border = BORDER_ALL
daily["I2"] = "Year:"
daily["I2"].font = LABEL_FONT
daily["J2"] = 2026
daily["J2"].font = Font(name=FONT_NAME, size=11, bold=True, color="B22222")
daily["J2"].fill = PatternFill("solid", fgColor="FFF2CC")
daily["J2"].alignment = Alignment(horizontal="center")
daily["J2"].border = BORDER_ALL
daily["A3"] = "Tip: Month is a number 1-12 (e.g. 8 = August)."
daily["A3"].font = NOTE_FONT
daily.merge_cells("A3:F3")

headers_dr = ["Day", "Date", "Total Income", "Total Expenses", "Net Profit", "Transactions"]
for i, h in enumerate(headers_dr, start=1):
    daily.cell(row=4, column=i, value=h)
style_header_row(daily, 4, 6, HEADER_FILL_SUMMARY)

DAILY_FIRST_ROW = 5
DAILY_LAST_ROW = DAILY_FIRST_ROW + 30  # 31 possible day rows (row 5..35)

for i in range(31):
    r_ = DAILY_FIRST_ROW + i
    day_num = i + 1
    daily.cell(row=r_, column=1, value=day_num).font = NORMAL_FONT
    date_formula = (f"=IF(A{r_}<=DAY(EOMONTH(DATE($J$2,$H$2,1),0)),DATE($J$2,$H$2,A{r_}),\"\")")
    daily.cell(row=r_, column=2, value=date_formula).number_format = DATE_FMT
    inc_formula = f"=IF($B{r_}=\"\",\"\",SUMIFS(Income!$E$4:$E${DATA_ROWS_END},Income!$A$4:$A${DATA_ROWS_END},$B{r_}))"
    exp_formula = f"=IF($B{r_}=\"\",\"\",SUMIFS(Expenses!$E$4:$E${DATA_ROWS_END},Expenses!$A$4:$A${DATA_ROWS_END},$B{r_}))"
    profit_formula = f"=IF($B{r_}=\"\",\"\",C{r_}-D{r_})"
    trans_formula = (f"=IF($B{r_}=\"\",\"\",COUNTIFS(Income!$A$4:$A${DATA_ROWS_END},$B{r_})"
                      f"+COUNTIFS(Expenses!$A$4:$A${DATA_ROWS_END},$B{r_}))")
    daily.cell(row=r_, column=3, value=inc_formula).number_format = CURRENCY_FMT
    daily.cell(row=r_, column=4, value=exp_formula).number_format = CURRENCY_FMT
    daily.cell(row=r_, column=5, value=profit_formula).number_format = CURRENCY_FMT
    daily.cell(row=r_, column=6, value=trans_formula)
    for c in range(1, 7):
        daily.cell(row=r_, column=c).border = BORDER_ALL
        if c != 1:
            daily.cell(row=r_, column=c).font = NORMAL_FONT

DAILY_TOTAL_ROW = DAILY_LAST_ROW + 1
daily.cell(row=DAILY_TOTAL_ROW, column=2, value="MONTH TOTAL").font = Font(name=FONT_NAME, size=12, bold=True)
daily.cell(row=DAILY_TOTAL_ROW, column=3, value=f"=SUM(C{DAILY_FIRST_ROW}:C{DAILY_LAST_ROW})").number_format = CURRENCY_FMT
daily.cell(row=DAILY_TOTAL_ROW, column=4, value=f"=SUM(D{DAILY_FIRST_ROW}:D{DAILY_LAST_ROW})").number_format = CURRENCY_FMT
daily.cell(row=DAILY_TOTAL_ROW, column=5, value=f"=SUM(E{DAILY_FIRST_ROW}:E{DAILY_LAST_ROW})").number_format = CURRENCY_FMT
daily.cell(row=DAILY_TOTAL_ROW, column=6, value=f"=SUM(F{DAILY_FIRST_ROW}:F{DAILY_LAST_ROW})")
for c in range(1, 7):
    daily.cell(row=DAILY_TOTAL_ROW, column=c).fill = TOTAL_FILL
    daily.cell(row=DAILY_TOTAL_ROW, column=c).font = Font(name=FONT_NAME, size=12, bold=True)
    daily.cell(row=DAILY_TOTAL_ROW, column=c).border = BORDER_ALL

daily.freeze_panes = "A5"
autosize(daily, {"A": 8, "B": 14, "C": 15, "D": 15, "E": 15, "F": 14})

protect_with_unlocked_cells(daily, ["H2", "J2"])

print("Daily Report sheet built.")

# =====================================================================
# 4. MONTHLY SUMMARY
# =====================================================================
ms = wb.create_sheet("Monthly Summary")
ms.sheet_view.showGridLines = False

ms["A1"] = f"{BUSINESS_NAME} — MONTHLY SUMMARY"
ms["A1"].font = Font(name=TITLE_FONT_NAME, size=15, bold=True, color=COL_BROWN)
ms.merge_cells("A1:F1")
ms["A2"] = "This fills in by itself from the Income and Expenses tabs. Change the Year in H2 if needed."
ms["A2"].font = NOTE_FONT
ms.merge_cells("A2:F2")

ms["G2"] = "Year:"
ms["G2"].font = LABEL_FONT
ms["H2"] = 2026
ms["H2"].font = Font(name=FONT_NAME, size=11, bold=True, color="B22222")
ms["H2"].fill = PatternFill("solid", fgColor="FFF2CC")
ms["H2"].alignment = Alignment(horizontal="center")
ms["H2"].border = BORDER_ALL

headers_ms = ["Month", "Total Income", "Total Expenses", "Net Profit", "Profit Margin %", "Transactions"]
for i, h in enumerate(headers_ms, start=1):
    ms.cell(row=3, column=i, value=h)
style_header_row(ms, 3, 6, HEADER_FILL_SUMMARY)

months = ["January", "February", "March", "April", "May", "June", "July", "August",
          "September", "October", "November", "December"]

for idx, mname in enumerate(months):
    r_ = 4 + idx
    m = idx + 1
    ms.cell(row=r_, column=1, value=mname).font = NORMAL_FONT
    inc_formula = (f"=SUMIFS(Income!$E$4:$E${DATA_ROWS_END},"
                   f"Income!$A$4:$A${DATA_ROWS_END},\">=\"&DATE($H$2,{m},1),"
                   f"Income!$A$4:$A${DATA_ROWS_END},\"<\"&DATE($H$2,{m}+1,1))")
    exp_formula = (f"=SUMIFS(Expenses!$E$4:$E${DATA_ROWS_END},"
                   f"Expenses!$A$4:$A${DATA_ROWS_END},\">=\"&DATE($H$2,{m},1),"
                   f"Expenses!$A$4:$A${DATA_ROWS_END},\"<\"&DATE($H$2,{m}+1,1))")
    ms.cell(row=r_, column=2, value=inc_formula).number_format = CURRENCY_FMT
    ms.cell(row=r_, column=3, value=exp_formula).number_format = CURRENCY_FMT
    ms.cell(row=r_, column=4, value=f"=B{r_}-C{r_}").number_format = CURRENCY_FMT
    ms.cell(row=r_, column=5, value=f"=IFERROR(D{r_}/B{r_},0)").number_format = PERCENT_FMT
    trans_formula = (f"=COUNTIFS(Income!$A$4:$A${DATA_ROWS_END},\">=\"&DATE($H$2,{m},1),"
                      f"Income!$A$4:$A${DATA_ROWS_END},\"<\"&DATE($H$2,{m}+1,1))"
                      f"+COUNTIFS(Expenses!$A$4:$A${DATA_ROWS_END},\">=\"&DATE($H$2,{m},1),"
                      f"Expenses!$A$4:$A${DATA_ROWS_END},\"<\"&DATE($H$2,{m}+1,1))")
    ms.cell(row=r_, column=6, value=trans_formula)
    for c in range(1, 7):
        ms.cell(row=r_, column=c).border = BORDER_ALL
        if c != 1:
            ms.cell(row=r_, column=c).font = NORMAL_FONT

yr_total_row = 4 + 12
ms.cell(row=yr_total_row, column=1, value="YEAR TOTAL").font = Font(name=FONT_NAME, size=12, bold=True)
ms.cell(row=yr_total_row, column=2, value=f"=SUM(B4:B{yr_total_row-1})").number_format = CURRENCY_FMT
ms.cell(row=yr_total_row, column=3, value=f"=SUM(C4:C{yr_total_row-1})").number_format = CURRENCY_FMT
ms.cell(row=yr_total_row, column=4, value=f"=SUM(D4:D{yr_total_row-1})").number_format = CURRENCY_FMT
ms.cell(row=yr_total_row, column=5, value=f"=IFERROR(D{yr_total_row}/B{yr_total_row},0)").number_format = PERCENT_FMT
ms.cell(row=yr_total_row, column=6, value=f"=SUM(F4:F{yr_total_row-1})")
for c in range(1, 7):
    ms.cell(row=yr_total_row, column=c).fill = TOTAL_FILL
    ms.cell(row=yr_total_row, column=c).font = Font(name=FONT_NAME, size=12, bold=True)
    ms.cell(row=yr_total_row, column=c).border = BORDER_ALL

autosize(ms, {"A": 14, "B": 15, "C": 15, "D": 15, "E": 15, "F": 13})

cat_start_row = yr_total_row + 3
ms.cell(row=cat_start_row, column=1, value="EXPENSES BY CATEGORY (Year total, H2)").font = Font(name=FONT_NAME, size=12, bold=True)
ms.merge_cells(start_row=cat_start_row, start_column=1, end_row=cat_start_row, end_column=2)
ms.cell(row=cat_start_row+1, column=1, value="Category").font = LABEL_FONT
ms.cell(row=cat_start_row+1, column=2, value="Amount").font = LABEL_FONT
for i, catname in enumerate(expense_categories):
    r_ = cat_start_row + 2 + i
    ms.cell(row=r_, column=1, value=catname).font = NORMAL_FONT
    formula = (f"=SUMIFS(Expenses!$E$4:$E${DATA_ROWS_END},Expenses!$C$4:$C${DATA_ROWS_END},A{r_},"
               f"Expenses!$A$4:$A${DATA_ROWS_END},\">=\"&DATE($H$2,1,1),"
               f"Expenses!$A$4:$A${DATA_ROWS_END},\"<\"&DATE($H$2+1,1,1))")
    ms.cell(row=r_, column=2, value=formula).number_format = CURRENCY_FMT
exp_cat_last_row = cat_start_row + 1 + len(expense_categories)

inc_cat_start_row = cat_start_row
ms.cell(row=inc_cat_start_row, column=4, value="INCOME BY CATEGORY (Year total, H2)").font = Font(name=FONT_NAME, size=12, bold=True)
ms.merge_cells(start_row=inc_cat_start_row, start_column=4, end_row=inc_cat_start_row, end_column=5)
ms.cell(row=inc_cat_start_row+1, column=4, value="Category").font = LABEL_FONT
ms.cell(row=inc_cat_start_row+1, column=5, value="Amount").font = LABEL_FONT
for i, catname in enumerate(income_categories):
    r_ = inc_cat_start_row + 2 + i
    ms.cell(row=r_, column=4, value=catname).font = NORMAL_FONT
    formula = (f"=SUMIFS(Income!$E$4:$E${DATA_ROWS_END},Income!$C$4:$C${DATA_ROWS_END},D{r_},"
               f"Income!$A$4:$A${DATA_ROWS_END},\">=\"&DATE($H$2,1,1),"
               f"Income!$A$4:$A${DATA_ROWS_END},\"<\"&DATE($H$2+1,1,1))")
    ms.cell(row=r_, column=5, value=formula).number_format = CURRENCY_FMT
inc_cat_last_row = inc_cat_start_row + 1 + len(income_categories)

protect_with_unlocked_cells(ms, ["H2"])

print("Monthly Summary sheet built.")

# =====================================================================
# 4b. YEARLY REPORT — multi-year history (fixed 2024-2040 span)
# =====================================================================
yearly = wb.create_sheet("Yearly Report")
yearly.sheet_view.showGridLines = False

yearly["A1"] = f"{BUSINESS_NAME} — YEARLY REPORT"
yearly["A1"].font = Font(name=TITLE_FONT_NAME, size=15, bold=True, color=COL_BROWN)
yearly.merge_cells("A1:F1")
yearly["A2"] = "Multi-year history, built automatically from the Income and Expenses tabs. Covers 2024-2040."
yearly["A2"].font = NOTE_FONT
yearly.merge_cells("A2:F2")

headers_yr = ["Year", "Total Income", "Total Expenses", "Net Profit", "Profit Margin %", "Transactions"]
for i, h in enumerate(headers_yr, start=1):
    yearly.cell(row=3, column=i, value=h)
style_header_row(yearly, 3, 6, HEADER_FILL_SUMMARY)

YEARLY_FIRST_ROW = 4
YEARLY_YEARS = list(range(2024, 2041))  # 2024-2040 inclusive, 17 years
YEARLY_LAST_ROW = YEARLY_FIRST_ROW + len(YEARLY_YEARS) - 1

for i, yr in enumerate(YEARLY_YEARS):
    r_ = YEARLY_FIRST_ROW + i
    yearly.cell(row=r_, column=1, value=yr).font = NORMAL_FONT
    inc_formula = (f"=SUMIFS(Income!$E$4:$E${DATA_ROWS_END},"
                    f"Income!$A$4:$A${DATA_ROWS_END},\">=\"&DATE({yr},1,1),"
                    f"Income!$A$4:$A${DATA_ROWS_END},\"<\"&DATE({yr}+1,1,1))")
    exp_formula = (f"=SUMIFS(Expenses!$E$4:$E${DATA_ROWS_END},"
                    f"Expenses!$A$4:$A${DATA_ROWS_END},\">=\"&DATE({yr},1,1),"
                    f"Expenses!$A$4:$A${DATA_ROWS_END},\"<\"&DATE({yr}+1,1,1))")
    yearly.cell(row=r_, column=2, value=inc_formula).number_format = CURRENCY_FMT
    yearly.cell(row=r_, column=3, value=exp_formula).number_format = CURRENCY_FMT
    yearly.cell(row=r_, column=4, value=f"=B{r_}-C{r_}").number_format = CURRENCY_FMT
    yearly.cell(row=r_, column=5, value=f"=IFERROR(D{r_}/B{r_},0)").number_format = PERCENT_FMT
    trans_formula = (f"=COUNTIFS(Income!$A$4:$A${DATA_ROWS_END},\">=\"&DATE({yr},1,1),"
                      f"Income!$A$4:$A${DATA_ROWS_END},\"<\"&DATE({yr}+1,1,1))"
                      f"+COUNTIFS(Expenses!$A$4:$A${DATA_ROWS_END},\">=\"&DATE({yr},1,1),"
                      f"Expenses!$A$4:$A${DATA_ROWS_END},\"<\"&DATE({yr}+1,1,1))")
    yearly.cell(row=r_, column=6, value=trans_formula)
    for c in range(1, 7):
        yearly.cell(row=r_, column=c).border = BORDER_ALL
        if c != 1:
            yearly.cell(row=r_, column=c).font = NORMAL_FONT

YEARLY_TOTAL_ROW = YEARLY_LAST_ROW + 1
yearly.cell(row=YEARLY_TOTAL_ROW, column=1, value="ALL YEARS TOTAL").font = Font(name=FONT_NAME, size=12, bold=True)
yearly.cell(row=YEARLY_TOTAL_ROW, column=2, value=f"=SUM(B{YEARLY_FIRST_ROW}:B{YEARLY_LAST_ROW})").number_format = CURRENCY_FMT
yearly.cell(row=YEARLY_TOTAL_ROW, column=3, value=f"=SUM(C{YEARLY_FIRST_ROW}:C{YEARLY_LAST_ROW})").number_format = CURRENCY_FMT
yearly.cell(row=YEARLY_TOTAL_ROW, column=4, value=f"=SUM(D{YEARLY_FIRST_ROW}:D{YEARLY_LAST_ROW})").number_format = CURRENCY_FMT
yearly.cell(row=YEARLY_TOTAL_ROW, column=5,
            value=f"=IFERROR(D{YEARLY_TOTAL_ROW}/B{YEARLY_TOTAL_ROW},0)").number_format = PERCENT_FMT
yearly.cell(row=YEARLY_TOTAL_ROW, column=6, value=f"=SUM(F{YEARLY_FIRST_ROW}:F{YEARLY_LAST_ROW})")
for c in range(1, 7):
    yearly.cell(row=YEARLY_TOTAL_ROW, column=c).fill = TOTAL_FILL
    yearly.cell(row=YEARLY_TOTAL_ROW, column=c).font = Font(name=FONT_NAME, size=12, bold=True)
    yearly.cell(row=YEARLY_TOTAL_ROW, column=c).border = BORDER_ALL

yearly.freeze_panes = "A4"
autosize(yearly, {"A": 10, "B": 15, "C": 15, "D": 15, "E": 15, "F": 14})

lock_whole_sheet(yearly)

print("Yearly Report sheet built.")

# =====================================================================
# 5. DASHBOARD
# =====================================================================
dash = wb.create_sheet("Dashboard")
dash.sheet_view.showGridLines = False

dash.merge_cells("A1:H1")
dash["A1"] = f"🧁 {BUSINESS_NAME} — DASHBOARD"
dash["A1"].font = Font(name=TITLE_FONT_NAME, size=20, bold=True, color="FFFFFF")
dash["A1"].fill = HEADER_FILL_DASH
dash["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
for c in range(1, 9):
    dash.cell(row=1, column=c).fill = HEADER_FILL_DASH
dash.row_dimensions[1].height = 32
gold_accent_bar(dash, 2, 8)

dash["A4"] = "Everything below updates by itself as you add rows to the Income and Expenses tabs."
dash["A4"].font = NOTE_FONT
dash.merge_cells("A4:H4")

def kpi_card(ws, row, col_start, label, formula, fmt, value_font=KPI_VALUE_FONT):
    col_end = col_start + 1
    ws.merge_cells(start_row=row, start_column=col_start, end_row=row, end_column=col_end)
    lbl = ws.cell(row=row, column=col_start, value=label)
    lbl.font = LABEL_FONT
    lbl.alignment = Alignment(horizontal="center")
    lbl.fill = KPI_FILL
    ws.cell(row=row, column=col_end).fill = KPI_FILL
    ws.merge_cells(start_row=row+1, start_column=col_start, end_row=row+2, end_column=col_end)
    val = ws.cell(row=row+1, column=col_start, value=formula)
    val.font = value_font
    val.number_format = fmt
    val.alignment = Alignment(horizontal="center", vertical="center")
    val.fill = KPI_FILL
    for rr in (row, row+1, row+2):
        for cc in range(col_start, col_end+1):
            ws.cell(row=rr, column=cc).border = BORDER_ALL
            ws.cell(row=rr, column=cc).fill = KPI_FILL
    return row, col_start  # anchor cell of the value

dash.cell(row=5, column=1, value="ALL-TIME PERFORMANCE").font = SECTION_FONT
dash.merge_cells("A5:H5")
for c in range(1, 9):
    dash.cell(row=5, column=c).fill = SECTION_FILL
dash.row_dimensions[5].height = 20

kpi_card(dash, 7, 1, "TOTAL INCOME", f"=Income!E{DATA_ROWS_END+1}", CURRENCY_FMT)
kpi_card(dash, 7, 3, "TOTAL EXPENSES", f"=Expenses!E{DATA_ROWS_END+1}", CURRENCY_FMT, KPI_VALUE_FONT_RED)
kpi_card(dash, 7, 5, "NET PROFIT", "=A8-C8", CURRENCY_FMT)
kpi_card(dash, 7, 7, "PROFIT MARGIN", "=IFERROR(E8/A8,0)", PERCENT_FMT)

# ---- TODAY'S LIVE VIEW (between All-Time and This Month) ----
dash.cell(row=11, column=1, value="TODAY'S LIVE VIEW").font = SECTION_FONT
dash.merge_cells("A11:H11")
for c in range(1, 9):
    dash.cell(row=11, column=c).fill = SECTION_FILL
dash.row_dimensions[11].height = 20

today_inc = f"=SUMIFS(Income!$E$4:$E${DATA_ROWS_END},Income!$A$4:$A${DATA_ROWS_END},TODAY())"
today_exp = f"=SUMIFS(Expenses!$E$4:$E${DATA_ROWS_END},Expenses!$A$4:$A${DATA_ROWS_END},TODAY())"

kpi_card(dash, 13, 1, "INCOME TODAY", today_inc, CURRENCY_FMT)
kpi_card(dash, 13, 3, "EXPENSES TODAY", today_exp, CURRENCY_FMT, KPI_VALUE_FONT_RED)
kpi_card(dash, 13, 5, "PROFIT TODAY", "=A14-C14", CURRENCY_FMT)

# Small helper table (chart source only) placed off to the side, out of the main reading flow
dash["K13"] = "Income"
dash["L13"] = "=A14"
dash["K14"] = "Expenses"
dash["L14"] = "=C14"
for addr in ("K13", "K14", "L13", "L14"):
    dash[addr].font = Font(name=FONT_NAME, size=9, color="B8ACA0")
dash["L13"].number_format = CURRENCY_FMT
dash["L14"].number_format = CURRENCY_FMT
dash["K12"] = "(chart source — do not edit)"
dash["K12"].font = NOTE_FONT

today_chart = BarChart()
today_chart.type = "col"
today_chart.title = "Today: Income vs Expenses"
today_chart.style = 10
today_chart.y_axis.title = "Amount (R)"
today_chart.height = 6.5
today_chart.width = 12
today_data = Reference(dash, min_col=12, min_row=13, max_row=14)
today_cats = Reference(dash, min_col=11, min_row=13, max_row=14)
today_chart.add_data(today_data, titles_from_data=False)
today_chart.set_categories(today_cats)
today_chart.legend = None
dash.add_chart(today_chart, "I11")

dash.cell(row=17, column=1, value="THIS MONTH").font = SECTION_FONT
dash.merge_cells("A17:H17")
for c in range(1, 9):
    dash.cell(row=17, column=c).fill = SECTION_FILL
dash.row_dimensions[17].height = 20

this_month_inc = (f"=SUMIFS(Income!$E$4:$E${DATA_ROWS_END},Income!$A$4:$A${DATA_ROWS_END},\">=\"&DATE(YEAR(TODAY()),MONTH(TODAY()),1),"
                   f"Income!$A$4:$A${DATA_ROWS_END},\"<\"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1))")
this_month_exp = (f"=SUMIFS(Expenses!$E$4:$E${DATA_ROWS_END},Expenses!$A$4:$A${DATA_ROWS_END},\">=\"&DATE(YEAR(TODAY()),MONTH(TODAY()),1),"
                   f"Expenses!$A$4:$A${DATA_ROWS_END},\"<\"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1))")

kpi_card(dash, 19, 1, "INCOME THIS MONTH", this_month_inc, CURRENCY_FMT)
kpi_card(dash, 19, 3, "EXPENSES THIS MONTH", this_month_exp, CURRENCY_FMT, KPI_VALUE_FONT_RED)
kpi_card(dash, 19, 5, "PROFIT THIS MONTH", "=A20-C20", CURRENCY_FMT)
kpi_card(dash, 19, 7, "ORDERS THIS MONTH",
         (f"=COUNTIFS(Income!$A$4:$A${DATA_ROWS_END},\">=\"&DATE(YEAR(TODAY()),MONTH(TODAY()),1),Income!$A$4:$A${DATA_ROWS_END},\"<\"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1))"),
         '0', KPI_VALUE_FONT)

# ---- THIS YEAR (between This Month and Quick Facts) ----
dash.cell(row=23, column=1, value="THIS YEAR").font = SECTION_FONT
dash.merge_cells("A23:H23")
for c in range(1, 9):
    dash.cell(row=23, column=c).fill = SECTION_FILL
dash.row_dimensions[23].height = 20

this_year_inc = (f"=SUMIFS(Income!$E$4:$E${DATA_ROWS_END},Income!$A$4:$A${DATA_ROWS_END},\">=\"&DATE(YEAR(TODAY()),1,1),"
                  f"Income!$A$4:$A${DATA_ROWS_END},\"<\"&DATE(YEAR(TODAY())+1,1,1))")
this_year_exp = (f"=SUMIFS(Expenses!$E$4:$E${DATA_ROWS_END},Expenses!$A$4:$A${DATA_ROWS_END},\">=\"&DATE(YEAR(TODAY()),1,1),"
                  f"Expenses!$A$4:$A${DATA_ROWS_END},\"<\"&DATE(YEAR(TODAY())+1,1,1))")

kpi_card(dash, 25, 1, "INCOME THIS YEAR", this_year_inc, CURRENCY_FMT)
kpi_card(dash, 25, 3, "EXPENSES THIS YEAR", this_year_exp, CURRENCY_FMT, KPI_VALUE_FONT_RED)
kpi_card(dash, 25, 5, "PROFIT THIS YEAR", "=A26-C26", CURRENCY_FMT)
kpi_card(dash, 25, 7, "ORDERS THIS YEAR",
         (f"=COUNTIFS(Income!$A$4:$A${DATA_ROWS_END},\">=\"&DATE(YEAR(TODAY()),1,1),Income!$A$4:$A${DATA_ROWS_END},\"<\"&DATE(YEAR(TODAY())+1,1,1))"),
         '0', KPI_VALUE_FONT)

dash.cell(row=29, column=1, value="QUICK FACTS").font = SECTION_FONT
dash.merge_cells("A29:H29")
for c in range(1, 9):
    dash.cell(row=29, column=c).fill = SECTION_FILL
dash.row_dimensions[29].height = 20

row0 = 31
dash.cell(row=row0-1, column=1, value="Fact").font = LABEL_FONT
dash.cell(row=row0-1, column=2, value="Value").font = LABEL_FONT
for c in (1, 2):
    dash.cell(row=row0-1, column=c).fill = TOTAL_FILL
    dash.cell(row=row0-1, column=c).border = BORDER_ALL

insight_labels = [
    "Best-selling category (this year)",
    "Biggest expense category (this year)",
    "Average order amount (all time)",
    "Biggest single order (all time)",
]
top_income_formula = (f"=INDEX('Monthly Summary'!$D${cat_start_row+2}:$D${inc_cat_last_row},"
                       f"MATCH(MAX('Monthly Summary'!$E${cat_start_row+2}:$E${inc_cat_last_row}),"
                       f"'Monthly Summary'!$E${cat_start_row+2}:$E${inc_cat_last_row},0))")
top_expense_formula = (f"=INDEX('Monthly Summary'!$A${cat_start_row+2}:$A${exp_cat_last_row},"
                        f"MATCH(MAX('Monthly Summary'!$B${cat_start_row+2}:$B${exp_cat_last_row}),"
                        f"'Monthly Summary'!$B${cat_start_row+2}:$B${exp_cat_last_row},0))")
avg_order_value = f"=IFERROR(Income!E{DATA_ROWS_END+1}/COUNT(Income!E4:E{DATA_ROWS_END}),0)"
highest_sale = f"=IFERROR(MAX(Income!E4:E{DATA_ROWS_END}),0)"

insight_values = [top_income_formula, top_expense_formula, avg_order_value, highest_sale]
insight_fmts = [None, None, CURRENCY_FMT, CURRENCY_FMT]

for i, (lbl, formula, fmt) in enumerate(zip(insight_labels, insight_values, insight_fmts)):
    r_ = row0 + i
    dash.cell(row=r_, column=1, value=lbl).font = NORMAL_FONT
    vcell = dash.cell(row=r_, column=2, value=formula)
    vcell.font = Font(name=FONT_NAME, size=11, bold=True)
    if fmt:
        vcell.number_format = fmt
    for c in (1, 2):
        dash.cell(row=r_, column=c).border = BORDER_ALL

autosize(dash, {"A": 20, "B": 16, "C": 16, "D": 16, "E": 16, "F": 16, "G": 16, "H": 16})

# ---- Charts ----
bar = BarChart()
bar.title = "Monthly Income vs Expenses"
bar.style = 10
bar.y_axis.title = "Amount (R)"
bar.x_axis.title = "Month"
bar.height = 8
bar.width = 18
data = Reference(ms, min_col=2, max_col=3, min_row=3, max_row=15)
cats = Reference(ms, min_col=1, min_row=4, max_row=15)
bar.add_data(data, titles_from_data=True)
bar.set_categories(cats)
dash.add_chart(bar, "A37")

line = LineChart()
line.title = "Net Profit Trend"
line.style = 12
line.y_axis.title = "Net Profit (R)"
line.x_axis.title = "Month"
line.height = 8
line.width = 18
data2 = Reference(ms, min_col=4, min_row=3, max_row=15)
line.add_data(data2, titles_from_data=True)
line.set_categories(cats)
dash.add_chart(line, "A55")

pie = PieChart()
pie.title = "Expenses by Category (Year)"
pie.height = 9
pie.width = 14
pdata = Reference(ms, min_col=2, min_row=cat_start_row+1, max_row=exp_cat_last_row)
pcats = Reference(ms, min_col=1, min_row=cat_start_row+2, max_row=exp_cat_last_row)
pie.add_data(pdata, titles_from_data=True)
pie.set_categories(pcats)
dash.add_chart(pie, "I37")

pie2 = PieChart()
pie2.title = "Income by Category (Year)"
pie2.height = 9
pie2.width = 14
pdata2 = Reference(ms, min_col=5, min_row=inc_cat_start_row+1, max_row=inc_cat_last_row)
pcats2 = Reference(ms, min_col=4, min_row=inc_cat_start_row+2, max_row=inc_cat_last_row)
pie2.add_data(pdata2, titles_from_data=True)
pie2.set_categories(pcats2)
dash.add_chart(pie2, "I55")

lock_whole_sheet(dash)
lock_whole_sheet(howto)

print("Dashboard sheet built.")

wb._sheets = [howto, dash, inc, exp, daily, ms, yearly, cat]
wb.active = 0

wb.properties.title = f"{BUSINESS_NAME} — Bookkeeping"
wb.properties.creator = BUSINESS_NAME

import os
OUTPUT_PATH = os.environ.get(
    "OUTPUT_PATH",
    os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                 "templates", "Premium_Bakery_Bookkeeping_TEMPLATE.xlsx")
)
wb.save(OUTPUT_PATH)
print(f"Saved to {OUTPUT_PATH}")
