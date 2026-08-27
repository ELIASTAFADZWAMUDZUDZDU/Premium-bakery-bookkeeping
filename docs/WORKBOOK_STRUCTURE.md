# Workbook Structure Reference

A quick structural map of `templates/Premium_Bakery_Bookkeeping_TEMPLATE.xlsx` — sheets, tab order, and row/column layout, without formula-level detail. See `SYSTEM_DOCUMENTATION.md` for exact formulas.

---

## Tab order

1. How To Use
2. Dashboard
3. Income
4. Expenses
5. Daily Report
6. Monthly Summary
7. Yearly Report
8. Categories

## Sheet protection at a glance

Only **Income**, **Expenses**, and **Categories** allow typing anywhere without unprotecting first. **Daily Report** and **Monthly Summary** each leave just their Month/Year selector cell(s) open. Everything else is fully locked. See `SYSTEM_DOCUMENTATION.md` §1a for the complete table.

---

## How To Use

| Rows | Content |
|---|---|
| 1 | Title banner |
| 2 | Gold accent stripe |
| 4–20 | 5 numbered instruction blocks |
| last row | Currency note |

---

## Dashboard

| Rows | Section |
|---|---|
| 1–2 | Title banner + gold accent stripe |
| 4 | Subtitle note |
| 5 | Section header: ALL-TIME PERFORMANCE |
| 7–9 | KPI cards: Total Income · Total Expenses · Net Profit · Profit Margin |
| 11 | Section header: TODAY'S LIVE VIEW |
| 13–15 | KPI cards: Income Today · Expenses Today · Profit Today, + live bar chart (anchored `I11`) |
| 17 | Section header: THIS MONTH |
| 19–21 | KPI cards: Income · Expenses · Profit · Orders (this month) |
| 23 | Section header: THIS YEAR |
| 25–27 | KPI cards: Income · Expenses · Profit · Orders (this calendar year) |
| 29 | Section header: QUICK FACTS |
| 30 | Table header (Fact / Value) |
| 31–34 | 4 fact rows (best-selling category, biggest expense category, average order, biggest single order) |
| 37 | Chart: Monthly Income vs Expenses (bar) |
| 37 (col I) | Chart: Expenses by Category (pie) |
| 55 | Chart: Net Profit Trend (line) |
| 55 (col I) | Chart: Income by Category (pie) |

Columns used for content: A–H. Columns K–L on rows 12–14 hold a small helper table feeding the "Today" chart only.

---

## Income / Expenses

Identical layout on both sheets.

| Rows | Content |
|---|---|
| 1 | Title banner |
| 2 | Instruction subtitle |
| 3 | Column headers |
| 4 | Example row (sample data) |
| 5–5002 | Data entry (~5,000 rows) |
| 5003 | Total row |

| Column | Field |
|---|---|
| A | Date |
| B | Description |
| C | Category (dropdown) |
| D | Payment Method (dropdown) |
| E | Amount (R) |
| F | Notes |

---

## Daily Report

| Rows | Content |
|---|---|
| 1 | Title banner |
| 2 | Subtitle note |
| 2 (col G–J) | Month / Year selectors |
| 3 | Tip note |
| 4 | Column headers |
| 5–35 | One row per calendar day (1–31, blank if the day doesn't exist in the selected month) |
| 36 | Month Total row |

| Column | Field |
|---|---|
| A | Day (1–31) |
| B | Date |
| C | Total Income |
| D | Total Expenses |
| E | Net Profit |
| F | Transactions |

---

## Monthly Summary

| Rows | Content |
|---|---|
| 1 | Title banner |
| 2 | Subtitle note + Year selector (col H) |
| 3 | Column headers |
| 4–15 | One row per month (January–December) |
| 16 | Year Total row |
| 19 | Section headers: EXPENSES BY CATEGORY (col A–B) / INCOME BY CATEGORY (col D–E) |
| 20 | Sub-headers (Category / Amount) |
| 21–31 | Expense category rows (11 categories) |
| 21–26 | Income category rows (6 categories, columns D–E) |

| Column | Field |
|---|---|
| A | Month |
| B | Total Income |
| C | Total Expenses |
| D | Net Profit |
| E | Profit Margin % |
| F | Transactions |

---

## Yearly Report

| Rows | Content |
|---|---|
| 1 | Title banner |
| 2 | Subtitle note |
| 3 | Column headers |
| 4–20 | One row per year, fixed span 2024–2040 (17 years) |
| 21 | All Years Total row |

| Column | Field |
|---|---|
| A | Year |
| B | Total Income |
| C | Total Expenses |
| D | Net Profit |
| E | Profit Margin % |
| F | Transactions |

No Month/Year selector — always shows every year in the fixed span at once. Fully locked.

---

## Categories

| Column | Content |
|---|---|
| A | Income categories (6) |
| B | Expense categories (11) |
| C | Payment methods (4) |

Named ranges: `IncomeCatList` (`A2:A20`), `ExpenseCatList` (`B2:B20`), `PaymentList` (`C2:C20`).
