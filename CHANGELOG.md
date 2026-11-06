# Changelog

All notable changes to this workbook are documented here.

## [1.5.0]
### Added
- `DEMO=1` option in `build_workbook.py` that fills both logs with fictional sample data; `demo/Premium_Bakery_DEMO.xlsx` and a dashboard screenshot.
- `tests/test_workbook.py`: recalculates the demo in LibreOffice and checks Dashboard, Monthly Summary and Yearly Report totals against an independent calculation.
- `CURRENCY_SYMBOL` / `CURRENCY_NAME` environment variables (default: South African Rand).
- MIT `LICENSE`.

### Changed
- Example rows are now generic sample entries (no names or shop names).
- Dashboard Quick Facts: wrapped labels and a wider value column so category names are no longer cut off.

## [1.4.0]
### Added
- **Sheet protection**: Income and Expenses are now the only sheets with open data-entry cells (rows 4–5002 across all 6 columns). Headers, total rows, and every other sheet are locked against accidental edits.
  - Daily Report and Monthly Summary stay locked overall, but their Month/Year selector cells remain unlocked so the displayed period can still be changed.
  - Categories is left fully unprotected, since renaming or adding a category is an intentional, occasional edit.
  - No password is set — this is accidental-edit prevention, not access control.
- New **Yearly Report** sheet: a fixed 2024–2040 multi-year history table (Total Income, Total Expenses, Net Profit, Profit Margin %, Transactions per year), plus an all-years total row. Fully locked, no selector — it always shows every year at once.
- New **This Year** section on the Dashboard, between This Month and Quick Facts: Income / Expenses / Profit / Orders for the current calendar year, recalculating automatically off `YEAR(TODAY())`.

### Changed
- Dashboard section order: All-Time Performance → Today's Live View → This Month → **This Year** → Quick Facts → Charts.
- Tab order: How To Use → Dashboard → Income → Expenses → Daily Report → Monthly Summary → **Yearly Report** → Categories.

## [1.3.0]
### Added
- New **Daily Report** sheet: day-by-day income/expense/profit/transaction breakdown for one selected month (Month/Year selectors), 31-row layout with automatic blank-out for days that don't exist in the selected month (e.g. day 30 in February).
- New **Today's Live View** section on the Dashboard (between All-Time Performance and This Month): live running totals of today's income, expenses, and profit, plus a bar chart comparing today's income vs. expenses. Recalculates automatically off `TODAY()` — no manual refresh needed.

### Changed
- Dashboard section order: All-Time Performance → Today's Live View → This Month → Quick Facts → Charts.

## [1.2.0]
### Changed
- Increased Income/Expenses row capacity from ~500 to ~5,000 rows per sheet (roughly 13+ years of daily entries), with full formatting (currency, date format, borders, dropdowns) applied to every row.
- Every formula range throughout Monthly Summary and Dashboard updated to match the new capacity.

## [1.1.0]
### Changed
- Localised to South Africa: currency changed from Indian Rupee (₹) to South African Rand (R).
- Renamed "Walk-in Counter Sales" income category to "Custom Orders".
- Simplified payment methods to Cash, Card, EFT / Bank Transfer, Other (removed UPI).
- Unified Income and Expenses sheets to an identical 6-column layout (Date, Description, Category, Payment Method, Amount, Notes) — dropped the separate Invoice No. and Vendor columns for consistency and simplicity.
- Applied "Premium Bakery" branding throughout: business name in sheet titles, espresso-brown/gold/cream color palette, Georgia display font for headings.

### Added
- New "How To Use" sheet: plain-language, numbered quick-start guide as the first tab.

## [1.0.0]
### Added
- Initial workbook: How To Use, Dashboard, Income, Expenses, Monthly Summary, Categories sheets.
- Dashboard KPIs: all-time totals, this-month totals, quick facts (top category, average order, biggest sale).
- Charts: monthly income vs. expenses (bar), net profit trend (line), expenses by category (pie), income by category (pie).
- Dropdown-driven category and payment method entry via named ranges.
