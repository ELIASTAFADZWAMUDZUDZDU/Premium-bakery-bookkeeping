"""Checks that the workbook's formulas agree with an independent calculation.

Builds the demo workbook (fictional data), recalculates it with LibreOffice,
then compares totals on the Dashboard, Monthly Summary and Yearly Report against
numbers computed directly from the Income and Expenses logs with plain Python.

Run:  python tests/test_workbook.py      (needs LibreOffice: `soffice` on PATH)
"""
import datetime, os, subprocess, sys, tempfile
import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_LAST_ROW = 5002

def build_and_recalc(tmp):
    built = os.path.join(tmp, "demo.xlsx")
    env = dict(os.environ, DEMO="1", OUTPUT_PATH=built)
    subprocess.run([sys.executable, os.path.join(ROOT, "src", "build_workbook.py")],
                   check=True, env=env, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    out = os.path.join(tmp, "recalc")
    os.makedirs(out)
    subprocess.run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", out, built],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return os.path.join(out, "demo.xlsx")

def read_log(ws):
    rows = []
    for r in range(4, DATA_LAST_ROW + 1):
        d, amt = ws.cell(r, 1).value, ws.cell(r, 5).value
        if isinstance(d, datetime.datetime) and isinstance(amt, (int, float)):
            rows.append((d.date(), float(amt)))
    return rows

def close(a, b):
    return abs(float(a) - float(b)) < 0.01

def main():
    failures = []
    def check(name, got, want):
        ok = close(got, want)
        print(("PASS " if ok else "FAIL ") + f"{name}: got {got}, expected {want}")
        if not ok: failures.append(name)

    with tempfile.TemporaryDirectory() as tmp:
        path = build_and_recalc(tmp)
        wb = openpyxl.load_workbook(path, data_only=True)

    inc, exp = read_log(wb["Income"]), read_log(wb["Expenses"])
    ti, te = sum(a for _, a in inc), sum(a for _, a in exp)

    check("Income sheet total row", wb["Income"].cell(DATA_LAST_ROW + 1, 5).value, ti)
    check("Expenses sheet total row", wb["Expenses"].cell(DATA_LAST_ROW + 1, 5).value, te)

    dash = wb["Dashboard"]
    check("Dashboard total income", dash["A8"].value, ti)
    check("Dashboard total expenses", dash["C8"].value, te)
    check("Dashboard net profit", dash["E8"].value, ti - te)
    check("Dashboard profit margin", dash["G8"].value, (ti - te) / ti if ti else 0)

    year = datetime.date.today().year
    yi = sum(a for d, a in inc if d.year == year)
    ye = sum(a for d, a in exp if d.year == year)
    yr = wb["Yearly Report"]
    for r in range(1, yr.max_row + 1):
        if yr.cell(r, 1).value == year:
            check(f"Yearly Report {year} income", yr.cell(r, 2).value, yi)
            check(f"Yearly Report {year} expenses", yr.cell(r, 3).value, ye)
            check(f"Yearly Report {year} net profit", yr.cell(r, 4).value, yi - ye)
            break
    else:
        failures.append("Yearly Report row not found")

    ms = wb["Monthly Summary"]
    ms_year = ms["H2"].value
    mi = [sum(a for d, a in inc if d.year == ms_year and d.month == m) for m in range(1, 13)]
    me = [sum(a for d, a in exp if d.year == ms_year and d.month == m) for m in range(1, 13)]
    months = ["January","February","March","April","May","June","July","August",
              "September","October","November","December"]
    got_i = got_e = 0.0
    for r in range(1, ms.max_row + 1):
        v = ms.cell(r, 1).value
        if v in months and isinstance(ms.cell(r, 2).value, (int, float)):
            m = months.index(v)
            got_i += ms.cell(r, 2).value
            got_e += ms.cell(r, 3).value
            check(f"Monthly Summary {v} income", ms.cell(r, 2).value, mi[m])
            check(f"Monthly Summary {v} expenses", ms.cell(r, 3).value, me[m])
    check(f"Monthly Summary {ms_year} total income", got_i, sum(mi))
    check(f"Monthly Summary {ms_year} total expenses", got_e, sum(me))

    print()
    print("ALL TESTS PASSED" if not failures else f"{len(failures)} FAILED: {failures}")
    sys.exit(1 if failures else 0)

if __name__ == "__main__":
    main()
