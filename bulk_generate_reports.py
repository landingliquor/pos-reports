from __future__ import annotations

import json
import os
import subprocess
from datetime import date, datetime, timedelta, time
from decimal import Decimal
from pathlib import Path

import pyodbc
from openpyxl import Workbook


# ================= CONFIG =================
DSN = "pos"
UID = "db"
PWD = "db"

BUSINESS_START = "06:30"

REPO_ROOT = Path(__file__).resolve().parent
REPORTS_DIR = REPO_ROOT / "reports"
INDEX_JSON = REPO_ROOT / "report_index.json"

TODAY = date.today()
START_DATE = date(TODAY.year - 1, 12, 1)
END_DATE = TODAY

# ==========================================


def d0(v):
    try:
        return Decimal(str(v or 0))
    except:
        return Decimal("0")


def fmt(v):
    return float(d0(v))


def business_window(day):
    hh, mm = map(int, BUSINESS_START.split(":"))
    start = datetime.combine(day, time(hh, mm))
    end = start + timedelta(days=1)
    return start, end


def db():
    return pyodbc.connect(f"DSN={DSN};UID={UID};PWD={PWD};", autocommit=True)


def ensure(p):
    p.mkdir(parents=True, exist_ok=True)


def write_html(path, content):
    ensure(path.parent)
    path.write_text(content, encoding="utf-8")


def write_xlsx(path, headers, rows):
    ensure(path.parent)
    wb = Workbook()
    ws = wb.active
    ws.append(headers)
    for r in rows:
        ws.append(r)
    wb.save(path)


# ================= DAILY =================
def daily(conn, start, end):

    cur = conn.cursor()

    rows = cur.execute("""
    SELECT TransType, GroupTransType, ReceiptN, SubCategoryID,
           Amount, DiscountAmount, TaxInclude,
           Tax1Amount, Tax2Amount, Tax3Amount, Tax4Amount
    FROM Journal
    WHERE DateR >= ? AND DateR < ?
      AND Status = 0
    """, start, end).fetchall()

    sales = [r for r in rows if int(r.TransType or 0) == 101]

    def tax(r):
        return d0(r.Tax1Amount) + d0(r.Tax2Amount) + d0(r.Tax3Amount) + d0(r.Tax4Amount)

    gross = sum(d0(r.Amount) + tax(r) * (1 - d0(r.TaxInclude)) for r in sales)
    gst = sum(d0(r.Tax1Amount) * (1 - d0(r.TaxInclude)) for r in sales)
    pst = sum(d0(r.Tax2Amount) * (1 - d0(r.TaxInclude)) for r in sales)
    liq = sum(d0(r.Tax3Amount) * (1 - d0(r.TaxInclude)) for r in sales)

    total_tax = gst + pst + liq
    net = gross - total_tax
    discount = sum(d0(r.DiscountAmount) for r in sales)

    customers = len({r.ReceiptN for r in rows if int(r.GroupTransType or 0) == 1})

    html = f"""
    <h3>Daily Summary</h3>
    <p>Total Sales: {gross:.2f}</p>
    <p>Net: {net:.2f}</p>
    <p>GST: {gst:.2f}</p>
    <p>PST: {pst:.2f}</p>
    <p>Liquor Tax: {liq:.2f}</p>
    <p>Discount: {discount:.2f}</p>
    <p>Customers: {customers}</p>
    """

    xlsx = [
        ["Total Sales", fmt(gross)],
        ["Net", fmt(net)],
        ["GST", fmt(gst)],
        ["PST", fmt(pst)],
        ["Liquor Tax", fmt(liq)],
        ["Discount", fmt(discount)],
        ["Customers", customers],
    ]

    return html, xlsx


# ================= CATEGORY =================
def category(conn, start, end):

    cur = conn.cursor()

    rows = cur.execute("""
    SELECT C.SubCategoryID,
           SUM(J.Amount) AS Amt,
           SUM(J.Quantity) AS Qty
    FROM Journal J
    LEFT JOIN Category C ON C.CategoryID = J.CategoryID
    WHERE J.DateR >= ? AND J.DateR <= ?
      AND J.Status = 0
      AND J.TransType IN (101,102,111,112)
    GROUP BY C.SubCategoryID
    """, start, end).fetchall()

    html = "<h3>Category Report</h3>"
    xlsx = [["Category", "Amount", "Qty"]]

    for r in rows:
        cat = r.SubCategoryID or "UNSPECIFIED"
        html += f"<p>{cat}: {r.Amt:.2f}</p>"
        xlsx.append([cat, fmt(r.Amt), int(r.Qty or 0)])

    return html, xlsx


# ================= INDEX =================
def update_index(dates):
    data = {
        "latest": dates[0],
        "dates": dates,
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    write_html(INDEX_JSON, json.dumps(data, indent=2))


# ================= GIT =================
def run(cmd):
    return subprocess.run(cmd, cwd=str(REPO_ROOT), text=True)


def git_push():
    run(["git", "add", "."])
    run(["git", "commit", "-m", "auto reports"])
    run(["git", "pull", "--rebase", "origin", "main"])
    run(["git", "push", "origin", "main"])


# ================= MAIN =================
def main():

    ensure(REPORTS_DIR)

    dates = []

    with db() as conn:

        d = START_DATE
        while d <= END_DATE:

            ds = d.strftime("%Y-%m-%d")
            dates.append(ds)

            start, end = business_window(d)
            folder = REPORTS_DIR / ds

            try:
                d_html, d_x = daily(conn, start, end)
                c_html, c_x = category(conn, start, end)

                write_html(folder / "summary_daily.html", d_html)
                write_html(folder / "category_report.html", c_html)

                # IMPORTANT: filenames match website
                write_xlsx(folder / "summary_daily.xlsx", ["Description", "Value"], d_x)
                write_xlsx(folder / "category_report.xlsx", ["Category", "Amount", "Qty"], c_x)

                print("OK", ds)

            except Exception as e:
                print("FAIL", ds, e)

            d += timedelta(days=1)

    dates.sort(reverse=True)
    update_index(dates)

    git_push()

    print("DONE")


if __name__ == "__main__":
    main()