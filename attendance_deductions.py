#!/usr/bin/env python3
"""
Attendance Deduction Calculator
================================

Reads an employee attendance Excel sheet, applies attendance deduction rules,
and produces:

  * a new Excel workbook that preserves every original sheet and column, writes
    the calculated amounts back into the native ``Latness``, ``Miss FB IN`` and
    ``Miss FB Out`` columns, adds audit columns, and adds a per-employee
    ``Deductions Summary`` sheet (including *overtime after deduction*);
  * an HTML deduction report covering every employee; and
  * (optionally) an e-mail of that report, sent over SMTP.

DEDUCTION RULES
---------------
1/2. Missing check-IN / check-OUT: the FIRST missed IN (and, separately, the
     first missed OUT) in the whole month is free; every later one costs 0.5
     working day. Counted per employee across the month; a "missed IN" is only
     counted when the person was present that day (an OUT punch exists), and a
     "missed OUT" only when an IN punch exists. Fully-blank days are ignored.
3-5. Late arrival: 9:20-9:30 -> 0.25, 9:31-9:45 -> 0.5, 9:46-10:00 -> 1.0 day.
     At/before 9:19 = grace (0). After 10:00 = no deduction.
6.   Type contains "Mission"/"Excuse"      -> no deduction that day.
7.   Note contains "Annual", "بدون بصمة" or
     "Electricity Issue"                    -> no deduction that day.

An exempt day does not consume the "first one is free" allowance.

OVERTIME AFTER DEDUCTION
------------------------
Overtime is recorded in hours in the ``Over Time H`` column. Per employee:

    overtime_days      = total overtime hours / WORKING_DAY_HOURS
    net overtime (days)= overtime_days - total_deduction_days
    net overtime (hrs) = total overtime hours - total_deduction_days * WORKING_DAY_HOURS

A negative net means the deductions outweigh the overtime.

USAGE
-----
    python attendance_deductions.py -i IN.xlsx -o OUT.xlsx
    python attendance_deductions.py -i IN.xlsx -o OUT.xlsx --report report.html

    # also e-mail the report (SMTP settings come from environment variables):
    export SMTP_HOST=smtp.gmail.com SMTP_PORT=587
    export SMTP_USER=you@gmail.com SMTP_PASSWORD='app-password'
    python attendance_deductions.py -i IN.xlsx -o OUT.xlsx \
        --send-email --email-to hr@company.com

    # per-employee e-mails (needs an e-mail column in the sheet):
    python attendance_deductions.py -i IN.xlsx -o OUT.xlsx \
        --send-email --email-mode individual --email-col Email
"""

from __future__ import annotations

import argparse
import datetime as dt
import mimetypes
import os
import smtplib
import ssl
from email.message import EmailMessage
from html import escape

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# CONFIGURATION  (edit these to adjust the policy -- no other code changes)
# ---------------------------------------------------------------------------

SHEET_NAME = "effects"          # worksheet holding the daily attendance rows

# Column names as they appear in the sheet (trailing/leading spaces match the
# real headers; matching is done on the stripped name so spaces don't matter).
COL_CODE = "HR Code"
COL_NAME = "name"
COL_DATE = "Date"
COL_IN = "IN"
COL_OUT = "OUT"
COL_TYPE = "Type"
COL_NOTE = "note"
COL_OT_HOURS = "Over Time H"    # overtime recorded in hours

# Native deduction columns that are UPDATED in place with the calculation.
COL_LATENESS = "Latness"        # <- late-arrival deduction (days)
COL_MISS_IN = "Miss FB IN"      # <- repeated missing-IN deduction (days)
COL_MISS_OUT = "Miss FB Out"    # <- repeated missing-OUT deduction (days)

# One working day equals this many hours (used to convert overtime hours).
WORKING_DAY_HOURS = 8.0

# Late-arrival tiers: (start, end_inclusive, deduction_in_days).
LATE_TIERS = [
    (dt.time(9, 20, 0), dt.time(9, 30, 59), 0.25),
    (dt.time(9, 31, 0), dt.time(9, 45, 59), 0.50),
    (dt.time(9, 46, 0), dt.time(10, 0, 0), 1.00),
]
GRACE_UNTIL = dt.time(9, 19, 59)          # at/before this = on time
MISSING_PUNCH_DEDUCTION = 0.5             # every repeated missing IN/OUT

EXEMPT_TYPES = ["mission", "excuse"]      # Type substrings (case-insensitive)
EXEMPT_NOTE_KEYWORDS = ["annual", "بدون بصمة", "electricity issue"]


# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------

def has_value(value) -> bool:
    if value is None:
        return False
    try:
        if pd.isna(value):
            return False
    except (TypeError, ValueError):
        pass
    if isinstance(value, str) and not value.strip():
        return False
    return True


def parse_time(value):
    """datetime.time for a cell, or None if blank/unparseable."""
    if not has_value(value):
        return None
    if isinstance(value, dt.time):
        return value
    if isinstance(value, dt.datetime):
        return value.time()
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime().time()
    text = str(value).strip()
    for fmt in ("%H:%M:%S", "%H:%M", "%I:%M:%S %p", "%I:%M %p", "%I:%M%p"):
        try:
            return dt.datetime.strptime(text, fmt).time()
        except ValueError:
            continue
    return None


def to_float(value) -> float:
    """Numeric value of a cell, or 0.0 when blank/unparseable."""
    if not has_value(value):
        return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def text_of(value) -> str:
    if not has_value(value):
        return ""
    return str(value).strip().lower()


def employee_key(code, name):
    """Stable per-employee key: HR Code if present, else the (stripped) name."""
    if has_value(code):
        return ("code", code)
    return ("name", str(name).strip() if has_value(name) else "")


def late_deduction(in_time) -> float:
    if in_time is None or in_time <= GRACE_UNTIL:
        return 0.0
    for start, end, amount in LATE_TIERS:
        if start <= in_time <= end:
            return amount
    return 0.0  # after 10:00 -> no deduction


def exempt_reason(type_cell, note_cell) -> str:
    type_text = text_of(type_cell)
    for token in EXEMPT_TYPES:
        if token in type_text:
            return type_cell if isinstance(type_cell, str) else token
    note_text = text_of(note_cell)
    for token in EXEMPT_NOTE_KEYWORDS:
        if token.lower() in note_text:
            return token
    return ""


# ---------------------------------------------------------------------------
# CORE CALCULATION
# ---------------------------------------------------------------------------

# Logical result columns produced per row.
CALC_COLS = ["late", "miss_in", "miss_out", "day_total",
             "in_occ", "out_occ", "reason"]

# Audit columns appended to the sheet (native columns are updated separately).
APPENDED_COLS = ["Day Deduction", "Missing IN #", "Missing OUT #",
                 "Exempt Reason"]


def compute(df: pd.DataFrame) -> pd.DataFrame:
    """Per-row deductions, aligned to df.index."""
    rows = {c: pd.Series(index=df.index, dtype="object") for c in CALC_COLS}
    miss_in_seen: dict = {}
    miss_out_seen: dict = {}

    has_date = df[COL_DATE].apply(has_value)
    work = df[has_date].copy()
    work["_order"] = range(len(work))
    work = work.sort_values([COL_CODE, COL_DATE, "_order"], kind="stable")

    for idx, row in work.iterrows():
        emp = employee_key(row[COL_CODE], row[COL_NAME])
        in_time = parse_time(row[COL_IN])
        out_time = parse_time(row[COL_OUT])
        in_present = in_time is not None
        out_present = out_time is not None
        reason = exempt_reason(row[COL_TYPE], row[COL_NOTE])

        late = miss_in = miss_out = 0.0
        in_occ = out_occ = ""

        if not reason:
            if in_present:
                late = late_deduction(in_time)
            if out_present and not in_present:          # missed IN (present)
                n = miss_in_seen.get(emp, 0) + 1
                miss_in_seen[emp] = n
                in_occ = n
                miss_in = 0.0 if n == 1 else MISSING_PUNCH_DEDUCTION
            if in_present and not out_present:           # missed OUT (present)
                n = miss_out_seen.get(emp, 0) + 1
                miss_out_seen[emp] = n
                out_occ = n
                miss_out = 0.0 if n == 1 else MISSING_PUNCH_DEDUCTION

        rows["late"][idx] = late
        rows["miss_in"][idx] = miss_in
        rows["miss_out"][idx] = miss_out
        rows["day_total"][idx] = round(late + miss_in + miss_out, 4)
        rows["in_occ"][idx] = in_occ
        rows["out_occ"][idx] = out_occ
        rows["reason"][idx] = reason

    for c in ("late", "miss_in", "miss_out", "day_total"):
        rows[c] = rows[c].fillna(0.0)
    for c in ("in_occ", "out_occ", "reason"):
        rows[c] = rows[c].fillna("")
    return pd.DataFrame(rows, index=df.index)


def summarize(df: pd.DataFrame, calc: pd.DataFrame) -> pd.DataFrame:
    """One row per employee, including overtime-after-deduction."""
    merged = df.copy()
    for c in CALC_COLS:
        merged[c] = calc[c]
    merged["ot_hours"] = merged[COL_OT_HOURS].apply(to_float) \
        if COL_OT_HOURS in merged.columns else 0.0
    merged = merged[merged[COL_DATE].apply(has_value)].copy()

    def first_value(s):
        for v in s:
            if has_value(v):
                return v
        return None

    def first_name(s):
        v = first_value(s)
        return str(v).strip() if v is not None else ""

    merged["_emp"] = [employee_key(c, n)
                      for c, n in zip(merged[COL_CODE], merged[COL_NAME])]
    g = merged.groupby("_emp", dropna=False)

    late = g["late"].sum()
    mi = g["miss_in"].sum()
    mo = g["miss_out"].sum()
    total = g["day_total"].sum()
    ot_h = g["ot_hours"].sum()
    ot_d = ot_h / WORKING_DAY_HOURS

    summary = pd.DataFrame({
        "HR Code": g[COL_CODE].apply(first_value).values,
        "Name": g[COL_NAME].apply(first_name).values,
        "Late (days)": late.round(2).values,
        "Missing IN (days)": mi.round(2).values,
        "Missing OUT (days)": mo.round(2).values,
        "Total Deduction (days)": total.round(2).values,
        "Overtime (hrs)": ot_h.round(2).values,
        "Overtime (days)": ot_d.round(2).values,
        "Net Overtime (days)": (ot_d - total).round(2).values,
        "Net Overtime (hrs)": (ot_h - total * WORKING_DAY_HOURS).round(2).values,
        "# Late": g["late"].apply(lambda s: int((s > 0).sum())).values,
        "# Missing IN": g["in_occ"].apply(
            lambda s: int(sum(1 for v in s if v != ""))).values,
        "# Missing OUT": g["out_occ"].apply(
            lambda s: int(sum(1 for v in s if v != ""))).values,
    })
    try:
        summary["HR Code"] = summary["HR Code"].astype("Int64")
    except (TypeError, ValueError):
        pass
    return summary.sort_values("Total Deduction (days)",
                               ascending=False).reset_index(drop=True)


# ---------------------------------------------------------------------------
# EXCEL OUTPUT
# ---------------------------------------------------------------------------

HEADER_FILL = PatternFill("solid", fgColor="4F46E5")
HEADER_FONT = Font(bold=True, color="FFFFFF")


def _header_columns(ws, header_row=1):
    """Map stripped header name -> column index for the header row."""
    cols = {}
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=header_row, column=col).value
        if val not in (None, ""):
            cols[str(val).strip()] = col
    return cols


def write_output(input_path, output_path, sheet_name, calc, summary):
    wb = load_workbook(input_path)       # keeps all sheets + formatting
    ws = wb[sheet_name]
    header_row = 1
    cols = _header_columns(ws, header_row)
    last_col = max(cols.values()) if cols else 0

    # Update native deduction columns in place (blank when 0, to match the
    # original "amount or blank" convention).
    native = {
        COL_LATENESS.strip(): "late",
        COL_MISS_IN.strip(): "miss_in",
        COL_MISS_OUT.strip(): "miss_out",
    }
    for header_name, calc_key in native.items():
        col = cols.get(header_name)
        if col is None:                   # column not present -> append it
            last_col += 1
            col = last_col
            c = ws.cell(row=header_row, column=col, value=header_name)
            c.fill, c.font = HEADER_FILL, HEADER_FONT
        for i in range(len(calc)):
            amount = calc[calc_key].iloc[i]
            # Assign .value directly so stale pre-existing values (e.g. the old
            # Latness figures) are CLEARED when the new amount is 0 -- the
            # cell(value=None) shortcut would leave them untouched.
            ws.cell(row=i + 2, column=col).value = amount if amount else None

    # Append audit columns.
    start = last_col + 1
    for off, name in enumerate(APPENDED_COLS):
        c = ws.cell(row=header_row, column=start + off, value=name)
        c.fill, c.font = HEADER_FILL, HEADER_FONT
    key_for = {"Day Deduction": "day_total", "Missing IN #": "in_occ",
               "Missing OUT #": "out_occ", "Exempt Reason": "reason"}
    for i in range(len(calc)):
        for off, name in enumerate(APPENDED_COLS):
            val = calc[key_for[name]].iloc[i]
            if name == "Day Deduction":
                val = float(val)
            elif val == "":
                val = None
            ws.cell(row=i + 2, column=start + off).value = val

    # (Re)build the summary sheet.
    if "Deductions Summary" in wb.sheetnames:
        del wb["Deductions Summary"]
    sm = wb.create_sheet("Deductions Summary")
    sm.append(list(summary.columns))
    for col in range(1, len(summary.columns) + 1):
        cell = sm.cell(row=1, column=col)
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
    for _, r in summary.iterrows():
        sm.append([None if pd.isna(v) else v for v in r.tolist()])
    for col in range(1, len(summary.columns) + 1):
        sm.column_dimensions[get_column_letter(col)].width = 18

    wb.save(output_path)


# ---------------------------------------------------------------------------
# HTML REPORT
# ---------------------------------------------------------------------------

def _fmt_date(v):
    if isinstance(v, (dt.datetime, pd.Timestamp)):
        return v.strftime("%Y-%m-%d")
    return str(v)


def build_report_html(df, calc, summary, today_str):
    dates = df[COL_DATE].dropna()
    period = (f"{_fmt_date(dates.min())} – {_fmt_date(dates.max())}"
              if len(dates) else "n/a")
    tot_ded = summary["Total Deduction (days)"].sum()
    tot_ot_h = summary["Overtime (hrs)"].sum()
    n_emp = len(summary)
    n_with = int((summary["Total Deduction (days)"] > 0).sum())

    def cls(v):
        return "neg" if v < 0 else ("pos" if v > 0 else "")

    # Summary table rows.
    srows = []
    for _, r in summary.iterrows():
        net = r["Net Overtime (days)"]
        srows.append(
            "<tr>"
            f"<td>{escape(str(r['HR Code']) if pd.notna(r['HR Code']) else '')}</td>"
            f"<td class='name'>{escape(str(r['Name']))}</td>"
            f"<td>{r['Late (days)']:.2f}</td>"
            f"<td>{r['Missing IN (days)']:.2f}</td>"
            f"<td>{r['Missing OUT (days)']:.2f}</td>"
            f"<td class='tot'>{r['Total Deduction (days)']:.2f}</td>"
            f"<td>{r['Overtime (hrs)']:.2f}</td>"
            f"<td>{r['Overtime (days)']:.2f}</td>"
            f"<td class='{cls(net)}'><b>{net:+.2f}</b></td>"
            "</tr>")

    # Per-employee detail (only days that caused a deduction).
    merged = df.copy()
    for c in CALC_COLS:
        merged[c] = calc[c]
    merged["_emp"] = [employee_key(c, n)
                      for c, n in zip(merged[COL_CODE], merged[COL_NAME])]
    details = []
    for _, r in summary[summary["Total Deduction (days)"] > 0].iterrows():
        key = employee_key(r["HR Code"] if pd.notna(r["HR Code"]) else None,
                           r["Name"])
        sub = merged[(merged["_emp"] == key) & (merged["day_total"] > 0)]
        if sub.empty:
            continue
        drows = []
        for _, d in sub.iterrows():
            parts = []
            if d["late"]:
                parts.append(f"late {d['late']:.2f}")
            if d["miss_in"]:
                parts.append(f"missing IN {d['miss_in']:.2f}")
            if d["miss_out"]:
                parts.append(f"missing OUT {d['miss_out']:.2f}")
            drows.append(
                "<tr>"
                f"<td>{_fmt_date(d[COL_DATE])}</td>"
                f"<td>{escape(str(d[COL_IN]) if has_value(d[COL_IN]) else '—')}</td>"
                f"<td>{escape(str(d[COL_OUT]) if has_value(d[COL_OUT]) else '—')}</td>"
                f"<td>{', '.join(parts)}</td>"
                f"<td class='tot'>{d['day_total']:.2f}</td>"
                "</tr>")
        details.append(
            f"<h3>{escape(str(r['Name']))} "
            f"<span class='sub'>(HR {escape(str(r['HR Code']) if pd.notna(r['HR Code']) else '—')}) "
            f"· total {r['Total Deduction (days)']:.2f} day(s) · "
            f"net overtime {r['Net Overtime (days)']:+.2f} day(s)</span></h3>"
            "<table class='detail'><thead><tr>"
            "<th>Date</th><th>IN</th><th>OUT</th><th>Reason</th><th>Deduction</th>"
            "</tr></thead><tbody>" + "".join(drows) + "</tbody></table>")

    return f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<title>Attendance Deduction Report</title>
<style>
  body{{font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif;color:#1a1f36;margin:0;padding:24px;background:#f7f8fc}}
  h1{{margin:0 0 4px}} .meta{{color:#6b7280;margin-bottom:18px}}
  .cards{{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:22px}}
  .card{{background:#fff;border:1px solid #e5e7eb;border-radius:10px;padding:12px 16px;min-width:150px}}
  .card .k{{color:#6b7280;font-size:12px}} .card .v{{font-size:22px;font-weight:700}}
  table{{border-collapse:collapse;width:100%;background:#fff;margin-bottom:26px;font-size:13px}}
  th,td{{border:1px solid #e5e7eb;padding:6px 9px;text-align:right}}
  th{{background:#4f46e5;color:#fff;position:sticky;top:0}}
  td.name,th:nth-child(2){{text-align:left}} td:first-child,th:first-child{{text-align:left}}
  td.tot{{font-weight:700}} .pos{{color:#15803d}} .neg{{color:#b91c1c}}
  h3{{margin:18px 0 6px}} h3 .sub{{font-weight:400;color:#6b7280;font-size:13px}}
  table.detail th{{background:#eef0fb;color:#1a1f36}}
  table.detail td:nth-child(4){{text-align:left}}
</style></head><body>
<h1>Attendance Deduction Report</h1>
<div class="meta">Period {period} &nbsp;·&nbsp; generated {today_str}</div>
<div class="cards">
  <div class="card"><div class="k">Employees</div><div class="v">{n_emp}</div></div>
  <div class="card"><div class="k">With deductions</div><div class="v">{n_with}</div></div>
  <div class="card"><div class="k">Total deduction</div><div class="v">{tot_ded:.2f} d</div></div>
  <div class="card"><div class="k">Total overtime</div><div class="v">{tot_ot_h:.0f} h</div></div>
</div>
<h2>Summary — all employees</h2>
<table><thead><tr>
  <th>HR</th><th>Name</th><th>Late</th><th>Miss IN</th><th>Miss OUT</th>
  <th>Total ded. (d)</th><th>OT (h)</th><th>OT (d)</th><th>Net OT (d)</th>
</tr></thead><tbody>{''.join(srows)}</tbody></table>
<h2>Detail — employees with deductions</h2>
{''.join(details)}
</body></html>"""


# ---------------------------------------------------------------------------
# EMAIL
# ---------------------------------------------------------------------------

def _smtp_config():
    host = os.environ.get("SMTP_HOST")
    if not host:
        return None
    return {
        "host": host,
        "port": int(os.environ.get("SMTP_PORT", "587")),
        "user": os.environ.get("SMTP_USER"),
        "password": os.environ.get("SMTP_PASSWORD"),
        "from": os.environ.get("SMTP_FROM", os.environ.get("SMTP_USER", "")),
    }


def _send(cfg, to_addr, subject, html, attachments=None):
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = cfg["from"]
    msg["To"] = to_addr
    msg.set_content("This report is formatted in HTML. Please view it in an "
                    "HTML-capable mail client.")
    msg.add_alternative(html, subtype="html")
    for path in attachments or []:
        ctype, _ = mimetypes.guess_type(path)
        maintype, subtype = (ctype or "application/octet-stream").split("/", 1)
        with open(path, "rb") as fh:
            msg.add_attachment(fh.read(), maintype=maintype, subtype=subtype,
                               filename=os.path.basename(path))
    context = ssl.create_default_context()
    with smtplib.SMTP(cfg["host"], cfg["port"]) as s:
        s.starttls(context=context)
        if cfg["user"]:
            s.login(cfg["user"], cfg["password"])
        s.send_message(msg)


def email_reports(args, df, calc, summary, today_str, attachments):
    cfg = _smtp_config()
    if cfg is None:
        print("  [email] SMTP_HOST not set — skipping send "
              "(set SMTP_HOST/PORT/USER/PASSWORD to enable).")
        return

    if args.email_mode == "individual":
        col = args.email_col
        if col not in df.columns:
            print(f"  [email] individual mode needs an e-mail column "
                  f"({col!r}) in the sheet — not found; nothing sent.")
            return
        merged = df.copy()
        for c in CALC_COLS:
            merged[c] = calc[c]
        merged["_emp"] = [employee_key(c, n)
                          for c, n in zip(merged[COL_CODE], merged[COL_NAME])]
        sent = 0
        for _, r in summary.iterrows():
            key = employee_key(r["HR Code"] if pd.notna(r["HR Code"]) else None,
                               r["Name"])
            addrs = merged[merged["_emp"] == key][col].dropna().unique()
            if len(addrs) == 0:
                continue
            one = summary[summary["Name"] == r["Name"]]
            html = build_report_html(df, calc, one, today_str)
            _send(cfg, str(addrs[0]),
                  f"Your attendance deductions — {r['Name']}", html)
            sent += 1
        print(f"  [email] sent {sent} individual report(s).")
    else:
        if not args.email_to:
            print("  [email] --email-to is required for summary mode; "
                  "nothing sent.")
            return
        html = build_report_html(df, calc, summary, today_str)
        _send(cfg, args.email_to, args.email_subject, html, attachments)
        print(f"  [email] summary report sent to {args.email_to}.")


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(
        description="Calculate attendance deductions from an Excel sheet.")
    ap.add_argument("-i", "--input", default="attendance.xlsx")
    ap.add_argument("-o", "--output", default="attendance_deductions_output.xlsx")
    ap.add_argument("--sheet", default=SHEET_NAME)
    ap.add_argument("--report", default=None,
                    help="path for the HTML report (default: <output>.html)")
    ap.add_argument("--send-email", action="store_true",
                    help="also e-mail the report (needs SMTP_* env vars)")
    ap.add_argument("--email-mode", choices=["summary", "individual"],
                    default="summary")
    ap.add_argument("--email-to", default=None,
                    help="recipient for the summary report e-mail")
    ap.add_argument("--email-from", default=None)
    ap.add_argument("--email-subject",
                    default="Attendance Deduction Report")
    ap.add_argument("--email-col", default="Email",
                    help="sheet column holding employee e-mails (individual mode)")
    args = ap.parse_args()
    if args.email_from:
        os.environ["SMTP_FROM"] = args.email_from
    today_str = dt.date.today().strftime("%Y-%m-%d")

    print(f"Reading {args.input!r} (sheet {args.sheet!r}) ...")
    df = pd.read_excel(args.input, sheet_name=args.sheet)
    print(f"  {len(df)} rows, {df[COL_CODE].nunique(dropna=True)} employees")

    calc = compute(df)
    summary = summarize(df, calc)
    total = summary["Total Deduction (days)"].sum()
    net_ot = summary["Net Overtime (days)"].sum()
    print(f"  total deduction: {total:.2f} working days "
          f"| net overtime after deduction: {net_ot:+.2f} days")

    write_output(args.input, args.output, args.sheet, calc, summary)
    print(f"  workbook written to {args.output!r}")

    report_path = args.report or (os.path.splitext(args.output)[0] + ".html")
    html = build_report_html(df, calc, summary, today_str)
    with open(report_path, "w", encoding="utf-8") as fh:
        fh.write(html)
    print(f"  HTML report written to {report_path!r}")

    if args.send_email:
        email_reports(args, df, calc, summary, today_str, [args.output])

    print("Done.")


if __name__ == "__main__":
    main()
