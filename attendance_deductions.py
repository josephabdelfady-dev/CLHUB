#!/usr/bin/env python3
"""
Attendance Deduction Calculator
================================

Reads an employee attendance Excel sheet, applies a set of attendance
deduction rules, and writes a new Excel file that:

  * preserves every original sheet and column, and
  * adds per-day deduction columns to the attendance sheet, plus
  * a "Deductions Summary" sheet with one row per employee.

RULES IMPLEMENTED
-----------------
1. Missing check-IN:  the FIRST missed IN in the whole month is free; every
   later missed IN costs 0.5 working day.
2. Missing check-OUT: the FIRST missed OUT in the whole month is free; every
   later missed OUT costs 0.5 working day.
   (IN and OUT occurrences are counted SEPARATELY, per employee, for the
    whole month, and the "first one is free" exemption is tracked
    independently for IN and for OUT.)

   A "missed IN" is only counted when the employee was actually present that
   day, i.e. there IS an OUT punch but the IN punch is blank. Likewise a
   "missed OUT" needs an IN punch present. A fully-blank day (no IN and no
   OUT -- weekend, holiday, or full absence) is NOT a fingerprint miss and is
   ignored by rules 1-2.

3. Late arrival 9:20-9:30  -> 0.25 working day (per occurrence).
4. Late arrival 9:31-9:45  -> 0.50 working day (per occurrence).
5. Late arrival 9:46-10:00 -> 1.00 working day (per occurrence).
   Arrivals at/before 9:19 are on time (grace). Arrivals AFTER 10:00 are NOT
   deducted at all.
6. Mission / Excuse: if the `Type` column contains "Mission" or "Excuse",
   NO deduction is applied for that day (late or missing).
7. Annual leave: if the `note` column contains "Annual", NO deduction is
   applied for that day.

   In addition (per the file owner's instruction) any day whose `note`
   contains "بدون بصمة" (no fingerprint) or "Electricity Issue" is waived and
   receives no deduction.

An exempt day (rules 6/7 or a waiver note) does NOT consume the "first one is
free" allowance -- only genuine, non-exempt fingerprint misses are counted.

USAGE
-----
    python attendance_deductions.py                      # uses defaults
    python attendance_deductions.py -i IN.xlsx -o OUT.xlsx
    python attendance_deductions.py --sheet effects -i IN.xlsx -o OUT.xlsx
"""

from __future__ import annotations

import argparse
import datetime as dt

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# CONFIGURATION  (edit these to adjust the policy -- no other code changes)
# ---------------------------------------------------------------------------

SHEET_NAME = "effects"          # worksheet holding the daily attendance rows

# Column names as they appear in the sheet (note the trailing/leading spaces
# are intentional -- they match the real headers).
COL_CODE = "HR Code"
COL_NAME = "name"
COL_DATE = "Date"
COL_IN = "IN"
COL_OUT = "OUT"
COL_TYPE = "Type"
COL_NOTE = "note"

# Late-arrival tiers. Each tuple is (start, end_inclusive, deduction_in_days).
# A punch time t is matched when start <= t <= end_inclusive.
LATE_TIERS = [
    (dt.time(9, 20, 0), dt.time(9, 30, 59), 0.25),
    (dt.time(9, 31, 0), dt.time(9, 45, 59), 0.50),
    (dt.time(9, 46, 0), dt.time(10, 0, 0), 1.00),
]
# Punches at/before this are on time (grace); punches strictly after the last
# tier's end (i.e. after 10:00:00) are NOT deducted.
GRACE_UNTIL = dt.time(9, 19, 59)

# Deduction for every REPEATED missing IN / OUT (the first is always free).
MISSING_PUNCH_DEDUCTION = 0.5

# Values in the `Type` column that fully exempt a day (case-insensitive,
# substring match).
EXEMPT_TYPES = ["mission", "excuse"]

# Substrings in the `note` column that fully exempt a day (case-insensitive).
#   - "annual"          -> rule 7 (annual leave)
#   - "بدون بصمة"        -> "no fingerprint" (device / punch not recorded)
#   - "electricity issue"-> power outage affecting the fingerprint device
# Add more keywords here (e.g. "grace", "no fingerprint deduction") to waive
# those days too.
EXEMPT_NOTE_KEYWORDS = ["annual", "بدون بصمة", "electricity issue"]


# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------

def parse_time(value):
    """Return a datetime.time for a cell, or None if blank / unparseable.

    Handles datetime.time, datetime.datetime, pandas.Timestamp, and common
    string formats like "9:30", "09:30:00", "9:30 AM".
    """
    if value is None:
        return None
    if isinstance(value, float) and pd.isna(value):
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    if isinstance(value, dt.time):
        return value
    if isinstance(value, dt.datetime):
        return value.time()
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime().time()

    text = str(value).strip()
    if not text:
        return None
    for fmt in ("%H:%M:%S", "%H:%M", "%I:%M:%S %p", "%I:%M %p", "%I:%M%p"):
        try:
            return dt.datetime.strptime(text, fmt).time()
        except ValueError:
            continue
    return None


def has_value(value) -> bool:
    """True when a cell holds a real (non-blank) value."""
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


def text_of(value) -> str:
    """Lower-cased, stripped text of a cell (empty string when blank)."""
    if not has_value(value):
        return ""
    return str(value).strip().lower()


def employee_key(code, name):
    """Stable per-employee grouping key.

    Uses the HR Code when present, otherwise falls back to the (stripped)
    name, so an employee with a blank HR Code is never dropped or merged
    with another code-less employee.
    """
    if has_value(code):
        return ("code", code)
    return ("name", str(name).strip() if has_value(name) else "")


def late_deduction(in_time) -> float:
    """Working-day deduction for a given IN time (0.0 when on time/after 10)."""
    if in_time is None:
        return 0.0
    if in_time <= GRACE_UNTIL:
        return 0.0
    for start, end, amount in LATE_TIERS:
        if start <= in_time <= end:
            return amount
    return 0.0  # after 10:00 -> no deduction


def exempt_reason(type_cell, note_cell) -> str:
    """Return why a day is exempt from all deductions, or '' if not exempt."""
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

# Output column names (appended to the attendance sheet).
OUT_COLS = [
    "Late Deduction",
    "Missing IN Deduction",
    "Missing OUT Deduction",
    "Day Deduction",
    "Missing IN #",
    "Missing OUT #",
    "Exempt Reason",
]


def compute(df: pd.DataFrame) -> pd.DataFrame:
    """Compute per-row deductions. Returns a DataFrame aligned to df.index."""
    rows = {c: pd.Series(index=df.index, dtype="object") for c in OUT_COLS}

    # Per-employee running counters (whole-month, counted in date order).
    miss_in_seen: dict = {}
    miss_out_seen: dict = {}

    # Iterate employee by employee, each in chronological (date) order, so the
    # "first missed IN/OUT is free" rule is applied to the earliest dates.
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
            # Rule 3-5: late arrival (needs an IN time).
            if in_present:
                late = late_deduction(in_time)

            # Rules 1-2: missed fingerprint while present.
            if out_present and not in_present:          # missed IN
                n = miss_in_seen.get(emp, 0) + 1
                miss_in_seen[emp] = n
                in_occ = n
                miss_in = 0.0 if n == 1 else MISSING_PUNCH_DEDUCTION
            if in_present and not out_present:           # missed OUT
                n = miss_out_seen.get(emp, 0) + 1
                miss_out_seen[emp] = n
                out_occ = n
                miss_out = 0.0 if n == 1 else MISSING_PUNCH_DEDUCTION

        day_total = round(late + miss_in + miss_out, 4)

        rows["Late Deduction"][idx] = late
        rows["Missing IN Deduction"][idx] = miss_in
        rows["Missing OUT Deduction"][idx] = miss_out
        rows["Day Deduction"][idx] = day_total
        rows["Missing IN #"][idx] = in_occ
        rows["Missing OUT #"][idx] = out_occ
        rows["Exempt Reason"][idx] = reason

    # Fill rows without a date (spacer/blank rows) with zeros / blanks.
    for c in ("Late Deduction", "Missing IN Deduction",
              "Missing OUT Deduction", "Day Deduction"):
        rows[c] = rows[c].fillna(0.0)
    for c in ("Missing IN #", "Missing OUT #", "Exempt Reason"):
        rows[c] = rows[c].fillna("")

    return pd.DataFrame(rows, index=df.index)


def summarize(df: pd.DataFrame, calc: pd.DataFrame) -> pd.DataFrame:
    """Build a one-row-per-employee summary of the month's deductions."""
    merged = df.copy()
    for c in OUT_COLS:
        merged[c] = calc[c]
    merged = merged[merged[COL_DATE].apply(has_value)]

    def first_value(s):
        for v in s:
            if has_value(v):
                return v
        return None

    def first_name(s):
        v = first_value(s)
        return str(v).strip() if v is not None else ""

    # Group by a stable key (HR Code, else name) so no employee is lost.
    merged = merged.copy()
    merged["_emp"] = [employee_key(c, n)
                      for c, n in zip(merged[COL_CODE], merged[COL_NAME])]
    grp = merged.groupby("_emp", dropna=False)
    summary = pd.DataFrame({
        "HR Code": grp[COL_CODE].apply(first_value).values,
        "Name": grp[COL_NAME].apply(first_name).values,
        "Late Deduction": grp["Late Deduction"].sum().round(2).values,
        "Missing IN Deduction": grp["Missing IN Deduction"].sum().round(2).values,
        "Missing OUT Deduction": grp["Missing OUT Deduction"].sum().round(2).values,
        "Total Deduction (days)": grp["Day Deduction"].sum().round(2).values,
        "# Late": grp["Late Deduction"].apply(lambda s: (s > 0).sum()).values,
        "# Missing IN (counted)": grp["Missing IN #"].apply(
            lambda s: sum(1 for v in s if v != "")).values,
        "# Missing OUT (counted)": grp["Missing OUT #"].apply(
            lambda s: sum(1 for v in s if v != "")).values,
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


def write_output(input_path: str, output_path: str, sheet_name: str,
                 calc: pd.DataFrame, summary: pd.DataFrame) -> None:
    """Copy the original workbook, append deduction columns, add summary."""
    wb = load_workbook(input_path)       # keeps all sheets + formatting
    ws = wb[sheet_name]

    # Find the first empty column in the header row to append after.
    header_row = 1
    last_col = 0
    for col in range(1, ws.max_column + 1):
        if ws.cell(row=header_row, column=col).value not in (None, ""):
            last_col = col
    start_col = last_col + 1

    # Write new headers.
    for offset, name in enumerate(OUT_COLS):
        c = ws.cell(row=header_row, column=start_col + offset, value=name)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT

    # Write values. pandas row index i -> worksheet row i + 2 (row 1 = header).
    for i, idx in enumerate(calc.index):
        excel_row = i + 2
        for offset, name in enumerate(OUT_COLS):
            val = calc[name].iloc[i]
            if isinstance(val, float) and val == 0.0 and name != "Day Deduction":
                val = 0
            ws.cell(row=excel_row, column=start_col + offset, value=val)

    # Add the summary sheet.
    if "Deductions Summary" in wb.sheetnames:
        del wb["Deductions Summary"]
    sm = wb.create_sheet("Deductions Summary")
    sm.append(list(summary.columns))
    for col in range(1, len(summary.columns) + 1):
        cell = sm.cell(row=1, column=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
    for _, r in summary.iterrows():
        sm.append([None if pd.isna(v) else v for v in r.tolist()])
    for col in range(1, len(summary.columns) + 1):
        sm.column_dimensions[get_column_letter(col)].width = 22

    wb.save(output_path)


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def main() -> None:
    ap = argparse.ArgumentParser(
        description="Calculate attendance deductions from an Excel sheet.")
    ap.add_argument("-i", "--input", default="attendance.xlsx",
                    help="path to the input Excel file")
    ap.add_argument("-o", "--output", default="attendance_deductions_output.xlsx",
                    help="path to write the results Excel file")
    ap.add_argument("--sheet", default=SHEET_NAME,
                    help=f"worksheet with daily rows (default: {SHEET_NAME})")
    args = ap.parse_args()

    print(f"Reading {args.input!r}  (sheet {args.sheet!r}) ...")
    df = pd.read_excel(args.input, sheet_name=args.sheet)
    print(f"  {len(df)} rows, {df[COL_CODE].nunique(dropna=True)} employees")

    calc = compute(df)
    summary = summarize(df, calc)

    total = summary["Total Deduction (days)"].sum()
    print(f"  total deduction across all employees: {total:.2f} working days")
    print(f"  employees with a deduction: "
          f"{(summary['Total Deduction (days)'] > 0).sum()}")

    write_output(args.input, args.output, args.sheet, calc, summary)
    print(f"Done. Results written to {args.output!r}")


if __name__ == "__main__":
    main()
