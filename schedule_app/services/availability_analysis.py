"""Shift availability parsing, Hope Drive grouping and weekly-capacity calculations.

Extracted from the supplied app; this module performs no page rendering on import.
"""

import pandas as pd
import re


def is_date_header_row(series, min_dates=3):
    """A row is a 'date header' if it has >= min_dates parsable dates in columns 2+."""
    parsed = pd.to_datetime(series, errors="coerce")
    return parsed.notna().sum() >= min_dates


def extract_names(cell: object) -> set[str]:
    """
    From a single Excel cell, return unique preceptor names.
    - Ignores any 'Closed' (case-insens., 'Closed?')
    - Requires '~' marker
    - Splits multiple entries on ';', '/', or 2+ spaces
    """
    if not isinstance(cell, str):
        return set()
    s = cell.replace("\r", " ").replace("\n", " ").strip()
    if re.search(r"\bclosed\b", s, flags=re.IGNORECASE):
        return set()
    if "~" not in s:
        return set()
    parts = re.split(r"[;/]| {2,}", s)
    return {p.replace("~", "").strip() for p in parts if p.replace("~", "").strip()}


def build_segmented_name_map(excel: pd.ExcelFile) -> dict:
    """
    Return dict: (site, date, shift_label) -> set(names), computed SEGMENT-BY-SEGMENT.
    A 'segment' begins at each date-header row and ends before the next date-header row.
    """
    bucket = {}
    for sheet in excel.sheet_names:
        df = pd.read_excel(excel, sheet_name=sheet, header=None)
        header_rows = [r for r in range(len(df)) if is_date_header_row(df.iloc[r, 1:])]
        if not header_rows:
            continue
        header_rows.append(len(df))  # sentinel end

        valid = (
            {"AM - ACUTES", "AM - CONTINUITY", "PM - ACUTES", "PM - CONTINUITY"}
            if sheet == "HOPE_DRIVE" else {"AM", "PM"}
        )

        for h in range(len(header_rows) - 1):
            date_row, end_row = header_rows[h], header_rows[h+1]
            dates = pd.to_datetime(df.iloc[date_row, 1:], errors="coerce")

            seg_bucket = {}
            for i in range(date_row + 1, end_row):
                label = str(df.iat[i, 0]).strip().upper()
                if label in valid:
                    for j, d in enumerate(dates, start=1):
                        if pd.isna(d):
                            continue
                        names = extract_names(df.iat[i, j])
                        if not names:
                            continue
                        key = (sheet, pd.Timestamp(d).date(), label)
                        seg_bucket.setdefault(key, set()).update(names)

            for k, s in seg_bucket.items():
                bucket.setdefault(k, set()).update(s)
    return bucket


def fold_hope_drive_rows(sub_df: pd.DataFrame):
    """
    For HOPE_DRIVE on a given date, combine:
      - AM = (AM - ACUTES) ∪ (AM - CONTINUITY)
      - PM = (PM - ACUTES) ∪ (PM - CONTINUITY)
    Return list of dict rows with Shift in {'AM','PM'}, Names list, Count.
    """
    am, pm = set(), set()
    for _, r in sub_df.iterrows():
        if r["Shift"].startswith("AM"):
            am |= set(r["Names"])
        elif r["Shift"].startswith("PM"):
            pm |= set(r["Names"])
    out = []
    if am:
        out.append({"Shift": "AM", "Names": sorted(am), "Count": len(am)})
    if pm:
        out.append({"Shift": "PM", "Names": sorted(pm), "Count": len(pm)})
    return out


def shift_order_for(site):
    return ["AM - ACUTES", "AM - CONTINUITY", "PM - ACUTES", "PM - CONTINUITY"] if site == "HOPE_DRIVE" else ["AM", "PM"]


def daily_caps_three_sites(day_df: pd.DataFrame) -> pd.DataFrame:
    df3 = day_df[day_df["Site"].isin({"ETOWN", "HOPE_DRIVE", "NYES"})].copy()
    df3["Weekday"] = df3["Date"].dt.weekday
    df3 = df3[df3["Weekday"] <= 4]  # Mon-Fri
    def ampm_bucketize(group):
        am_cap = group.loc[group["Shift"].str.startswith("AM"), "Count"].sum()
        pm_cap = group.loc[group["Shift"].str.startswith("PM"), "Count"].sum()
        return pd.Series({"AM_Capacity": am_cap, "PM_Capacity": pm_cap})
    dc = df3.groupby("Date").apply(ampm_bucketize).reset_index()
    dc["WeekStart"] = dc["Date"] - pd.to_timedelta(dc["Date"].dt.weekday, unit="D")
    return dc


def weekly_student_capacity(g):
    am_vals = sorted(g["AM_Capacity"].tolist())      # Mon..Fri
    pm_vals = sorted(g["PM_Capacity"].tolist())      # Mon..Fri
    S_am = am_vals[0] if am_vals else 0              # AM: all 5 days → min
    S_pm = pm_vals[1] if len(pm_vals) >= 2 else (pm_vals[0] if pm_vals else 0)  # PM: drop 1 → second-smallest
    return pd.Series({"AM_students_max": S_am, "PM_students_max": S_pm, "Total_students_max": S_am + S_pm})
