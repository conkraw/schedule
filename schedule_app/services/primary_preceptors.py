"""Weekly primary/fragmented preceptor rules and the Power Automate Excel report.

Extracted from the supplied app; this module performs no page rendering on import.
"""

from collections import Counter
from datetime import datetime
from datetime import timedelta
from io import BytesIO
from schedule_app.settings import FOCUS_SITES
from schedule_app.settings import PRECEPTOR_EMAIL_MAP
from schedule_app.settings import REPORT_COLUMNS
import pandas as pd
import re
import xlsxwriter


def _normalize_name(value):
    """Normalize names for reliable hard-coded email matching."""
    return re.sub(r"\s+", " ", str(value or "").strip()).casefold()


def _normalize_site(value):
    """Convert site labels such as 'Hope Drive' to 'HOPE_DRIVE'."""
    cleaned = re.sub(r"[^A-Za-z0-9]+", "_", str(value or "").strip().upper())
    return cleaned.strip("_")


def _parse_excel_date(value):
    """Return a Python date from a typical openpyxl/pandas date value."""
    if value is None or value == "":
        return None

    if isinstance(value, datetime):
        return value.date()

    if hasattr(value, "date") and not isinstance(value, str):
        try:
            return value.date()
        except Exception:
            pass

    try:
        parsed = pd.to_datetime(value, errors="coerce")
        if pd.notna(parsed):
            return parsed.date()
    except Exception:
        pass

    return None


def _week_monday_from_date_row(ws, date_row):
    """Find/infer Monday from the B:H date cells for a weekly block."""
    for day_offset, col in enumerate(range(2, 9)):
        parsed = _parse_excel_date(ws.cell(row=date_row, column=col).value)
        if parsed is not None:
            return parsed - timedelta(days=day_offset)
    return None


def _student_name_from_sheet(ws):
    """Prefer the displayed student name; fall back to the worksheet title."""
    for coordinate in ("B1", "B2"):
        value = ws[coordinate].value
        if value is not None and str(value).strip():
            return str(value).strip()
    return str(ws.title).strip()


def _parse_schedule_assignment(value):
    """Parse a student-schedule cell formatted as 'Preceptor - [SITE]'."""
    if not isinstance(value, str) or not value.strip():
        return None, None

    match = re.match(r"^\s*(.*?)\s*-\s*\[\s*([^\]]+)\s*\]\s*$", value)
    if not match:
        return None, None

    preceptor = re.sub(r"\s+", " ", match.group(1).strip())
    site = _normalize_site(match.group(2))
    return preceptor, site


def build_preceptor_assignment_report(master_wb):
    """
    Build one row per student/preceptor/week for HOPE_DRIVE, NYES, and ETOWN.

    Primary logic:
      - Every student/week represented in this report receives exactly one
        primary preceptor.
      - Prefer the preceptor with the largest session count among those with
        at least 3 sessions.
      - If nobody reaches 3 sessions, select the highest-session preceptor
        anyway and flag that primary assignment for review.
      - The same preceptor may be primary for more than one student. Repeated
        primary assignments within the same week are also flagged.
      - Ties are resolved alphabetically for stable output.

    Fragmentation logic:
      - YES when that preceptor has <3 sessions with the student that week.
      - A count of exactly 3 is not fragmented.
    """
    NORMALIZED_EMAIL_MAP = {
        _normalize_name(preceptor): str(email).strip()
        for preceptor, email in PRECEPTOR_EMAIL_MAP.items()
        if str(preceptor).strip()
    }
    weekly_rows = []

    # Rows created by create_ms_schedule_template:
    # week 1 dates/AM/PM = 4/6/7; then every 8 rows for weeks 2-4.
    week_layout = [
        {"date_row": 4, "session_rows": (6, 7)},
        {"date_row": 12, "session_rows": (14, 15)},
        {"date_row": 20, "session_rows": (22, 23)},
        {"date_row": 28, "session_rows": (30, 31)},
    ]

    for ws in master_wb.worksheets:
        student_name = _student_name_from_sheet(ws)

        for layout in week_layout:
            monday_date = _week_monday_from_date_row(ws, layout["date_row"])
            if monday_date is None:
                continue

            preceptor_counts = Counter()

            for session_row in layout["session_rows"]:
                for col in range(2, 9):  # Monday-Sunday, B:H
                    preceptor, site = _parse_schedule_assignment(
                        ws.cell(row=session_row, column=col).value
                    )
                    if not preceptor or site not in FOCUS_SITES:
                        continue
                    preceptor_counts[preceptor] += 1

            if not preceptor_counts:
                continue

            # Every student/week must have one primary. Prefer >=3 sessions;
            # otherwise choose the best available preceptor and flag it.
            eligible_primary = [
                (preceptor, count)
                for preceptor, count in preceptor_counts.items()
                if count >= 3
            ]
            primary_pool = eligible_primary or list(preceptor_counts.items())
            primary_pool.sort(
                key=lambda item: (-item[1], _normalize_name(item[0]))
            )
            primary_name, primary_count = primary_pool[0]
            below_threshold = primary_count < 3

            for preceptor, count in preceptor_counts.items():
                is_primary = preceptor == primary_name
                flag_reasons = []
                if is_primary and below_threshold:
                    flag_reasons.append("SELECTED PRIMARY HAS FEWER THAN 3 SESSIONS")

                weekly_rows.append(
                    {
                        "preceptor_name": preceptor,
                        "student_name": student_name,
                        "no_of_sessions": int(count),
                        "monday_date": monday_date,
                        "primary_preceptor": "YES" if is_primary else "NO",
                        "fragmented_preceptor": "YES" if count < 3 else "NO",
                        "primary_preceptor_flag": "YES" if flag_reasons else "NO",
                        "primary_preceptor_flag_reason": "; ".join(flag_reasons),
                        "email": NORMALIZED_EMAIL_MAP.get(
                            _normalize_name(preceptor), ""
                        ),
                    }
                )

    report_df = pd.DataFrame(weekly_rows, columns=REPORT_COLUMNS)

    if not report_df.empty:
        # Flag a provider who is serving as primary for multiple students in
        # the same week. Reuse is allowed; the flag simply makes it visible.
        primary_mask = report_df["primary_preceptor"].eq("YES")
        primary_rows = report_df.loc[
            primary_mask,
            ["monday_date", "preceptor_name", "student_name"],
        ].copy()
        primary_rows["_normalized_preceptor"] = primary_rows[
            "preceptor_name"
        ].map(_normalize_name)

        repeated_keys = set(
            primary_rows.groupby(
                ["monday_date", "_normalized_preceptor"], dropna=False
            )["student_name"]
            .nunique()
            .loc[lambda counts: counts > 1]
            .index.tolist()
        )

        for row_idx in report_df.index[primary_mask]:
            key = (
                report_df.at[row_idx, "monday_date"],
                _normalize_name(report_df.at[row_idx, "preceptor_name"]),
            )
            if key not in repeated_keys:
                continue

            reason = str(
                report_df.at[row_idx, "primary_preceptor_flag_reason"] or ""
            ).strip()
            repeat_reason = "PRECEPTOR IS PRIMARY FOR MULTIPLE STUDENTS THIS WEEK"
            report_df.at[row_idx, "primary_preceptor_flag"] = "YES"
            report_df.at[row_idx, "primary_preceptor_flag_reason"] = (
                f"{reason}; {repeat_reason}" if reason else repeat_reason
            )

        report_df["_primary_sort"] = report_df["primary_preceptor"].map(
            {"YES": 0, "NO": 1}
        )
        report_df = (
            report_df.sort_values(
                by=[
                    "monday_date",
                    "student_name",
                    "_primary_sort",
                    "no_of_sessions",
                    "preceptor_name",
                ],
                ascending=[True, True, True, False, True],
                kind="stable",
            )
            .drop(columns=["_primary_sort"])
            .reset_index(drop=True)
        )

    return report_df


def build_preceptor_report_workbook(report_df):
    """
    Create a Power Automate-ready .xlsx workbook using XlsxWriter.

    The report is written as a genuine Excel table without adding a second,
    overlapping worksheet AutoFilter. Avoiding that overlap prevents Excel's
    "We found a problem with some content" repair warning.
    """
    output = BytesIO()
    workbook = xlsxwriter.Workbook(
        output,
        {
            "in_memory": True,
            "strings_to_formulas": False,
            "strings_to_urls": False,
        },
    )

    worksheet = workbook.add_worksheet("Preceptor Assignments")
    worksheet.freeze_panes(1, 0)
    worksheet.set_zoom(90)
    worksheet.set_column("A:A", 28)
    worksheet.set_column("B:B", 28)
    worksheet.set_column("C:C", 16)
    worksheet.set_column("D:D", 15)
    worksheet.set_column("E:E", 20)
    worksheet.set_column("F:F", 23)
    worksheet.set_column("G:G", 23)
    worksheet.set_column("H:H", 58)
    worksheet.set_column("I:I", 38)

    header_format = workbook.add_format(
        {
            "bold": True,
            "font_color": "#FFFFFF",
            "bg_color": "#1F4E78",
            "align": "center",
            "valign": "vcenter",
            "border": 1,
            "border_color": "#D9E2F3",
        }
    )
    body_format = workbook.add_format(
        {
            "valign": "top",
            "bottom": 1,
            "bottom_color": "#D9E2F3",
        }
    )
    integer_format = workbook.add_format(
        {
            "valign": "top",
            "align": "center",
            "num_format": "0",
            "bottom": 1,
            "bottom_color": "#D9E2F3",
        }
    )
    date_format = workbook.add_format(
        {
            "valign": "top",
            "align": "center",
            "num_format": "mm/dd/yyyy",
            "bottom": 1,
            "bottom_color": "#D9E2F3",
        }
    )
    primary_formats = [
        workbook.add_format(
            {
                "valign": "top",
                "bg_color": "#E2F0D9",
                "bottom": 1,
                "bottom_color": "#D9E2F3",
                **({"align": "center", "num_format": "0"} if col == 2 else {}),
                **({"align": "center", "num_format": "mm/dd/yyyy"} if col == 3 else {}),
            }
        )
        for col in range(len(REPORT_COLUMNS))
    ]
    fragmented_formats = [
        workbook.add_format(
            {
                "valign": "top",
                "bg_color": "#FFF2CC",
                "bottom": 1,
                "bottom_color": "#D9E2F3",
                **({"align": "center", "num_format": "0"} if col == 2 else {}),
                **({"align": "center", "num_format": "mm/dd/yyyy"} if col == 3 else {}),
            }
        )
        for col in range(len(REPORT_COLUMNS))
    ]
    flagged_primary_formats = [
        workbook.add_format(
            {
                "valign": "top",
                "bg_color": "#FCE4D6",
                "font_color": "#9C0006",
                "bottom": 1,
                "bottom_color": "#D9E2F3",
                **({"align": "center", "num_format": "0"} if col == 2 else {}),
                **({"align": "center", "num_format": "mm/dd/yyyy"} if col == 3 else {}),
            }
        )
        for col in range(len(REPORT_COLUMNS))
    ]
    missing_email_format = workbook.add_format(
        {
            "valign": "top",
            "bg_color": "#FCE4D6",
            "bottom": 1,
            "bottom_color": "#D9E2F3",
        }
    )

    # Write headers explicitly. The table is added after the data is written.
    for col_idx, header in enumerate(REPORT_COLUMNS):
        worksheet.write(0, col_idx, header, header_format)

    for row_offset, row in enumerate(
        report_df.itertuples(index=False, name=None), start=1
    ):
        is_primary = str(row[4]).strip().upper() == "YES"
        is_fragmented = str(row[5]).strip().upper() == "YES"
        is_flagged_primary = str(row[6]).strip().upper() == "YES"

        for col_idx, value in enumerate(row):
            if is_flagged_primary:
                cell_format = flagged_primary_formats[col_idx]
            elif is_primary:
                cell_format = primary_formats[col_idx]
            elif is_fragmented:
                cell_format = fragmented_formats[col_idx]
            elif col_idx == 2:
                cell_format = integer_format
            elif col_idx == 3:
                cell_format = date_format
            else:
                cell_format = body_format

            if col_idx == len(REPORT_COLUMNS) - 1 and not str(value or "").strip():
                cell_format = missing_email_format

            if col_idx == 3:
                parsed_date = _parse_excel_date(value)
                if parsed_date is not None:
                    worksheet.write_datetime(
                        row_offset,
                        col_idx,
                        datetime.combine(parsed_date, datetime.min.time()),
                        cell_format,
                    )
                else:
                    worksheet.write_blank(row_offset, col_idx, None, cell_format)
            elif value is None or (isinstance(value, float) and pd.isna(value)):
                worksheet.write_blank(row_offset, col_idx, None, cell_format)
            else:
                worksheet.write(row_offset, col_idx, value, cell_format)

    # Power Automate needs a named Excel table. Only the table owns the filter;
    # do not also call worksheet.autofilter() on the same range.
    if not report_df.empty:
        worksheet.add_table(
            0,
            0,
            len(report_df),
            len(REPORT_COLUMNS) - 1,
            {
                "name": "PreceptorAssignmentTable",
                "style": "Table Style Medium 2",
                "columns": [{"header": header} for header in REPORT_COLUMNS],
            },
        )

    notes = workbook.add_worksheet("Definitions")
    notes.set_column("A:A", 24)
    notes.set_column("B:B", 90)
    notes_format = workbook.add_format({"valign": "top", "text_wrap": True})
    notes_label_format = workbook.add_format(
        {"bold": True, "valign": "top", "text_wrap": True}
    )
    definitions = [
        ("Report scope", "HOPE_DRIVE, NYES, and ETOWN only"),
        (
            "Primary preceptor",
            "Exactly one per student/week. The app prefers the highest-session "
            "preceptor with at least 3 sessions. If nobody reaches 3, the "
            "highest-session preceptor is still selected and flagged.",
        ),
        (
            "Repeated primary",
            "The same preceptor may be primary for multiple students. Those "
            "primary rows are flagged for visibility.",
        ),
        ("Fragmented preceptor", "YES when no_of_sessions < 3."),
        (
            "Primary preceptor flag",
            "YES when the selected primary has fewer than 3 sessions or the "
            "same preceptor is primary for multiple students that week.",
        ),
        (
            "Email mapping",
            "Emails come from PRECEPTOR_EMAIL_MAP in the Streamlit source code.",
        ),
    ]
    for row_idx, (label, definition) in enumerate(definitions):
        notes.write(row_idx, 0, label, notes_label_format)
        notes.write(row_idx, 1, definition, notes_format)

    missing_names = (
        sorted(
            report_df.loc[
                report_df["email"].fillna("").astype(str).str.strip().eq(""),
                "preceptor_name",
            ]
            .dropna()
            .unique()
            .tolist(),
            key=_normalize_name,
        )
        if not report_df.empty
        else []
    )

    if missing_names:
        missing_ws = workbook.add_worksheet("Missing Emails")
        missing_ws.freeze_panes(1, 0)
        missing_ws.set_column("A:A", 32)
        missing_ws.set_column("B:B", 70)
        missing_ws.write(0, 0, "preceptor_name", header_format)
        missing_ws.write(0, 1, "action_needed", header_format)
        for row_idx, name in enumerate(missing_names, start=1):
            missing_ws.write(row_idx, 0, name, body_format)
            missing_ws.write(
                row_idx,
                1,
                "Add this name and email to PRECEPTOR_EMAIL_MAP in schedule_app/settings.py",
                body_format,
            )

    workbook.close()
    output.seek(0)
    return output, missing_names
