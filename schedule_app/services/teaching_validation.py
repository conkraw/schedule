"""Block teaching reports with unresolved clinical work-type conflicts.

The archive can still be scanned for diagnostics. Validation is applied to the
selected reporting dates before any report/percentage/chart is produced. Only
source coordinates and provider metadata are retained; no student names.
"""
from datetime import date
from schedule_app.services.teaching_priority import require_outpatient_priority_data
from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.reporting_periods import teaching_period

STRICT_CONFLICT_SCHEMA_VERSION = 1
STRICT_REPORT_VERSION = 1
CONFLICT_CSV_COLUMNS = (
    "conflict_id", "issue", "preceptor_name", "date", "shift",
    "conflicting_work_types", "rotation_start", "archive_file", "archive_path",
    "worksheet", "cell", "listed_work_type", "student_assigned", "github_blob_sha",
)


class TeachingConflictError(OPDArchiveError):
    """Structured, safe diagnostics that the UI can display without a traceback."""
    def __init__(self, rows):
        self.rows = rows
        self.conflict_count = len({row["conflict_id"] for row in rows})
        super().__init__(
            f"Reports blocked: {self.conflict_count} conflicting clinical half-day(s) "
            "were found in the selected reporting period. Correct the source OPDs "
            "listed below, re-upload them, then click Load / refresh archived OPDs. "
            "No teaching report, chart, or report ZIP was generated."
        )


def require_conflict_source_data(scan):
    require_outpatient_priority_data(scan)
    if scan.get("strict_conflict_source_version") != STRICT_CONFLICT_SCHEMA_VERSION:
        raise OPDArchiveError(
            "This saved scan predates source-level conflict checking. Click Load / "
            "refresh archived OPDs once so any conflict can identify the exact OPD, "
            "worksheet and cell. No new report was generated."
        )


def _in_selected_period(scan, day, selected_years):
    period = teaching_period(scan)
    if period is not None:
        return period.start.year in selected_years and period.start <= day <= period.end
    return (day.year if day.month >= 7 else day.year - 1) in selected_years


def teaching_conflict_rows(scan, selected_years):
    """One diagnostic row per source cell, grouped by a stable conflict ID.

    Check every provider in the selected dates, including providers/categories
    omitted from positive-teaching tables. A blank student field does not resolve
    a provider being listed in two different work types at once.
    """
    require_conflict_source_data(scan)
    years = {int(year) for year in selected_years}
    rows, seen = [], set()
    for item in sorted(scan.get("clinical_shift_conflicts", []),
                       key=lambda row: (row["date"], row["shift"], row["preceptor_name"].casefold())):
        day = date.fromisoformat(item["date"])
        if not _in_selected_period(scan, day, years):
            continue
        identity = (item["preceptor_name"], item["date"], item["shift"])
        if identity in seen:
            continue
        seen.add(identity)
        types = "; ".join(sorted(set(item["work_types"])))
        sources = item.get("sources", [])
        if not sources or any(not row.get("worksheet") or not row.get("cell")
                              or not row.get("archive_file") for row in sources):
            raise OPDArchiveError(
                "Conflict details are missing their source cells. Click Load / "
                "refresh archived OPDs again; do not use a previous report."
            )
        conflict_id = f"C{len(seen):04d}"
        for source in sorted(sources, key=lambda row: (row.get("rotation_start", ""),
                              row.get("archive_file", ""), row["worksheet"], row["cell"])):
            rows.append({
                "conflict_id": conflict_id,
                "issue": "Same preceptor/date/AM-or-PM listed in different clinical experiences",
                "preceptor_name": item["preceptor_name"], "date": item["date"], "shift": item["shift"],
                "conflicting_work_types": types,
                **{field: source.get(field, "") for field in (
                    "rotation_start", "archive_file", "archive_path", "worksheet", "cell", "github_blob_sha")},
                "listed_work_type": source.get("work_type", ""),
                "student_assigned": "YES" if source.get("has_student") else "NO",
            })
    # Defensive check: a cross-setting student assignment must also have the
    # corresponding provider-level conflict. Do not make up missing diagnostics.
    for item in scan.get("work_type_conflicts", []):
        if _in_selected_period(scan, date.fromisoformat(item["date"]), years):
            if (item["preceptor_name"], item["date"], item["shift"]) not in seen:
                raise OPDArchiveError("A teaching work-type conflict has no matching clinical-source details. "
                                      "Refresh the archived OPDs; reports are blocked.")
    return rows


def validate_teaching_report(scan, selected_years):
    """Raise before exporting rather than publishing any withheld/N/A percentages."""
    rows = teaching_conflict_rows(scan, selected_years)
    if rows:
        raise TeachingConflictError(rows)
