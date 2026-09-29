"""Explicit local reporting rule: clinic teaching takes precedence over nursery.

This is not a general conflict-resolution heuristic. Only HOPE_DRIVE, ETOWN and
NYES displace PSHCH_NURSERY for the SAME preceptor/date/AM-or-PM. Other services
keep the existing strict validation. Nothing is written to the original OPDs.
"""
from datetime import date
import re
from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.reporting_periods import teaching_period

OUTPATIENT_PRIORITY_VERSION = 1
ACADEMIC_PEDIATRICS_SITES = frozenset({"HOPE_DRIVE", "ETOWN", "NYES"})
PSHCH_NURSERY_SITE = "PSHCH_NURSERY"
OUTPATIENT_PRIORITY_NOTE = (
    "For the same preceptor, date and AM/PM shift, Academic Pediatrics takes priority "
    "over PSHCH Nursery. The overlapping nursery listing and its student assignments "
    "are excluded from clinical and educational hours. Only students actually listed "
    "in clinic receive outpatient teaching credit."
)
PRIORITY_AUDIT_COLUMNS = (
    "adjustment_id", "preceptor_name", "date", "shift", "rotation_start",
    "archive_file", "archive_path", "worksheet", "cell", "listed_work_type",
    "student_assigned", "report_action", "github_blob_sha",
)


def priority_site_key(value):
    return re.sub(r"[^A-Z0-9]+", "_", str(value or "").upper()).strip("_")


def nursery_sites_to_exclude(sites):
    normalized = {priority_site_key(site) for site in sites}
    if PSHCH_NURSERY_SITE in normalized and normalized & ACADEMIC_PEDIATRICS_SITES:
        return {PSHCH_NURSERY_SITE}
    return set()


def require_outpatient_priority_data(scan):
    if scan.get("outpatient_priority_version") != OUTPATIENT_PRIORITY_VERSION:
        raise OPDArchiveError(
            "This scan predates the outpatient-over-nursery rule. Click Load / refresh "
            "archived OPDs once, then regenerate reports. No new report was generated.")


def selected_priority_adjustments(scan, selected_years, *, preceptor_name=None):
    """Read-only, date-filtered adjustment log; contains no student identifiers."""
    years = {int(year) for year in selected_years}
    period = teaching_period(scan)
    rows = []
    for item in scan.get("outpatient_priority_adjustments", []):
        day = date.fromisoformat(item["date"])
        if period:
            included = period.start.year in years and period.start <= day <= period.end
        else:
            included = (day.year if day.month >= 7 else day.year - 1) in years
        if included and (preceptor_name is None or item["preceptor_name"] == preceptor_name):
            rows.append(item)
    return sorted(rows, key=lambda item: (item["date"], item["shift"], item["preceptor_name"].casefold()))


def outpatient_priority_audit_rows(scan, selected_years):
    """One row per relevant original source cell; the same ID means one half-day."""
    rows = []
    for number, item in enumerate(selected_priority_adjustments(scan, selected_years), start=1):
        for source in sorted(item["sources"], key=lambda row: (
                row.get("rotation_start", ""), row["worksheet"], row["cell"])):
            site = priority_site_key(source["worksheet"])
            if site != PSHCH_NURSERY_SITE and site not in ACADEMIC_PEDIATRICS_SITES:
                continue
            rows.append({
                "adjustment_id": f"P{number:04d}", "preceptor_name": item["preceptor_name"],
                "date": item["date"], "shift": item["shift"],
                **{key: source.get(key, "") for key in (
                    "rotation_start", "archive_file", "archive_path", "worksheet", "cell", "github_blob_sha")},
                "listed_work_type": source["work_type"],
                "student_assigned": "YES" if source["has_student"] else "NO",
                "report_action": ("EXCLUDED: overlapping nursery coverage" if site == PSHCH_NURSERY_SITE
                                  else "RETAINED: Academic Pediatrics"),
            })
    return rows


def outpatient_priority_report_note(scan, selected_years, *, preceptor_name=None):
    rows = selected_priority_adjustments(scan, selected_years, preceptor_name=preceptor_name)
    if not rows:
        return ""
    return (f"Outpatient priority applied to {len(rows):,} overlapping half-day(s). "
            "PSHCH Nursery hours and nursery student assignments are excluded for those half-days; "
            "Academic Pediatrics uses only its own recorded student assignments. Other nursery shifts remain included.")
