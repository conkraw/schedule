"""Unique students and distinct teaching days for the individual Word report.

One in-memory record per preceptor/student contains dates only. The scan uses
run-local keyed identifiers to assemble these groups, then discards the keys,
identifiers and student names. Date groups are never written to GitHub or reports.
Separate students with identical dates MUST remain separate list entries.
"""
from collections import Counter, defaultdict
from datetime import date
from typing import Any, Iterable, Mapping

from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.reporting_periods import ReportingPeriod, teaching_report_bounds

STUDENT_CONTINUITY_SCHEMA_VERSION = 1
STUDENT_CONTINUITY_NOTE = (
    "Each student is counted once within this reporting period, across all work types. "
    "Three or more days means three distinct calendar dates, not three shifts; "
    "AM and PM on the same date count as one day. Days do not need to be consecutive. "
    "Only assignments within the displayed dates are included."
)
STUDENT_MATCHING_NOTE = (
    "Unique-student counts use the student names recorded in the OPDs, ignoring "
    "capitalization, extra spaces and comma spacing. Different spellings are not "
    "automatically merged; different people with the same recorded name cannot "
    "be distinguished. Student names and individual date groups are not exported."
)


def build_student_continuity_data(
    assignments: Iterable[Mapping[str, Any]], provider_names: Mapping[str, str]
) -> dict[str, Any]:
    """Use only deduplicated assignments retained AFTER outpatient priority."""
    days_by_student = defaultdict(set)
    for item in assignments:
        days_by_student[(item["provider_key"], item["student_key"])].add(item["day"].isoformat())
    # Intentionally use a list, not a set: two students may have identical dates.
    groups = [
        {"preceptor_name": provider_names[key], "dates": sorted(days)}
        for (key, _), days in days_by_student.items()
    ]
    groups.sort(key=lambda row: (row["preceptor_name"].casefold(), tuple(row["dates"])))
    return {"student_continuity_version": STUDENT_CONTINUITY_SCHEMA_VERSION,
            "student_day_groups": groups}


def require_student_continuity_data(scan: Mapping[str, Any]) -> None:
    """Refuse old/inconsistent scans rather than invent zero or estimate counts."""
    if (scan.get("student_continuity_version") != STUDENT_CONTINUITY_SCHEMA_VERSION
            or not isinstance(scan.get("student_day_groups"), list)):
        raise OPDArchiveError(
            "This teaching scan predates unique-student reporting. Click Load / "
            "refresh archived OPDs once, then generate the individual reports again. "
            "Unique students cannot be recovered from old shift totals alone."
        )
    try:
        unique_days, shifts = Counter(), Counter()
        period = scan.get("reporting_period")
        for row in scan["student_day_groups"]:
            if set(row) != {"preceptor_name", "dates"}:
                raise ValueError("unexpected fields")
            name, days = row["preceptor_name"], row["dates"]
            if (not isinstance(name, str) or not name.strip() or not isinstance(days, list)
                    or not days or any(not isinstance(day, str) for day in days)
                    or days != sorted(set(days))):
                raise ValueError("invalid student date group")
            for day in days:
                parsed = date.fromisoformat(day)
                if parsed.isoformat() != day or not 1970 <= parsed.year <= 2100:
                    raise ValueError("invalid date")
                if period and not period["start_date"] <= day <= period["end_date"]:
                    raise ValueError("date outside report period")
                unique_days[(name, day)] += 1
        for row in scan["daily_by_work_type"]:
            count = row["no_of_shifts"]
            if type(count) is not int or count <= 0:
                raise ValueError("invalid daily shifts")
            shifts[(row["preceptor_name"], row["date"])] += count
        # Every assigned date must be represented; unique students on a date
        # cannot outnumber that date's retained student-shifts.
        if unique_days.keys() != shifts.keys() or any(
            total > shifts[key] for key, total in unique_days.items()
        ):
            raise ValueError("student dates do not reconcile")
    except (KeyError, TypeError, ValueError, AttributeError):
        raise OPDArchiveError(
            "Unique-student date details are incomplete or inconsistent with the "
            "teaching shifts. Click Load / refresh archived OPDs; no individual "
            "report was generated from these details."
        ) from None


def filter_student_continuity_dates(
    source: Mapping[str, Any], target: dict[str, Any], period: ReportingPeriod
) -> None:
    """Keep one group per learner with dates inside the inclusive custom period."""
    require_student_continuity_data(source)
    start, end = period.start.isoformat(), period.end.isoformat()
    groups = []
    for row in source["student_day_groups"]:
        days = [day for day in row["dates"] if start <= day <= end]
        if days:
            groups.append({"preceptor_name": row["preceptor_name"], "dates": days})
    target["student_continuity_version"] = STUDENT_CONTINUITY_SCHEMA_VERSION
    target["student_day_groups"] = groups
    require_student_continuity_data(target)


def student_continuity_counts(
    scan: Mapping[str, Any], preceptor_name: str, group_year: int
) -> dict[str, int]:
    """Two overall counts for ONE report section, never a sum of monthly uniques.

    A custom date range is a single group even across July/month/rotation
    boundaries. Standard academic-year sections each use their own exact bounds.
    The caller applies existing conflict validation before building reports.
    """
    require_student_continuity_data(scan)
    start, end = (day.isoformat() for day in teaching_report_bounds(scan, group_year))
    total = three_plus = 0
    for row in scan["student_day_groups"]:
        if row["preceptor_name"] != preceptor_name:
            continue
        count = sum(start <= day <= end for day in row["dates"])
        total += int(count > 0)
        three_plus += int(count >= 3)
    return {"unique_students": total, "unique_students_3plus_days": three_plus}
