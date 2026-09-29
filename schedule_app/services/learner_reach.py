"""Learner Reach: distinct recorded OPD shifts with learners / all recorded shifts.

This is a schedule-based coverage measure, not a teaching-quality score, a
student-capacity calculation, or proof of actual clinical hours/attendance.
The existing student-weighted educational-hour metric remains separate.
"""
from collections import Counter, defaultdict
from datetime import date
from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.reporting_periods import teaching_period, teaching_report_label
from schedule_app.settings import TEACHING_HOURS_PER_STUDENT_SHIFT, TEACHING_WORK_TYPE_REVIEW
import re

LEARNER_REACH_SCHEMA_VERSION = 1
# A display/export change only: existing complete scans can be reused safely.
PARTICIPATION_REPORT_VERSION = 1
PARTICIPATION_SCOPE_NOTE = (
    "Only preceptors and work types with student assignments during the selected period are listed. "
    "Overall Learner Reach uses all recorded shifts for these preceptors, including shifts without students; "
    "category percentages use all shifts in that category."
)
REACH_DETAIL_TOTAL_NOTE = (
    "Overall OPD hours also include shifts in unlisted settings or awaiting work-type review. "
    "Detail tables omit settings with no student assignments, so their OPD-hour subtotals "
    "may be lower than the overall total."
)
LEARNER_REACH_COLUMNS = (
    "recorded_clinical_shifts", "recorded_clinical_hours", "shifts_with_students",
    "hours_with_students", "shifts_without_students", "hours_without_students",
    "learner_reach_pct", "months_scheduled", "availability_review_shifts", "learner_reach_note",
)
REACH_COUNT_FIELDS = ("recorded_clinical_shifts", "shifts_with_students", "shifts_without_students")
REACH_DEFINITION = (
    "Learner Reach = recorded OPD shifts with at least one student / all recorded OPD shifts x 100. "
    "A preceptor/date/AM-or-PM counts once, even with multiple students or repeated rows. "
    "Both filled and blank student fields are included in the denominator."
)
REACH_SCOPE_NOTE = (
    "OPD hours are four-hour equivalents of recorded sessions, not verified total clinical work. "
    "Blank student fields mean no student is recorded, not that the preceptor declined teaching or could accept another student. "
    "Unlisted clinical work, other learners, missing rotations and actual attendance are not measured. "
    "Repeated template provider listings can overstate availability; verify the source schedules before interpreting percentages."
)


def nonclinical_provider(value):
    """Reject explicit nonclinical/closed labels, not similarly named people."""
    text = re.sub(r"\s+", " ", str(value or "").strip()).upper()
    if text in {"", "NAN", "NONE", "NULL", "NA", "N/A", "-", "--", "~"}:
        return True
    return bool(re.match(
        r"^(?:CLOSED|CANCELLED|CANCELED|OFF|PTO|VACATION|HOLIDAY|SICK|ADMIN|ADMINISTRATIVE|"
        r"MEETING|BLOCKED|NO CLINIC|CLINIC CANCELLED|CLINIC CANCELED)(?:\b|$)", text))


class ClinicalShiftAccumulator:
    """Temporary, provider-level distinct shift registry; contains no students."""
    def __init__(self):
        self.shifts = {}
        self.duplicate_listings = 0

    def add(self, provider_key, row, work_type, site):
        identity = (provider_key, row["day"], row["shift"])
        if identity in self.shifts:
            self.duplicate_listings += 1
        item = self.shifts.setdefault(identity, {
            "work_types": set(), "sites": set(), "has_student": False,
        })
        item["work_types"].add(work_type)
        item["sites"].add(site)
        item["has_student"] |= bool(row["has_student"])

    def finish(self, names):
        overall, typed = {}, {}
        conflicts = []
        for (key, day, shift), item in sorted(self.shifts.items()):
            name = names[key]
            kind = next(iter(item["work_types"])) if len(item["work_types"]) == 1 else TEACHING_WORK_TYPE_REVIEW
            for bucket, index in ((overall, (name, day)), (typed, (name, day, kind))):
                entry = bucket.setdefault(index, {
                    "preceptor_name": name, "date": day.isoformat(),
                    "recorded_clinical_shifts": 0, "shifts_with_students": 0,
                    "shifts_without_students": 0, "source_sites": set(),
                    **({"work_type": kind} if bucket is typed else {}),
                })
                entry["recorded_clinical_shifts"] += 1
                entry["shifts_with_students"] += int(item["has_student"])
                entry["shifts_without_students"] += int(not item["has_student"])
                entry["source_sites"].update(item["sites"])
            if len(item["work_types"]) > 1:
                conflicts.append({"preceptor_name": name, "date": day.isoformat(), "shift": shift,
                                  "work_types": sorted(item["work_types"]), "source_sites": sorted(item["sites"]),
                                  "has_student": bool(item["has_student"])})
        def finish(rows):
            return [dict(row, source_sites=sorted(row["source_sites"])) for row in rows.values()]
        return {"learner_reach_version": LEARNER_REACH_SCHEMA_VERSION,
                "clinical_daily": finish(overall), "clinical_daily_by_work_type": finish(typed),
                "clinical_shift_conflicts": conflicts,
                "duplicate_clinical_listings_removed": self.duplicate_listings}


def require_learner_reach_data(scan):
    if scan.get("learner_reach_version") != LEARNER_REACH_SCHEMA_VERSION:
        raise OPDArchiveError("This saved scan has no clinical availability data. Click Load / refresh archived OPDs once to calculate Learner Reach.")
    expected, actual = Counter(), Counter()
    try:
        for field, counter in (("clinical_daily", expected), ("clinical_daily_by_work_type", actual)):
            seen = set()
            for row in scan[field]:
                day = date.fromisoformat(row["date"])
                identity = (row["preceptor_name"], day, row.get("work_type", ""))
                if not row["preceptor_name"] or identity in seen:
                    raise ValueError("invalid or duplicate clinical day")
                seen.add(identity)
                values = [row[key] for key in REACH_COUNT_FIELDS]
                if any(type(n) is not int or n < 0 for n in values) or values[0] <= 0 or values[0] > 2 or values[1] + values[2] != values[0]:
                    raise ValueError("invalid clinical shift count")
                if not isinstance(row.get("source_sites"), list) or not row["source_sites"]:
                    raise ValueError("missing clinical source site")
                if field.endswith("by_work_type") and not row.get("work_type"):
                    raise ValueError("missing clinical work type")
                for key in REACH_COUNT_FIELDS:
                    counter[(row["preceptor_name"], day, key)] += row[key]
        if expected != actual:
            raise ValueError("clinical subtotals disagree")
        teaching = Counter()
        for row in scan.get("daily_by_work_type", []):
            teaching[(row["preceptor_name"], date.fromisoformat(row["date"]))] += row["no_of_shifts"]
        clinical = {(row["preceptor_name"], date.fromisoformat(row["date"])): row for row in scan["clinical_daily"]}
        for key in set(teaching) | set(clinical):
            with_students = clinical.get(key, {}).get("shifts_with_students", 0)
            if bool(teaching[key]) != bool(with_students) or with_students > teaching[key]:
                raise ValueError("clinical and educational totals disagree")
    except (KeyError, TypeError, ValueError):
        raise OPDArchiveError("Clinical availability data did not reconcile. Refresh the archived OPDs; no partial percentage was generated.") from None


def reach_group_year(scan, day):
    period = teaching_period(scan)
    return period.start.year if period else day.year if day.month >= 7 else day.year - 1


def reach_totals(entries):
    """Use a ratio of summed shifts, NEVER an average of percentages."""
    rows = list(entries)
    counts = {key: sum(int(row.get(key, 0)) for row in rows) for key in REACH_COUNT_FIELDS}
    total, with_students = counts["recorded_clinical_shifts"], counts["shifts_with_students"]
    if counts["shifts_without_students"] + with_students != total or with_students > total:
        raise OPDArchiveError("Clinical shift totals do not reconcile; a percentage cannot be reported.")
    review = sum(int(row.get("availability_review_shifts", 0)) for row in rows)
    note = "Clinical work-type overlap: review the affected shifts before interpreting this percentage." if review else ""
    return {**counts,
            "recorded_clinical_hours": total * TEACHING_HOURS_PER_STUDENT_SHIFT,
            "hours_with_students": with_students * TEACHING_HOURS_PER_STUDENT_SHIFT,
            "hours_without_students": counts["shifts_without_students"] * TEACHING_HOURS_PER_STUDENT_SHIFT,
            "learner_reach_pct": round(100 * with_students / total, 1) if total and not review else None,
            "availability_review_shifts": review,
            "learner_reach_note": note or ("No recorded clinical shifts in this category." if not total else "")}


def reach_percent(value):
    return "N/A" if value is None else f"{value:.1f}%"


def learner_reach_rows(scan, selected_years, *, by_work_type=False, monthly=False):
    """Joinable rows for all recorded providers, including zero teaching assignments."""
    require_learner_reach_data(scan)
    from schedule_app.services.teaching_analysis import teaching_name_key, teaching_work_type_sort, teaching_month_label
    years = {int(year) for year in selected_years}
    grouped = defaultdict(list)
    for row in scan["clinical_daily_by_work_type" if by_work_type else "clinical_daily"]:
        day = date.fromisoformat(row["date"])
        year = reach_group_year(scan, day)
        if year in years:
            key = (row["preceptor_name"], year, row.get("work_type", "") if by_work_type else "",
                   day.replace(day=1).isoformat() if monthly else "")
            grouped[key].append(row)
    reviews = Counter()
    review_sites = defaultdict(set)
    if by_work_type:
        for row in scan.get("clinical_shift_conflicts", []):
            day = date.fromisoformat(row["date"])
            year = reach_group_year(scan, day)
            if year not in years:
                continue
            # Known dates, but no defensible allocation of the single clinical
            # half-day to multiple concurrent settings. Never count it twice.
            for kind in set(row["work_types"]) | {TEACHING_WORK_TYPE_REVIEW}:
                key = (row["preceptor_name"], year, kind, day.replace(day=1).isoformat() if monthly else "")
                grouped.setdefault(key, [])
                reviews[key] += 1
                review_sites[key].update(site for site in row["source_sites"]
                                         if kind == TEACHING_WORK_TYPE_REVIEW
                                         or scan.get("site_work_type_mapping", {}).get(site) == kind)
    results = []
    for (name, year, kind, month), items in sorted(grouped.items(), key=lambda p: (teaching_name_key(p[0][0]), p[0][1], teaching_work_type_sort(p[0][2]), p[0][3])):
        key = (name, year, kind, month)
        totals = reach_totals(items)
        if reviews[key]:
            totals["availability_review_shifts"] = reviews[key]
            totals["learner_reach_pct"] = None
            totals["learner_reach_note"] = "Concurrent work types: affected clinical shifts are counted once under Work type needs review. Listed category hours exclude those unresolved shifts; category percentage withheld."
        months = sorted({date.fromisoformat(row["date"]).replace(day=1) for row in items})
        sites = {site for row in items for site in row["source_sites"]} | review_sites[key]
        results.append({"preceptor_name": name, "academic_start_year": year,
                        "academic_year": teaching_report_label(scan, year), **totals,
                        "months_scheduled": "; ".join(teaching_month_label(day) for day in months),
                        "source_sites": "; ".join(sorted(sites)),
                        **({"work_type": kind} if by_work_type else {}),
                        **({"month": month} if monthly else {})})
    return results


def teaching_participation_keys(scan, selected_years, *, by_work_type=False):
    """Positive teaching eligibility per period, never per individual month.

    The caller supplies an exact-date projection for custom dates. These keys
    control visibility only: do not use them to delete blank clinical shifts
    from the raw scan or from an included preceptor's overall denominator.
    """
    years = {int(year) for year in selected_years}
    source = "monthly_by_work_type" if by_work_type else "monthly"
    return {
        (row["preceptor_name"], int(row["academic_start_year"]),
         row["work_type"] if by_work_type else "")
        for row in scan[source]
        if int(row["academic_start_year"]) in years and int(row["no_of_shifts"]) > 0
    }


def participating_reach_rows(scan, selected_years, *, by_work_type=False, monthly=False):
    """Report rows for teaching contributors, retaining their zero-teaching months.

    Whole-period eligibility is used even for the monthly CSV. A month with no
    students must not be removed from an included provider/category denominator.
    """
    years = tuple(int(year) for year in selected_years)
    eligible = teaching_participation_keys(scan, years, by_work_type=by_work_type)
    return [row for row in learner_reach_rows(scan, years, by_work_type=by_work_type, monthly=monthly)
            if (row["preceptor_name"], row["academic_start_year"],
                row["work_type"] if by_work_type else "") in eligible]


def enrich_teaching_rows(scan, selected_years, teaching_rows, *, by_work_type=False):
    """Enrich positive teaching rows without creating availability-only entries.

    The overall join still uses ALL recorded clinical shifts for that preceptor
    in the period, including blank student fields and hidden work types.
    """
    from schedule_app.services.teaching_analysis import teaching_name_key, teaching_work_type_sort
    def key(row):
        return (row["preceptor_name"], row["academic_year"], row.get("work_type", "") if by_work_type else "")
    result = {key(row): dict(row) for row in teaching_rows if int(row["no_of_shifts"]) > 0}
    for reach in learner_reach_rows(scan, selected_years, by_work_type=by_work_type):
        item = result.get(key(reach))
        if item is None:
            continue  # Hide zero-teaching people/categories, not their source data.
        item.update({field: reach.get(field, "") for field in LEARNER_REACH_COLUMNS})
        if by_work_type:
            item["source_sites"] = "; ".join(sorted(set(filter(None, item.get("source_sites", "").split("; "))) | set(filter(None, reach["source_sites"].split("; ")))))
    for row in result.values():
        if "recorded_clinical_shifts" not in row:
            row.update(reach_totals([]), months_scheduled="")
        # A teaching assignment without a corresponding clinical session must
        # not be presented as an ordinary zero or a reliable percentage.
        if row["no_of_shifts"] and not row["shifts_with_students"] and not row.get("availability_review_shifts"):
            raise OPDArchiveError("A teaching assignment has no matching clinical shift. Refresh the archive before exporting.")
    return sorted(result.values(), key=lambda row: (teaching_name_key(row["preceptor_name"]), row["academic_year"], teaching_work_type_sort(row.get("work_type", ""))))


def filter_reach_dates(source_scan, target_scan, period):
    """Project denominators with the same exact inclusive dates as the numerator."""
    require_learner_reach_data(source_scan)
    within = lambda row: period.start.isoformat() <= row["date"] <= period.end.isoformat()
    for field in ("clinical_daily", "clinical_daily_by_work_type", "clinical_shift_conflicts"):
        target_scan[field] = [dict(row, source_sites=list(row["source_sites"])) for row in source_scan[field] if within(row)]
    names = {row["preceptor_name"] for row in target_scan["clinical_daily"]}
    target_scan["unresolved_preceptor_labels"] = [name for name in source_scan.get("unresolved_preceptor_labels", []) if name in names]
    target_scan["name_variants"] = {name: vals for name, vals in source_scan.get("name_variants", {}).items() if name in names}
    sites = {site for row in target_scan["clinical_daily"] for site in row["source_sites"]}
    target_scan["site_work_type_mapping"] = {site: label for site, label in source_scan["site_work_type_mapping"].items() if site in sites}
    require_learner_reach_data(target_scan)
