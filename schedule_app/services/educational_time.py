"""Public teaching-time measures: four hours per preceptor AM/PM, never per learner.

The raw scan retains legacy per-student assignment fields for deduplication and
student-continuity validation. Those fields are NOT hours earned by a preceptor.
All published educational hours come from validated distinct clinical shifts.
No source OPD, learner identifier, or encrypted GitHub record is modified here.
"""
from schedule_app.services.report_diagnostics import checked_report_reach

EDUCATIONAL_TIME_REPORT_VERSION = 2
HOURS_CALCULATION_BASIS = "distinct_preceptor_date_am_pm"
TIME_DEFINITION = (
    "Total scheduled availability includes all recorded AM and PM shifts in the selected dates, "
    "with or without students, including weekends. Each shift represents four hours. "
    "Educational hours include only shifts with at least one student; two or more students "
    "in the same shift still count as four hours. Learner Reach is educational hours "
    "divided by total scheduled availability."
)
TIME_SCOPE_NOTE = (
    "These are scheduled four-hour equivalents from saved OPDs, not verified total clinical work or attendance. "
    "Repeated listings count once. Existing outpatient-over-nursery priority is applied before counting. "
    "Unlisted work is not measured; future scheduled shifts within the selected dates are included."
)
TIME_CSV_COLUMNS = (
    "preceptor_name", "academic_year", "total_scheduled_availability_hours", "educational_hours",
    "learner_reach_pct", "unique_students", "unique_students_3plus_days",
    "months_with_students", "months_scheduled", "scheduled_shifts", "teaching_shifts",
    "student_counts_start_date", "student_counts_end_date", "student_counts_status",
    "eligible_students", "minimum_shifts",
)
TIME_WORK_TYPE_CSV_COLUMNS = (
    "preceptor_name", "academic_year", "work_type", "total_scheduled_availability_hours",
    "educational_hours", "learner_reach_pct", "months_with_students", "months_scheduled",
    "scheduled_shifts", "teaching_shifts", "source_sites",
)
TIME_MONTHLY_CSV_COLUMNS = (
    "preceptor_name", "academic_year", "work_type", "month", "total_scheduled_availability_hours",
    "educational_hours", "learner_reach_pct", "scheduled_shifts", "teaching_shifts", "source_sites",
)
TIME_PREVIEW_LABELS = {
    "preceptor_name": "Preceptor", "academic_year": "Reporting period", "work_type": "Clinical experience",
    "total_scheduled_availability_hours": "Total scheduled availability (hours)",
    "educational_hours": "Educational hours", "learner_reach_pct": "Learner Reach (%)",
    "unique_students": "Unique students", "unique_students_3plus_days": "Students assigned on 3+ days",
    "student_counts_start_date": "Student-count start date", "student_counts_end_date": "Student-count cutoff",
    "student_counts_status": "Student identity basis", "eligible_students": "Students meeting minimum shifts",
    "minimum_shifts": "Minimum shifts", "months_with_students": "Months with students", "months_scheduled": "Scheduled months",
    "scheduled_shifts": "Scheduled shifts", "teaching_shifts": "Shifts with students", "source_sites": "OPD sites",
}


def with_educational_time(row, **context):
    """Recalculate from distinct shifts, ignoring legacy student-weighted hours.

    Counts and pre-existing clinical display fields must still reconcile. This
    is a deliberate metric change, not a fallback for invalid/missing counts.
    """
    result = checked_report_reach(row, **context)
    result.update({
        "total_scheduled_availability_hours": result["recorded_clinical_hours"],
        "educational_hours": result["hours_with_students"],
        "scheduled_shifts": result["recorded_clinical_shifts"],
        "teaching_shifts": result["shifts_with_students"],
    })
    return result


def teaching_time_rows(scan, selected_years, *, by_work_type=False, monthly=False):
    """Simple report/CSV rows with the same denominator and numerator everywhere.

    Monthly detail retains zero-teaching months for included preceptor/services.
    Unique-student counts are overall only, never copied to service/month rows.
    """
    from schedule_app.services.teaching_analysis import teaching_annual_rows, teaching_work_type_rows
    from schedule_app.services.learner_reach import participating_reach_rows
    from schedule_app.services.student_continuity import student_continuity_summary
    from schedule_app.services.reporting_periods import teaching_report_label
    years = tuple(sorted({int(year) for year in selected_years}))
    if monthly:
        if not by_work_type:
            raise ValueError("Monthly teaching-time rows require by_work_type=True.")
        rows = participating_reach_rows(scan, years, by_work_type=True, monthly=True)
    else:
        rows = teaching_work_type_rows(scan, years) if by_work_type else teaching_annual_rows(scan, years)
    year_by_label = {teaching_report_label(scan, year): year for year in years}
    result = []
    for row in rows:
        value = with_educational_time(row, report="Teaching time summary")
        value["months_with_students"] = row.get("months_worked", "")
        if not by_work_type:
            value.update(student_continuity_summary(scan, row["preceptor_name"], year_by_label[row["academic_year"]]))
        result.append(value)
    return result
