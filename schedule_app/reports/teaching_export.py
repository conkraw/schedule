"""Teaching CSV outputs and the complete reports ZIP.

Extracted from the supplied app; this module performs no page rendering on import.
"""

from schedule_app.services.teaching_priority import (
    OUTPATIENT_PRIORITY_NOTE, PRIORITY_AUDIT_COLUMNS, outpatient_priority_audit_rows,
    selected_priority_adjustments,
)
from schedule_app.services.teaching_validation import validate_teaching_report
from schedule_app.services.educational_time import (
    teaching_time_rows, TIME_CSV_COLUMNS, TIME_WORK_TYPE_CSV_COLUMNS, TIME_MONTHLY_CSV_COLUMNS,
    TIME_DEFINITION, TIME_SCOPE_NOTE, HOURS_CALCULATION_BASIS,
)
from schedule_app.services.report_diagnostics import report_step
from schedule_app.services.student_continuity import (
    require_student_continuity_data, STUDENT_CONTINUITY_NOTE, STUDENT_MATCHING_NOTE,
)
from collections import defaultdict
from datetime import date
from schedule_app.services.learner_reach import (
    LEARNER_REACH_COLUMNS, REACH_DEFINITION, REACH_SCOPE_NOTE,
    participating_reach_rows, teaching_participation_keys, reach_group_year, require_learner_reach_data,
    PARTICIPATION_SCOPE_NOTE, REACH_DETAIL_TOTAL_NOTE,
)
from io import BytesIO
from schedule_app.reports.chair_summary import teaching_make_chair_summary, teaching_chair_summary_data
from schedule_app.reports.learner_reach_charts import teaching_clinical_charts, CHART_DATA_COLUMNS
from schedule_app.reports.individual_teaching import teaching_make_docx
from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.reporting_periods import (
    teaching_report_label, teaching_period, reporting_period_json,
)
from schedule_app.services.teaching_analysis import teaching_annual_rows
from schedule_app.services.teaching_analysis import teaching_name_key
from schedule_app.services.teaching_analysis import teaching_work_type_rows
from schedule_app.settings import TEACHING_CHAIR_SUMMARY_FILENAME
from schedule_app.settings import TEACHING_CSV_COLUMNS
from schedule_app.settings import TEACHING_HOURS_PER_STUDENT_SHIFT
from schedule_app.settings import TEACHING_WORK_TYPE_CSV_COLUMNS
from zipfile import ZIP_DEFLATED
from zipfile import ZipFile
import csv
import io
import re


def teaching_csv_bytes(rows, columns=TIME_CSV_COLUMNS):
    stream = io.StringIO(newline="")
    writer = csv.DictWriter(stream, fieldnames=list(columns), extrasaction="ignore", lineterminator="\r\n")
    writer.writeheader()
    for row in rows:
        cleaned = {}
        for column in columns:
            value = row.get(column, "")
            # Keep name text from becoming an Excel formula when CSV is opened.
            if isinstance(value, str) and value.lstrip().startswith(("=", "+", "-", "@")):
                value = "'" + value
            cleaned[column] = value
        writer.writerow(cleaned)
    return stream.getvalue().encode("utf-8-sig")


def teaching_build_zip(scan, selected_years, *, oasis_feedback=None):
    """Return a ZIP and annual preview. Does not call GitHub or write plaintext there."""
    require_learner_reach_data(scan)
    years = sorted({int(year) for year in selected_years})
    validate_teaching_report(scan, years)
    require_student_continuity_data(scan)
    annual = teaching_annual_rows(scan, years)
    typed = teaching_work_type_rows(scan, years)
    if not annual:
        raise OPDArchiveError("No student assignments were found for the selected reporting period(s); no empty report was generated.")
    monthly_by_name = defaultdict(list)
    for item in scan["monthly"]:
        if item["academic_start_year"] in years and int(item["no_of_shifts"]) > 0:
            monthly_by_name[item["preceptor_name"]].append(item)
    eligible = teaching_participation_keys(scan, years)
    active_names = {name for name, _, _ in eligible}
    active_sites = {site for row in typed for site in row.get("source_sites", "").split("; ") if site}
    labels = ", ".join(teaching_report_label(scan, year) for year in years)
    output = BytesIO()
    notes = [
        "PRECEPTOR TEACHING SUMMARY", f"Selected academic year(s): {labels}",
        f"Archive retrieved: {scan['generated_at']}", f"Repository snapshot: {scan['commit']}",
        f"Current OPD files read: {len(scan['sources'])}", "",
        "Source: every current OPD_YYYY-MM-DD.xlsx.enc in the configured archive folder.",
        "Superseded Git history is not counted. The archived files are not changed.",
        "All recognized OPD site worksheets are scanned. Only entries with student assignments in the selected period are listed.",
        "This report describes scheduled teaching time and student continuity, not patient encounters or confirmed attendance.",
        "One row in preceptor_teaching_summary.csv = one preceptor in one academic year (unchanged).",
        "One row in preceptor_teaching_by_work_type.csv = one preceptor / academic year / work type.",
        "The chair and individual Word reports contain work-type subtotals plus overall totals.",
        "Academic Pediatrics combines HOPE_DRIVE, ETOWN and NYES. Ward A, PSHCH Nursery and Complex Care are separate.",
        "Other worksheets are kept as separate work types unless explicitly mapped in TEACHING_WORK_TYPE_MAP.",
        "Classification uses the site of each assignment, not the preceptor's usual specialty or home division.",
        "Every setting uses four hours per distinct preceptor/date/AM-or-PM shift; multiple students do not multiply time.",
        "Educational hours use the same distinct shifts as the Learner Reach numerator, overall and by clinical experience.",
        "Academic Pediatrics takes priority over PSHCH Nursery on the same preceptor/date/AM-or-PM; other unresolved work-type conflicts block reports.",
        f"{TEACHING_CHAIR_SUMMARY_FILENAME} = one combined Word summary for the chair.",
        "The chair summary lists named preceptors alphabetically and keeps unresolved provider labels separate.",
        "Academic year is July 1 through June 30 of the following calendar year.",
        "The actual session date, not the rotation's start date, determines the month and academic year.",
        "scheduled_shifts counts all recorded preceptor/date/AM-or-PM shifts; teaching_shifts counts those with at least one student.",
        f"educational_hours = teaching_shifts x {TEACHING_HOURS_PER_STUDENT_SHIFT}.",
        "Two or more students in one AM/PM session still represent one teaching shift and four educational hours.",
        "An identical preceptor/student/date/AM-or-PM duplicate counts once, even across overlapping OPDs.",
        "Provider availability with no student does not count. All filled student assignments are treated equally.",
        "Names are matched case-insensitively with normalized whitespace and comma spacing; no fuzzy identity matching.",
        "Student names and decrypted workbooks are not included in this ZIP.",
        "CHAIR AND INDIVIDUAL REPORTS: UNIQUE STUDENTS AND CONTINUITY",
        STUDENT_CONTINUITY_NOTE,
        STUDENT_MATCHING_NOTE,
        "The chair's Student continuity by preceptor table combines all work types for each person, "
        "matching their individual report. Columns are not summed across preceptors because the same "
        "student can be assigned to more than one preceptor.",
        "Unique counts use only retained assignments after outpatient/nursery priority. "
        "These counts do not change student-shifts, educational hours, Learner Reach, or clinical-experience pies.",
        "Future scheduled assignments in the selected academic year(s) are included.",
        "Providers with no assignments in the selected year(s) do not receive a report.",
        "The summary remains a snapshot until Load / refresh archived OPDs is clicked again.", "",
        "DATA QUALITY (the following diagnostics cover ALL scanned academic years)",
        f"Exact duplicate student-shifts removed: {scan['duplicate_assignments_removed']}",
        f"Future student-shifts in the entire archive when scanned: {scan['future_assignments_in_archive']}",
    ]
    notes += ["", "SIMPLIFIED HOURS (2026-09-30)", TIME_DEFINITION, TIME_SCOPE_NOTE,
              f"Calculation basis: {HOURS_CALCULATION_BASIS}.",
              "The old student-weighted educational-hours measure is no longer exported.",
              "Teaching-summary CSV headers now use total_scheduled_availability_hours, educational_hours, "
              "scheduled_shifts and teaching_shifts. Raw no_of_shifts is not exported in summary CSVs.",
              "Source audit counts describe student assignments only; they are not educational hours."]
    period = teaching_period(scan)
    if period:
        substitutions = {
            f"Selected academic year(s): {labels}": f"Reporting label (academic_year in CSVs): {period.label}",
            "One row in preceptor_teaching_summary.csv = one preceptor in one academic year (unchanged).":
                "One row in preceptor_teaching_summary.csv = one preceptor in the selected custom reporting period.",
            "One row in preceptor_teaching_by_work_type.csv = one preceptor / academic year / work type.":
                "One row in preceptor_teaching_by_work_type.csv = one preceptor / custom reporting period / work type.",
            "Academic year is July 1 through June 30 of the following calendar year.":
                "The custom reporting period uses the selected inclusive start and end dates; it is not split at July 1.",
            "The actual session date, not the rotation's start date, determines the month and academic year.":
                "The actual session date determines inclusion. Each included assignment uses the user-entered academic_year label.",
            "Future scheduled assignments in the selected academic year(s) are included.":
                "Future scheduled assignments within the selected dates are included.",
            "Providers with no assignments in the selected year(s) do not receive a report.":
                "Providers with no assignments within the selected dates do not receive a report.",
            "DATA QUALITY (the following diagnostics cover ALL scanned academic years)":
                "DATA QUALITY (archive-wide counts below cover ALL scanned dates, not just the selected period)",
        }
        notes = [substitutions.get(line, line) for line in notes]
        notes[2:2] = [f"Exact reporting dates: {period.start.isoformat()} through {period.end.isoformat()} (inclusive).",
                      "Partial months and partial rotations are filtered by actual assignment date.",
                      "Reporting_Period.json can reload these dates and the label in the app."]
        notes += ["", "Archive_Sources.csv describes full source files, including any outside the selected period.",
                  "File-level counts and missing-provider warnings are full-rotation diagnostics, not date-filtered teaching totals."]
    review_names = [name for name in scan["unresolved_preceptor_labels"] if name in active_names]
    if review_names:
        notes += ["Participating provider/site/slot labels needing review (not assigned to a guessed individual):"]
        notes += ["  " + name for name in review_names]
    for item in scan["warnings"]:
        notes.append(f"Rotation {item['rotation_start']}: {item['issue']}: {item['details']}")
    notes += ["", "WORK-TYPE GROUPING FOR ASSIGNED OPD SITES"]
    for site, work_type in scan["site_work_type_mapping"].items():
        if site in active_sites:
            notes.append(f"  {site} -> {work_type}")
    selected_labels = {teaching_report_label(scan, year) for year in years}
    conflicts = [row for row in scan.get("work_type_conflicts", []) if row["academic_year"] in selected_labels]
    if conflicts:
        notes += ["", "WORK-TYPE REVIEW", "See Work_Type_Review.csv. These assignments are included once in the review category."]
    notes = [line for line in notes if not line.startswith((
        "Provider availability with no student does not count.", "Providers with no assignments"))]
    notes += ["", "REPORT INCLUSION", PARTICIPATION_SCOPE_NOTE, REACH_DETAIL_TOTAL_NOTE,
        "A preceptor with no student assignments in the selected period receives no report or CSV entry.",
        "Work-type entries are included only when that preceptor has student assignments in that category during the selected period.",
        "", "LEARNER REACH", REACH_DEFINITION, REACH_SCOPE_NOTE,
        "total_scheduled_availability_hours = scheduled_shifts x 4, including weekends and unassigned shifts.",
        "educational_hours = teaching_shifts x 4; two simultaneous students still count as four hours.",
        "The difference between total scheduled availability and educational hours is scheduled time without a student.",
        "learner_reach_pct is a number on the 0-100 scale (80 means 80%), not the fraction 0.8.",
        "All published hours count a preceptor/date/AM-or-PM once. Unique-student and 3+ day counts are separate, not hour multipliers.",
        "No student listed: keep that clinical shift in the denominator for an included preceptor. Do not reduce their denominator to teaching days only.",
        "Wholly blank cells, explicit closed/off/nonclinical labels and cells without a '~' marker are not counted as clinical shifts.",
        "Nonempty cells lacking the marker and nonclinical labels are logged by coordinate; names or shifts are not guessed.",
        OUTPATIENT_PRIORITY_NOTE,
        "Nursery student assignments are excluded, not transferred to clinic. An empty clinic student field remains an unassigned clinical shift.",
        "Outpatient_Priority_Adjustments.csv lists the retained/excluded source cells for the selected reporting dates, without learner names.",
        "After outpatient priority, other provider/date/AM-or-PM conflicts block the entire selected-period report until corrected.",
        "Every reported Learner Reach uses a nonzero validated denominator. No unresolved clinical work-type conflict is accepted.",
        "One pie per included clinical experience compares recorded hours with students against recorded hours without students.",
        "The pies use the exact educational hours and total scheduled availability shown in the chair summary.",
        "Chart PNGs are in Learner_Reach_Charts; clinical_experience_learner_reach.csv contains their numbers.",
        "Overall percentages use total shifts with learners / total recorded shifts, never an average of provider percentages.",
        "preceptor_learner_reach_monthly.csv includes all recorded months for participating preceptor/work-type entries, including zero-teaching months. It excludes wholly nonparticipating entries.",
        f"Duplicate clinical provider listings combined across the full archive: {scan.get('duplicate_clinical_listings_removed', 0)}.",
        "Clinical shifts with no identifiable provider cannot be attributed. Generic provider/slot labels are not verified individuals.",
    ]
    clinical_review = [dict(row, academic_year=teaching_report_label(scan, reach_group_year(scan, date.fromisoformat(row["date"]))),
                            work_types="; ".join(row["work_types"]), source_sites="; ".join(row["source_sites"]))
                       for row in scan.get("clinical_shift_conflicts", [])
                       if (row["preceptor_name"], reach_group_year(scan, date.fromisoformat(row["date"])), "") in eligible]
    if clinical_review:
        notes += ["Clinical_Shift_Review.csv lists concurrent clinical work-type conflicts without student identifiers."]
    source_columns = (
        "rotation_start", "last_scheduled_date", "archive_file", "github_blob_sha", "name_order",
        "assigned_student_shifts_read", "assigned_student_shifts_counted",
        "duplicate_student_shifts_removed", "missing_provider_cells",
        "clinical_provider_listings_read", "ignored_session_cells",
        "nursery_student_assignment_listings_excluded",
    )
    priority_rows = outpatient_priority_audit_rows(scan, years)
    notes += [f"Outpatient-priority half-days in selected reporting dates: {len(selected_priority_adjustments(scan, years)):,}.",
              f"Unique student-shifts removed by outpatient priority across the full archive: {scan.get('student_shifts_removed_by_outpatient_priority', 0):,}.",
              "Archive_Sources.csv counts satisfy: assignments read = retained unique credit + retained duplicates + excluded nursery listings."]
    with report_step("Chair summary calculations"):
        summaries = teaching_chair_summary_data(scan, years)
    with report_step("Clinical experience pie charts"):
        charts = teaching_clinical_charts(scan, summaries)
    with report_step("Chair Word report"):
        chair_bytes = teaching_make_chair_summary(scan, years, charts=charts)
    if oasis_feedback is not None:
        from schedule_app.services.teaching_evaluations import feedback_for_preceptor
        # Validate every included document's period before packaging anything.
        for item in scan["monthly"]:
            if item["academic_start_year"] in years and int(item["no_of_shifts"]) > 0:
                feedback_for_preceptor(oasis_feedback, scan, item["preceptor_name"], item["academic_start_year"])
        notes += ["", "LINKED OASIS EVALUATIONS",
                  "Only explicitly username-linked teaching preceptors receive OASIS sections in their individual Word reports.",
                  "OASIS-only educators are ignored. Unlinked/unmatched preceptors retain teaching-only reports.",
                  "OPD teaching uses assignment dates; OASIS feedback uses Submit Date within the same exact boundaries.",
                  "Evaluation counts and question averages come from the selected saved summary snapshot, not recomputed here.",
                  "Questions use full source wording (the verified source dictionary supports older known-question summaries).",
                  "Qualitative comments remain verbatim and may contain identifying details."]
        for group in oasis_feedback["periods"].values():
            notes.append(f"OASIS period: {group['start_date']} through {group['end_date']}; "
                         f"source: {group['summary_filename']}; blob: {group['summary_sha']}.")
    with ZipFile(output, "w", compression=ZIP_DEFLATED) as zf:
        if priority_rows:
            zf.writestr("Outpatient_Priority_Adjustments.csv", teaching_csv_bytes(priority_rows, PRIORITY_AUDIT_COLUMNS))
        if period:
            zf.writestr("Reporting_Period.json", reporting_period_json(period))
            zf.writestr("Reporting_Period.csv", teaching_csv_bytes([{
                "academic_year": period.label, "start_date": period.start.isoformat(),
                "end_date": period.end.isoformat(), "both_dates_included": "YES",
            }], ("academic_year", "start_date", "end_date", "both_dates_included")))
        zf.writestr(TEACHING_CHAIR_SUMMARY_FILENAME, chair_bytes)
        for chart in charts:
            zf.writestr(chart["filename"], chart["png"])
        zf.writestr("clinical_experience_learner_reach.csv", teaching_csv_bytes(
            [chart["data"] for chart in charts], CHART_DATA_COLUMNS))
        zf.writestr("preceptor_teaching_summary.csv", teaching_csv_bytes(
            teaching_time_rows(scan, years), TIME_CSV_COLUMNS))
        zf.writestr("preceptor_teaching_by_work_type.csv", teaching_csv_bytes(
            teaching_time_rows(scan, years, by_work_type=True), TIME_WORK_TYPE_CSV_COLUMNS))
        monthly_reach = teaching_time_rows(scan, years, by_work_type=True, monthly=True)
        zf.writestr("preceptor_learner_reach_monthly.csv", teaching_csv_bytes(monthly_reach, TIME_MONTHLY_CSV_COLUMNS))
        if clinical_review:
            zf.writestr("Clinical_Shift_Review.csv", teaching_csv_bytes(clinical_review,
                ("preceptor_name", "academic_year", "date", "shift", "work_types", "source_sites", "has_student")))
        if conflicts:
            zf.writestr("Work_Type_Review.csv", teaching_csv_bytes(conflicts,
                ("preceptor_name", "academic_year", "date", "shift", "conflicting_work_types", "source_sites", "no_of_student_shifts")))
        used = set()
        for name in sorted(monthly_by_name, key=teaching_name_key):
            base = re.sub(r"[^A-Za-z0-9._-]+", "_", name).strip("._")[:110] or "preceptor"
            safe = base
            number = 1
            while safe.casefold() in used:
                number += 1
                safe = f"{base}_{number}"
            used.add(safe.casefold())
            with report_step("Individual preceptor Word report", preceptor_name=name, academic_year=labels):
                if oasis_feedback is None:
                    individual_bytes = teaching_make_docx(name, monthly_by_name[name], scan)
                else:
                    individual_bytes = teaching_make_docx(name, monthly_by_name[name], scan, oasis_feedback=oasis_feedback)
            zf.writestr(f"Preceptor_Reports/{safe}_Teaching_Report.docx", individual_bytes)
        if scan.get("assessment_completion") is not None:
            from schedule_app.services.assessment_completion import COLUMNS, METHOD_NOTE, SCOPE_NOTE, completion_rows
            bundle = scan["assessment_completion"]
            assessment_rows = [row for year in years for row in completion_rows(bundle, scan, year)]
            zf.writestr("preceptor_student_assessment_completion.csv", teaching_csv_bytes(assessment_rows, COLUMNS))
            if bundle["warnings"]:
                zf.writestr("Evaluation_Completeness_Alerts.csv", teaching_csv_bytes(bundle["warnings"],
                    ("preceptor_name", "academic_year", "username", "direction", "issue", "action")))
            notes += ["", "STUDENT ASSESSMENT COMPLETION", METHOD_NOTE, SCOPE_NOTE,
                      "The 3+ shifts denominator differs from the existing 3+ distinct days continuity measure.",
                      "No student names, external IDs, grades or assessment comments are exported in these completion tables.",
                      "An unverified/unchecked value is blank in CSV; it is not zero."]
        zf.writestr("Report_Notes.txt", "\n".join(notes).encode("utf-8"))
        zf.writestr("Archive_Sources.csv", teaching_csv_bytes(scan["sources"], source_columns))
    return output.getvalue(), annual
