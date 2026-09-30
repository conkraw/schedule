"""Safe, actionable diagnostics for teaching-report generation.

Only approved report labels and aggregate counts enter these diagnostics. Never
attach the scan, source workbook, credentials or individual student details.
No invalid percentage is changed to zero or allowed into a published report.
"""
from contextlib import contextmanager
import math
from numbers import Real
from typing import Mapping

from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.settings import TEACHING_HOURS_PER_STUDENT_SHIFT

REPORT_OUTPUT_VERSION = 3
REPORT_BUILD_ID = "2026-09-30-simple-educational-hours-1"
REPORT_ISSUE_COLUMNS = (
    "report", "section", "preceptor_name", "work_type", "academic_year",
    "recorded_clinical_shifts", "shifts_with_students", "shifts_without_students",
    "learner_reach_pct", "issue", "action",
)
COUNT_FIELDS = ("recorded_clinical_shifts", "shifts_with_students", "shifts_without_students")
DISPLAY_FIELDS = ("recorded_clinical_hours", "hours_with_students", "hours_without_students", "learner_reach_pct")
METRIC_FIELDS = COUNT_FIELDS + DISPLAY_FIELDS + ("availability_review_shifts",)
LABEL_FIELDS = ("report", "section", "preceptor_name", "work_type", "academic_year")


def _safe_value(value):
    if value is None:
        return "missing"
    if isinstance(value, str):
        return value[:400]
    if isinstance(value, Real):
        return str(value) if not math.isfinite(value) else value
    return "<invalid " + type(value).__name__ + ">"


class ReportDataError(OPDArchiveError):
    """A single report problem with downloadable, student-free details."""
    def __init__(self, issue, *, metrics=None, action=None, **context):
        row = {field: "" for field in REPORT_ISSUE_COLUMNS}
        if isinstance(metrics, Mapping):
            for field in (*LABEL_FIELDS, *COUNT_FIELDS, "learner_reach_pct"):
                if field in row and field in metrics:
                    row[field] = _safe_value(metrics[field])
        for field in LABEL_FIELDS:
            if context.get(field) is not None:
                row[field] = _safe_value(context[field])
        row["issue"] = str(issue)
        row["action"] = action or (
            "Restart the app after replacing all update files, then refresh archived OPDs and retry. "
            "If this persists, share Learner_Reach_Report_Issues.csv. "
            "This report check alone does not establish that an OPD needs editing."
        )
        self.rows = [row]
        labels = [str(row[key]) for key in LABEL_FIELDS if row[key] != ""]
        location = " (" + " / ".join(dict.fromkeys(labels)) + ")" if labels else ""
        super().__init__(f"Report generation stopped{location}: {issue} No report ZIP was retained.")


def validated_shift_counts(metrics, **context):
    """Require actual counts; missing summary fields must not become fake zeros."""
    if not isinstance(metrics, Mapping):
        raise ReportDataError("Clinical metrics are not a valid report row.", **context)
    for field in (*COUNT_FIELDS, "availability_review_shifts"):
        if field == "availability_review_shifts" and field not in metrics:
            continue
        value = metrics.get(field)
        if (isinstance(value, bool) or not isinstance(value, Real)
                or not math.isfinite(value) or value < 0 or int(value) != value):
            issue = (f"Required clinical count '{field}' is missing."
                     if field not in metrics or value is None
                     else f"Clinical count '{field}' must be a nonnegative whole number.")
            raise ReportDataError(issue, metrics=metrics, **context)
    counts = {key: int(metrics[key]) for key in COUNT_FIELDS}
    if counts["shifts_with_students"] + counts["shifts_without_students"] != counts["recorded_clinical_shifts"]:
        raise ReportDataError(
            "Clinical shift totals do not reconcile: with-student plus without-student shifts "
            "must equal all recorded clinical shifts.", metrics=metrics, **context)
    if int(metrics.get("availability_review_shifts", 0)):
        raise ReportDataError(
            "An unresolved clinical work-type conflict is still present. Reports are blocked.",
            metrics=metrics, **context)
    return counts


def checked_report_reach(metrics, **context):
    """Derive display values from complete validated counts; never guess capacity.

    A missing/None derived percentage or hours can be reconstructed exactly from
    the three counts. Contradictory, nonfinite, negative or zero-denominator data
    remains a blocking error. Input dictionaries and source scans are unchanged.
    """
    counts = validated_shift_counts(metrics, **context)
    total = counts["recorded_clinical_shifts"]
    if total <= 0:
        raise ReportDataError(
            "This report entry has no recorded clinical shifts, so Learner Reach has no denominator. "
            "An empty group must not be published as 0% or N/A.", metrics=metrics, **context)
    expected = {
        **counts,
        "recorded_clinical_hours": total * TEACHING_HOURS_PER_STUDENT_SHIFT,
        "hours_with_students": counts["shifts_with_students"] * TEACHING_HOURS_PER_STUDENT_SHIFT,
        "hours_without_students": counts["shifts_without_students"] * TEACHING_HOURS_PER_STUDENT_SHIFT,
        "learner_reach_pct": round(100 * counts["shifts_with_students"] / total, 1),
        "availability_review_shifts": 0,
    }
    for field in DISPLAY_FIELDS:
        supplied = metrics.get(field)
        if supplied is None:
            continue  # Missing display value only: counts fully determine it.
        if (isinstance(supplied, bool) or not isinstance(supplied, Real)
                or not math.isfinite(supplied)
                or not math.isclose(float(supplied), float(expected[field]), rel_tol=0, abs_tol=1e-8)):
            raise ReportDataError(
                f"The stored '{field}' does not match the clinical shift counts "
                f"(expected {expected[field]}).", metrics=metrics, **context)
    return {**metrics, **expected}


@contextmanager
def report_step(report, **context):
    """Add the failed export stage without exposing traceback locals or learners."""
    try:
        yield
    except ReportDataError:
        raise
    except OPDArchiveError as exc:
        # Preserve the existing source/cell table for real scheduling conflicts.
        from schedule_app.services.teaching_validation import TeachingConflictError
        if isinstance(exc, TeachingConflictError):
            raise
        raise ReportDataError(str(exc), report=report, **context) from None
    except Exception as exc:
        # Arbitrary exception text can include input data. Export its type only.
        raise ReportDataError(
            f"An unexpected {type(exc).__name__} occurred while creating this report component.",
            report=report, **context) from None
