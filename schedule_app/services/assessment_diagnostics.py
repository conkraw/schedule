"""Read-only, per-preceptor explanation of unavailable completion percentages.

Names stay in the protected matching UI. No report bundle, CSV, log or GitHub
write receives this diagnostic object. Calculations use the engine's existing
rows and unresolved-eligibility list; this does not relax identity safeguards.
"""
from __future__ import annotations
import hashlib
import json

DIAGNOSTICS_UI_VERSION = "2026-10-01-completion-review-1"
FOCUS_KEY = "assessment_completion_diagnostics_focus"


def unavailable_completion_choices(bundle):
    """Only true unresolved checks, not calculated or zero-eligibility rows."""
    choices = {}
    for row in bundle.get("rows", []):
        status = str(row.get("assessment_status", ""))
        if status == "Calculated" or status.startswith("No eligible students"):
            continue
        token = hashlib.sha256(json.dumps([
            row.get("preceptor_name"), row.get("group_year"),
            row.get("report_start_date"), row.get("report_end_date"),
        ], ensure_ascii=True).encode()).hexdigest()[:24]
        choices[token] = row
    return dict(sorted(choices.items(), key=lambda item: (
        str(item[1].get("preceptor_name", "")).casefold(),
        str(item[1].get("report_start_date", "")), item[0])))


def unresolved_students_for_row(row, unmatched):
    """Return only eligible, unresolved OPD names for the chosen preceptor/period.

    Use the engine's returned unmatched records, NOT every unrecognized name in
    the archive and not the complete set of students this educator has assessed.
    """
    found = {}
    for item in unmatched:
        if (item.get("preceptor_name") != row.get("preceptor_name")
                or item.get("academic_year") != row.get("academic_year")):
            continue
        # Optional exact boundaries make this forward compatible with richer
        # engine diagnostics without changing the existing result schema.
        if any(item.get(k) is not None and item.get(k) != row.get(k)
               for k in ("group_year", "report_start_date", "report_end_date")):
            continue
        key = item.get("name_key")
        if not key:
            continue
        found[key] = {"name_key": key, "student_name": item.get("student_name", ""),
                      "assigned_shifts": item.get("assigned_shifts"),
                      "issue": item.get("issue", "OPD name needs a matching OASIS student name")}
    return sorted(found.values(), key=lambda item: str(item["student_name"]).casefold())


def completion_review_detail(row, unmatched):
    """An explicit distinction between completed forms and eligible students."""
    missing = unresolved_students_for_row(row, unmatched)
    return {
        "preceptor_name": row.get("preceptor_name", ""),
        "academic_year": row.get("academic_year", ""),
        "report_start_date": row.get("report_start_date", ""),
        "report_end_date": row.get("report_end_date", ""),
        "record_id": row.get("record_id", ""),
        "minimum_shifts": row.get("minimum_shifts"),
        "eligible_students": row.get("eligible_students"),
        "assessment_status": row.get("assessment_status", ""),
        "clinical_forms_submitted": row.get("clinical_forms_submitted"),
        "hp_forms_submitted": row.get("hp_forms_submitted"),
        "unresolved_students": missing,
    }


def make_student_review_focus(bundle, row):
    """Provider-and-period filter, not a saved identity decision or a flag waiver."""
    return {"context": bundle.get("context"), "preceptor_name": row.get("preceptor_name"),
            "group_year": row.get("group_year"),
            "academic_year": row.get("academic_year"),
            "report_start_date": row.get("report_start_date"),
            "report_end_date": row.get("report_end_date")}


def focused_missing_names(missing, unmatched, focus, context):
    """Filter the correction queue dynamically so verified saves clear it.

    Return None when the report's scope changed; the caller restores the normal
    all-unmatched queue. A valid empty dictionary means this teacher is resolved.
    """
    if not isinstance(focus, dict) or focus.get("context") != context:
        return None
    keys = {item["name_key"] for item in unresolved_students_for_row(focus, unmatched)}
    return {key: value for key, value in missing.items() if key in keys}
