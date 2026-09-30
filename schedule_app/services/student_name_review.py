"""Presentation data for explicit OPD-name -> OASIS Student External ID matches.

No fuzzy matching, new username inference, or changes to completion calculations.
These names/IDs stay in the protected review UI; never attach this object to a
report bundle, CSV, Word document, log, or plaintext repository file.
"""
from __future__ import annotations
from collections import defaultdict
import hashlib
import json

from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.oasis_educator_reports import name_key
from schedule_app.services.student_assessment_links import student_name_key
from schedule_app.services.teaching_evaluations import active_periods

STUDENT_REVIEW_UI_VERSION = 1


def oasis_student_choices(prepared):
    """One selectable name/ID pair per observed association; never invent an ID.

    The crosswalk already omits the recognized MD-class suffix. Two IDs attached
    to the same name remain separate choices and require explicit confirmation.
    Aliases with the same ID may appear under each observed name.
    """
    choices = {}
    names = prepared.get("student_names", {})
    for key, ids in sorted(prepared.get("name_ids", {}).items()):
        display = str(names.get(key) or key).strip()
        clean_ids = sorted({str(sid).strip() for sid in ids if str(sid).strip()})
        for sid in clean_ids:
            if sid.casefold() in {"nan", "n/a", "none", "null", "-", "--"}:
                continue
            token = hashlib.sha256(json.dumps([key, sid], ensure_ascii=True).encode()).hexdigest()
            choices[token] = {"oasis_student_name": display, "external_id": sid,
                              "ambiguous_name": len(clean_ids) > 1}
    return dict(sorted(choices.items(), key=lambda pair: (
        student_name_key(pair[1]["oasis_student_name"]), pair[1]["external_id"])))


def selected_student_id(choices, selection):
    """Reject missing/stale choices rather than persisting an arbitrary ID."""
    if not selection or selection not in choices:
        raise OPDArchiveError("Choose the matching OASIS student before saving. Nothing was changed.")
    return choices[selection]["external_id"]


def student_name_review(inputs, scan, years, unmatched):
    """Return all current assigned names and a missing-only correction queue.

    Includes below-threshold names for proactive correction; only the existing
    `unmatched` results determine which flags affect completion percentages.
    A confirmed saved ID is reused even if no assessment exists for it in this
    period. Absence of an assessment is not absence of an identity match.
    """
    review = {name_key(name) for name in scan.get("unresolved_preceptor_labels", [])}
    active = {}
    for year, start, end, label in active_periods(scan, years):
        teachers = {row["preceptor_name"] for row in scan["monthly"]
                    if row["academic_start_year"] == year and row["no_of_shifts"] > 0
                    and name_key(row["preceptor_name"]) not in review}
        for row in inputs["assignments"]:
            if row["preceptor_name"] not in teachers or not start.isoformat() <= row["date"] <= end.isoformat():
                continue
            key = student_name_key(row["student"])
            if not key:
                continue
            item = active.setdefault(key, {"student_name": row["student"].strip(),
                                            "preceptors": set(), "periods": set()})
            item["preceptors"].add(row["preceptor_name"])
            item["periods"].add(label)
    affects = defaultdict(set)
    for row in unmatched:
        affects[row["name_key"]].add(row["preceptor_name"])
    entries = inputs["student_links"]["entries"]
    name_ids = inputs["prepared"]["name_ids"]
    result, missing = {}, {}
    for key, item in sorted(active.items()):
        ids = name_ids.get(key, [])
        saved_id = entries.get(key, {}).get("external_id", "")
        state = "Saved match" if saved_id else "Exact normalized name match" if len(ids) == 1 else "Needs review"
        row = {"name_key": key, "student_name": item["student_name"],
               "preceptors": sorted(item["preceptors"], key=name_key),
               "periods": sorted(item["periods"]), "match_status": state,
               "affects_completion": bool(affects[key])}
        result[key] = row
        if state == "Needs review":
            missing[key] = {**row, "issue": (
                "This OASIS name has multiple Student External IDs; choose the correct student"
                if len(ids) > 1 else
                "No unique OASIS name/Student External ID match was found for this OPD name")}
    return {"active": result, "missing": missing,
            "eligible_missing_count": sum(row["affects_completion"] for row in missing.values())}


def review_table_rows(missing):
    """Readable, one-row-per-name table without implying zero assessments."""
    return [{"OPD student name": row["student_name"], "Issue": row["issue"],
             "Preceptor(s)": "; ".join(row["preceptors"]),
             "Reporting period(s)": "; ".join(row["periods"]),
             "Affects completion percentage": "YES" if row["affects_completion"] else "NO — below 3-shift threshold"}
            for row in missing.values()]


def names_for_external_id(choices, external_id):
    """Display current OASIS names for an already saved ID without reassigning it."""
    return sorted({row["oasis_student_name"] for row in choices.values()
                   if row["external_id"] == external_id}, key=student_name_key)
