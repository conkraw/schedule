"""Presentation data for explicit OPD-name -> OASIS Student External ID matches.

No fuzzy matching, new username inference, or changes to completion calculations.
These names/IDs stay in the protected review UI; never attach this object to a
report bundle, CSV, Word document, log, or plaintext repository file.
"""
from __future__ import annotations
from collections import defaultdict
import hashlib
import json

from schedule_app.services.assessment_settings import DEFAULT_MINIMUM_SHIFTS, validate_minimum_shifts
from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.oasis_educator_reports import name_key
from schedule_app.services.student_assessment_links import student_name_key
from schedule_app.services.teaching_evaluations import active_periods
from schedule_app.services.student_name_matching import (
    StudentNameMatcher, student_matching_key, student_name_without_designations,
)

from schedule_app.services.assessment_progress import (
    AssessmentIdentityResolver, assessment_as_of, MATCHED, ABSENT, REVIEW,
)

STUDENT_REVIEW_UI_VERSION = 4


def oasis_student_choices(prepared):
    """One selectable name/ID pair per observed association; never invent an ID.

    Known trailing program/class designations are omitted from the display.
    Distinct records sharing that name remain ambiguous in the name-only selector.
    Aliases with the same ID may appear under each observed name.
    """
    choices = {}
    names = prepared.get("student_names", {})
    for key, ids in sorted(prepared.get("name_ids", {}).items()):
        display = student_name_without_designations(names.get(key) or key)
        clean_ids = sorted({str(sid).strip() for sid in ids if str(sid).strip()})
        for sid in clean_ids:
            if sid.casefold() in {"nan", "n/a", "none", "null", "-", "--"}:
                continue
            token = hashlib.sha256(json.dumps([key, sid], ensure_ascii=True).encode()).hexdigest()
            choices[token] = {"oasis_student_name": display, "external_id": sid,
                              "ambiguous_name": len(clean_ids) > 1}
    return dict(sorted(choices.items(), key=lambda pair: (
        student_name_key(pair[1]["oasis_student_name"]), pair[1]["external_id"])))



def name_only_student_choices(choices):
    """Return one visible option per OASIS name, carrying its record internally.

    A name linked to several external IDs is shown once as needing source review,
    not as several identical-looking choices. It cannot be saved from the name-only
    selector. The existing pair-level choices remain available to resolve already
    saved links and never change the encrypted catalog schema.
    """
    grouped = defaultdict(list)
    for token, row in choices.items():
        grouped[student_matching_key(row["oasis_student_name"])].append((token, row))
    result = {}
    for key, pairs in sorted(grouped.items()):
        pairs = sorted(pairs, key=lambda pair: pair[0])
        ids = {row["external_id"] for _, row in pairs}
        # A unique choice keeps its existing token so unrelated data refreshes
        # do not change which record it represents. Ambiguity gets a new token.
        if len(ids) == 1 and not any(row.get("ambiguous_name", False) for _, row in pairs):
            token, row = pairs[0]
            result[token] = dict(row)
        else:
            token = hashlib.sha256(json.dumps(["ambiguous_name", key], ensure_ascii=True).encode()).hexdigest()
            result[token] = {"oasis_student_name": pairs[0][1]["oasis_student_name"],
                             "external_id": "", "ambiguous_name": True}
    return result


def selected_name_student_id(choices, selection):
    """Resolve the selected *name* internally; never guess among duplicate names."""
    if not selection or selection not in choices:
        raise OPDArchiveError("Choose the matching OASIS student before saving. Nothing was changed.")
    row = choices[selection]
    if row.get("ambiguous_name") or not row.get("external_id"):
        raise OPDArchiveError("More than one OASIS student record uses this name. Review the source records in OER; no match was saved.")
    return row["external_id"]


def selected_student_id(choices, selection):
    """Reject missing/stale choices rather than persisting an arbitrary ID."""
    if not selection or selection not in choices:
        raise OPDArchiveError("Choose the matching OASIS student before saving. Nothing was changed.")
    return choices[selection]["external_id"]


def student_name_review(inputs, scan, years, unmatched, *, as_of=None):
    """Return all current assigned names and a missing-only correction queue.

    Possible name discrepancies form the routine queue. Absent OASIS names stay
    separate and need no confirmation. Includes below-threshold names for proactive
    correction; only the existing
    `unmatched` results determine which flags affect completion percentages.
    A confirmed saved ID is reused even if no assessment exists for it in this
    period. Absence of an assessment is not absence of an identity match.
    """
    cutoff = assessment_as_of(as_of)
    review = {name_key(name) for name in scan.get("unresolved_preceptor_labels", [])}
    active = {}
    for year, start, end, label in active_periods(scan, years):
        teachers = {row["preceptor_name"] for row in scan["monthly"]
                    if row["academic_start_year"] == year and row["no_of_shifts"] > 0
                    and name_key(row["preceptor_name"]) not in review}
        for row in inputs["assignments"]:
            if row["preceptor_name"] not in teachers or not start.isoformat() <= row["date"] <= min(end, cutoff).isoformat():
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
    matcher = AssessmentIdentityResolver(inputs["prepared"], inputs["student_links"]["entries"])
    result, missing, absent = {}, {}, {}
    for key, item in sorted(active.items()):
        match = matcher.resolve(item["student_name"])
        ids, state = match.candidate_ids, match.status
        row = {"name_key": key, "student_name": item["student_name"],
               "preceptors": sorted(item["preceptors"], key=name_key),
               "periods": sorted(item["periods"]), "match_status": state,
               "affects_completion": bool(affects[key])}
        result[key] = row
        if state == REVIEW:
            missing[key] = {**row, "issue": (
                "More than one OASIS student record uses this name; review the source records in OER"
                if len(ids) > 1 else
                "Possible spelling/name difference; confirm only if these are the same person"),
                "suggested_names": list(match.suggestions)}
        elif state == ABSENT:
            absent[key] = {**row, "issue": "No OASIS student record found; no confirmation required"}
    return {"active": result, "missing": missing, "absent": absent,
            "eligible_missing_count": sum(row["affects_completion"] for row in missing.values())}


def review_table_rows(missing, *, minimum_shifts=DEFAULT_MINIMUM_SHIFTS):
    """Readable, one-row-per-name table without implying zero assessments."""
    minimum_shifts = validate_minimum_shifts(minimum_shifts)
    return [{"OPD student name": row["student_name"], "Issue": row["issue"],
             "Preceptor(s)": "; ".join(row["preceptors"]),
             "Reporting period(s)": "; ".join(row["periods"]),
             "Affects completion percentage": "YES" if row["affects_completion"] else f"NO — below {minimum_shifts}-shift threshold"}
            for row in missing.values()]


def names_for_external_id(choices, external_id):
    """Display current OASIS names for an already saved ID without reassigning it."""
    return sorted({row["oasis_student_name"] for row in choices.values()
                   if row["external_id"] == external_id}, key=student_name_key)
