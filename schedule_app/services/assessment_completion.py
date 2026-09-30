"""Read-only, unique-student assessment completion for teaching reports.

Denominator: unique students with >=3 distinct retained AM/PM assignments to a
preceptor IN the selected period. Numerators: those same eligible students with
at least one submitted target form, also IN the period. Each form type is
separate; repeated questions, exports, or forms for a student cannot inflate it.
No answer content, grades or comments are read into this calculation.
"""
from __future__ import annotations
from collections import defaultdict
import csv
from datetime import date, datetime, timezone
import hashlib
import io
import json
import re

from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.oasis_evaluations import _CSV_LOCK, OASIS_MAX_FIELD_CHARS
from schedule_app.services.oasis_student_evaluations import (
    GitHubOASISStudentEvaluations, inspect_student_oasis_csv, _form_key,
)
from schedule_app.services.oasis_educator_reports import name_key, email_username
from schedule_app.services.preceptor_oasis_links import GitHubPreceptorOASISLinks, period_key
from schedule_app.services.student_assessment_links import GitHubStudentAssessmentLinks, student_name_key
from schedule_app.services.teaching_evaluations import active_periods, parse_saved_summary
from schedule_app.services.oasis_workflow import GitHubOASISSummaries
from schedule_app.services.teaching_analysis import teaching_scan_archives

ASSESSMENT_VERSION = 1
CLINICAL = "*Clinical Assessment of Student"
HP = "*PEDS History Taking & Physical Exam"
TARGETS = {_form_key(CLINICAL): "clinical", _form_key(HP): "hp"}
MAX_SOURCES = 100
MAX_TOTAL_BYTES = 64 * 1024 * 1024
METHOD_NOTE = (
    "Eligible students were assigned to this preceptor on at least 3 distinct AM/PM shifts in the reporting period. "
    "Each eligible student counts once per form type, even with several submitted forms. "
    "AM and PM on the same day are two shifts. Assignment dates and OASIS Submit Date both use the displayed period."
)
SCOPE_NOTE = (
    "This is a record-completeness measure, not a judgment of teaching quality. "
    "An absent record may reflect an incomplete export or an unmatched identifier. "
    "The two form types are shown separately; this report does not establish that both are required for every student."
)
COLUMNS = ("preceptor_name", "academic_year", "report_start_date", "report_end_date", "record_id",
           "eligible_students_3plus_shifts", "clinical_students_evaluated", "clinical_completion_pct",
           "hp_students_evaluated", "hp_completion_pct", "either_students_evaluated", "either_completion_pct",
           "clinical_forms_submitted", "hp_forms_submitted", "student_feedback_evaluations", "assessment_status")


def _parse_submit(value):
    value = str(value or "").strip()
    if not value:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
                "%m/%d/%Y %H:%M:%S", "%m/%d/%Y %H:%M", "%m/%d/%Y",
                "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S.%f"):
        try:
            return datetime.strptime(value, fmt).date().isoformat()
        except ValueError:
            pass
    raise ValueError("Invalid Submit Date")


def prepare_student_assessments(exports):
    """Collapse question-level rows into unique form metadata across snapshots.

    Missing or conflicting identifiers are surfaced, not converted to guessed
    identities or zero. Blank fields can be filled by a consistent nonblank copy.
    Only the two requested form types contribute completion counts. All three
    student-form types may provide a name-to-external-ID crosswalk.
    """
    grouped = {}
    name_ids, names = defaultdict(set), {}
    issues, courses = [], set()
    source_count = 0
    for filename, raw in exports:
        source_count += 1
        details = inspect_student_oasis_csv(raw)
        codec = {"UTF-8": "utf-8-sig", "UTF-16": "utf-16", "Windows-1252": "cp1252"}[details["encoding"]]
        with _CSV_LOCK:
            before = csv.field_size_limit()
            csv.field_size_limit(OASIS_MAX_FIELD_CHARS)
            try:
                reader = csv.DictReader(io.StringIO(raw.decode(codec), newline=""), strict=True)
                reader.fieldnames = [x.strip().lstrip("\ufeff") for x in reader.fieldnames]
                needed = {"Student External ID", "Student", "Evaluator Email", "Evaluation", "Form Record", "Submit Date", "Course ID"}
                if not needed.issubset(reader.fieldnames):
                    raise OPDArchiveError("Student-assessment export is missing required identity columns: "
                                          + ", ".join(sorted(needed - set(reader.fieldnames))) + ".")
                for line, row in enumerate(reader, start=2):
                    if not any(str(x or "").strip() for x in row.values()):
                        continue
                    course = row["Course ID"].strip()
                    student, sid = row["Student"].strip(), row["Student External ID"].strip()
                    if sid.casefold() in {"nan", "n/a", "none", "null", "-", "--"}:
                        sid = ""
                    key = student_name_key(student)
                    if key and sid:
                        name_ids[key].add(sid)
                        names.setdefault(key, re.sub(r";\s*MD\s*\d{4}\s*$", "", student, flags=re.I).strip())
                    target = TARGETS.get(_form_key(row["Evaluation"]))
                    if target is None:
                        continue
                    courses.add(course)
                    fid = row["Form Record"].strip()
                    if not course or not fid:
                        issues.append({"issue": "Missing Course ID or Form Record", "source": filename,
                                       "csv_record": line, "form_record": fid, "username": "", "submit_date": ""})
                        continue
                    identity = (course, target, fid)
                    group = grouped.setdefault(identity, {"course": course, "target": target, "form_record": fid,
                        "emails": set(), "sids": set(), "dates": set(), "sources": set(), "bad_dates": False,
                        "evaluator_names": set(), "evaluator_username_hints": set()})
                    group["sources"].add(filename)
                    email = row["Evaluator Email"].strip().lower()
                    if email:
                        group["emails"].add(email)
                    if sid:
                        group["sids"].add(sid)
                    if row.get("Evaluator", "").strip():
                        group["evaluator_names"].add(row["Evaluator"].strip())
                    hint = row.get("Evaluator Username", "").strip().lower()
                    if hint:
                        group["evaluator_username_hints"].add(hint)
                    try:
                        submitted = _parse_submit(row["Submit Date"])
                        if submitted:
                            group["dates"].add(submitted)
                    except ValueError:
                        group["bad_dates"] = True
            except (csv.Error, UnicodeError):
                raise OPDArchiveError("An archived student-assessment CSV could not be parsed safely.") from None
            finally:
                csv.field_size_limit(before)
    forms, unsubmitted = [], 0
    for group in grouped.values():
        if not group["dates"] and not group["bad_dates"]:
            unsubmitted += 1
            continue
        usernames = set()
        bad_email = False
        for email in group["emails"]:
            try:
                username = email_username(email)
                if not username:
                    bad_email = True
                else:
                    usernames.add(username)
            except (ValueError, OPDArchiveError):
                bad_email = True
        problems = []
        if group["bad_dates"] or len(group["dates"]) != 1:
            problems.append("Invalid or conflicting Submit Date")
        if len(group["sids"]) != 1:
            problems.append("Missing or conflicting Student External ID")
        if bad_email or len(usernames) != 1:
            problems.append("Missing or conflicting Evaluator Email username")
        common = {"form_record": group["form_record"], "course": group["course"], "target": group["target"],
                  "username": next(iter(usernames)) if len(usernames) == 1 else "",
                  "submit_date": next(iter(group["dates"])) if len(group["dates"]) == 1 else "",
                  "source": "; ".join(sorted(group["sources"])), "csv_record": ""}
        if problems:
            issues.append({**common, "issue": "; ".join(problems),
                           "evaluator_names": sorted(group["evaluator_names"]),
                           "evaluator_username_hints": sorted(group["evaluator_username_hints"])})
            continue
        forms.append({**common, "external_id": next(iter(group["sids"]))})
    return {"forms": forms, "name_ids": {key: sorted(ids) for key, ids in name_ids.items()},
            "student_names": names, "issues": issues, "courses": sorted(courses),
            "source_count": source_count, "unsubmitted_forms": unsubmitted}


def load_completion_inputs(archive, scan, years, *, progress=None):
    """Read assessment snapshots once; replay OPDs at their displayed snapshot.

    Raw workbooks and assessment scores/comments are never stored in the result.
    This temporary input object DOES hold matching names/IDs in session memory;
    only aggregated results are passed into reports and CSV downloads.
    """
    commit = archive._head()
    service = GitHubOASISStudentEvaluations(archive)
    listing = service.list_exports(commit=commit)
    filenames = listing["filenames"]
    if len(filenames) > MAX_SOURCES:
        raise OPDArchiveError("More than 100 student-assessment snapshots: no completion percentages were calculated from a partial archive.")
    manifests, total = [], 0
    def sources():
        nonlocal total
        for number, filename in enumerate(filenames, start=1):
            loaded = service.load(filename, commit=commit)
            total += len(loaded["raw"])
            if total > MAX_TOTAL_BYTES:
                raise OPDArchiveError("Student-assessment snapshots exceed 64 MiB; no partial completion calculation was made.")
            manifests.append({"archive_file": filename, "github_blob_sha": loaded["sha"]})
            if progress:
                progress(number, len(filenames))
            yield filename, loaded["raw"]
    prepared = prepare_student_assessments(sources())
    catalog = GitHubPreceptorOASISLinks(archive).load(commit=commit)
    student_links = GitHubStudentAssessmentLinks(archive).load(commit=commit)
    assignments = []
    replay = teaching_scan_archives(archive, scan["default_name_order"], commit=scan["commit"],
                                   assessment_collector=assignments.extend)
    # Pin to and verify the displayed OPD source set, even when other catalog
    # writes have moved Git HEAD since the OPDs were scanned.
    expected = sorted((r["archive_file"], r["github_blob_sha"]) for r in scan["sources"])
    found = sorted((r["archive_file"], r["github_blob_sha"]) for r in replay["sources"])
    if expected != found:
        raise OPDArchiveError("OPD source records changed. Refresh the teaching scan before evaluating completion.")
    summaries, feedback_issues = {}, {}
    for year, start, end, label in active_periods(scan, years):
        link = catalog["report_links"].get(period_key(start, end))
        if not link:
            feedback_issues[str(year)] = "Not checked: no exact-date educator-feedback summary linked"
            continue
        try:
            parsed = parse_saved_summary(GitHubOASISSummaries(archive).load(link["summary_filename"], commit=commit))
            if (parsed["details"]["start_date"], parsed["details"]["end_date"]) != (start, end):
                raise OPDArchiveError("Linked educator-feedback summary dates differ from this reporting period.")
            summaries[str(year)] = parsed
        except OPDArchiveError:
            feedback_issues[str(year)] = "Not checked: linked educator-feedback summary could not be verified; refresh its link"
    return {"prepared": prepared, "assignments": assignments, "catalog": catalog, "student_links": student_links,
            "summaries": summaries, "feedback_issues": feedback_issues, "sources": manifests,
            "opd_commit": scan["commit"], "commit": commit,
            "retrieved_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")}


def _warning(name, label, rid, direction, issue, action):
    return {"preceptor_name": name, "academic_year": label, "username": rid,
            "direction": direction, "issue": issue, "action": action}


def completion_context(scan, years):
    return {"opd_commit": scan["commit"], "periods": [
        [year, start.isoformat(), end.isoformat(), label]
        for year, start, end, label in active_periods(scan, years)]}


def build_completion_bundle(inputs, scan, years, *, courses=None):
    """Return provider-level metrics/warnings only (no student names or IDs).

    A missing external ID never shrinks the denominator. That preceptor's rates
    remain Not verified until resolved; report generation itself is not blocked.
    """
    prepared = inputs["prepared"]
    chosen = set(courses) if courses is not None else set(prepared["courses"])
    if courses is None and len(chosen) > 1:
        raise OPDArchiveError("Choose the student-assessment course(s) for this teaching report; multiple courses are archived.")
    if not chosen.issubset(set(prepared["courses"])):
        raise OPDArchiveError("Refresh the assessment data: a selected course is no longer available.")
    overrides = inputs["student_links"]["entries"]
    catalog = inputs["catalog"]
    review = {name_key(n) for n in scan.get("unresolved_preceptor_labels", [])}
    result_rows, warnings, unmatched = [], [], []
    for year, start, end, label in active_periods(scan, years):
        begin, finish = start.isoformat(), end.isoformat()
        active = sorted({r["preceptor_name"] for r in scan["monthly"]
                         if r["academic_start_year"] == year and r["no_of_shifts"] > 0
                         and name_key(r["preceptor_name"]) not in review}, key=name_key)
        pairs, name_lookup = defaultdict(set), {}
        for a in inputs["assignments"]:
            if a["preceptor_name"] not in active or not begin <= a["date"] <= finish:
                continue
            skey = student_name_key(a["student"])
            ids = [overrides[skey]["external_id"]] if skey in overrides else prepared["name_ids"].get(skey, [])
            if len(ids) == 1:
                identity = ("external_id", ids[0])
            else:
                identity = ("unresolved", skey)
                name_lookup[(a["preceptor_name"], identity)] = (a["student"], len(ids))
            pairs[(a["preceptor_name"], identity)].add((a["date"], a["shift"]))
        period_forms = [f for f in prepared["forms"] if f["course"] in chosen and begin <= f["submit_date"] <= finish]
        period_issues = [i for i in prepared["issues"] if (not i.get("course") or i["course"] in chosen)
                         and (not i.get("submit_date") or begin <= i["submit_date"] <= finish)]
        summary = inputs["summaries"].get(str(year))
        active_keys = {name_key(n) for n in active}
        active_usernames = {catalog["entries"].get(k, {}).get("record_id", "") for k in active_keys} - {""}
        unknown_issues = [i for i in period_issues if not i.get("username")
                          and not ({name_key(n) for n in i.get("evaluator_names", [])} & active_keys)
                          and not (set(i.get("evaluator_username_hints", [])) & active_usernames)]
        if unknown_issues:
            warnings.append(_warning("Unattributed / other evaluator", label, "", "Preceptor → student",
                f"{len(unknown_issues)} submitted form(s) have unverified evaluator identity and are not credited",
                "Review the source-metadata table. Percentages use identifiable matched records only; they can change after correction."))
        for name in active:
            rid = catalog["entries"].get(name_key(name), {}).get("record_id", "")
            eligible = {identity for (n, identity), slots in pairs.items() if n == name and len(slots) >= 3}
            resolved = {key for kind, key in eligible if kind == "external_id"}
            missing = [identity for identity in eligible if identity[0] == "unresolved"]
            provider_forms = [f for f in period_forms if rid and f["username"] == rid]
            provider_issues = [i for i in period_issues if (rid and i.get("username") == rid)
                               or (not i.get("username") and (
                                   name_key(name) in {name_key(n) for n in i.get("evaluator_names", [])}
                                   or (rid and rid in i.get("evaluator_username_hints", []))))]
            totals = {t: len([f for f in provider_forms if f["target"] == t]) for t in ("clinical", "hp")}
            completed = {t: {f["external_id"] for f in provider_forms if f["target"] == t} & resolved
                         for t in ("clinical", "hp")}
            clinical, hp = completed["clinical"], completed["hp"]
            available = bool(prepared["source_count"] and chosen)
            if not rid:
                status = "Not checked: preceptor username missing"
            elif not available:
                status = "Not checked: no student-assessment sources/course selected"
            elif provider_issues:
                status = "Not verified: assessment source metadata needs review"
            elif missing:
                status = f"Not verified: {len(missing)} eligible student name(s) need an external ID match"
            elif not eligible:
                status = "No eligible students (3+ shifts)"
            else:
                status = "Calculated"
            row = {"preceptor_name": name, "academic_year": label, "report_start_date": begin,
                   "report_end_date": finish, "record_id": rid, "eligible_students_3plus_shifts": len(eligible),
                   "clinical_students_evaluated": len(clinical) if status == "Calculated" else None,
                   "hp_students_evaluated": len(hp) if status == "Calculated" else None,
                   "either_students_evaluated": len(clinical | hp) if status == "Calculated" else None,
                   "clinical_forms_submitted": totals["clinical"] if rid and available and not provider_issues else None,
                   "hp_forms_submitted": totals["hp"] if rid and available and not provider_issues else None,
                   "assessment_status": status, "group_year": year,
                   "student_feedback_evaluations": None}
            for prefix in ("clinical", "hp", "either"):
                row[prefix + "_completion_pct"] = (round(100 * row[prefix + "_students_evaluated"] / len(eligible), 1)
                                                   if status == "Calculated" else None)
            if rid and summary is not None:
                row["student_feedback_evaluations"] = summary["rows_by_id"].get(rid, {}).get("evaluation_count", 0)
                if not row["student_feedback_evaluations"]:
                    warnings.append(_warning(name, label, rid, "Student → educator",
                        "No educator evaluation by a student found in the linked summary",
                        "Review the selected OASIS feedback summary, Submit Dates and username link."))
            else:
                warnings.append(_warning(name, label, rid, "Student → educator",
                    "Not checked: preceptor username missing" if not rid else inputs["feedback_issues"].get(str(year), "Not checked: no linked summary"),
                    "Save a username and an exact-date educator-feedback summary link above, then refresh this check."))
            if rid and available and not provider_issues:
                if not provider_forms:
                    warnings.append(_warning(name, label, rid, "Preceptor → student",
                        "No submitted Clinical Assessment or History & Physical forms found in these dates",
                        "Review/archive the student-assessment export and check Evaluator Email and Submit Date."))
                else:
                    for t, form_label in (("clinical", "Clinical Assessment of Student"), ("hp", "History Taking & Physical Exam")):
                        if not totals[t]:
                            warnings.append(_warning(name, label, rid, "Preceptor → student",
                                f"No submitted {form_label} forms found", "Review the student-assessment source; a different form type does not replace this one."))
            else:
                warnings.append(_warning(name, label, rid, "Preceptor → student", status,
                                         "Review source availability, the username link and source diagnostics."))
            if missing:
                warnings.append(_warning(name, label, rid, "Student identity", status,
                    "Resolve exact OPD-name → Student External ID links below; the denominator has not been reduced."))
                for identity in missing:
                    display, candidates = name_lookup[(name, identity)]
                    unmatched.append({"preceptor_name": name, "academic_year": label, "student_name": display,
                                      "name_key": identity[1], "assigned_shifts": len(pairs[(name, identity)]),
                                      "issue": "Name has multiple Student External IDs" if candidates else "No Student External ID found for OPD name"})
            result_rows.append(row)
    return {"version": ASSESSMENT_VERSION, "context": completion_context(scan, years), "rows": result_rows,
            "warnings": warnings, "source_count": prepared["source_count"], "sources": inputs["sources"],
            "retrieved_at": inputs["retrieved_at"], "assessment_commit": inputs["commit"],
            "username_mapping_sha": catalog.get("sha"), "student_mapping_sha": inputs["student_links"].get("sha"),
            "course_ids": sorted(chosen)}, unmatched


def unverified_bundle(scan, years, reason="Not checked: load evaluation completeness"):
    rows, warnings = [], []
    for year, start, end, label in active_periods(scan, years):
        review = set(scan.get("unresolved_preceptor_labels", []))
        names = sorted({r["preceptor_name"] for r in scan["monthly"] if r["academic_start_year"] == year
                        and r["no_of_shifts"] > 0 and r["preceptor_name"] not in review}, key=name_key)
        for name in names:
            row = {key: None for key in COLUMNS}
            row.update(preceptor_name=name, academic_year=label, report_start_date=start.isoformat(),
                       report_end_date=end.isoformat(), record_id="", assessment_status=reason, group_year=year)
            rows.append(row)
            for direction in ("Student → educator", "Preceptor → student"):
                warnings.append(_warning(name, label, "", direction, reason,
                    "Load/refresh the evaluation completeness check. Teaching reports can still be generated."))
    return {"version": ASSESSMENT_VERSION, "context": completion_context(scan, years), "rows": rows,
            "warnings": warnings, "source_count": 0, "sources": [], "retrieved_at": "Not checked", "course_ids": []}


def completion_signature(bundle):
    return hashlib.sha256(json.dumps(bundle, sort_keys=True, ensure_ascii=True).encode()).hexdigest() if bundle else None


def completion_rows(bundle, scan, year):
    if bundle is None:
        return []
    if bundle.get("version") != ASSESSMENT_VERSION or bundle.get("context", {}).get("opd_commit") != scan["commit"]:
        raise OPDArchiveError("Refresh evaluation completeness: its OPD snapshot differs from this report.")
    wanted = completion_context(scan, [year])["periods"]
    if any(p not in bundle["context"]["periods"] for p in wanted):
        raise OPDArchiveError("Refresh evaluation completeness: its dates differ from this report.")
    return [row for row in bundle["rows"] if row["group_year"] == year]


def completion_display(row, prefix):
    pct = row.get(prefix + "_completion_pct")
    if pct is None:
        return "No eligible students" if row["assessment_status"].startswith("No eligible") else "Not verified"
    return f"{row[prefix + '_students_evaluated']} / {row['eligible_students_3plus_shifts']} ({pct:.1f}%)"
