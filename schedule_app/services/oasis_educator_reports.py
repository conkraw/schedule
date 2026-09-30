"""Question-level OASIS CSV -> one educator row, without retaining student columns.

Question IDs (not changing Question Numbers) keep the old/new forms aligned.
Distinct submitted Course ID / Evaluation / Form Record keys count evaluations.
Repeated identical question responses count once. Conflicting copies stop export;
there is no hidden first/latest-snapshot policy and no average across different questions.
"""
from __future__ import annotations

import csv
import hashlib
import io
import json
import re
from collections import defaultdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Any
from zipfile import ZipFile, ZIP_DEFLATED

from schedule_app.services.oasis_evaluations import (
    inspect_oasis_csv, _CSV_LOCK, OASIS_MAX_FIELD_CHARS,
)
from schedule_app.services.opd_archive import OPDArchiveError

OASIS_REPORT_VERSION = 1
MAX_TOTAL_BYTES = 64 * 1024 * 1024
MAX_TOTAL_ROWS = 500_000
MAX_EXPORTS = 100
MAX_OUTPUT_BYTES = 64 * 1024 * 1024
SUMMARY_FILENAME = "oasis_educator_summary.csv"
# Stable Question ID -> exact wording seen in the supplied OASIS export.
# q1286 is a duration-category code, NOT a teaching-quality rating or weeks.
QUESTION_TEXTS = {
    "587": "Checked to see where my current knowledge and skills were.",
    "588": "Built on my knowledge and skill base.",
    "589": "Demonstrated respect for me as a learner.",
    "590": "Demonstrated respect for patients, staff, care providers, and other specialties.",
    "591": "Used high value, cost conscious care considerations in clinical decision making.",
    "592": "Encouraged me to integrate high value, cost conscious care in my clinical decision making (e.g., note writing, assessment and plan, presentations).",
    "593": "Created a safe environment for me to ask questions and voice uncertainty.",
    "594": "Asked me to include my differential diagnosis, assessment, and plan in my case presentations.",
    "595": "Asked me to provide the rationale for my clinical decisions in my case presentations.",
    "596": "Asked me to investigate a relevant clinical topic and report back.",
    "597": "Directly observed me (e.g., taking portions of a history, doing portions of a physical exam, communicating with patients).",
    "598": "Provided feedback by giving specific examples of what I did well.",
    "599": "Provided feedback by giving specific examples of how I could improve.",
    "600": "Helped me develop a plan to improve my knowledge or skills.",
    "1285": "My preceptor was a positive role model for me.",
    "2774": "(FQ) Overall, this faculty/preceptor/resident helped me to further my clinical learning.",
    "1286": "Please indicate the amount of time you worked with this preceptor in this rotation.",
}
STRENGTHS_QUESTION = "Please indicate this educator's strengths"
IMPROVEMENT_QUESTION = "Areas for Improvement"
COMMENT_TEXTS = {"172": STRENGTHS_QUESTION, "437": IMPROVEMENT_QUESTION}
REPORT_REQUIRED = (
    "Evaluator", "Form Record", "Question ID", "Question", "Answer text",
    "Multiple Choice Value", "Multiple Choice Label", "Submit Date",
)
META_FIELDS = ("Evaluator", "Evaluator Email", "Evaluator Username", "Evaluator External ID",
               "Submit Date", "Start Date", "End Date")
ISSUE_COLUMNS = ("issue", "source", "csv_record", "course_id", "evaluation", "form_record", "question_id")
NON_SCORED_LABELS = {"n/a", "na", "not applicable", "not observed", "unable to assess"}


class OASISReportError(OPDArchiveError):
    """Diagnostics never include student columns, answers, or raw secret/error values."""
    def __init__(self, message: str, issues: list[dict] | None = None):
        super().__init__(message)
        self.issues = issues or []


def name_key(value: str) -> str:
    return re.sub(r"\s*,\s*", ", ", re.sub(r"\s+", " ", str(value or "").strip())).casefold()


def question_key(value: str) -> str:
    return re.sub(r"\s+", " ", str(value or "").replace("\u2019", "'").strip()).rstrip(". :").casefold()


def normalize_email(value: str) -> str:
    value = str(value or "").strip().lower()
    if len(value) > 254 or not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._+\-]*@[A-Za-z0-9](?:[A-Za-z0-9.\-]*[A-Za-z0-9])?\.[A-Za-z]{2,}", value):
        return ""
    return value


def validate_username(value: str) -> str:
    if not isinstance(value, str) or not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._\-]{0,63}", value.strip()):
        raise OASISReportError("Enter the username only (before @): 1-64 letters, numbers, dots, underscores, or hyphens. No email address or spaces.")
    return value.strip().lower()


def email_username(value: str) -> str:
    email = normalize_email(value)
    if not email:
        return ""
    try:
        return validate_username(email.split("@", 1)[0])
    except OASISReportError:
        return ""


def _date(value: str) -> date | None:
    value = str(value or "").strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d", "%m/%d/%Y %H:%M:%S", "%m/%d/%Y", "%m-%d-%Y"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            pass
    return None


def _issue(message, source, row_number, key=("", "", ""), qid=""):
    return dict(zip(ISSUE_COLUMNS, (message, source, row_number, *key, qid)))


def _metadata_value(field, value):
    value = str(value or "").strip()
    if field == "Evaluator":
        return name_key(value)
    if field in ("Evaluator Email", "Evaluator Username", "Evaluator External ID"):
        return value.lower()
    return value


def _score(value: str, label: str):
    """The exported Multiple Choice Value is the only numeric source."""
    value = value.strip()
    if not value or question_key(value) in NON_SCORED_LABELS or question_key(label) in NON_SCORED_LABELS:
        return None
    try:
        number = Decimal(value)
    except InvalidOperation:
        raise OASISReportError("Non-numeric Multiple Choice Value; review the indicated export/form/question.") from None
    if not number.is_finite() or number.copy_abs() > Decimal("1000000"):
        raise OASISReportError("Invalid or oversized Multiple Choice Value; review the indicated export/form/question.")
    return number


def parse_export(raw: bytes, source: str) -> tuple[list[dict], dict]:
    """Discard every Student/Who Completed column immediately after CSV parsing."""
    details = inspect_oasis_csv(raw)
    encoding = {"UTF-8": "utf-8-sig", "UTF-16": "utf-16", "Windows-1252": "cp1252"}[details["encoding"]]
    with _CSV_LOCK:
        old = csv.field_size_limit()
        csv.field_size_limit(OASIS_MAX_FIELD_CHARS)
        try:
            reader = csv.DictReader(io.StringIO(raw.decode(encoding), newline=""))
            reader.fieldnames = [v.strip().lstrip("\ufeff") for v in reader.fieldnames or []]
            missing = [v for v in REPORT_REQUIRED if v not in reader.fieldnames]
            if missing:
                raise OASISReportError("The report needs these columns: " + ", ".join(missing))
            result = []
            for rownum, row in enumerate(reader, 2):
                if all(not str(v or "").strip() for v in row.values()):
                    continue
                key = tuple(row[c].strip() for c in ("Course ID", "Evaluation", "Form Record"))
                qid = row["Question ID"].strip()
                if not all(key) or not re.fullmatch(r"[0-9]+", qid) or not row["Question"].strip():
                    raise OASISReportError("Missing form/course/evaluation/question identifier. No report was created.",
                                           [_issue("Missing/invalid required identifier or Question text", source, rownum, key, qid)])
                try:
                    score = _score(row["Multiple Choice Value"], row["Multiple Choice Label"])
                except OASISReportError as exc:
                    raise OASISReportError(str(exc), [_issue(str(exc), source, rownum, key, qid)]) from None
                result.append({"form_key": key, "qid": qid, "question": row["Question"].strip(),
                               "answer": row["Answer text"], "score": score,
                               "choice_value": row["Multiple Choice Value"].strip(),
                               "choice_label": row["Multiple Choice Label"].strip(),
                               "metadata": {v: row.get(v, "").strip() for v in META_FIELDS},
                               "source": source, "csv_record": rownum})
            return result, details
        finally:
            csv.field_size_limit(old)


def prepare_reports(exports: list[tuple[str, bytes]]) -> dict:
    """Validate selected snapshots, collapse identical responses, resolve identities.

    Conflicting copies of one response or form metadata block all selected exports;
    the user can select one intended snapshot instead of combining conflicting ones.
    """
    if not exports or len(exports) > MAX_EXPORTS or sum(len(raw) for _, raw in exports) > MAX_TOTAL_BYTES:
        raise OASISReportError("Choose 1-100 exports totaling at most 64 MiB for one report.")
    forms, catalog, duplicates, source_rows = {}, {}, 0, []
    rows_total = 0
    for source, raw in exports:
        records, details = parse_export(raw, source)
        rows_total += len(records)
        if rows_total > MAX_TOTAL_ROWS:
            raise OASISReportError("Selected exports exceed the 500,000 response-row limit. Select fewer exports.")
        source_rows.append({"source": source, "response_rows": len(records), "sha256": hashlib.sha256(raw).hexdigest()})
        for record in records:
            key, qid = record["form_key"], record["qid"]
            norm_question = question_key(record["question"])
            expected = (QUESTION_TEXTS | COMMENT_TEXTS).get(qid)
            if expected and question_key(expected) != norm_question:
                raise OASISReportError("A known Question ID has changed wording. Review the question mapping before exporting.",
                    [_issue("Known Question ID has different wording", source, record["csv_record"], key, qid)])
            q = catalog.setdefault(qid, {"question_id": qid, "question": record["question"],
                                        "choice_options": set(), "has_choice": False, "has_text": False})
            if question_key(q["question"]) != norm_question:
                raise OASISReportError("One Question ID has conflicting question text across selected exports.",
                    [_issue("Conflicting text for Question ID", source, record["csv_record"], key, qid)])
            if record["choice_label"] or record["choice_value"]:
                q["has_choice"] = True
                q["choice_options"].add((record["choice_value"], record["choice_label"]))
            q["has_text"] |= bool(record["answer"].strip())
            form = forms.setdefault(key, {"key": key, "metadata": {v: "" for v in META_FIELDS},
                                          "responses": {}, "locations": []})
            location = (source, record["csv_record"], qid)
            for field, value in record["metadata"].items():
                previous = form["metadata"][field]
                if previous and value and _metadata_value(field, previous) != _metadata_value(field, value):
                    raise OASISReportError("Conflicting metadata for the same evaluation. Select the intended export or correct the source.",
                        [_issue("Conflicting " + field + " for one Form Record", source, record["csv_record"], key, qid)])
                if not previous and value:
                    form["metadata"][field] = value
            fingerprint = (record["score"], record["choice_value"], record["choice_label"], record["answer"])
            previous = form["responses"].get(qid)
            if previous:
                prior_fp = (previous["score"], previous["choice_value"], previous["choice_label"], previous["answer"])
                if prior_fp != fingerprint:
                    refs = previous["locations"] + [location]
                    raise OASISReportError("Different answers exist for the same evaluation/question. Choose the intended snapshot; no averages or comments were combined.",
                        [_issue("Conflicting copies of one response", s, n, key, qi) for s, n, qi in refs])
                duplicates += 1
                previous["locations"].append(location)
            else:
                form["responses"][qid] = {k: record[k] for k in ("score", "choice_value", "choice_label", "answer")}
                form["responses"][qid]["locations"] = [location]
            form["locations"].append(location)
    submitted, unsubmitted = [], 0
    for form in forms.values():
        meta = form["metadata"]
        if not meta["Submit Date"]:
            unsubmitted += 1
            continue
        if _date(meta["Submit Date"]) is None or not meta["Evaluator"]:
            source, num, qid = form["locations"][0]
            raise OASISReportError("A submitted evaluation has a missing educator or unreadable Submit Date.",
                [_issue("Missing educator or invalid Submit Date", source, num, form["key"], qid)])
        submitted.append(form)
    _assign_educators(submitted)
    # JSON-compatible sorted choice keys keep UI cache signatures stable.
    for q in catalog.values():
        q["choice_options"] = sorted(q["choice_options"])
    return {"version": OASIS_REPORT_VERSION, "forms": submitted, "questions": catalog,
            "sources": source_rows, "response_rows": rows_total, "duplicates_removed": duplicates,
            "unsubmitted_forms_excluded": unsubmitted}


def _assign_educators(forms):
    """Prefer exported External ID, then exported Username, then full email.

    Email-local-part is the output record ID, NOT an automatic identity merge key.
    A name-only record may join a uniquely matching name; ambiguous names stop.
    """
    ext_by_user, ext_by_email, user_by_email = defaultdict(set), defaultdict(set), defaultdict(set)
    for form in forms:
        m = form["metadata"]
        ext, user, email = (m["Evaluator External ID"].lower(), m["Evaluator Username"].lower(), normalize_email(m["Evaluator Email"]))
        if user and ext:
            ext_by_user[user].add(ext)
        if email and ext:
            ext_by_email[email].add(ext)
        if email and user:
            user_by_email[email].add(user)
    for mapping in (ext_by_user, ext_by_email, user_by_email):
        if any(len(v) > 1 for v in mapping.values()):
            raise OASISReportError("An educator email or username links to multiple exported educator identifiers. Correct the identity conflict in the selected exports before reporting.")
    names_to_keys = defaultdict(set)
    for form in forms:
        m = form["metadata"]
        user, email = m["Evaluator Username"].lower(), normalize_email(m["Evaluator Email"])
        ext = m["Evaluator External ID"].lower() or next(iter(ext_by_user.get(user, set())), "") or next(iter(ext_by_email.get(email, set())), "")
        user = user or next(iter(user_by_email.get(email, set())), "")
        key = "ext:" + ext if ext else "user:" + user if user else "email:" + email if email else ""
        form["educator_key"] = key
        if key:
            names_to_keys[name_key(m["Evaluator"])].add(key)
    for form in forms:
        if not form["educator_key"]:
            key = name_key(form["metadata"]["Evaluator"])
            candidates = names_to_keys[key]
            if len(candidates) > 1:
                raise OASISReportError("A name-only educator matches more than one identity. Add a source identifier before combining these exports.")
            form["educator_key"] = next(iter(candidates), "name:" + key)


def filtered_forms(prepared: dict, evaluation_types=None, courses=None,
                   start_date: date | None = None, end_date: date | None = None,
                   date_field: str = "Submit Date") -> list[dict]:
    if prepared.get("version") != OASIS_REPORT_VERSION:
        raise OASISReportError("Reload the selected evaluations after the report-code update.")
    if date_field not in ("Submit Date", "Start Date", "End Date"):
        raise OASISReportError("Choose Submit Date, Start Date, or End Date for the date filter.")
    if (start_date is None) != (end_date is None) or (start_date is not None and start_date > end_date):
        raise OASISReportError("Choose both dates, with the end date on or after the start date.")
    rows = []
    for form in prepared["forms"]:
        if evaluation_types is not None and form["key"][1] not in evaluation_types:
            continue
        if courses is not None and form["key"][0] not in courses:
            continue
        if start_date is not None:
            day = _date(form["metadata"][date_field])
            if day is None:
                s, n, q = form["locations"][0]
                raise OASISReportError("A selected evaluation has an unreadable date for the filter; no evaluation was silently dropped.",
                    [_issue("Invalid " + date_field, s, n, form["key"], q)])
            if not start_date <= day <= end_date:
                continue
        rows.append(form)
    return rows


def question_columns(prepared: dict) -> list[str]:
    ids = list(QUESTION_TEXTS)
    for qid in sorted(prepared["questions"], key=int):
        q = prepared["questions"][qid]
        if q["has_choice"] and qid not in ids and comment_kind(q["question"]) is None:
            ids.append(qid)
    return ids


def comment_kind(question: str) -> str | None:
    normalized = question_key(question)
    if normalized == question_key(STRENGTHS_QUESTION):
        return "strengths_comments"
    if normalized == question_key(IMPROVEMENT_QUESTION):
        return "areas_for_improvement_comments"
    return None


def educator_summary(prepared: dict, overrides=None, **filters) -> dict:
    """Return numeric preview and ID issues. Export refuses unresolved record IDs."""
    overrides = overrides or {}
    forms = filtered_forms(prepared, **filters)
    groups = defaultdict(list)
    for form in forms:
        groups[form["educator_key"]].append(form)
    qids = question_columns(prepared)
    columns = ["record_id", "educator_name", "evaluator_email", "evaluation_count"]
    for qid in qids:
        columns += [f"q{qid}_mean", f"q{qid}_n"]
    columns += ["strengths_comments", "areas_for_improvement_comments", "email_missing", "record_id_source"]
    rows, educators, issues, detail = [], [], [], []
    for key, items in sorted(groups.items(), key=lambda x: name_key(x[1][0]["metadata"]["Evaluator"])):
        items.sort(key=lambda f: (_date(f["metadata"]["Submit Date"]), f["metadata"]["Submit Date"], f["key"]))
        names = sorted({f["metadata"]["Evaluator"].strip() for f in items}, key=name_key)
        emails = sorted({normalize_email(f["metadata"]["Evaluator Email"]) for f in items} - {""})
        usernames = sorted({f["metadata"]["Evaluator Username"] for f in items} - {""})
        exported_emails = sorted({f["metadata"]["Evaluator Email"] for f in items} - {""})
        # Recover a known email from another included evaluation for the same educator.
        email = emails[0] if len(emails) == 1 else ""
        supplied = overrides.get(key, {})
        supplied = supplied.get("record_id", "") if isinstance(supplied, dict) else supplied
        record_id = validate_username(supplied) if supplied else email_username(email)
        id_source = "manual_username" if supplied else "evaluator_email" if record_id else "unresolved"
        reasons = []
        if len(emails) > 1 and not supplied:
            reasons.append("Multiple Evaluator Email addresses; confirm the intended username.")
        if not record_id:
            reasons.append("Missing/invalid Evaluator Email. Enter the educator's username; no value was guessed.")
        row = {"record_id": record_id, "educator_name": names[0], "evaluator_email": "; ".join(emails),
               "evaluation_count": len(items), "email_missing": "NO" if emails else "YES", "record_id_source": id_source}
        for qid in qids:
            numbers = [f["responses"][qid]["score"] for f in items if qid in f["responses"] and f["responses"][qid]["score"] is not None]
            mean = ((sum(numbers) / len(numbers)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP) if numbers else None)
            row[f"q{qid}_mean"], row[f"q{qid}_n"] = (format(mean, ".2f") if mean is not None else ""), len(numbers)
            question = prepared["questions"].get(qid, {}).get("question", QUESTION_TEXTS.get(qid, ""))
            detail.append({"record_id": record_id, "educator_name": names[0], "question_id": qid,
                           "question": question, "mean_value": row[f"q{qid}_mean"], "response_count": len(numbers),
                           "evaluation_count": len(items)})
        comments = {"strengths_comments": [], "areas_for_improvement_comments": []}
        for form in items:
            for qid, response in sorted(form["responses"].items(), key=lambda kv: int(kv[0])):
                kind = comment_kind(prepared["questions"][qid]["question"])
                if kind and response["answer"].strip():
                    comments[kind].append(response["answer"].strip())
        for field, parts in comments.items():
            row[field] = "\n\n".join(f"{i}. {comment}" for i, comment in enumerate(parts, 1))
        rows.append(row)
        educators.append({"educator_key": key, "educator_name": names[0], "name_variants": "; ".join(names),
                          "evaluator_email": "; ".join(exported_emails), "exported_username": "; ".join(usernames),
                          "record_id": record_id, "record_id_source": id_source, "email_missing": row["email_missing"]})
        if reasons:
            issues.append({"educator_key": key, "educator_name": names[0], "record_id": record_id,
                           "issue": " ".join(reasons)})
    owners = defaultdict(list)
    for educator in educators:
        if educator["record_id"]:
            owners[educator["record_id"]].append(educator)
    for rid, entries in owners.items():
        if len(entries) > 1:
            for e in entries:
                issues.append({"educator_key": e["educator_key"], "educator_name": e["educator_name"], "record_id": rid,
                               "issue": "Duplicate record_id assigned to different educator identities. Correct the usernames or source identity before exporting."})
    return {"rows": rows, "columns": columns, "educators": educators, "issues": issues, "detail": detail,
            "evaluation_count": len(forms), "question_ids": qids}


def csv_bytes(rows, columns) -> bytes:
    """UTF-8 BOM, correct quoting and newline handling. Guard spreadsheet formulas."""
    output = io.StringIO(newline="")
    writer = csv.DictWriter(output, list(columns), extrasaction="ignore", lineterminator="\r\n")
    writer.writeheader()
    for row in rows:
        values = {}
        for field in columns:
            value = row.get(field, "")
            if isinstance(value, str) and value.lstrip().startswith(("=", "+", "-", "@")) and not re.fullmatch(r"-?[0-9]+(?:\.[0-9]+)?", value):
                value = "'" + value
            values[field] = value
        writer.writerow(values)
    data = output.getvalue().encode("utf-8-sig")
    if len(data) > MAX_OUTPUT_BYTES:
        raise OASISReportError("This CSV exceeds the 64 MiB output limit. Narrow the reporting dates or selected exports.")
    return data


def question_key_rows(prepared):
    rows = []
    for qid in question_columns(prepared):
        q = prepared["questions"].get(qid, {})
        text = q.get("question", QUESTION_TEXTS.get(qid, ""))
        note = ("Mean of duration-category codes, not a teaching-quality score and not weeks."
                if qid == "1286" else "Mean of exported Multiple Choice Value; blank/N/A ratings excluded.")
        rows.append({"mean_column": f"q{qid}_mean", "count_column": f"q{qid}_n", "question_id": qid,
                     "question": text, "choice_values_seen": "; ".join(f"{v or '(blank)'} = {lab}" for v, lab in q.get("choice_options", [])),
                     "note": note})
    for qid, text in COMMENT_TEXTS.items():
        rows.append({"mean_column": comment_kind(text), "count_column": "", "question_id": qid,
                     "question": text, "choice_values_seen": "", "note": "Nonblank comments combined, numbered; repeated text from different forms retained."})
    return rows


def build_report_downloads(prepared, summary, *, filter_description="All submitted evaluations in the selected exports"):
    if summary["issues"]:
        raise OASISReportError("Resolve the missing/duplicate record_id issues before downloading the final CSV.")
    if not summary["rows"]:
        raise OASISReportError("No submitted evaluations match the chosen filters.")
    main = csv_bytes(summary["rows"], summary["columns"])
    key_rows = question_key_rows(prepared)
    key_csv = csv_bytes(key_rows, ("mean_column", "count_column", "question_id", "question", "choice_values_seen", "note"))
    details = csv_bytes(summary["detail"], ("record_id", "educator_name", "question_id", "question", "mean_value", "response_count", "evaluation_count"))
    notes = [
        "OASIS EDUCATOR EVALUATION REPORT", filter_description,
        "One row per educator. The source Evaluator field is treated as the educator, as requested.",
        "record_id = the part before @ in Evaluator Email, unless a username override is explicitly entered.",
        "Evaluator Username and Evaluator External ID help link records but never automatically replace a missing email username.",
        "evaluation_count = distinct submitted Course ID / Evaluation / Form Record keys, not question rows or unique students.",
        f"Exact duplicate question responses removed: {prepared['duplicates_removed']}",
        f"Forms without Submit Date excluded: {prepared['unsubmitted_forms_excluded']}",
        "Conflicting copies of an evaluation/question stop processing; no hidden latest/first-copy selection.",
        "Question IDs and wording align form versions. Question Number is not used for scoring.",
        "q<ID>_mean averages numeric Multiple Choice Value. q<ID>_n is the number of scored responses used.",
        "Blank and non-scored N/A labels are excluded, not replaced by zero. No scored answers means blank mean and n=0.",
        "The duration question q1286 is a mean of category codes, not weeks and not a quality score. No overall composite score is generated.",
        "Comments are numbered in submission order. All nonblank comments, including literal n/a, are retained.",
        "Identical comment text from different forms is retained; repeated copies of the same form/question are counted once.",
        "No structured student names/IDs/emails are exported. Comment text is not anonymized and may itself identify someone.",
        "CSV cells are quoted correctly; spreadsheet-formula-like text is prefixed with an apostrophe for safety.",
        "A missing email remains flagged email_missing=YES after an explicit manual username is supplied; no email is fabricated.",
        "The original encrypted exports are never changed by this report. The downloadable CSV/ZIP is unencrypted; do not commit it publicly.",
    ]
    ignored = [q for q in prepared["questions"].values() if q["has_text"] and not comment_kind(q["question"]) and not q["has_choice"]]
    if ignored:
        notes.append("Additional text-only questions not requested in this report: " + "; ".join(q["question_id"] + " " + q["question"] for q in ignored))
    output = io.BytesIO()
    with ZipFile(output, "w", compression=ZIP_DEFLATED) as zf:
        zf.writestr(SUMMARY_FILENAME, main)
        zf.writestr("oasis_question_key.csv", key_csv)
        zf.writestr("oasis_educator_question_detail.csv", details)
        zf.writestr("OASIS_Report_Notes.txt", "\n".join(notes).encode("utf-8"))
    return {"csv": main, "question_key": key_csv, "zip": output.getvalue()}
