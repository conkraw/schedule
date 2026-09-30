"""Read saved OASIS summary CSVs and explicitly join them to teaching preceptors.

No fuzzy name matching, averaging of averages, or inclusion of OASIS-only people.
A summary is already aggregated: exact period boundaries must match. OPD dates
measure assignments; OASIS dates measure submissions, not those same encounters.
"""
from __future__ import annotations

import csv
import hashlib
import io
import json
import re
from datetime import date
from decimal import Decimal, InvalidOperation

from schedule_app.services.oasis_evaluations import _CSV_LOCK
from schedule_app.services.oasis_educator_reports import (
    OASISReportError, QUESTION_TEXTS, name_key, question_key, validate_username,
)
from schedule_app.services.oasis_workflow import GitHubOASISSummaries, inspect_summary_csv, MAX_SUMMARY_BYTES
from schedule_app.services.preceptor_oasis_links import GitHubPreceptorOASISLinks, period_key
from schedule_app.services.reporting_periods import teaching_report_bounds, teaching_report_label

TEACHING_OASIS_REPORT_VERSION = 1


def word_text(value: str) -> str:
    """Remove only the CSV writer's formula-safety apostrophe; never interpret markup."""
    value = str(value or "")
    if value.startswith("'") and value[1:].lstrip().startswith(("=", "+", "-", "@")):
        return value[1:]
    return value


def _word_compatible(value: str) -> bool:
    return not any((ord(c) < 32 and c not in "\n\r\t") or 0xD800 <= ord(c) <= 0xDFFF
                   or c in "\ufffe\uffff" for c in value)


def parse_saved_summary(loaded: dict) -> dict:
    """Read and validate an existing encrypted-output receipt. No names/comments in errors."""
    raw, filename = loaded["raw"], loaded["filename"]
    details = inspect_summary_csv(raw, filename)
    with _CSV_LOCK:
        previous = csv.field_size_limit()
        csv.field_size_limit(MAX_SUMMARY_BYTES)
        try:
            reader = csv.DictReader(io.StringIO(raw.decode("utf-8-sig"), newline=""), strict=True)
            fields = reader.fieldnames or []
            required = {"strengths_comments", "areas_for_improvement_comments"}
            if not required.issubset(fields):
                raise OASISReportError("The saved OASIS summary lacks the combined-comment columns. Rebuild it in OASIS Evaluations.")
            qids = [m[1] for field in fields if (m := re.fullmatch(r"q([0-9]+)_mean", field))]
            count_ids = [m[1] for field in fields if (m := re.fullmatch(r"q([0-9]+)_n", field))]
            if set(count_ids) != set(qids):
                raise OASISReportError("A saved question average is missing its response count (or vice versa). Rebuild the OASIS output.")
            rows = list(reader)
        except (csv.Error, UnicodeError, OverflowError):
            raise OASISReportError("The saved OASIS summary could not be read. Rebuild it before linking.") from None
        finally:
            csv.field_size_limit(previous)
    questions = {}
    for qid in qids:
        col = f"q{qid}_question"
        if col in fields:
            variants = {word_text(row[col]).strip() for row in rows}
            if len(variants) != 1 or not next(iter(variants)):
                raise OASISReportError(f"Question {qid} has missing or inconsistent wording in the saved summary. Rebuild it in OASIS Evaluations.")
            text = next(iter(variants))
            if qid in QUESTION_TEXTS and question_key(text) != question_key(QUESTION_TEXTS[qid]):
                raise OASISReportError(f"Question {qid} has changed wording. Review the OASIS question mapping before linking.")
        else:
            # Backward compatibility is limited to verified wording from the
            # user's supplied export. Never invent a label for an unknown ID.
            text = QUESTION_TEXTS.get(qid, "")
            if not text:
                raise OASISReportError(f"Question {qid} has no full wording in this older summary. Open OASIS Evaluations and rebuild this period with the updated app.")
        if not _word_compatible(text):
            raise OASISReportError(f"Question {qid} contains text that cannot be written to Word. Correct the source and rebuild.")
        questions[qid] = text
    by_id = {}
    for row in rows:
        rid = validate_username(row["record_id"])
        count = int(row["evaluation_count"])
        scores = []
        for qid in qids:
            n_value, mean_value = row[f"q{qid}_n"], row[f"q{qid}_mean"]
            if not re.fullmatch(r"0|[1-9][0-9]*", n_value) or int(n_value) > count:
                raise OASISReportError(f"Saved OASIS response count is invalid for username {rid}, question {qid}. Rebuild the OASIS summary.")
            n = int(n_value)
            if n == 0:
                if mean_value.strip():
                    raise OASISReportError(f"Username {rid}, question {qid} has a mean but no scored responses. Rebuild the OASIS summary.")
                mean = "Not scored"
            else:
                try:
                    number = Decimal(mean_value)
                    if not number.is_finite() or abs(number) > Decimal("1000000"):
                        raise ValueError()
                except (InvalidOperation, ValueError):
                    raise OASISReportError(f"Saved OASIS average is invalid for username {rid}, question {qid}. Rebuild the summary.") from None
                mean = format(number.quantize(Decimal("0.01")), ".2f")
            scores.append({"question": questions[qid], "mean": mean, "response_count": n,
                           "duration_category": qid == "1286"})
        texts = {key: word_text(row[key]) for key in ("educator_name", "strengths_comments", "areas_for_improvement_comments")}
        if any(not _word_compatible(text) for text in texts.values()):
            raise OASISReportError(f"Saved comments or educator text for username {rid} contain unsupported Word characters. Correct the source and rebuild.")
        by_id[rid] = {"record_id": rid, "evaluation_count": count, "questions": scores, **texts}
    return {"rows_by_id": by_id, "details": details, "filename": filename,
            "sha": loaded["sha"], "commit": loaded["commit"]}


def active_periods(scan, selected_years):
    years = {int(y) for y in selected_years}
    present = {row["academic_start_year"] for row in scan["monthly"]
               if row["academic_start_year"] in years and int(row["no_of_shifts"]) > 0}
    return [(year, *teaching_report_bounds(scan, year), teaching_report_label(scan, year))
            for year in sorted(present)]


def join_feedback(scan, selected_years, catalog, summaries_by_year, *, allow_missing_summaries=False):
    """Link a report-period summary to only explicitly mapped active named preceptors."""
    periods, status = {}, []
    review = {name_key(n) for n in scan.get("unresolved_preceptor_labels", [])}
    for year, start, end, label in active_periods(scan, selected_years):
        summary = summaries_by_year.get(year)
        if summary is None and allow_missing_summaries:
            periods[str(year)] = {"start_date": start.isoformat(), "end_date": end.isoformat(),
                                 "oasis_label": label, "summary_filename": "Not linked", "summary_sha": "",
                                 "snapshot": "", "preceptors": {}}
            for name in sorted({r["preceptor_name"] for r in scan["monthly"]
                                if r["academic_start_year"] == year and int(r["no_of_shifts"]) > 0}, key=name_key):
                status.append({"preceptor_name": name, "academic_year": label,
                    "record_id": catalog["entries"].get(name_key(name), {}).get("record_id", ""),
                    "oasis_educator_name": "", "evaluation_count": "",
                    "status": "No exact-date OASIS summary linked; teaching-only report"})
            continue
        if summary is None:
            raise OASISReportError(f"Choose and save an OASIS summary for {start} through {end}, or turn off linked evaluations.")
        if (summary["details"]["start_date"], summary["details"]["end_date"]) != (start, end):
            raise OASISReportError("OASIS summary dates do not exactly match the teaching report. Aggregated averages cannot be filtered to different dates. Rebuild OASIS for the same boundaries.")
        binding = catalog["report_links"].get(period_key(start, end), {})
        if binding.get("summary_filename") != summary["filename"]:
            raise OASISReportError("Save the selected OASIS summary link for this period before generating reports.")
        names = sorted({r["preceptor_name"] for r in scan["monthly"]
                        if r["academic_start_year"] == year and int(r["no_of_shifts"]) > 0}, key=name_key)
        matched = {}
        for name in names:
            key = name_key(name)
            entry = catalog["entries"].get(key)
            rid = entry["record_id"] if entry else ""
            feedback = summary["rows_by_id"].get(rid) if rid and key not in review else None
            if key in review:
                outcome = "Provider label needs review; teaching-only report"
            elif not rid:
                outcome = "No username assigned; teaching-only report"
            elif feedback is None:
                outcome = "Username not in selected OASIS summary; teaching-only report"
            else:
                outcome = "Linked evaluations will be included"
                matched[key] = feedback
            status.append({"preceptor_name": name, "academic_year": label, "record_id": rid,
                           "oasis_educator_name": feedback["educator_name"] if feedback else "",
                           "evaluation_count": feedback["evaluation_count"] if feedback else "", "status": outcome})
        periods[str(year)] = {"start_date": start.isoformat(), "end_date": end.isoformat(),
                             "oasis_label": word_text(summary["details"]["label"]),
                             "summary_filename": summary["filename"], "summary_sha": summary["sha"],
                             "snapshot": summary["commit"], "preceptors": matched}
    return {"version": TEACHING_OASIS_REPORT_VERSION, "mapping_sha": catalog.get("sha"),
            "periods": periods, "status": status}


def feedback_signature(bundle):
    if bundle is None:
        return None
    # Includes usernames and values. Do not put credentials, raw comments or names in filenames/logs.
    material = {**bundle, "periods": {key: {k: v for k, v in value.items() if k != "snapshot"}
                                     for key, value in bundle["periods"].items()}}
    return hashlib.sha256(json.dumps(material, sort_keys=True, ensure_ascii=True).encode()).hexdigest()


def feedback_for_preceptor(bundle, scan, name, year):
    if bundle is None:
        return None
    if not isinstance(bundle, dict) or bundle.get("version") != TEACHING_OASIS_REPORT_VERSION:
        raise OASISReportError("Refresh teaching/OASIS links after this app update.")
    group = bundle.get("periods", {}).get(str(year))
    if group is None:
        raise OASISReportError("Linked OASIS feedback is missing this reporting period. Refresh the saved links.")
    start, end = teaching_report_bounds(scan, year)
    if (group["start_date"], group["end_date"]) != (start.isoformat(), end.isoformat()):
        raise OASISReportError("Linked evaluations belong to different dates. Refresh the selected OASIS summary.")
    record = group["preceptors"].get(name_key(name))
    if record is None:
        return None
    return {**record, **{k: v for k, v in group.items() if k != "preceptors"}}


def load_feedback_bundle(archive, scan, years, expected_catalog, expected_summaries, *, allow_missing_summaries=False):
    """Recheck the saved links and each selected summary at ONE current commit.

    Updated summaries/maps require a refresh so the displayed match preview and
    generated documents agree. No OPD or OASIS file is written during this read.
    """
    commit = archive._head()
    catalog = GitHubPreceptorOASISLinks(archive).load(commit=commit)
    if (expected_catalog.get("scope") != archive.config.signature()
            or catalog["sha"] != expected_catalog.get("sha")):
        raise OASISReportError("Teaching/OASIS links changed. Click Refresh links and OASIS summaries, review the matches, and generate again.")
    service = GitHubOASISSummaries(archive)
    summaries = {}
    for year, start, end, _ in active_periods(scan, years):
        binding = catalog["report_links"].get(period_key(start, end))
        if allow_missing_summaries and year not in expected_summaries:
            continue
        if not binding:
            raise OASISReportError(f"No saved OASIS summary link for {start} through {end}.")
        loaded = service.load(binding["summary_filename"], commit=commit)
        expected = expected_summaries.get(year, {})
        if expected.get("sha") != loaded["sha"] or expected.get("filename") != loaded["filename"]:
            raise OASISReportError("A saved OASIS summary was updated. Click Refresh links and OASIS summaries to use the new evaluations.")
        summaries[year] = parse_saved_summary(loaded)
    return join_feedback(scan, years, catalog, summaries, allow_missing_summaries=allow_missing_summaries)
