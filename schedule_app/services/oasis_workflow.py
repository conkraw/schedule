"""Combined OASIS workflow: archived originals -> one encrypted summary CSV.

All snapshots are a cumulative union, not a concatenation of educator summaries.
Identical Course ID / Evaluation / Form Record / Question ID responses count once.
Conflicting versions retain the existing fail-visible policy; no answer is guessed.
No ZIP, Word, question-key, or detail report is produced by this workflow.
"""
from __future__ import annotations

import base64
import csv
import hashlib
import hmac
import io
import json
import re
from dataclasses import dataclass
from datetime import date
from urllib.parse import quote

from cryptography.fernet import InvalidToken

from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.oasis_evaluations import (
    GitHubOASISEvaluations, _CSV_LOCK,
)
from schedule_app.services.oasis_educator_reports import (
    OASISReportError, prepare_reports, educator_summary, csv_bytes,
    MAX_EXPORTS, MAX_TOTAL_BYTES, validate_username,
)
from schedule_app.services.oasis_educator_usernames import GitHubOASISUsernames
from schedule_app.services.reporting_periods import ReportingPeriod

WORKFLOW_VERSION = 1
OUTPUT_FOLDER = "oasis_reports"
MAX_SUMMARY_BYTES = 10 * 1024 * 1024
MAX_SUMMARY_CIPHER = 15 * 1024 * 1024
SUMMARY_RE = re.compile(r"OASIS_Summary_(\d{4}-\d{2}-\d{2})_to_(\d{4}-\d{2}-\d{2})_([0-9a-f]{20})\.csv\.enc")
PERIOD_COLUMNS = ("academic_year", "report_start_date", "report_end_date", "date_basis")


@dataclass(frozen=True)
class OASISOutputScope:
    period: ReportingPeriod
    courses: tuple[str, ...]
    evaluation_types: tuple[str, ...]

    def __post_init__(self):
        if not isinstance(self.period, ReportingPeriod):
            raise OASISReportError("Apply a valid reporting period first.")
        for field in ("courses", "evaluation_types"):
            values = getattr(self, field)
            if (not isinstance(values, (tuple, list)) or not values or
                    any(not isinstance(v, str) or not v.strip() for v in values)):
                raise OASISReportError("Select at least one course and evaluation form.")
            object.__setattr__(self, field, tuple(sorted(set(values))))

    @property
    def filename(self) -> str:
        # Label changes replace the same period's output, not a second copy.
        # Scope suffix separates courses/forms even when reporting dates match.
        material = json.dumps([self.courses, self.evaluation_types], ensure_ascii=True,
                              separators=(",", ":")).encode()
        scope_id = hashlib.sha256(material).hexdigest()[:20]
        return (f"OASIS_Summary_{self.period.start.isoformat()}_to_"
                f"{self.period.end.isoformat()}_{scope_id}.csv.enc")

    def filters(self) -> dict:
        return {"courses": self.courses, "evaluation_types": self.evaluation_types,
                "start_date": self.period.start, "end_date": self.period.end,
                "date_field": "Submit Date"}


def load_cumulative_evaluations(client: GitHubOASISEvaluations, progress=None) -> dict:
    """Read every saved source at one commit. Never silently choose just the new CSV."""
    snapshot = client.list_exports()
    filenames = snapshot["filenames"]
    if not filenames:
        raise OASISReportError("No OASIS exports are archived yet. Upload your original CSV above.")
    if len(filenames) > MAX_EXPORTS:
        raise OASISReportError(f"The cumulative archive exceeds {MAX_EXPORTS} source snapshots. "
                               "An archive-compaction update is needed; no partial output was created.")
    exports, total = [], 0
    for index, filename in enumerate(filenames, 1):
        loaded = client.load(filename, commit=snapshot["commit"])
        total += len(loaded["raw"])
        if total > MAX_TOTAL_BYTES:
            raise OASISReportError("The cumulative OASIS sources exceed 64 MiB. "
                                   "No sources were omitted and no partial output was created.")
        exports.append((filename, loaded["raw"]))
        if progress:
            progress(index, len(filenames))
    prepared = prepare_reports(exports)
    prepared["archive_filenames"] = sorted(filenames)
    prepared["archive_commit"] = snapshot["commit"]
    prepared["archive_scope"] = client.config.signature()
    prepared["workflow_version"] = WORKFLOW_VERSION
    # Only normalized form data (not the Student / Who Completed columns) persist.
    return prepared


def make_period_summary(prepared: dict, scope: OASISOutputScope, overrides=None) -> dict:
    if prepared.get("workflow_version") != WORKFLOW_VERSION:
        raise OASISReportError("Refresh the cumulative OASIS data after this app update.")
    summary = educator_summary(prepared, overrides, **scope.filters())
    # Preserve every existing report column, with period metadata appended.
    summary["columns"] = list(summary["columns"]) + list(PERIOD_COLUMNS)
    for row in summary["rows"]:
        row.update(academic_year=scope.period.label,
                   report_start_date=scope.period.start.isoformat(),
                   report_end_date=scope.period.end.isoformat(), date_basis="Submit Date")
    return summary


def summary_csv(summary: dict) -> bytes:
    if summary["issues"]:
        raise OASISReportError("Resolve the missing or duplicate educator usernames before the output CSV is saved.")
    if not summary["rows"]:
        raise OASISReportError("No submitted evaluations match this period. No empty output replaces a saved report.")
    raw = csv_bytes(summary["rows"], summary["columns"])
    if len(raw) > MAX_SUMMARY_BYTES:
        raise OASISReportError("The encrypted-output workflow supports a summary CSV up to 10 MiB. Narrow the reporting period.")
    return raw


def _filename_dates(filename: str) -> tuple[date, date]:
    match = SUMMARY_RE.fullmatch(str(filename))
    if not match:
        raise OASISReportError("Choose a recognized OASIS summary CSV from the saved-output list.")
    try:
        start, end = date.fromisoformat(match[1]), date.fromisoformat(match[2])
    except ValueError:
        raise OASISReportError("The saved OASIS summary filename has invalid dates.") from None
    if start > end:
        raise OASISReportError("The saved OASIS summary filename has a reversed date range.")
    return start, end


def inspect_summary_csv(raw: bytes, filename: str) -> dict:
    """Validate decrypted CSV without exposing comments or educator identities in errors."""
    start, end = _filename_dates(filename)
    if not isinstance(raw, bytes) or not raw or len(raw) > MAX_SUMMARY_BYTES:
        raise OASISReportError("The OASIS summary CSV is empty or exceeds 10 MiB.")
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeError:
        raise OASISReportError("The decrypted OASIS summary is not a UTF-8 CSV.") from None
    with _CSV_LOCK:
        old_limit = csv.field_size_limit()
        # A combined comment can be larger than any one source response.
        csv.field_size_limit(MAX_SUMMARY_BYTES)
        try:
            reader = csv.DictReader(io.StringIO(text, newline=""), strict=True)
            fields = reader.fieldnames or []
            required = {"record_id", "educator_name", "evaluation_count", *PERIOD_COLUMNS}
            if not fields or fields[0] != "record_id" or len(fields) != len(set(fields)) or not required.issubset(fields):
                raise OASISReportError("The decrypted OASIS summary has unsupported columns.")
            count, evaluations, ids, labels = 0, 0, set(), set()
            for row in reader:
                if None in row or any(value is None for value in row.values()):
                    raise OASISReportError("The decrypted OASIS summary has malformed CSV rows.")
                username = validate_username(row["record_id"])
                if username != row["record_id"] or username in ids or not row["educator_name"].strip():
                    raise OASISReportError("The OASIS summary contains invalid/duplicate educator usernames.")
                ids.add(username)
                if (row["report_start_date"] != start.isoformat() or row["report_end_date"] != end.isoformat()
                        or row["date_basis"] != "Submit Date" or not row["academic_year"].strip()
                        or not re.fullmatch(r"[1-9][0-9]*", row["evaluation_count"])):
                    raise OASISReportError("The OASIS summary has inconsistent dates or evaluation counts.")
                labels.add(row["academic_year"])
                count += 1
                evaluations += int(row["evaluation_count"])
            if count == 0 or len(labels) != 1:
                raise OASISReportError("The OASIS summary has no educators or inconsistent report labels.")
            return {"educator_count": count, "evaluation_count": evaluations,
                    "label": next(iter(labels)), "start_date": start, "end_date": end}
        except (csv.Error, UnicodeError, OverflowError):
            raise OASISReportError("The decrypted OASIS summary could not be parsed as a CSV.") from None
        finally:
            csv.field_size_limit(old_limit)


class GitHubOASISSummaries:
    """Only summary CSV ciphertext is written under oasis_reports.

    Same reporting dates + courses + forms replace one current output. Original
    sources, OPDs, and username/date catalogs are never edited by this class.
    """
    def __init__(self, archive):
        self.archive = archive
        self.config = archive.config
        self.cipher = self.config.cipher()
        self.folder = f"{self.config.folder}/{OUTPUT_FOLDER}"

    def path_for(self, filename: str) -> str:
        _filename_dates(filename)
        return f"{self.folder}/{filename}"

    def load(self, filename: str, *, commit=None, missing_ok=False) -> dict | None:
        path = self.path_for(filename)
        commit = commit or self.archive._head()
        route = "/contents/" + quote(path, safe="/")
        meta = self.archive._request("GET", route, params={"ref": commit}, missing_ok=True)
        if meta is None:
            if missing_ok:
                return None
            raise OASISReportError("That output CSV was not found. Refresh the saved-output list.")
        if (not isinstance(meta, dict) or meta.get("type") != "file" or meta.get("target")
                or meta.get("submodule_git_url") or meta.get("path", path) != path):
            raise OASISReportError("The OASIS output path is not a regular file. Nothing was overwritten.")
        size = meta.get("size")
        if type(size) is not int or not 0 < size <= MAX_SUMMARY_CIPHER:
            raise OASISReportError("The encrypted output CSV has an unsupported size.")
        if meta.get("encoding") == "base64" and meta.get("content"):
            content = meta["content"]
            if not isinstance(content, str) or len(content) > MAX_SUMMARY_CIPHER * 2:
                raise OASISReportError("GitHub returned invalid encrypted summary contents.")
            try:
                token = base64.b64decode("".join(content.split()), validate=True)
            except (ValueError, TypeError):
                raise OASISReportError("GitHub returned invalid encrypted summary contents.") from None
        else:
            token = self.archive._request("GET", route, params={"ref": commit}, raw=True)
        if not isinstance(token, bytes) or len(token) != size:
            raise OASISReportError("The encrypted output CSV size could not be verified.")
        sha = hashlib.sha1(b"blob " + str(len(token)).encode() + b"\0" + token).hexdigest()
        if sha != meta.get("sha"):
            raise OASISReportError("The encrypted output CSV did not match its GitHub identifier.")
        try:
            raw = self.cipher.decrypt(token)
        except (InvalidToken, ValueError):
            raise OASISReportError("The saved output could not be decrypted. Keep/restore the correct encryption key; nothing was overwritten.") from None
        details = inspect_summary_csv(raw, filename)
        return {"raw": raw, "sha": sha, "path": path, "filename": filename,
                "commit": commit, "details": details}

    def list_outputs(self) -> dict:
        commit = self.archive._head()
        result = self.archive._request("GET", "/contents/" + quote(self.folder, safe="/"),
                                       params={"ref": commit}, missing_ok=True)
        if result is None:
            return {"filenames": [], "commit": commit}
        if isinstance(result, dict) and result.get("type", "dir") != "dir":
            raise OASISReportError("The OASIS output location is not a directory.")
        entries = result.get("entries") if isinstance(result, dict) else result
        if not isinstance(entries, list) or len(entries) >= 1000:
            raise OASISReportError("A complete OASIS output listing could not be obtained. No partial list is shown.")
        names = []
        for item in entries:
            if not isinstance(item, dict):
                raise OASISReportError("GitHub returned an invalid output listing.")
            name = item.get("name", "")
            if item.get("type") == "file" and isinstance(name, str) and SUMMARY_RE.fullmatch(name):
                self.path_for(name)
                names.append(name)
        return {"filenames": sorted(set(names), reverse=True), "commit": commit}

    def _check_sources(self, prepared, username_catalog, *, commit=None) -> dict:
        if prepared.get("archive_scope") != self.config.signature():
            raise OASISReportError("The archive settings changed. Refresh cumulative evaluations before saving the report.")
        sources = GitHubOASISEvaluations(self.archive).list_exports(commit=commit)
        if sorted(sources["filenames"]) != prepared.get("archive_filenames"):
            raise OASISReportError("The OASIS source archive changed in another session. Refresh cumulative evaluations; no output is confirmed current.")
        names = GitHubOASISUsernames(self.archive).load(commit=sources["commit"])
        if (not isinstance(username_catalog, dict) or username_catalog.get("scope") != self.config.signature()
                or names["sha"] != username_catalog.get("sha")):
            raise OASISReportError("Educator usernames changed in another session. Refresh cumulative evaluations/usernames before saving the output.")
        return sources

    def save(self, scope: OASISOutputScope, raw: bytes, *, prepared: dict, username_catalog: dict) -> dict:
        details = inspect_summary_csv(raw, scope.filename)
        expected_label = scope.period.label
        if (expected_label.lstrip().startswith(("=", "+", "-", "@"))
                and not re.fullmatch(r"-?[0-9]+(?:\.[0-9]+)?", expected_label)):
            expected_label = "'" + expected_label
        if details["label"] != expected_label:
            raise OASISReportError("The output CSV label does not match the applied reporting period.")
        snapshot = self._check_sources(prepared, username_catalog)
        current = self.load(scope.filename, commit=snapshot["commit"], missing_ok=True)
        path = self.path_for(scope.filename)
        if current is not None and hmac.compare_digest(current["raw"], raw):
            return {**current, "action": "unchanged", "details": details,
                    "source_count": len(prepared["archive_filenames"])}
        token = self.cipher.encrypt(raw)
        if len(token) > MAX_SUMMARY_CIPHER:
            raise OASISReportError("The encrypted summary exceeds the supported size.")
        body = {"message": "Update encrypted OASIS educator summary CSV", "branch": self.config.branch,
                "content": base64.b64encode(token).decode("ascii")}
        if current is not None:
            body["sha"] = current["sha"]
        result = self.archive._request("PUT", "/contents/" + quote(path, safe="/"), body=body)
        try:
            commit = result["commit"]["sha"]
            if not isinstance(commit, str) or not re.fullmatch(r"[0-9a-f]{40,64}", commit):
                raise ValueError()
        except (KeyError, ValueError, TypeError):
            raise OASISReportError("GitHub did not confirm the output save. Retry to verify the existing saved file.") from None
        verified = self.load(scope.filename, commit=commit)
        if not hmac.compare_digest(verified["raw"], raw):
            raise OASISReportError("The saved output CSV did not match the generated report. No success is confirmed.")
        # A different source can be added while a Contents API write is in flight.
        # Verify inputs at the actual save commit before claiming a current result.
        self._check_sources(prepared, username_catalog, commit=commit)
        return {**verified, "action": "updated" if current else "created",
                "source_count": len(prepared["archive_filenames"])}
