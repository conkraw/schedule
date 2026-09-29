"""Archive original OASIS CSV snapshots beside the OPDs, never mixed into them.

Each different byte-for-byte export is retained. An identical re-upload is a
no-op. The public filename contains course dates and a keyed, opaque identifier,
not student/provider names, free-text comments, or an unkeyed plaintext hash.
Fernet encrypts ONLY the original CSV bytes; decrypting returns the original CSV.
No evaluation scores, learner identities, or comments are analyzed by this module.
"""
from __future__ import annotations

import base64
import csv
import hashlib
import hmac
import io
import re
from datetime import date, datetime
from threading import RLock
from typing import Any
from urllib.parse import quote

from cryptography.fernet import InvalidToken

from schedule_app.services.opd_archive import GitHubOPDArchive, OPDArchiveError

OASIS_ARCHIVE_VERSION = 1
OASIS_SUBFOLDER = "oasis_evaluations"
OASIS_MAX_BYTES = 10 * 1024 * 1024
OASIS_MAX_ENCRYPTED_BYTES = 15 * 1024 * 1024
OASIS_MAX_FIELD_CHARS = 2 * 1024 * 1024
OASIS_MAX_ROWS = 250_000
OASIS_REQUIRED_COLUMNS = (
    "Course ID", "Start Date", "End Date", "Evaluator", "Evaluation",
    "Form Record", "Question ID", "Question",
)
# Filename dates are an organizational hint, NOT an overwrite/merge key.
OASIS_FILE_RE = re.compile(
    r"^OASIS_(?P<coverage>\d{4}-\d{2}-\d{2}_to_\d{4}-\d{2}-\d{2}|undated)"
    r"_(?P<export_id>[0-9a-f]{64})\.csv\.enc$"
)
_CSV_LOCK = RLock()
_ID_CONTEXT = b"schedule-app:oasis-evaluation-export-id:v1"


class OASISArchiveError(OPDArchiveError):
    """Safe messages only: never echo response bodies, row values, or secrets."""


def _course_date(text: str) -> date | None:
    text = text.strip()
    if not text:
        return None
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m-%d-%Y", "%Y-%m-%d %H:%M:%S"):
        try:
            parsed = datetime.strptime(text, fmt).date()
            return parsed if 1900 <= parsed.year <= 2100 else None
        except ValueError:
            pass
    return None


def inspect_oasis_csv(raw: bytes) -> dict[str, Any]:
    """Validate structure and return non-personal metadata without rewriting data.

    Blank/unreadable course dates do not prevent preserving the original; the
    archive uses an 'undated' filename if any row's course dates are uncertain.
    CSV rows represent question responses, not a count of distinct evaluations.
    """
    if not isinstance(raw, bytes) or not raw:
        raise OASISArchiveError("The OASIS CSV is empty. Upload the original exported CSV.")
    if len(raw) > OASIS_MAX_BYTES:
        raise OASISArchiveError("OASIS exports must be 10 MiB or smaller in this version.")
    if raw.startswith((b"PK\x03\x04", b"\xd0\xcf\x11\xe0", b"%PDF")):
        raise OASISArchiveError("Upload the OASIS comma-separated CSV, not an Excel workbook or PDF.")
    try:
        if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
            text, encoding = raw.decode("utf-16"), "UTF-16"
        else:
            try:
                text, encoding = raw.decode("utf-8-sig"), "UTF-8"
            except UnicodeDecodeError:
                text, encoding = raw.decode("cp1252"), "Windows-1252"
    except UnicodeError:
        raise OASISArchiveError("The CSV text could not be decoded. Re-export as a UTF-8 CSV.") from None
    if "\x00" in text:
        raise OASISArchiveError("The file contains binary/null characters. Re-export as a UTF-8 CSV.")
    # csv.field_size_limit is process-wide; hold the lock and restore its old value.
    with _CSV_LOCK:
        previous_limit = csv.field_size_limit()
        csv.field_size_limit(OASIS_MAX_FIELD_CHARS)
        try:
            reader = csv.reader(io.StringIO(text, newline=""), strict=True)
            headers = next(reader, None)
            if not headers:
                raise OASISArchiveError("The CSV has no header row.")
            normalized = [item.strip().lstrip("\ufeff") for item in headers]
            if (any(not item for item in normalized)
                    or len(set(normalized)) != len(normalized)):
                raise OASISArchiveError("The CSV contains blank or duplicate column headers.")
            missing = [item for item in OASIS_REQUIRED_COLUMNS if item not in normalized]
            if missing:
                # Only expected, hard-coded header names are shown.
                raise OASISArchiveError("This is not the expected OASIS evaluation export. Missing columns: "
                                       + ", ".join(missing) + ".")
            start_idx, end_idx = normalized.index("Start Date"), normalized.index("End Date")
            row_count = missing_dates = invalid_dates = reversed_dates = 0
            earliest = latest = None
            for row in reader:
                if not row or all(not item.strip() for item in row):
                    continue
                row_count += 1
                if row_count > OASIS_MAX_ROWS:
                    raise OASISArchiveError("The export exceeds this version's 250,000-row limit.")
                if len(row) != len(headers):
                    raise OASISArchiveError(
                        f"CSV record {row_count + 1} has {len(row)} fields; the header has {len(headers)}. "
                        "Re-export the CSV; nothing has been saved."
                    )
                start_text, end_text = row[start_idx].strip(), row[end_idx].strip()
                start, end = _course_date(start_text), _course_date(end_text)
                missing_dates += int(not start_text) + int(not end_text)
                invalid_dates += int(bool(start_text) and start is None) + int(bool(end_text) and end is None)
                reversed_dates += int(start is not None and end is not None and start > end)
                if start is not None:
                    earliest = min(earliest, start) if earliest else start
                if end is not None:
                    latest = max(latest, end) if latest else end
        except csv.Error:
            raise OASISArchiveError(
                "The CSV has invalid quoting or an oversized field. Re-export the original OASIS CSV."
            ) from None
        finally:
            csv.field_size_limit(previous_limit)
    if not row_count:
        raise OASISArchiveError("The export has column headers but no response rows. Nothing was saved.")
    complete_dates = (not (missing_dates or invalid_dates or reversed_dates)
                      and earliest is not None and latest is not None and earliest <= latest)
    coverage = f"{earliest.isoformat()}_to_{latest.isoformat()}" if complete_dates else "undated"
    return {
        "byte_count": len(raw), "row_count": row_count, "column_count": len(headers),
        "encoding": encoding, "coverage": coverage,
        "course_start": earliest.isoformat() if earliest else None,
        "course_end": latest.isoformat() if latest else None,
        "date_range_complete": bool(complete_dates),
        "missing_date_cells": missing_dates, "invalid_date_cells": invalid_dates,
        "reversed_date_rows": reversed_dates,
    }


def _file_parts(filename: str) -> dict[str, str]:
    match = OASIS_FILE_RE.fullmatch(filename) if isinstance(filename, str) else None
    if not match:
        raise OASISArchiveError("Select an OASIS export from this archive's saved-file list.")
    result = match.groupdict()
    if result["coverage"] != "undated":
        try:
            start, end = result["coverage"].split("_to_")
            if date.fromisoformat(start) > date.fromisoformat(end):
                raise ValueError()
        except ValueError:
            raise OASISArchiveError("The archived OASIS filename contains invalid course dates.") from None
    return result


def oasis_export_label(filename: str) -> str:
    entry = _file_parts(filename)
    coverage = entry["coverage"].replace("_to_", " to ")
    if coverage == "undated":
        coverage = "Course dates need review"
    return f"{coverage} | export {entry['export_id'][:12]}"


class GitHubOASISEvaluations:
    """Append-only CSV snapshots using the already-configured OPD archive client.

    Each successful save is read at its returned commit, decrypted, and compared
    byte-for-byte. No catalog is needed, so partial catalog updates cannot orphan
    exports. A retry after a failed verification safely recognizes the saved file.
    """

    def __init__(self, archive: GitHubOPDArchive):
        self.archive = archive
        self.config = archive.config
        self.cipher = self.config.cipher()
        self.folder = f"{self.config.folder}/{OASIS_SUBFOLDER}"

    def _head(self) -> str:
        try:
            return self.archive._head()
        except OPDArchiveError as exc:
            raise OASISArchiveError("OASIS archive could not be reached. " + str(exc)) from None

    def _call(self, method: str, route: str, **kwargs: Any) -> Any:
        try:
            return self.archive._request(method, route, **kwargs)
        except OPDArchiveError as exc:
            raise OASISArchiveError("OASIS archive operation was not confirmed. " + str(exc)) from None

    def _candidate_names(self, raw: bytes, details: dict[str, Any]) -> list[str]:
        names = []
        for encoded_key in (self.config.encryption_key, *self.config.previous_encryption_keys):
            # Domain separation: never use the Fernet encryption key directly as
            # the per-file identifier key. Keep unkeyed plaintext hashes private.
            key = base64.urlsafe_b64decode(encoded_key.encode("ascii"))
            id_key = hmac.new(key, _ID_CONTEXT, hashlib.sha256).digest()
            identity = hmac.new(id_key, raw, hashlib.sha256).hexdigest()
            filename = f"OASIS_{details['coverage']}_{identity}.csv.enc"
            if filename not in names:
                names.append(filename)
        return names

    def path_for(self, filename: str) -> str:
        _file_parts(filename)
        return f"{self.folder}/{filename}"

    def load(self, filename: str, *, commit: str | None = None,
             missing_ok: bool = False) -> dict[str, Any] | None:
        path = self.path_for(filename)
        route = "/contents/" + quote(path, safe="/")
        commit = commit or self._head()
        metadata = self._call("GET", route, params={"ref": commit}, missing_ok=True)
        if metadata is None:
            if missing_ok:
                return None
            raise OASISArchiveError("That OASIS export was not found. Refresh the saved export list.")
        if (not isinstance(metadata, dict) or metadata.get("type") != "file"
                or metadata.get("target") or metadata.get("submodule_git_url")
                or metadata.get("path", path) != path):
            raise OASISArchiveError("The OASIS archive path is not a regular file. Nothing was overwritten.")
        size = metadata.get("size")
        if type(size) is not int or not 0 < size <= OASIS_MAX_ENCRYPTED_BYTES:
            raise OASISArchiveError("The encrypted OASIS export has an invalid or unsupported size.")
        if metadata.get("encoding") == "base64" and metadata.get("content"):
            content = metadata["content"]
            if not isinstance(content, str) or len(content) > OASIS_MAX_ENCRYPTED_BYTES * 2:
                raise OASISArchiveError("GitHub returned invalid encoded OASIS contents.")
            try:
                token = base64.b64decode("".join(content.split()), validate=True)
            except (ValueError, TypeError):
                raise OASISArchiveError("GitHub returned invalid encoded OASIS contents.") from None
        else:
            # GitHub omits inline base64 for files >1 MB. Use the same API route
            # and pinned commit, not a public or expiring download URL.
            token = self._call("GET", route, params={"ref": commit}, raw=True)
        if not isinstance(token, bytes) or len(token) != size or len(token) > OASIS_MAX_ENCRYPTED_BYTES:
            raise OASISArchiveError("The OASIS download size could not be verified. Retry loading it.")
        sha = hashlib.sha1(b"blob " + str(len(token)).encode() + b"\0" + token).hexdigest()
        if sha != metadata.get("sha"):
            raise OASISArchiveError("The OASIS download did not match its GitHub file identifier.")
        try:
            raw = self.cipher.decrypt(token)  # no TTL; historical exports remain recoverable
        except (InvalidToken, ValueError):
            raise OASISArchiveError(
                "This OASIS export cannot be decrypted with the configured key(s), or its contents "
                "were altered. Keep/restore the correct key in Streamlit Secrets. Nothing was overwritten."
            ) from None
        details = inspect_oasis_csv(raw)
        if not any(hmac.compare_digest(filename, candidate)
                   for candidate in self._candidate_names(raw, details)):
            raise OASISArchiveError("The decrypted OASIS export does not match its archive identifier.")
        return {"filename": filename, "path": path, "commit": commit,
                "sha": sha, "details": details, "raw": raw}

    def save(self, raw: bytes) -> dict[str, Any]:
        details = inspect_oasis_csv(raw)  # reject invalid inputs before any GitHub operation
        names = self._candidate_names(raw, details)
        commit = self._head()
        for name in names:
            current = self.load(name, commit=commit, missing_ok=True)
            if current is None:
                continue
            if not hmac.compare_digest(current["raw"], raw):
                raise OASISArchiveError("The existing OASIS export could not be verified. It was not changed.")
            return {"action": "unchanged", "filename": name, "path": current["path"],
                    "sha": current["sha"], "details": details, "commit": commit}
        filename = names[0]
        path = self.path_for(filename)
        token = self.cipher.encrypt(raw)
        if len(token) > OASIS_MAX_ENCRYPTED_BYTES:
            raise OASISArchiveError("The encrypted export exceeds this version's archive size limit.")
        result = self._call("PUT", "/contents/" + quote(path, safe="/"), body={
            "message": "Archive encrypted OASIS evaluation export",
            "branch": self.config.branch,
            "content": base64.b64encode(token).decode("ascii"),
            # No sha: this operation may CREATE, never overwrite, a snapshot.
        })
        try:
            saved_commit = result["commit"]["sha"]
            if not isinstance(saved_commit, str) or not re.fullmatch(r"[0-9a-f]{40,64}", saved_commit):
                raise ValueError()
        except (KeyError, TypeError, ValueError):
            raise OASISArchiveError("GitHub did not confirm the OASIS save. Retry to verify the saved copy.") from None
        verified = self.load(filename, commit=saved_commit)
        if verified is None or not hmac.compare_digest(verified["raw"], raw):
            raise OASISArchiveError("The OASIS file could not be verified after saving. Retry; no success is confirmed.")
        return {"action": "created", "filename": filename, "path": path,
                "sha": verified["sha"], "details": details, "commit": saved_commit}

    def list_exports(self) -> dict[str, Any]:
        commit = self._head()
        response = self._call("GET", "/contents/" + quote(self.folder, safe="/"),
                              params={"ref": commit}, missing_ok=True)
        if response is None:
            return {"filenames": [], "commit": commit}
        if isinstance(response, dict) and response.get("type", "dir") != "dir":
            raise OASISArchiveError("The configured OASIS archive path is not a folder.")
        entries = response.get("entries") if isinstance(response, dict) else response
        if not isinstance(entries, list):
            raise OASISArchiveError("GitHub did not return a valid OASIS archive list.")
        if len(entries) >= 1000:
            raise OASISArchiveError("The OASIS directory listing limit was reached. A complete list cannot "
                                   "be shown; ask for an archive-pagination update before continuing.")
        filenames = []
        for entry in entries:
            if not isinstance(entry, dict):
                raise OASISArchiveError("GitHub returned an invalid OASIS directory entry.")
            name = entry.get("name", "")
            if (entry.get("type") != "file" or not isinstance(name, str)
                    or not OASIS_FILE_RE.fullmatch(name)):
                continue
            _file_parts(name)
            expected_path = f"{self.folder}/{name}"
            if entry.get("path", expected_path) != expected_path or entry.get("submodule_git_url"):
                raise OASISArchiveError("An OASIS export is not located in the expected archive folder.")
            filenames.append(name)
        if len(filenames) != len(set(filenames)):
            raise OASISArchiveError("GitHub returned duplicate OASIS filenames. Refresh the archive list.")
        def order(name: str):
            parts = _file_parts(name)
            coverage = parts["coverage"]
            if coverage == "undated":
                return (0, "", "", parts["export_id"])
            start, end = coverage.split("_to_")
            return (1, end, start, parts["export_id"])
        return {"filenames": sorted(filenames, key=order, reverse=True), "commit": commit}
