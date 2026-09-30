"""Explicit teaching-preceptor -> OASIS record_id links, encrypted in GitHub.

One catalog also remembers the summary selected for each exact reporting period.
This never rewrites source OPDs, OASIS originals, or OASIS username overrides.
"""
from __future__ import annotations

import base64
import hashlib
import json
from datetime import date, datetime, timezone
from urllib.parse import quote

from cryptography.fernet import InvalidToken
from schedule_app.services.oasis_educator_reports import OASISReportError, validate_username, name_key
from schedule_app.services.oasis_workflow import _filename_dates

FILENAME = "preceptor_oasis_links.json.enc"
KIND = "teaching_preceptor_oasis_links"
LINK_VERSION = 1
MAX_PLAIN = 256 * 1024
MAX_CIPHER = 384 * 1024


def period_key(start: date, end: date) -> str:
    if type(start) is not date or type(end) is not date or start > end:
        raise OASISReportError("Select a valid teaching-report date range before linking an OASIS summary.")
    return f"{start.isoformat()}_to_{end.isoformat()}"


def _empty():
    return {"kind": KIND, "schema_version": LINK_VERSION, "entries": {}, "report_links": {}}


def _timestamp():
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _valid_timestamp(value):
    try:
        datetime.strptime(value, "%Y-%m-%dT%H:%M:%SZ")
    except (TypeError, ValueError):
        raise OASISReportError("Saved teaching/OASIS links contain an invalid update date.") from None


def _valid_name(value):
    if (not isinstance(value, str) or not value.strip() or len(value) > 256
            or any(ord(c) < 32 or ord(c) == 127 or 0xD800 <= ord(c) <= 0xDFFF or c in "\ufffe\uffff" for c in value)):
        raise OASISReportError("Choose a valid individual preceptor name from the teaching summary.")
    return value.strip()


def _validated(data):
    if (not isinstance(data, dict) or set(data) != set(_empty())
            or data.get("kind") != KIND or type(data.get("schema_version")) is not int
            or data["schema_version"] != LINK_VERSION
            or not isinstance(data.get("entries"), dict) or len(data["entries"]) > 2000
            or not isinstance(data.get("report_links"), dict) or len(data["report_links"]) > 1000):
        raise OASISReportError("Saved teaching/OASIS links have an unsupported format. Nothing was changed.")
    usernames = set()
    for key, entry in data["entries"].items():
        if not isinstance(entry, dict) or set(entry) != {"preceptor_name", "record_id", "updated_at"}:
            raise OASISReportError("Saved teaching/OASIS links contain an invalid preceptor entry.")
        name = _valid_name(entry["preceptor_name"])
        rid = validate_username(entry["record_id"])
        if key != name_key(name) or rid != entry["record_id"] or rid in usernames:
            raise OASISReportError("Saved teaching/OASIS links contain invalid or duplicate usernames. Nothing was changed.")
        usernames.add(rid)
        _valid_timestamp(entry["updated_at"])
    for key, entry in data["report_links"].items():
        if not isinstance(entry, dict) or set(entry) != {"start_date", "end_date", "summary_filename", "updated_at"}:
            raise OASISReportError("Saved teaching/OASIS links contain an invalid reporting-period entry.")
        try:
            start, end = date.fromisoformat(entry["start_date"]), date.fromisoformat(entry["end_date"])
        except (ValueError, TypeError):
            raise OASISReportError("Saved teaching/OASIS links contain invalid reporting dates.") from None
        if period_key(start, end) != key or _filename_dates(entry["summary_filename"]) != (start, end):
            raise OASISReportError("A saved OASIS selection does not match its teaching-report dates.")
        _valid_timestamp(entry["updated_at"])
    return data


class GitHubPreceptorOASISLinks:
    """Verified ciphertext-only catalog writes with stale-edit protection."""
    def __init__(self, archive):
        self.archive, self.config = archive, archive.config
        self.cipher = self.config.cipher()
        self.path = f"{self.config.folder}/{FILENAME}"
        self.route = "/contents/" + quote(self.path, safe="/")

    def load(self, *, commit=None):
        commit = commit or self.archive._head()
        meta = self.archive._request("GET", self.route, params={"ref": commit}, missing_ok=True)
        if meta is None:
            return {**_empty(), "sha": None, "commit": commit, "scope": self.config.signature()}
        if (not isinstance(meta, dict) or meta.get("type") != "file" or meta.get("target")
                or meta.get("submodule_git_url") or meta.get("path", self.path) != self.path
                or meta.get("encoding") != "base64" or not isinstance(meta.get("content"), str)):
            raise OASISReportError("The teaching/OASIS link path is not a regular encrypted catalog. Nothing was changed.")
        if (type(meta.get("size")) is not int or not 0 < meta["size"] <= MAX_CIPHER
                or len(meta["content"]) > MAX_CIPHER * 2):
            raise OASISReportError("The teaching/OASIS link catalog exceeds its size limit.")
        try:
            token = base64.b64decode("".join(meta["content"].split()), validate=True)
        except (ValueError, TypeError):
            raise OASISReportError("The encrypted teaching/OASIS link catalog could not be read.") from None
        sha = hashlib.sha1(b"blob " + str(len(token)).encode() + b"\0" + token).hexdigest()
        if len(token) != meta["size"] or sha != meta.get("sha"):
            raise OASISReportError("The teaching/OASIS link download could not be verified. Refresh and retry.")
        try:
            raw = self.cipher.decrypt(token)
        except (InvalidToken, ValueError):
            raise OASISReportError("Teaching/OASIS links could not be decrypted. Restore the correct key; nothing was overwritten.") from None
        if len(raw) > MAX_PLAIN:
            raise OASISReportError("The decrypted teaching/OASIS link catalog is too large.")
        try:
            data = _validated(json.loads(raw.decode("utf-8")))
        except (ValueError, UnicodeError, RecursionError):
            raise OASISReportError("The decrypted teaching/OASIS link catalog is invalid. Nothing was changed.") from None
        return {**data, "sha": sha, "commit": commit, "scope": self.config.signature()}

    def _current(self, expected):
        if not isinstance(expected, dict) or expected.get("scope") != self.config.signature() or "sha" not in expected:
            raise OASISReportError("Load/refresh teaching/OASIS links before making changes.")
        current = self.load()
        if current["sha"] != expected["sha"]:
            raise OASISReportError("Teaching/OASIS links changed in another session. Refresh links, review the current values, and retry. Nothing was overwritten.")
        return current

    def _write(self, data, current):
        data = _validated(data)
        if all(data[k] == current[k] for k in data):
            return current
        raw = json.dumps(data, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode()
        if len(raw) > MAX_PLAIN:
            raise OASISReportError("Too many saved teaching/OASIS links; nothing was written.")
        token = self.cipher.encrypt(raw)
        body = {"message": "Update encrypted teaching OASIS links", "branch": self.config.branch,
                "content": base64.b64encode(token).decode("ascii")}
        if current["sha"] is not None:
            body["sha"] = current["sha"]
        result = self.archive._request("PUT", self.route, body=body)
        try:
            commit = result["commit"]["sha"]
            if not isinstance(commit, str) or not commit:
                raise ValueError()
        except (KeyError, TypeError, ValueError):
            raise OASISReportError("GitHub did not confirm the teaching/OASIS link save. Refresh to verify before retrying.") from None
        verified = self.load(commit=commit)
        if any(verified[k] != data[k] for k in data):
            raise OASISReportError("The saved teaching/OASIS links could not be verified. Refresh before retrying.")
        return verified

    def save_username(self, preceptor_name, record_id, *, expected):
        name, rid = _valid_name(preceptor_name), validate_username(record_id)
        current = self._current(expected)
        key = name_key(name)
        if any(k != key and row["record_id"] == rid for k, row in current["entries"].items()):
            raise OASISReportError("That username is already linked to another preceptor name. Remove/correct that link, or combine spelling aliases in TEACHING_PRECEPTOR_NAME_MAP before linking.")
        if current["entries"].get(key, {}).get("record_id") == rid:
            return current
        data = {k: current[k] for k in _empty()}
        data["entries"] = {**current["entries"], key: {"preceptor_name": name, "record_id": rid, "updated_at": _timestamp()}}
        return self._write(data, current)

    def remove_username(self, preceptor_name, *, expected):
        current = self._current(expected)
        data = {k: current[k] for k in _empty()}
        data["entries"] = {k: v for k, v in current["entries"].items() if k != name_key(preceptor_name)}
        return self._write(data, current)

    def save_report(self, start, end, filename, *, expected):
        key = period_key(start, end)
        if _filename_dates(filename) != (start, end):
            raise OASISReportError("Choose an OASIS summary with the same exact start and end dates as this teaching report.")
        current = self._current(expected)
        if current["report_links"].get(key, {}).get("summary_filename") == filename:
            return current
        data = {k: current[k] for k in _empty()}
        data["report_links"] = {**current["report_links"], key: {"start_date": start.isoformat(), "end_date": end.isoformat(),
                   "summary_filename": filename, "updated_at": _timestamp()}}
        return self._write(data, current)

    def remove_report(self, start, end, *, expected):
        current = self._current(expected)
        data = {k: current[k] for k in _empty()}
        data["report_links"] = {k: v for k, v in current["report_links"].items() if k != period_key(start, end)}
        return self._write(data, current)
