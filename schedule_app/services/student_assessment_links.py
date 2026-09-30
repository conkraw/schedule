"""Explicit OPD student-name to Student External ID corrections, encrypted at rest.

The catalog is separate from preceptor usernames and does not modify any source.
Aliases may point to the same external ID. Identity is never fuzzy-matched.
"""
from __future__ import annotations
import base64
import hashlib
import json
import re
from datetime import datetime, timezone
from urllib.parse import quote
from cryptography.fernet import InvalidToken
from schedule_app.services.opd_archive import OPDArchiveError

FILENAME = "student_assessment_id_links.json.enc"
KIND = "student_assessment_id_links"
MAX_BYTES = 512 * 1024


def student_name_key(value):
    text = re.sub(r"\s+", " ", str(value or "").strip())
    # The supplied OASIS Student field explicitly appends the MD class year.
    # Strip only this recognized suffix, not arbitrary content after a semicolon.
    text = re.sub(r";\s*MD\s*\d{4}\s*$", "", text, flags=re.I)
    return re.sub(r"\s*,\s*", ", ", text).strip().casefold()


def _text(value, label):
    if (not isinstance(value, str) or not value.strip() or len(value) > 256
            or any(ord(c) < 32 or ord(c) == 127 or 0xD800 <= ord(c) <= 0xDFFF or c in "\ufffe\uffff" for c in value)):
        raise OPDArchiveError(f"Enter a valid {label}.")
    return value.strip()


def _empty():
    return {"kind": KIND, "version": 1, "entries": {}}


def _validate(data):
    if (not isinstance(data, dict) or set(data) != set(_empty())
            or data.get("kind") != KIND or type(data.get("version")) is not int
            or data["version"] != 1 or not isinstance(data.get("entries"), dict)
            or len(data["entries"]) > 3000):
        raise OPDArchiveError("The saved student ID links have an unsupported format. Nothing was changed.")
    for key, row in data["entries"].items():
        if not isinstance(row, dict) or set(row) != {"student_name", "external_id", "updated_at"}:
            raise OPDArchiveError("The saved student ID links contain an invalid entry.")
        if key != student_name_key(_text(row["student_name"], "student name")):
            raise OPDArchiveError("A saved student ID link has an invalid name key.")
        _text(row["external_id"], "Student External ID")
        try:
            datetime.strptime(row["updated_at"], "%Y-%m-%dT%H:%M:%SZ")
        except (ValueError, TypeError):
            raise OPDArchiveError("A saved student ID link has an invalid update date.") from None
    return data


class GitHubStudentAssessmentLinks:
    def __init__(self, archive):
        self.archive, self.config = archive, archive.config
        self.path = f"{self.config.folder}/{FILENAME}"
        self.route = "/contents/" + quote(self.path, safe="/")
        self.cipher = self.config.cipher()

    def load(self, *, commit=None):
        commit = commit or self.archive._head()
        meta = self.archive._request("GET", self.route, params={"ref": commit}, missing_ok=True)
        if meta is None:
            return {**_empty(), "sha": None, "commit": commit, "scope": self.config.signature()}
        if (not isinstance(meta, dict) or meta.get("type") != "file" or meta.get("target")
                or meta.get("submodule_git_url") or meta.get("path", self.path) != self.path
                or meta.get("encoding") != "base64" or type(meta.get("size")) is not int
                or not 0 < meta["size"] <= MAX_BYTES
                or not isinstance(meta.get("content"), str) or len(meta["content"]) > 2 * MAX_BYTES):
            raise OPDArchiveError("Student ID links are not a supported encrypted catalog.")
        try:
            token = base64.b64decode("".join(meta["content"].split()), validate=True)
            sha = hashlib.sha1(b"blob " + str(len(token)).encode() + b"\0" + token).hexdigest()
            if len(token) != meta["size"] or sha != meta.get("sha"):
                raise ValueError()
            raw = self.cipher.decrypt(token)
            if len(raw) > MAX_BYTES:
                raise ValueError()
            data = _validate(json.loads(raw.decode("utf-8")))
        except (ValueError, TypeError, UnicodeError, InvalidToken, RecursionError):
            raise OPDArchiveError("Student ID links could not be verified/decrypted. Restore the correct key; nothing was overwritten.") from None
        return {**data, "sha": sha, "commit": commit, "scope": self.config.signature()}

    def _write(self, entries, *, expected):
        if not isinstance(expected, dict) or expected.get("scope") != self.config.signature() or "sha" not in expected:
            raise OPDArchiveError("Load student ID links before editing them.")
        current = self.load()
        if current["sha"] != expected["sha"]:
            raise OPDArchiveError("Student ID links changed in another session. Refresh and review before retrying.")
        if entries == current["entries"]:
            return current
        data = _validate({**_empty(), "entries": entries})
        raw = json.dumps(data, ensure_ascii=True, sort_keys=True, separators=(",", ":")).encode()
        token = self.cipher.encrypt(raw)
        if len(token) > MAX_BYTES:
            raise OPDArchiveError("The student ID catalog is too large; nothing was written.")
        body = {"message": "Update encrypted student assessment links", "branch": self.config.branch,
                "content": base64.b64encode(token).decode("ascii")}
        if current["sha"] is not None:
            body["sha"] = current["sha"]
        result = self.archive._request("PUT", self.route, body=body)
        commit = result.get("commit", {}).get("sha") if isinstance(result, dict) else None
        if not commit:
            raise OPDArchiveError("Student ID link save was not confirmed. Refresh before retrying.")
        verified = self.load(commit=commit)
        if verified["entries"] != entries:
            raise OPDArchiveError("Student ID link save could not be verified. Refresh before retrying.")
        return verified

    def save_link(self, student_name, external_id, *, expected):
        name, sid = _text(student_name, "student name"), _text(external_id, "Student External ID")
        if sid.casefold() in {"nan", "n/a", "none", "null", "-", "--"}:
            raise OPDArchiveError("Enter the actual Student External ID, not a missing-value label.")
        key = student_name_key(name)
        if not key:
            raise OPDArchiveError("Select a student name.")
        previous = expected.get("entries", {}).get(key, {})
        if previous.get("external_id") == sid:
            return self._write(expected["entries"], expected=expected)
        entries = {**expected["entries"], key: {"student_name": name, "external_id": sid,
                   "updated_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")}}
        return self._write(entries, expected=expected)

    def remove_link(self, student_name, *, expected):
        entries = {k: v for k, v in expected["entries"].items() if k != student_name_key(student_name)}
        return self._write(entries, expected=expected)
