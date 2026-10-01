"""Explicit, reversible OPD student-entry exclusions, encrypted in GitHub.

Original OPDs, OASIS records and saved name matches are never changed. Every
catalog write uses the prior revision and a verified encrypted read-back.
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

FILENAME = "pts_ignored_student_entries.json.enc"
KIND = "pts_ignored_student_entries"
MAX_BYTES = 512 * 1024


def ignored_entry_key(value):
    """Exact OPD entry match, ignoring case, whitespace, and comma spacing only.

    Keep program designations, initials, punctuation and spelling intact here.
    The identity-matching rules are separate from a user's explicit exclusion.
    """
    text = re.sub(r"\s+", " ", str(value or "").strip())
    return re.sub(r"\s*,\s*", ", ", text).casefold()


def exclusion_signature(names=()):
    keys = sorted({ignored_entry_key(name) for name in names if ignored_entry_key(name)})
    return hashlib.sha256(json.dumps(keys, ensure_ascii=True, separators=(",", ":")).encode()).hexdigest()


EMPTY_EXCLUSION_SIGNATURE = exclusion_signature()
EXCLUSION_VERSION = 1


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
        raise OPDArchiveError("The saved ignored student entries have an unsupported format. Nothing was changed.")
    for key, row in data["entries"].items():
        if not isinstance(row, dict) or set(row) != {"student_entry", "updated_at"}:
            raise OPDArchiveError("The saved ignored student entries contain an invalid entry.")
        if key != ignored_entry_key(_text(row["student_entry"], "student name")):
            raise OPDArchiveError("A saved ignored student entry has an invalid name key.")
        try:
            datetime.strptime(row["updated_at"], "%Y-%m-%dT%H:%M:%SZ")
        except (ValueError, TypeError):
            raise OPDArchiveError("A saved ignored student entry has an invalid update date.") from None
    return data


class GitHubIgnoredStudentEntries:
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
            raise OPDArchiveError("Ignored student entries are not a supported encrypted catalog.")
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
            raise OPDArchiveError("Ignored student entries could not be verified/decrypted. Restore the correct key; nothing was overwritten.") from None
        return {**data, "sha": sha, "commit": commit, "scope": self.config.signature()}

    def _write(self, entries, *, expected):
        if not isinstance(expected, dict) or expected.get("scope") != self.config.signature() or "sha" not in expected:
            raise OPDArchiveError("Load ignored student entries before editing them.")
        current = self.load()
        if current["sha"] != expected["sha"]:
            raise OPDArchiveError("Ignored student entries changed in another session. Refresh and review before retrying.")
        if entries == current["entries"]:
            return current
        data = _validate({**_empty(), "entries": entries})
        raw = json.dumps(data, ensure_ascii=True, sort_keys=True, separators=(",", ":")).encode()
        token = self.cipher.encrypt(raw)
        if len(token) > MAX_BYTES:
            raise OPDArchiveError("The ignored student-entry catalog is too large; nothing was written.")
        body = {"message": "Update encrypted PTS student exclusions", "branch": self.config.branch,
                "content": base64.b64encode(token).decode("ascii")}
        if current["sha"] is not None:
            body["sha"] = current["sha"]
        result = self.archive._request("PUT", self.route, body=body)
        commit = result.get("commit", {}).get("sha") if isinstance(result, dict) else None
        if not commit:
            raise OPDArchiveError("Ignored student entry save was not confirmed. Refresh before retrying.")
        verified = self.load(commit=commit)
        if verified["entries"] != entries:
            raise OPDArchiveError("Ignored student entry save could not be verified. Refresh before retrying.")
        return verified

    def ignore_entries(self, student_entries, *, expected):
        if not isinstance(expected, dict) or not isinstance(expected.get("entries"), dict):
            raise OPDArchiveError("Load ignored student entries before editing them.")
        if not isinstance(student_entries, (list, tuple)) or not student_entries:
            raise OPDArchiveError("Select at least one OPD student entry to ignore.")
        entries = {**expected.get("entries", {})}
        for value in student_entries:
            name = _text(value, "OPD student entry")
            key = ignored_entry_key(name)
            if not key:
                raise OPDArchiveError("Select a nonempty OPD student entry.")
            if key not in entries:
                entries[key] = {"student_entry": name,
                                "updated_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")}
        return self._write(entries, expected=expected)

    def restore_entries(self, student_entries, *, expected):
        if not isinstance(expected, dict) or not isinstance(expected.get("entries"), dict):
            raise OPDArchiveError("Load ignored student entries before editing them.")
        if not isinstance(student_entries, (list, tuple)) or not student_entries:
            raise OPDArchiveError("Select an ignored student entry to restore.")
        keys = {ignored_entry_key(_text(value, "OPD student entry")) for value in student_entries}
        entries = {key: value for key, value in expected.get("entries", {}).items() if key not in keys}
        return self._write(entries, expected=expected)


def require_matching_exclusions(scan, catalog):
    """Do not combine an old teaching denominator with a different ignore list."""
    signature = exclusion_signature(catalog["entries"])
    if scan.get("student_exclusions_signature", EMPTY_EXCLUSION_SIGNATURE) != signature:
        raise OPDArchiveError("The ignored-student list changed. Refresh the archived OPDs and evaluation completeness before generating reports.")
    return signature
