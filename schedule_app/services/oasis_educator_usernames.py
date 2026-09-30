"""Explicit educator username corrections, encrypted in the existing archive repo.

Does not change any original OASIS CSV, OPD, date preset, or teaching report.
Only the fixed educator-mapping catalog is writable. Optimistic concurrency plus
read/decrypt verification protect other sessions' edits and damaged catalogs.
"""
from __future__ import annotations

import base64
import hashlib
import json
from datetime import datetime, timezone
from urllib.parse import quote

from cryptography.fernet import InvalidToken
from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.oasis_educator_reports import OASISReportError, validate_username

KIND = "oasis_educator_username_overrides"
SCHEMA = 1
FILENAME = "oasis_educator_usernames.json.enc"
MAX_PLAIN = 256 * 1024
MAX_CIPHER = 384 * 1024


def _empty():
    return {"kind": KIND, "schema_version": SCHEMA, "entries": {}}


def _validated(data):
    if (not isinstance(data, dict) or set(data) != {"kind", "schema_version", "entries"}
            or data.get("kind") != KIND or type(data.get("schema_version")) is not int
            or data["schema_version"] != SCHEMA or not isinstance(data.get("entries"), dict)
            or len(data["entries"]) > 2000):
        raise OASISReportError("Saved OASIS usernames have an unsupported format. The catalog was not changed.")
    ids = set()
    for key, entry in data["entries"].items():
        if (not isinstance(key, str) or len(key) > 320 or not key.startswith(("ext:", "user:", "email:", "name:"))
                or not isinstance(entry, dict) or set(entry) != {"educator_name", "record_id", "updated_at"}
                or not isinstance(entry.get("educator_name"), str) or not entry["educator_name"].strip()
                or len(entry["educator_name"]) > 256):
            raise OASISReportError("Saved OASIS usernames contain an invalid entry. The catalog was not changed.")
        rid = validate_username(entry["record_id"])
        if rid != entry["record_id"] or rid in ids:
            raise OASISReportError("Saved OASIS usernames contain duplicate or invalid record IDs.")
        ids.add(rid)
        try:
            datetime.strptime(entry["updated_at"], "%Y-%m-%dT%H:%M:%SZ")
        except (TypeError, ValueError):
            raise OASISReportError("Saved OASIS usernames contain an invalid update date.") from None
    return data


class GitHubOASISUsernames:
    def __init__(self, archive):
        self.archive = archive
        self.config = archive.config
        self.cipher = self.config.cipher()
        self.path = f"{self.config.folder}/{FILENAME}"
        self.route = "/contents/" + quote(self.path, safe="/")

    def load(self, *, commit=None):
        commit = commit or self.archive._head()
        metadata = self.archive._request("GET", self.route, params={"ref": commit}, missing_ok=True)
        if metadata is None:
            return {**_empty(), "sha": None, "commit": commit, "scope": self.config.signature()}
        if (not isinstance(metadata, dict) or metadata.get("type") != "file" or metadata.get("target")
                or metadata.get("submodule_git_url") or metadata.get("path", self.path) != self.path
                or metadata.get("encoding") != "base64" or not isinstance(metadata.get("content"), str)):
            raise OASISReportError("The saved username path is not a regular encrypted catalog. Nothing was changed.")
        size = metadata.get("size")
        if type(size) is not int or not 0 < size <= MAX_CIPHER or len(metadata["content"]) > MAX_CIPHER * 2:
            raise OASISReportError("The encrypted username catalog exceeds the supported size.")
        try:
            encrypted = base64.b64decode("".join(metadata["content"].split()), validate=True)
        except (ValueError, TypeError):
            raise OASISReportError("GitHub returned invalid username catalog content.") from None
        sha = hashlib.sha1(b"blob " + str(len(encrypted)).encode() + b"\0" + encrypted).hexdigest()
        if sha != metadata.get("sha") or len(encrypted) != size:
            raise OASISReportError("The username catalog download could not be verified. Refresh saved usernames.")
        try:
            plain = self.cipher.decrypt(encrypted)
        except (InvalidToken, ValueError):
            raise OASISReportError("Saved OASIS usernames could not be decrypted. Keep/restore the correct key; the catalog was not overwritten.") from None
        if len(plain) > MAX_PLAIN:
            raise OASISReportError("The decrypted username catalog is too large.")
        try:
            data = _validated(json.loads(plain.decode("utf-8")))
        except (ValueError, UnicodeError, RecursionError):
            raise OASISReportError("The decrypted username catalog is invalid. It was not changed.") from None
        return {**data, "sha": sha, "commit": commit, "scope": self.config.signature()}

    def _current(self, expected):
        if not isinstance(expected, dict) or expected.get("scope") != self.config.signature() or "sha" not in expected:
            raise OASISReportError("Load/refresh the saved OASIS usernames before saving changes.")
        current = self.load()
        if current["sha"] != expected["sha"]:
            raise OASISReportError("OASIS usernames changed in another session. Refresh saved usernames, review the current values, and retry. No overwrite was made.")
        return current

    def _write(self, entries, current):
        data = _validated({**_empty(), "entries": entries})
        if current["entries"] == entries:
            return current
        plain = json.dumps(data, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode()
        if len(plain) > MAX_PLAIN:
            raise OASISReportError("Too many saved username corrections; nothing was written.")
        encrypted = self.cipher.encrypt(plain)
        body = {"message": "Update encrypted OASIS educator usernames", "branch": self.config.branch,
                "content": base64.b64encode(encrypted).decode("ascii")}
        if current["sha"] is not None:
            body["sha"] = current["sha"]
        result = self.archive._request("PUT", self.route, body=body)
        try:
            commit = result["commit"]["sha"]
            if not isinstance(commit, str) or not commit:
                raise ValueError()
        except (KeyError, TypeError, ValueError):
            raise OASISReportError("GitHub did not confirm the username update. Refresh before retrying.") from None
        verified = self.load(commit=commit)
        if verified["entries"] != data["entries"]:
            raise OASISReportError("The saved username change could not be verified. Refresh before retrying.")
        return verified

    def save(self, educator_key, educator_name, record_id, *, expected):
        record_id = validate_username(record_id)
        current = self._current(expected)
        entries = {k: dict(v) for k, v in current["entries"].items()}
        existing = entries.get(educator_key, {})
        if existing.get("record_id") == record_id and existing.get("educator_name") == educator_name:
            return current
        entries[educator_key] = {"educator_name": educator_name, "record_id": record_id,
                                 "updated_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")}
        return self._write(entries, current)

    def remove(self, educator_key, *, expected):
        current = self._current(expected)
        entries = {k: dict(v) for k, v in current["entries"].items() if k != educator_key}
        return self._write(entries, current)
