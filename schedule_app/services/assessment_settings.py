"""One shared, encrypted PTS assessment-completion threshold.

An absent file means the initial default (3). An unreadable file never means
"use 3": fail visibly rather than replacing a setting we could not verify.
No preceptor/student data or credentials are stored in this small catalog.
"""
from __future__ import annotations
import base64
import hashlib
import json
from urllib.parse import quote
from cryptography.fernet import InvalidToken
from schedule_app.services.opd_archive import OPDArchiveError

DEFAULT_MINIMUM_SHIFTS = 3
MAX_MINIMUM_SHIFTS = 10000
FILENAME = "pts_assessment_settings.json.enc"
KIND = "pts_assessment_settings"
VERSION = 1
MAX_BYTES = 8192


def validate_minimum_shifts(value):
    """Reject booleans/floats/strings rather than silently rounding a threshold."""
    if type(value) is not int or not 1 <= value <= MAX_MINIMUM_SHIFTS:
        raise OPDArchiveError(f"Minimum shifts must be a whole number from 1 to {MAX_MINIMUM_SHIFTS:,}.")
    return value


def _data(value=DEFAULT_MINIMUM_SHIFTS):
    return {"kind": KIND, "version": VERSION,
            "minimum_shifts": validate_minimum_shifts(value)}


def _validate(data):
    if (not isinstance(data, dict) or set(data) != {"kind", "version", "minimum_shifts"}
            or data.get("kind") != KIND or type(data.get("version")) is not int
            or data["version"] != VERSION):
        raise OPDArchiveError("PTS assessment settings have an unsupported format. No setting was changed.")
    validate_minimum_shifts(data["minimum_shifts"])
    return data


def _unique_keys(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError("Duplicate settings field")
        result[key] = value
    return result


class GitHubAssessmentSettings:
    """Read-at-commit, SHA-checked updates, and verified ciphertext-only writes."""
    def __init__(self, archive):
        self.archive, self.config = archive, archive.config
        self.path = f"{self.config.folder}/{FILENAME}"
        self.route = "/contents/" + quote(self.path, safe="/")
        self.cipher = self.config.cipher()

    def load(self, *, commit=None):
        commit = commit or self.archive._head()
        meta = self.archive._request("GET", self.route, params={"ref": commit}, missing_ok=True)
        if meta is None:
            return {**_data(), "sha": None, "commit": commit, "scope": self.config.signature()}
        if (not isinstance(meta, dict) or meta.get("type") != "file" or meta.get("target")
                or meta.get("submodule_git_url") or meta.get("path", self.path) != self.path
                or meta.get("encoding") != "base64" or type(meta.get("size")) is not int
                or not 0 < meta["size"] <= MAX_BYTES
                or not isinstance(meta.get("content"), str) or len(meta["content"]) > 2 * MAX_BYTES):
            raise OPDArchiveError("PTS assessment settings are not a supported encrypted file. Nothing was changed.")
        try:
            token = base64.b64decode("".join(meta["content"].split()), validate=True)
            sha = hashlib.sha1(b"blob " + str(len(token)).encode() + b"\0" + token).hexdigest()
            if len(token) != meta["size"] or sha != meta.get("sha"):
                raise ValueError()
            raw = self.cipher.decrypt(token)
            if len(raw) > MAX_BYTES:
                raise ValueError()
            data = _validate(json.loads(raw.decode("utf-8"), object_pairs_hook=_unique_keys))
        except (ValueError, TypeError, UnicodeError, InvalidToken, RecursionError):
            raise OPDArchiveError("PTS assessment settings could not be verified/decrypted. "
                                  "Check the encryption key or saved file; the last setting was not overwritten.") from None
        return {**data, "sha": sha, "commit": commit, "scope": self.config.signature()}

    def save(self, minimum_shifts, *, expected):
        value = validate_minimum_shifts(minimum_shifts)
        if (not isinstance(expected, dict) or expected.get("scope") != self.config.signature()
                or "sha" not in expected):
            raise OPDArchiveError("Load the saved minimum-shifts setting before changing it.")
        current = self.load()
        if current["sha"] != expected["sha"]:
            raise OPDArchiveError("The minimum-shifts setting changed in another session. "
                                  "Reload the saved minimum, review it, then make your change again.")
        if current["minimum_shifts"] == value and current["sha"] is not None:
            return current
        raw = json.dumps(_data(value), sort_keys=True, separators=(",", ":")).encode("utf-8")
        token = self.cipher.encrypt(raw)
        body = {"message": "Update encrypted PTS assessment settings", "branch": self.config.branch,
                "content": base64.b64encode(token).decode("ascii")}
        if current["sha"] is not None:
            body["sha"] = current["sha"]
        result = self.archive._request("PUT", self.route, body=body)
        commit = result.get("commit", {}).get("sha") if isinstance(result, dict) else None
        if not commit:
            raise OPDArchiveError("The minimum-shifts save was not confirmed. Reload the saved minimum to verify it.")
        verified = self.load(commit=commit)
        if verified["minimum_shifts"] != value:
            raise OPDArchiveError("The minimum-shifts save could not be verified. Reload the saved minimum before continuing.")
        return verified
