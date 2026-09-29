"""Encrypted, shared reporting-date presets in the existing GitHub archive repo.

Only this one catalog can be written. No OPDs or generated reports are changed.
All writes require the catalog revision that the user actually reviewed. Each
successful change is re-read at its returned commit and decrypted for verification.
"""
from copy import deepcopy
from datetime import date, datetime, timezone
from typing import Any
from urllib.parse import quote
from uuid import uuid4
import base64
import hashlib
import json
import re
import unicodedata

from cryptography.fernet import InvalidToken

from schedule_app.services.opd_archive import GitHubOPDArchive, OPDArchiveError
from schedule_app.services.reporting_periods import ReportingPeriod

PRESET_SCHEMA_VERSION = 1
PRESET_FILENAME = "reporting_date_presets.json.enc"
PRESET_KIND = "pediatric_clerkship_reporting_date_presets"
MAX_PRESETS = 250
MAX_PLAINTEXT_BYTES = 256 * 1024
MAX_ENCRYPTED_BYTES = 384 * 1024
CONCURRENT_MESSAGE = (
    "Saved date presets changed in GitHub after this list was loaded. "
    "No replacement was made. Click Refresh saved presets, review the current "
    "selection, and save or delete again."
)


class ReportingPresetError(OPDArchiveError):
    """Safe error text for the date-preset controls; contains no credentials."""


def normalize_preset_name(value: str) -> str:
    if not isinstance(value, str):
        raise ReportingPresetError("Enter a preset name of 1-80 characters.")
    value = unicodedata.normalize("NFKC", value).strip()
    if (not value or len(value) > 80 or any(
        ord(c) < 32 or ord(c) == 127 or c in "\ufffe\uffff" or 0xD800 <= ord(c) <= 0xDFFF
        for c in value
    )):
        raise ReportingPresetError("Enter a preset name of 1-80 characters without control characters.")
    return re.sub(r"\s+", " ", value)


def preset_name_key(value: str) -> str:
    return normalize_preset_name(value).casefold()


def _timestamp() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _valid_timestamp(value: Any) -> bool:
    if not isinstance(value, str):
        return False
    try:
        return datetime.strptime(value, "%Y-%m-%dT%H:%M:%SZ").strftime("%Y-%m-%dT%H:%M:%SZ") == value
    except ValueError:
        return False


def period_from_preset(preset: dict) -> ReportingPeriod:
    """Use the same exact-date/label validation as the existing reports."""
    try:
        values = preset["period"]
        if not isinstance(values, dict) or set(values) != {"label", "start_date", "end_date"}:
            raise ValueError()
        start_text, end_text = values["start_date"], values["end_date"]
        start, end = date.fromisoformat(start_text), date.fromisoformat(end_text)
        if start.isoformat() != start_text or end.isoformat() != end_text:
            raise ValueError()
        return ReportingPeriod(values["label"], start, end)
    except (KeyError, TypeError, ValueError, OPDArchiveError):
        raise ReportingPresetError(
            "A saved preset has invalid dates or a report label. The catalog was not changed."
        ) from None


def _validate_catalog(data: Any) -> dict:
    """Reject unknown/corrupt formats, rather than replacing them with an empty list."""
    if (not isinstance(data, dict)
        or set(data) != {"kind", "schema_version", "presets"}
        or data.get("kind") != PRESET_KIND
        or type(data.get("schema_version")) is not int
        or data["schema_version"] != PRESET_SCHEMA_VERSION
        or not isinstance(data.get("presets"), list)
        or len(data["presets"]) > MAX_PRESETS):
        raise ReportingPresetError("The saved date-preset catalog has an unsupported format. It was not changed.")
    ids, names = set(), set()
    result = []
    for entry in data["presets"]:
        if (not isinstance(entry, dict)
            or set(entry) != {"id", "name", "period", "created_at", "updated_at"}
            or not isinstance(entry.get("id"), str)
            or not re.fullmatch(r"[0-9a-f]{32}", entry["id"])
            or not _valid_timestamp(entry.get("created_at"))
            or not _valid_timestamp(entry.get("updated_at"))):
            raise ReportingPresetError("The saved date-preset catalog contains an invalid entry. It was not changed.")
        name = normalize_preset_name(entry["name"])
        period = period_from_preset(entry)
        key = preset_name_key(name)
        if entry["id"] in ids or key in names:
            raise ReportingPresetError("The saved date-preset catalog contains duplicate names or identifiers. It was not changed.")
        ids.add(entry["id"])
        names.add(key)
        result.append({**entry, "name": name, "period": period.as_dict()})
    # Stable order makes the dropdown easy to use and verification deterministic.
    result.sort(key=lambda item: (preset_name_key(item["name"]), item["id"]))
    return {"kind": PRESET_KIND, "schema_version": PRESET_SCHEMA_VERSION, "presets": result}


def _empty_catalog() -> dict:
    return {"kind": PRESET_KIND, "schema_version": PRESET_SCHEMA_VERSION, "presets": []}


class GitHubReportingPresets:
    """Reuse the configured repo, branch, token and Fernet key; no new Secrets."""

    def __init__(self, archive: GitHubOPDArchive):
        self.archive = archive
        self.config = archive.config
        self.cipher = self.config.cipher()
        # A fixed filename cannot be influenced by a typed preset name or label.
        self.path = f"{self.config.folder}/{PRESET_FILENAME}"
        self.route = "/contents/" + quote(self.path, safe="/")

    def _call(self, method, route, **kwargs):
        try:
            return self.archive._request(method, route, **kwargs)
        except OPDArchiveError as exc:
            raise ReportingPresetError(
                "GitHub date-preset operation was not confirmed. " + str(exc) +
                " Refresh saved presets before retrying. Your OPDs were not changed."
            ) from None

    def _head(self):
        try:
            return self.archive._head()
        except OPDArchiveError as exc:
            raise ReportingPresetError("Saved date presets could not be reached. " + str(exc)) from None

    def load(self, *, commit=None) -> dict:
        commit = commit or self._head()
        metadata = self._call("GET", self.route, params={"ref": commit}, missing_ok=True)
        if metadata is None:
            return {**_empty_catalog(), "sha": None, "commit": commit, "scope": self.config.signature()}
        if (not isinstance(metadata, dict) or metadata.get("type") != "file"
            or metadata.get("submodule_git_url") or metadata.get("target")
            or metadata.get("path", self.path) != self.path):
            raise ReportingPresetError("The date-preset path is not a regular catalog file. It was not changed.")
        size = metadata.get("size")
        if type(size) is not int or size <= 0 or size > MAX_ENCRYPTED_BYTES:
            raise ReportingPresetError("The saved date-preset file has an invalid size. It was not changed.")
        if metadata.get("encoding") != "base64" or not isinstance(metadata.get("content"), str):
            # Our small catalog never needs raw-download URLs or larger-file fallback.
            raise ReportingPresetError("GitHub did not return the expected date-preset contents. Refresh saved presets.")
        content = metadata["content"]
        if len(content) > MAX_ENCRYPTED_BYTES * 2:
            raise ReportingPresetError("The saved date-preset file exceeds this app's size limit.")
        try:
            token = base64.b64decode("".join(content.split()), validate=True)
        except (ValueError, TypeError):
            raise ReportingPresetError("GitHub returned invalid date-preset contents. It was not changed.") from None
        if len(token) != size or len(token) > MAX_ENCRYPTED_BYTES:
            raise ReportingPresetError("The date-preset download size could not be verified. Refresh saved presets.")
        sha = hashlib.sha1(b"blob " + str(len(token)).encode() + b"\0" + token).hexdigest()
        if sha != metadata.get("sha"):
            raise ReportingPresetError("The date-preset download did not match its GitHub identifier. Refresh saved presets.")
        try:
            plain = self.cipher.decrypt(token)
        except (InvalidToken, ValueError):
            raise ReportingPresetError(
                "Saved date presets could not be decrypted. Check the existing encryption key in "
                "Streamlit Secrets; the catalog was not overwritten."
            ) from None
        if len(plain) > MAX_PLAINTEXT_BYTES:
            raise ReportingPresetError("The decrypted date-preset catalog is too large. It was not changed.")
        try:
            data = json.loads(plain.decode("utf-8"))
        except (ValueError, UnicodeError, RecursionError):
            raise ReportingPresetError("The decrypted date-preset catalog is invalid. It was not changed.") from None
        return {**_validate_catalog(data), "sha": sha, "commit": commit, "scope": self.config.signature()}

    def _current_for_write(self, expected: dict) -> dict:
        if (not isinstance(expected, dict) or expected.get("scope") != self.config.signature()
            or "sha" not in expected):
            raise ReportingPresetError("Refresh saved presets for this repository before making changes.")
        current = self.load()
        if current["sha"] != expected["sha"]:
            raise ReportingPresetError(CONCURRENT_MESSAGE)
        return current

    def _write(self, entries: list, current: dict) -> dict:
        data = _validate_catalog({**_empty_catalog(), "presets": entries})
        plain = json.dumps(data, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")
        if len(plain) > MAX_PLAINTEXT_BYTES:
            raise ReportingPresetError("There are too many saved date settings. Delete unused presets first.")
        token = self.cipher.encrypt(plain)
        if len(token) > MAX_ENCRYPTED_BYTES:
            raise ReportingPresetError("The encrypted date-preset catalog is too large to save.")
        body = {"message": "Update encrypted reporting-date presets", "branch": self.config.branch,
                "content": base64.b64encode(token).decode("ascii")}
        if current["sha"] is not None:
            body["sha"] = current["sha"]
        result = self._call("PUT", self.route, body=body)
        try:
            commit = result["commit"]["sha"]
            if not isinstance(commit, str) or not re.fullmatch(r"[0-9a-f]{40,64}", commit):
                raise ValueError()
        except (KeyError, TypeError, ValueError):
            raise ReportingPresetError(
                "GitHub did not confirm the date-preset change. Refresh saved presets before retrying."
            ) from None
        verified = self.load(commit=commit)
        if verified["sha"] is None or verified["presets"] != data["presets"]:
            raise ReportingPresetError(
                "The date-preset change could not be verified after saving. Refresh saved presets before retrying."
            )
        return verified

    def save(self, name: str, period: ReportingPeriod, expected: dict, *, replace_id=None) -> dict:
        name = normalize_preset_name(name)
        if not isinstance(period, ReportingPeriod):
            raise ReportingPresetError("Choose valid dates and a report label before saving a preset.")
        current = self._current_for_write(expected)
        entries = deepcopy(current["presets"])
        same_name = next((row for row in entries if preset_name_key(row["name"]) == preset_name_key(name)), None)
        if replace_id is None:
            if same_name:
                raise ReportingPresetError("That preset name already exists. Confirm replacement or choose a different name.")
            if len(entries) >= MAX_PRESETS:
                raise ReportingPresetError(f"A maximum of {MAX_PRESETS} date presets is supported. Delete an unused preset first.")
            stamp = _timestamp()
            entry = {"id": uuid4().hex, "name": name, "period": period.as_dict(),
                     "created_at": stamp, "updated_at": stamp}
            entries.append(entry)
            action = "created"
        else:
            entry = next((row for row in entries if row["id"] == replace_id), None)
            if not entry:
                raise ReportingPresetError("The preset to replace is no longer available. Refresh saved presets.")
            if same_name is not None and same_name["id"] != replace_id:
                raise ReportingPresetError("Another preset already uses that name. Choose a different name.")
            if entry["name"] == name and entry["period"] == period.as_dict():
                return {"action": "unchanged", "preset_id": entry["id"], "snapshot": current}
            entry.update(name=name, period=period.as_dict(), updated_at=_timestamp())
            action = "updated"
        verified = self._write(entries, current)
        return {"action": action, "preset_id": entry["id"], "snapshot": verified}

    def delete(self, preset_id: str, expected: dict) -> dict:
        current = self._current_for_write(expected)
        if not any(row["id"] == preset_id for row in current["presets"]):
            raise ReportingPresetError("The selected preset is no longer available. Refresh saved presets.")
        entries = [row for row in current["presets"] if row["id"] != preset_id]
        # Keep an empty encrypted catalog after deleting the last entry. This
        # avoids any wildcard/folder deletion and preserves revision protection.
        verified = self._write(entries, current)
        return {"action": "deleted", "preset_id": preset_id, "snapshot": verified}
