import io
import streamlit as st
import pandas as pd
import numpy as np 
import re
import xlsxwriter
import random
from openpyxl import load_workbook # Ensure load_workbook is imported
import io, zipfile
from docx import Document
from docx.shared import Pt
from docx.enum.section import WD_ORIENT
from datetime import timedelta
from xlsxwriter import Workbook as Workbook
from collections import defaultdict
from datetime import datetime, timedelta
from collections import Counter
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
import re
from collections import defaultdict
from copy import copy
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import column_index_from_string
from openpyxl.styles import Font, PatternFill, Color
from collections import deque

# =============================================================================
# ENCRYPTED OPD ARCHIVE - whole-file encryption, one current file per rotation
# Only ciphertext is sent to GitHub. No credentials or plaintext index is saved.
# Configure [opd_archive] in Streamlit Secrets; see SETUP_OPD_ARCHIVE.md.
# =============================================================================
import base64
import hashlib
import hmac
from dataclasses import dataclass, field
from datetime import date as CalendarDate
from urllib.parse import quote
import requests
from cryptography.fernet import Fernet, MultiFernet, InvalidToken
from openpyxl.utils.datetime import from_excel

OPD_ARCHIVE_VERSION = 1
OPD_MAX_BYTES = 10 * 1024 * 1024
OPD_MAX_ENCRYPTED_BYTES = 15 * 1024 * 1024
OPD_XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
OPD_NAME_ORDERS = (
    "Auto-detect using rotation list",
    "Student ~ Preceptor",
    "Preceptor ~ Student",
)


class OPDArchiveError(RuntimeError):
    """A safe, user-facing archive error; never include a token or workbook data."""


def _opd_secret_settings():
    try:
        return dict(st.secrets.get("opd_archive", {}))
    except Exception:
        return {}


@dataclass(frozen=True)
class OPDArchiveConfig:
    owner: str
    repo: str
    branch: str
    folder: str
    github_token: str = field(repr=False)
    encryption_key: str = field(repr=False)
    previous_encryption_keys: tuple = field(default=(), repr=False)

    def cipher(self):
        try:
            keys = (self.encryption_key,) + self.previous_encryption_keys
            return MultiFernet([Fernet(k.encode("ascii")) for k in keys])
        except (ValueError, TypeError, UnicodeError):
            raise OPDArchiveError("An archive encryption key is invalid. Use a Fernet "
                                  "key from generate_opd_secrets.py.") from None

    def signature(self):
        # Session-only identity. Includes credential changes to invalidate receipts.
        material = "\0".join((self.owner, self.repo, self.branch, self.folder,
                              self.github_token, self.encryption_key,
                              *self.previous_encryption_keys))
        return hashlib.sha256(material.encode()).hexdigest()


def get_opd_archive_config():
    values = _opd_secret_settings()
    required = ("owner", "repo", "github_token", "encryption_key")
    if any(not str(values.get(k, "")).strip() for k in required):
        raise OPDArchiveError("Archive setup is incomplete. Add owner, repo, "
                              "github_token and encryption_key under [opd_archive] "
                              "in Streamlit Secrets.")
    owner, repo = str(values["owner"]).strip(), str(values["repo"]).strip()
    if not re.fullmatch(r"[A-Za-z0-9-]+", owner) or not re.fullmatch(r"[A-Za-z0-9_.-]+", repo):
        raise OPDArchiveError("Use the GitHub account name and repository name only, not a URL.")
    branch = str(values.get("branch", "main")).strip()
    folder = str(values.get("folder", "opd_archive")).strip("/")
    if not branch or not folder or any(
        not re.fullmatch(r"[A-Za-z0-9_-]+", part) for part in folder.split("/")
    ):
        raise OPDArchiveError("Check the archive branch and folder in Streamlit Secrets.")
    previous = values.get("previous_encryption_keys", [])
    if not isinstance(previous, (tuple, list)):
        raise OPDArchiveError("previous_encryption_keys must be a TOML list, or omitted.")
    config = OPDArchiveConfig(owner, repo, branch, folder,
                              str(values["github_token"]).strip(),
                              str(values["encryption_key"]).strip(),
                              tuple(str(k).strip() for k in previous))
    config.cipher()
    return config


def validate_opd_xlsx_bytes(raw):
    """Bound memory use and reject non-XLSX inputs without modifying the bytes."""
    if not isinstance(raw, bytes) or not raw:
        raise OPDArchiveError("The OPD upload is empty.")
    if len(raw) > OPD_MAX_BYTES:
        raise OPDArchiveError("OPD uploads must be 10 MB or smaller.")
    try:
        with ZipFile(BytesIO(raw)) as zf:
            members = zf.infolist()
            if len(members) > 20000 or sum(i.file_size for i in members) > 100 * 1024 * 1024:
                raise OPDArchiveError("This workbook is too large to process safely.")
            names = set(zf.namelist())
            if "xl/workbook.xml" not in names or "[Content_Types].xml" not in names:
                raise OPDArchiveError("Upload an actual Excel .xlsx workbook, not an .xls file.")
            if any(i.flag_bits & 1 for i in members) or "xl/vbaProject.bin" in names:
                raise OPDArchiveError("Use a normal .xlsx file without password protection or macros.")
            if zf.testzip() is not None:
                raise OPDArchiveError("The uploaded workbook failed its ZIP integrity check.")
    except zipfile.BadZipFile:
        raise OPDArchiveError("The file could not be read as an Excel .xlsx workbook.") from None


def _opd_date(value, epoch):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, CalendarDate):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            dt = from_excel(value, epoch)
            return dt.date() if isinstance(dt, datetime) and 1970 <= dt.year <= 2100 else None
        except (ValueError, OverflowError):
            return None
    if isinstance(value, str):
        for fmt in ("%m/%d/%Y", "%m-%d-%Y", "%Y-%m-%d", "%B %d, %Y", "%b %d, %Y"):
            try:
                return datetime.strptime(value.strip(), fmt).date()
            except ValueError:
                pass
    return None


def inspect_opd_rotation(raw):
    """Find the first scheduled Monday in site-tab OPDs, using calendar headers.

    The filename and names around '~' play no role in the archive identifier.
    Different date ranges across site worksheets are rejected, not guessed.
    """
    validate_opd_xlsx_bytes(raw)
    try:
        wb = load_workbook(BytesIO(raw), read_only=True, data_only=True)
    except Exception:
        raise OPDArchiveError("The OPD workbook could not be opened. Save it as .xlsx and retry.") from None
    try:
        expected_days = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
        schedules = {}
        for ws in wb.worksheets:
            rows = list(ws.iter_rows(min_row=1, max_row=min(ws.max_row or 0, 1200),
                                     max_col=8, values_only=True))
            if not rows or str(rows[0][0] or "").strip().rstrip(":").casefold() != "site":
                continue
            has_sessions = any(re.match(r"^\s*(AM|PM)\b", str(row[0] or ""), re.I) for row in rows)
            if not has_sessions:
                continue
            mondays = []
            for r, row in enumerate(rows[:-1]):
                labels = [str(v or "").strip().casefold() for v in row[1:8]]
                if labels != expected_days:
                    continue
                values = [_opd_date(v, wb.epoch) for v in rows[r + 1][1:8]]
                if not all(values) or values[0].weekday() != 0 or any(
                    day != values[0] + timedelta(days=i) for i, day in enumerate(values)
                ):
                    raise OPDArchiveError(f"The date row on worksheet '{ws.title}' near row {r+2} "
                                          "must contain seven consecutive Monday-Sunday dates.")
                mondays.append(values[0])
            if not mondays:
                raise OPDArchiveError(f"No readable Monday-Sunday date headers were found on '{ws.title}'.")
            if any(d != mondays[0] + timedelta(days=7*i) for i, d in enumerate(mondays)):
                raise OPDArchiveError(f"The calendar weeks on '{ws.title}' are not consecutive.")
            schedules[ws.title] = tuple(mondays)
        if not schedules:
            raise OPDArchiveError("This is not a recognized OPD site-tab workbook. Upload the original "
                                  "OPD.xlsx, not the one-tab-per-student MS_Schedule.xlsx.")
        first = next(iter(schedules.values()))
        if any(weeks != first for weeks in schedules.values()):
            raise OPDArchiveError("The OPD site worksheets have different rotation dates. "
                                  "Correct the dates before archiving; no file was replaced.")
        return {"rotation_start": first[0], "week_mondays": first,
                "site_names": tuple(schedules), "sha256": hashlib.sha256(raw).hexdigest()}
    finally:
        wb.close()


class GitHubOPDArchive:
    """GitHub Contents API transport; only authenticated ciphertext is written.

    Reads are pinned to a commit. Updates use the existing blob SHA. A competing
    write causes an explicit retry prompt rather than a silent overwrite/rebase.
    """
    def __init__(self, config, transport=None):
        self.config = config
        self.transport = transport or requests
        self.base = f"https://api.github.com/repos/{config.owner}/{config.repo}"
        self.cipher = config.cipher()

    def _request(self, method, route, *, params=None, body=None, raw=False, missing_ok=False):
        headers = {"Authorization": f"Bearer {self.config.github_token}",
                   "Accept": ("application/vnd.github.raw+json" if raw else
                              "application/vnd.github.object+json" if method == "GET" else "application/vnd.github+json"),
                   "X-GitHub-Api-Version": "2026-03-10",
                   "User-Agent": "OPD-Encrypted-Archive/1.0"}
        try:
            response = self.transport.request(method, self.base + route, headers=headers,
                                              params=params, json=body, timeout=(10, 45),
                                              allow_redirects=False, stream=raw)
        except requests.RequestException:
            raise OPDArchiveError("GitHub could not be reached. The save is not confirmed. "
                                  "Retry when the connection is available.") from None
        status = response.status_code
        if status == 404 and missing_ok:
            response.close()
            return None
        if status not in (200, 201):
            response.close()
            messages = {
                401: "GitHub authentication failed. Check or renew the access token in Streamlit Secrets.",
                403: "GitHub denied access or rate-limited this request. Check token permissions, organization approval and repository rules, then retry later.",
                404: "GitHub repository, branch or archive file was not found. Check owner, repo and branch.",
                409: "Another save changed the repository during this operation. Retry explicitly; the app will re-read the current version first.",
                422: "GitHub rejected the write. Check branch rules and token permissions, then retry.",
                429: "GitHub rate-limited this request. Wait before retrying.",
            }
            raise OPDArchiveError(messages.get(status, f"GitHub request failed (HTTP {status}); the save is not confirmed."))
        if raw:
            try:
                chunks, size = [], 0
                for chunk in response.iter_content(chunk_size=65536):
                    size += len(chunk)
                    if size > OPD_MAX_ENCRYPTED_BYTES:
                        raise OPDArchiveError("The encrypted archive file exceeds this app's size limit.")
                    chunks.append(chunk)
                return b"".join(chunks)
            except requests.RequestException:
                raise OPDArchiveError("The encrypted file download was interrupted. Retry loading it.") from None
            finally:
                response.close()
        try:
            return response.json()
        except (ValueError, requests.RequestException):
            raise OPDArchiveError("GitHub returned an unexpected response. No save is confirmed.") from None
        finally:
            response.close()

    def _head(self):
        data = self._request("GET", "/branches/" + quote(self.config.branch, safe=""))
        try:
            return data["commit"]["sha"]
        except (KeyError, TypeError):
            raise OPDArchiveError("GitHub did not return a valid branch. Initialize the archive repository with a README.") from None

    def path_for(self, rotation_start):
        if not isinstance(rotation_start, CalendarDate) or rotation_start.weekday() != 0:
            raise OPDArchiveError("The archive identifier must be the rotation's first Monday.")
        return f"{self.config.folder}/OPD_{rotation_start.isoformat()}.xlsx.enc"

    def _read_at(self, path, commit):
        route = "/contents/" + quote(path, safe="/")
        metadata = self._request("GET", route, params={"ref": commit}, missing_ok=True)
        if metadata is None:
            return None
        if not isinstance(metadata, dict) or metadata.get("type") != "file" or metadata.get("submodule_git_url"):
            raise OPDArchiveError("The archive path is not a regular file. No replacement was made.")
        if int(metadata.get("size", 0)) > OPD_MAX_ENCRYPTED_BYTES:
            raise OPDArchiveError("The encrypted archive file is too large for this app.")
        if metadata.get("encoding") == "base64" and metadata.get("content"):
            try:
                token = base64.b64decode("".join(metadata["content"].split()), validate=True)
            except (ValueError, TypeError):
                raise OPDArchiveError("GitHub returned invalid encoded file contents.") from None
        else:
            token = self._request("GET", route, params={"ref": commit}, raw=True)
        if len(token) > OPD_MAX_ENCRYPTED_BYTES:
            raise OPDArchiveError("The encrypted archive file is too large for this app.")
        blob_sha = hashlib.sha1(b"blob " + str(len(token)).encode() + b"\0" + token).hexdigest()
        if blob_sha != metadata.get("sha"):
            raise OPDArchiveError("The archive download did not match its GitHub file identifier. Retry loading it.")
        try:
            plaintext = self.cipher.decrypt(token)  # no TTL: historical files remain reloadable
        except (InvalidToken, ValueError):
            raise OPDArchiveError("This OPD cannot be decrypted with the configured key(s), or its encrypted "
                                  "contents were altered. The existing file will NOT be overwritten. "
                                  "Restore the correct encryption key in Streamlit Secrets.") from None
        details = inspect_opd_rotation(plaintext)
        if path != self.path_for(details["rotation_start"]):
            raise OPDArchiveError("The decrypted OPD's rotation does not match its archive filename.")
        return {"raw": plaintext, "encrypted": token, "details": details,
                "sha": metadata["sha"], "path": path, "commit": commit}

    def load(self, rotation_start, commit=None):
        found = self._read_at(self.path_for(rotation_start), commit or self._head())
        if found is None:
            raise OPDArchiveError("No archived OPD was found for that rotation. Refresh the archive list.")
        return found

    def save(self, raw):
        details = inspect_opd_rotation(raw)
        path = self.path_for(details["rotation_start"])
        current = self._read_at(path, self._head())
        if current is not None and hmac.compare_digest(current["raw"], raw):
            return {"action": "unchanged", "path": path, "sha": current["sha"],
                    "sha256": details["sha256"], "rotation_start": details["rotation_start"]}
        token = self.cipher.encrypt(raw)
        body = {"message": "Update encrypted OPD archive", "branch": self.config.branch,
                "content": base64.b64encode(token).decode("ascii")}
        if current is not None:
            body["sha"] = current["sha"]
        result = self._request("PUT", "/contents/" + quote(path, safe="/"), body=body)
        try:
            commit = result["commit"]["sha"]
        except (KeyError, TypeError):
            raise OPDArchiveError("GitHub did not return a save confirmation. Retry to verify the current archive.") from None
        verified = self._read_at(path, commit)
        if verified is None or not hmac.compare_digest(verified["raw"], raw):
            raise OPDArchiveError("The uploaded OPD could not be verified after saving. Retry before creating schedules.")
        return {"action": "replaced" if current else "created", "path": path,
                "sha": verified["sha"], "sha256": details["sha256"],
                "rotation_start": details["rotation_start"]}

    def list_rotations(self, commit=None):
        commit = commit or self._head()
        route = "/contents/" + quote(self.config.folder, safe="/")
        data = self._request("GET", route, params={"ref": commit}, missing_ok=True)
        if data is None:
            return []
        entries = data.get("entries") if isinstance(data, dict) else data
        if not isinstance(entries, list):
            raise OPDArchiveError("The configured archive folder is not a directory.")
        if len(entries) >= 1000:
            raise OPDArchiveError("The GitHub directory listing limit was reached. The app cannot safely "
                                  "show a complete archive; organize the archive before continuing.")
        rotations = []
        for item in entries:
            match = re.fullmatch(r"OPD_(\d{4}-\d{2}-\d{2})\.xlsx\.enc", str(item.get("name", "")))
            if item.get("type") != "file" or not match:
                continue
            try:
                day = CalendarDate.fromisoformat(match.group(1))
            except ValueError:
                continue
            if day.weekday() == 0:
                rotations.append(day)
        return sorted(set(rotations), reverse=True)


def _opd_scope_archive_session(config):
    signature = config.signature()
    if st.session_state.get("opd_archive_scope") != signature:
        for key in ("opd_archive_list", "opd_archive_loaded", "opd_archive_upload_state",
                    "opd_generated_master", "opd_master_signature"):
            st.session_state.pop(key, None)
        st.session_state["opd_archive_scope"] = signature


def _opd_upload_changed():
    for key in ("opd_archive_upload_state", "opd_generated_master", "opd_master_signature"):
        st.session_state.pop(key, None)


def _opd_archive_upload_ui(uploaded, client):
    """Save each upload event once; retries are explicit. A widget rerun is not a new save."""
    raw = uploaded.getvalue()
    details = inspect_opd_rotation(raw)
    signature = (client.config.signature(), getattr(uploaded, "file_id", None),
                 hashlib.sha256(raw).hexdigest())
    status = st.session_state.get("opd_archive_upload_state")
    if not status or status.get("signature") != signature:
        status = None
    if status is None:
        with st.spinner("Encrypting and verifying the original OPD in GitHub..."):
            try:
                receipt = client.save(raw)
                status = {"signature": signature, "receipt": receipt}
                st.session_state.pop("opd_archive_list", None)
                st.session_state.pop("opd_archive_loaded", None)
            except OPDArchiveError as exc:
                status = {"signature": signature, "error": str(exc)}
        st.session_state["opd_archive_upload_state"] = status
    if status.get("error"):
        st.error("OPD was NOT confirmed archived. " + status["error"])
        st.button("Retry encrypted OPD save", key="opd_retry_archive",
                  on_click=_opd_upload_changed)
        return raw, details, False
    receipt = status["receipt"]
    message = {"created": "New rotation saved", "replaced": "Current copy for this rotation replaced",
               "unchanged": "Identical original already archived; no extra commit created"}[receipt["action"]]
    st.success(f"OPD archived and verified - {details['rotation_start']:%B %d, %Y}. {message}.")
    return raw, details, True


def _opd_archive_picker(client, prefix):
    refresh = st.button("Refresh archive list", key=f"{prefix}_refresh")
    if refresh:
        st.session_state.pop("opd_archive_list", None)
        st.session_state.pop("opd_archive_loaded", None)
    if "opd_archive_list" not in st.session_state:
        try:
            st.session_state["opd_archive_list"] = client.list_rotations()
        except OPDArchiveError as exc:
            st.error(str(exc))
            return None
    rotations = st.session_state["opd_archive_list"]
    if not rotations:
        st.info("No archived rotations yet. Upload an OPD in Create Student Schedule to save the first one.")
        return None
    loaded = st.session_state.get("opd_archive_loaded")
    previous = loaded["details"]["rotation_start"] if loaded else None
    default_index = rotations.index(previous) if previous in rotations else 0
    selection_key = f"{prefix}_rotation"
    if st.session_state.get(selection_key) not in rotations:
        st.session_state.pop(selection_key, None)
    selected = st.selectbox("Rotation beginning", rotations, index=default_index,
                            format_func=lambda d: d.strftime("%B %d, %Y"), key=selection_key)
    if st.button("Load / decrypt selected OPD", key=f"{prefix}_load"):
        st.session_state.pop("opd_archive_loaded", None)
        try:
            with st.spinner("Loading and decrypting the latest archived OPD..."):
                st.session_state["opd_archive_loaded"] = client.load(selected)
        except OPDArchiveError as exc:
            st.error(str(exc))
    loaded = st.session_state.get("opd_archive_loaded")
    if loaded and loaded["details"]["rotation_start"] == selected:
        st.success("Archived original loaded. Reloading does not overwrite the archive.")
        st.download_button("Download original OPD.xlsx", loaded["raw"],
                           file_name=f"OPD_{selected.isoformat()}.xlsx", mime=OPD_XLSX_MIME,
                           key=f"{prefix}_download")
        return loaded
    return None


def _opd_use_loaded_callback():
    loaded = st.session_state.get("opd_archive_loaded")
    if loaded:
        st.session_state["schedule_archive_rotation"] = loaded["details"]["rotation_start"]
    st.session_state["opd_source_choice"] = "Reload archived OPD"
    st.session_state["schedule_app_mode"] = "Create Student Schedule"
    st.session_state.pop("opd_generated_master", None)


def render_opd_archive_page():
    st.subheader("Encrypted OPD Archive")
    st.caption("One current original OPD per rotation. Only encrypted workbook bytes are stored in GitHub.")
    try:
        client = GitHubOPDArchive(get_opd_archive_config())
        _opd_scope_archive_session(client.config)
    except OPDArchiveError as exc:
        st.error(str(exc))
        return
    loaded = _opd_archive_picker(client, "archive_page")
    if loaded:
        st.button("Use this OPD to create student schedules", on_click=_opd_use_loaded_callback,
                  key="opd_use_loaded_in_schedule")
    st.info("A newer upload with the same first Monday replaces the current copy. Older encrypted "
            "versions remain in Git history. Rotation dates, file sizes and commit metadata are public.")


def _opd_name_key(value):
    return re.sub(r"\s+", " ", str(value or "").strip()).casefold()


def split_opd_assignment(value, roster, order=OPD_NAME_ORDERS[0]):
    """Return (preceptor, canonical_student), ignoring spaces around '~'.

    Auto mode identifies the student using the rotation list, not an assumption
    about which side contains the preceptor. Unknown students are reported by the
    caller; a provider-only cell with an empty opposite side is not assigned.
    """
    if not isinstance(value, str) or "~" not in value:
        return None, None
    left, right = (part.strip() for part in value.split("~", 1))
    left_student, right_student = roster.get(_opd_name_key(left)), roster.get(_opd_name_key(right))
    if order == "Student ~ Preceptor":
        return (right, left_student) if left_student else (None, None)
    if order == "Preceptor ~ Student":
        return (left, right_student) if right_student else (None, None)
    if left_student and right_student:
        raise OPDArchiveError("Both sides of an assignment match the student list. Choose an explicit '~' name order.")
    if left_student:
        return right, left_student
    if right_student:
        return left, right_student
    return None, None


def collect_opd_assignments(raw, students, order=OPD_NAME_ORDERS[0]):
    """Read the existing four-week AM/PM layout without changing the original OPD."""
    roster = {_opd_name_key(name): name for name in students}
    wb = load_workbook(BytesIO(raw), data_only=True)
    assignments, unmatched = [], []
    try:
        for ws in wb.worksheets:
            for shift in ("AM", "PM"):
                rows = [cell.row for cell in ws["A"] if re.match(rf"^\s*{shift}\b", str(cell.value or ""), re.I)]
                blocks, block = [], []
                for r in rows:
                    if block and r != block[-1] + 1:
                        blocks.append(block)
                        block = []
                    block.append(r)
                if block:
                    blocks.append(block)
                for week, block in enumerate(blocks[:4]):
                    for col in range(2, 9):
                        for row in block:
                            cell = ws.cell(row=row, column=col)
                            preceptor, student = split_opd_assignment(cell.value, roster, order)
                            if student is not None:
                                assignments.append({"student": student, "preceptor": preceptor,
                                                    "site": ws.title, "week": week, "shift": shift,
                                                    "column": col, "coordinate": cell.coordinate})
                            elif isinstance(cell.value, str) and "~" in cell.value:
                                if all(part.strip() for part in cell.value.split("~", 1)):
                                    unmatched.append(f"{ws.title}!{cell.coordinate}")
        return assignments, unmatched
    finally:
        wb.close()


def populate_ms_schedule(blank_bytes, assignments):
    wb = load_workbook(BytesIO(blank_bytes))
    try:
        sheets = {_opd_name_key(ws["B1"].value): ws for ws in wb.worksheets}
        for item in assignments:
            ws = sheets.get(_opd_name_key(item["student"]))
            if ws is None:
                raise OPDArchiveError("A student could not be matched to the generated schedule tabs.")
            row = (6 if item["shift"] == "AM" else 7) + 8 * item["week"]
            ws.cell(row=row, column=item["column"], value=f"{item['preceptor']} - [{item['site']}]")
        output = BytesIO()
        wb.save(output)
        return output.getvalue()
    finally:
        wb.close()


# =============================================================================
# PRECEPTOR TEACHING SUMMARY - READ-ONLY analysis of current encrypted OPDs
# No workbook/student data is written back to GitHub by this section.
# =============================================================================
import csv
import json
import secrets as _teaching_secrets
from datetime import timezone as _teaching_timezone
from zoneinfo import ZoneInfo as _TeachingZoneInfo

TEACHING_REPORT_VERSION = 2
TEACHING_HOURS_PER_STUDENT_SHIFT = 4
TEACHING_CSV_COLUMNS = (
    "preceptor_name", "academic_year", "no_of_shifts",
    "months_worked", "educational_hours",
)
TEACHING_MONTH_NAMES = (
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
)
TEACHING_NAME_ORDERS = ("Preceptor ~ Student", "Student ~ Preceptor")

# Optional, explicit spelling aliases. Never guess that two different names are
# the same person. Matching ignores case, repeated spaces and spaces by commas.
# Do NOT map a rotating generic slot (e.g. SJR_1) to one person for every date.
TEACHING_PRECEPTOR_NAME_MAP = {
    # "Smith, J.": "Smith, Jane",
}

# Optional exceptions for a rotation whose name order differs from the selection
# on the summary page. Keys are the first Monday shown INSIDE each OPD.
TEACHING_OPD_NAME_ORDER_OVERRIDES = {
    # "2026-08-03": "Preceptor ~ Student",
}

# Labels that do not identify an assigned student. Case-insensitive exact match.
TEACHING_EMPTY_STUDENT_LABELS = {
    "", "nan", "none", "n/a", "na", "tbd", "unassigned", "no student",
    "no students", "open", "available", "off", "vacation", "holiday",
}


def teaching_name_key(value):
    cleaned = re.sub(r"\s+", " ", str(value or "").strip())
    return re.sub(r"\s*,\s*", ", ", cleaned).casefold()


def teaching_display_name(value):
    return re.sub(r"\s*,\s*", ", ", re.sub(r"\s+", " ", str(value or "").strip()))


def teaching_local_today():
    # Streamlit Cloud may run on UTC. Keep the academic-year boundary local.
    try:
        return datetime.now(_TeachingZoneInfo("America/New_York")).date()
    except Exception:
        return datetime.now(_teaching_timezone.utc).date()


def teaching_academic_start(day):
    return day.year if day.month >= 7 else day.year - 1


def teaching_academic_label(start_year):
    return f"{start_year % 100:02d}-{(start_year + 1) % 100:02d}"


def teaching_month_label(month):
    return f"{TEACHING_MONTH_NAMES[month.month - 1]} {month.year}"


def teaching_label_needs_review(name, site_names=()):
    key = re.sub(r"[^A-Z0-9]+", "_", str(name).upper()).strip("_")
    site_keys = {re.sub(r"[^A-Z0-9]+", "_", str(s).upper()).strip("_") for s in site_names}
    return (
        key in site_keys or key in {"HAMPDEN_NURSERY", "PSHCH_NURSERY", "SJR_HOSPITALIST", "WARD_A"}
        or bool(re.fullmatch(r"(?:SJR|AAC|HOPE_DRIVE|NYES|ETOWN|LANCASTER|HAMPDEN_NURSERY|PSHCH_NURSERY)_?\d+", key))
        or key in {"TBD", "UNKNOWN", "PRECEPTOR", "PROVIDER", "TO_BE_ASSIGNED"}
        or bool(re.search(r"[;&|]|\s/\s|\s+and\s+", str(name), re.I))
    )


def teaching_split_assignment(value, order):
    """Return (provider, [students]) groups without splitting 'Last, First'.

    Names are used only temporarily to count separate students and recognize
    exact duplicates. They never enter the exported CSV, Word files or notes.
    Multiple students: separate OPD rows, or ; / newline / | / ' & ' / ' and '
    on the student side. Repeated full Provider ~ Student pairs may be separated
    with a newline, semicolon or |. A bare repeated ~ is ambiguous and rejected.
    """
    if order not in TEACHING_NAME_ORDERS:
        raise OPDArchiveError("Choose Preceptor ~ Student or Student ~ Preceptor for the teaching summary.")
    if not isinstance(value, str) or "~" not in value:
        return []
    text = value.strip()
    if text.count("~") == 1:
        segments = [text]
    else:
        segments = [part.strip() for part in re.split(r"[\r\n;|]+", text) if part.strip()]
        if any(part.count("~") != 1 for part in segments):
            raise OPDArchiveError("An assignment contains an ambiguous repeated '~'. Use separate OPD rows "
                                  "or a semicolon-separated list of students on the student side.")
    parsed = []
    for segment in segments:
        left, right = (part.strip() for part in segment.split("~", 1))
        provider, student_text = (left, right) if order == "Preceptor ~ Student" else (right, left)
        students = [teaching_display_name(part) for part in
                    re.split(r"[\r\n;|]+|\s+&\s+|\s+and\s+", student_text, flags=re.I)]
        students = [name for name in students if teaching_name_key(name) not in TEACHING_EMPTY_STUDENT_LABELS]
        parsed.append((teaching_display_name(provider), students))
    return parsed


def teaching_extract_assignments(raw, details, order):
    """Read AM/PM cells under each *actual* date, including hidden OPD rows.

    Returns temporary records containing student names. The caller aggregates
    and discards these records before storing anything in Streamlit session_state.
    """
    wb = load_workbook(BytesIO(raw), read_only=True, data_only=False)
    records, missing_provider_cells = [], []
    expected_days = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
    try:
        for sheet_name in details["site_names"]:
            ws = wb[sheet_name]
            if (ws.max_row or 0) > 1200:
                raise OPDArchiveError(f"Worksheet '{sheet_name}' exceeds the supported 1,200-row OPD layout; "
                                      "no partial teaching report was generated.")
            rows = list(ws.iter_rows(max_col=8))
            current_dates = None
            for row_index, row in enumerate(rows):
                values = [cell.value for cell in row]
                if [str(v or "").strip().casefold() for v in values[1:8]] == expected_days:
                    if row_index + 1 >= len(rows):
                        raise OPDArchiveError(f"Missing dates on worksheet '{sheet_name}'.")
                    current_dates = [_opd_date(cell.value, wb.epoch) for cell in rows[row_index + 1][1:8]]
                    if not all(current_dates) or any(
                        day != current_dates[0] + timedelta(days=i) for i, day in enumerate(current_dates)
                    ):
                        raise OPDArchiveError(f"Invalid date row on worksheet '{sheet_name}'.")
                    continue
                match = re.match(r"^\s*(AM|PM)\b", str(values[0] or ""), re.I)
                if not match:
                    continue
                for col_index, cell in enumerate(row[1:8]):
                    if cell.data_type == "f":
                        raise OPDArchiveError(f"{sheet_name}!{cell.coordinate} contains a formula in a session cell. "
                                              "Use assignment values rather than formulas for this summary.")
                    try:
                        groups = teaching_split_assignment(cell.value, order)
                    except OPDArchiveError as exc:
                        raise OPDArchiveError(f"{sheet_name}!{cell.coordinate}: {exc}") from None
                    for provider, students in groups:
                        if not students:
                            continue  # provider availability alone earns no teaching hours
                        if teaching_name_key(provider) in TEACHING_EMPTY_STUDENT_LABELS:
                            missing_provider_cells.append(f"{sheet_name}!{cell.coordinate}")
                            continue
                        if current_dates is None:
                            raise OPDArchiveError(f"{sheet_name}!{cell.coordinate} has an assignment before a date header.")
                        for student in students:
                            records.append({
                                "preceptor_name": provider, "student": student,
                                "day": current_dates[col_index], "shift": match.group(1).upper(),
                                "site": sheet_name, "cell": cell.coordinate,
                            })
        return records, missing_provider_cells
    finally:
        wb.close()


def teaching_scan_archives(client, default_order=TEACHING_NAME_ORDERS[0], progress=None):
    """Use one repository snapshot; read/decrypt every current OPD once.

    An unreadable file aborts the run. Never return partial totals as complete.
    Files are streamed one rotation at a time; no raw workbooks or students are
    retained in the returned object. Student IDs for deduplication are keyed with
    a fresh, run-local random key and then discarded, not saved or exported.
    """
    if default_order not in TEACHING_NAME_ORDERS:
        raise OPDArchiveError("Select a supported OPD name order.")
    commit = client._head()
    rotations = client.list_rotations(commit=commit)
    counts = Counter()
    names, name_variants, review_labels, manifests, warnings = {}, defaultdict(set), set(), [], []
    seen_assignments = set()
    salt = _teaching_secrets.token_bytes(32)
    aliases = {teaching_name_key(k): teaching_display_name(v)
               for k, v in TEACHING_PRECEPTOR_NAME_MAP.items() if str(k).strip() and str(v).strip()}
    duplicates = 0
    future_assignments = 0
    today = teaching_local_today()
    for number, rotation in enumerate(rotations, start=1):
        order = TEACHING_OPD_NAME_ORDER_OVERRIDES.get(rotation.isoformat(), default_order)
        if order not in TEACHING_NAME_ORDERS:
            raise OPDArchiveError(f"Invalid teaching-summary name-order override for rotation {rotation.isoformat()}.")
        try:
            loaded = client.load(rotation, commit=commit)
            records, missing_cells = teaching_extract_assignments(loaded["raw"], loaded["details"], order)
        except OPDArchiveError as exc:
            raise OPDArchiveError(f"Rotation {rotation.isoformat()}: {exc} No ZIP was generated from partial data.") from None
        counted = removed = 0
        for item in records:
            original_name = item["preceptor_name"]
            name = aliases.get(teaching_name_key(original_name), original_name)
            key = teaching_name_key(name)
            # Source/site are deliberately absent: the same provider + student +
            # date + half-day duplicated in another site/rotation is one assignment.
            # Different students in the SAME half-day have DIFFERENT identifiers.
            identity = json.dumps([key, teaching_name_key(item["student"]),
                                   item["day"].isoformat(), item["shift"]], ensure_ascii=False)
            digest = hmac.new(salt, identity.encode("utf-8"), hashlib.sha256).digest()
            if digest in seen_assignments:
                duplicates += 1
                removed += 1
                continue
            seen_assignments.add(digest)
            names.setdefault(key, name)
            name_variants[key].add(original_name)
            month = item["day"].replace(day=1)
            counts[(key, month)] += 1
            counted += 1
            future_assignments += int(item["day"] > today)
            if teaching_label_needs_review(name, loaded["details"]["site_names"]):
                review_labels.add(key)
        if missing_cells:
            warnings.append({"rotation_start": rotation.isoformat(), "issue": "Missing provider; not attributed",
                             "details": ", ".join(sorted(set(missing_cells)))})
        manifests.append({
            "rotation_start": rotation.isoformat(),
            "last_scheduled_date": (loaded["details"]["week_mondays"][-1] + timedelta(days=6)).isoformat(),
            "archive_file": loaded["path"].rsplit("/", 1)[-1],
            "github_blob_sha": loaded["sha"], "name_order": order,
            "assigned_student_shifts_read": len(records),
            "assigned_student_shifts_counted": counted,
            "duplicate_student_shifts_removed": removed,
            "missing_provider_cells": len(set(missing_cells)),
        })
        # Only provider-level aggregates remain after each rotation is processed.
        del records, loaded
        if progress:
            progress(number, len(rotations))
    monthly = [{"preceptor_name": names[key], "academic_year": teaching_academic_label(teaching_academic_start(month)),
                "academic_start_year": teaching_academic_start(month), "month": month.isoformat(),
                "no_of_shifts": int(count), "educational_hours": int(count * TEACHING_HOURS_PER_STUDENT_SHIFT)}
               for (key, month), count in sorted(counts.items(), key=lambda item: (item[0][0], item[0][1]))]
    return {
        "version": TEACHING_REPORT_VERSION, "commit": commit,
        "generated_at": datetime.now(_teaching_timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
        "default_name_order": default_order, "monthly": monthly, "sources": manifests,
        "warnings": warnings, "duplicate_assignments_removed": duplicates,
        "future_assignments_in_archive": future_assignments,
        "unresolved_preceptor_labels": sorted((names[k] for k in review_labels), key=teaching_name_key),
        "name_variants": {names[k]: sorted(v) for k, v in name_variants.items() if len(v) > 1},
    }


def teaching_annual_rows(scan, selected_years):
    selected = {int(year) for year in selected_years}
    grouped = defaultdict(list)
    for item in scan["monthly"]:
        if item["academic_start_year"] in selected:
            grouped[(item["preceptor_name"], item["academic_start_year"])].append(item)
    rows = []
    for (name, start_year), items in sorted(grouped.items(), key=lambda item: (teaching_name_key(item[0][0]), item[0][1])):
        ordered = sorted(items, key=lambda item: item["month"])
        total = sum(item["no_of_shifts"] for item in ordered)
        rows.append({"preceptor_name": name, "academic_year": teaching_academic_label(start_year),
                     "no_of_shifts": total,
                     "months_worked": "; ".join(teaching_month_label(CalendarDate.fromisoformat(item["month"])) for item in ordered),
                     "educational_hours": total * TEACHING_HOURS_PER_STUDENT_SHIFT})
    return rows


def teaching_csv_bytes(rows, columns=TEACHING_CSV_COLUMNS):
    stream = io.StringIO(newline="")
    writer = csv.DictWriter(stream, fieldnames=list(columns), extrasaction="ignore", lineterminator="\r\n")
    writer.writeheader()
    for row in rows:
        cleaned = {}
        for column in columns:
            value = row.get(column, "")
            # Keep name text from becoming an Excel formula when CSV is opened.
            if isinstance(value, str) and value.lstrip().startswith(("=", "+", "-", "@")):
                value = "'" + value
            cleaned[column] = value
        writer.writerow(cleaned)
    return stream.getvalue().encode("utf-8-sig")


def teaching_make_docx(name, monthly, scan):
    """One editable Word document per preceptor; one page per academic year."""
    from docx.shared import Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Inches(8.5), Inches(11)
    section.top_margin = section.bottom_margin = Inches(0.7)
    section.left_margin = section.right_margin = Inches(0.8)
    normal = doc.styles["Normal"]
    normal.font.name, normal.font.size = "Calibri", Pt(11)
    normal.paragraph_format.space_after = Pt(7)
    for style_name in ("Title", "Heading 1", "Heading 2"):
        doc.styles[style_name].font.name = "Calibri"
        doc.styles[style_name].font.color.rgb = RGBColor.from_string("24466B")
    doc.styles["Title"].font.size = Pt(24)
    doc.styles["Heading 1"].font.size = Pt(16)
    header = section.header.paragraphs[0]
    header.text = "PENN STATE  |  PEDIATRIC CLERKSHIP"
    header.runs[0].font.size = Pt(9)
    header.runs[0].font.color.rgb = RGBColor.from_string("526475")
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    footer.add_run("Third-year student teaching  |  Page ").font.size = Pt(9)
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    footer._p.append(field)
    doc.core_properties.title = f"Preceptor teaching report - {name}"
    doc.core_properties.author = "Pediatric Clerkship"
    doc.core_properties.subject = "Scheduled student-shifts and student-weighted educational hours"

    grouped = defaultdict(list)
    for row in monthly:
        grouped[row["academic_start_year"]].append(row)
    for index, (year, rows) in enumerate(sorted(grouped.items())):
        if index:
            doc.add_page_break()
        doc.add_paragraph("Preceptor teaching report", style="Subtitle")
        doc.add_paragraph(name, style="Title")
        doc.add_heading(f"Academic year {teaching_academic_label(year)}", level=1)
        doc.add_paragraph(f"July 1, {year} - June 30, {year + 1}")
        if name in scan["unresolved_preceptor_labels"]:
            warning = doc.add_paragraph()
            run = warning.add_run("Review required: this is a site, slot, or combined provider label, not a verified individual preceptor.")
            run.bold = True
            run.font.color.rgb = RGBColor.from_string("8C3B25")
        count = sum(row["no_of_shifts"] for row in rows)
        p = doc.add_paragraph()
        p.add_run("Assigned student-shifts: ").bold = True
        p.add_run(f"{count:,}")
        p = doc.add_paragraph()
        p.add_run("Student-weighted educational hours: ").bold = True
        p.add_run(f"{count * TEACHING_HOURS_PER_STUDENT_SHIFT:,}")
        doc.add_heading("Monthly assignments", level=2)
        table = doc.add_table(rows=1, cols=3)
        table.style = "Light Shading Accent 1"
        table.autofit = False
        widths = (Inches(2.5), Inches(2.2), Inches(2.2))
        for cell, title, width in zip(table.rows[0].cells, ("Month", "Student-shifts", "Educational hours"), widths):
            cell.text, cell.width = title, width
            for run in cell.paragraphs[0].runs:
                run.bold = True
        repeat_header = OxmlElement("w:tblHeader")
        table.rows[0]._tr.get_or_add_trPr().append(repeat_header)
        for row in sorted(rows, key=lambda item: item["month"]):
            values = (teaching_month_label(CalendarDate.fromisoformat(row["month"])),
                      f"{row['no_of_shifts']:,}", f"{row['educational_hours']:,}")
            for cell, value, width in zip(table.add_row().cells, values, widths):
                cell.text, cell.width = value, width
        for cell, value, width in zip(table.add_row().cells,
                                      ("Total", f"{count:,}", f"{count * TEACHING_HOURS_PER_STUDENT_SHIFT:,}"), widths):
            cell.text, cell.width = value, width
            for run in cell.paragraphs[0].runs:
                run.bold = True
        for row in table.rows:
            row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
            for i, cell in enumerate(row.cells):
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.space_before = Pt(3)
                    paragraph.paragraph_format.space_after = Pt(3)
                    if i:
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        doc.add_paragraph()
        note = doc.add_paragraph(
            f"Counting method: one student assigned to one AM or PM shift counts as one student-shift and "
            f"{TEACHING_HOURS_PER_STUDENT_SHIFT} educational hours. Two different students in the same shift count "
            "as two student-shifts and eight hours. These are student-weighted scheduled educational hours, "
            "not distinct clock hours or confirmed attendance. Identical duplicate assignments count once."
        )
        for run in note.runs:
            run.font.size = Pt(9)
        source = doc.add_paragraph(
            "Source: current encrypted OPD archive files, decrypted for this report. "
            f"Archive snapshot: {scan['commit'][:12]}. Retrieved: {scan['generated_at']}. "
            "Only months with assignments are shown. Student names are omitted."
        )
        for run in source.runs:
            run.font.size = Pt(9)
    output = BytesIO()
    doc.save(output)
    return output.getvalue()


# The chair summary is an additional output; the CSV and individual reports use
# the same existing aggregates and counting rules as before.
TEACHING_CHAIR_SUMMARY_FILENAME = "Pediatric_Clerkship_Educational_Effort_Summary.docx"


def teaching_brief_months(month_values):
    """Compact month labels without implying teaching occurred in missing months."""
    months = sorted({CalendarDate.fromisoformat(value).replace(day=1) for value in month_values})
    if not months:
        return "Not recorded"
    groups, first, last = [], months[0], months[0]
    for month in months[1:]:
        if month.year == last.year and month.month == last.month + 1:
            last = month
        else:
            groups.append((first, last))
            first = last = month
    groups.append((first, last))
    labels = []
    for first, last in groups:
        label = TEACHING_MONTH_NAMES[first.month - 1][:3]
        if first != last:
            label += "-" + TEACHING_MONTH_NAMES[last.month - 1][:3]
        labels.append(f"{label} {first.year}")
    return "; ".join(labels)


def teaching_chair_summary_data(scan, selected_years):
    """Year-by-year summaries using the exact rows used for the existing CSV.

    Generic site/slot/combined labels remain in overall assignment totals, but
    are separated from named preceptors so they are not presented as people.
    """
    years = sorted({int(year) for year in selected_years})
    annual = teaching_annual_rows(scan, years)
    review_keys = {teaching_name_key(name) for name in scan.get("unresolved_preceptor_labels", [])}
    summaries = []
    for year in years:
        label = teaching_academic_label(year)
        monthly = [row for row in scan["monthly"] if row["academic_start_year"] == year]
        by_name = defaultdict(list)
        for row in monthly:
            by_name[row["preceptor_name"]].append(row["month"])
        named, unresolved = [], []
        for row in annual:
            if row["academic_year"] != label:
                continue
            entry = dict(row)
            entry["months_brief"] = teaching_brief_months(by_name[row["preceptor_name"]])
            target = unresolved if teaching_name_key(row["preceptor_name"]) in review_keys else named
            target.append(entry)
        start, end = CalendarDate(year, 7, 1), CalendarDate(year + 1, 6, 30)
        relevant_sources = [source for source in scan.get("sources", [])
                            if CalendarDate.fromisoformat(source["rotation_start"]) <= end
                            and CalendarDate.fromisoformat(source["last_scheduled_date"]) >= start]
        rows = named + unresolved
        summaries.append({
            "academic_start_year": year,
            "academic_year": label,
            "named_preceptors": sorted(named, key=lambda row: teaching_name_key(row["preceptor_name"])),
            "unresolved_labels": sorted(unresolved, key=lambda row: teaching_name_key(row["preceptor_name"])),
            "named_preceptor_count": len(named),
            "no_of_shifts": sum(row["no_of_shifts"] for row in rows),
            "educational_hours": sum(row["educational_hours"] for row in rows),
            "months_brief": teaching_brief_months(row["month"] for row in monthly),
            "source_count": len(relevant_sources),
            # Cells may contain more than one learner; this is not a shift count.
            "has_missing_provider": any(int(source.get("missing_provider_cells", 0)) > 0
                                        for source in relevant_sources),
        })
    return summaries


def teaching_make_chair_summary(scan, selected_years):
    """One editable, chair-friendly Word report covering all selected years.

    Uses scheduled student-shifts, not unique students, elapsed hours, patient
    encounters, or a claim that all planned teaching has actually occurred.
    """
    from docx.shared import Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    summaries = teaching_chair_summary_data(scan, selected_years)
    if not summaries or not any(item["no_of_shifts"] for item in summaries):
        raise OPDArchiveError("No student assignments were found for the selected academic year(s); no empty chair summary was generated.")

    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Inches(8.5), Inches(11)
    section.top_margin = section.bottom_margin = Inches(0.7)
    section.left_margin = section.right_margin = Inches(0.8)
    section.header_distance = section.footer_distance = Inches(0.3)
    normal = doc.styles["Normal"]
    normal.font.name, normal.font.size = "Calibri", Pt(11)
    normal.paragraph_format.space_after = Pt(5)
    normal.paragraph_format.line_spacing = 1.05
    for name, size in (("Title", 23), ("Heading 1", 16), ("Heading 2", 12)):
        style = doc.styles[name]
        style.font.name, style.font.size = "Calibri", Pt(size)
        style.font.color.rgb = RGBColor.from_string("24466B")
        style.paragraph_format.keep_with_next = True
        style.paragraph_format.space_before = Pt(8 if name == "Heading 2" else 0)
        style.paragraph_format.space_after = Pt(4)
        properties = style.element.find(qn("w:pPr"))
        if properties is not None:
            borders = properties.find(qn("w:pBdr"))
            if borders is not None:
                properties.remove(borders)
    subtitle = doc.styles["Subtitle"]
    subtitle.font.name, subtitle.font.size = "Calibri", Pt(12)
    subtitle.font.italic = False
    subtitle.font.color.rgb = RGBColor.from_string("526475")
    subtitle.paragraph_format.space_before = Pt(0)
    subtitle.paragraph_format.space_after = Pt(12)
    subtitle.paragraph_format.keep_with_next = True
    header = section.header.paragraphs[0]
    header.text = "PENN STATE  |  PEDIATRIC CLERKSHIP"
    header.runs[0].font.size = Pt(9)
    header.runs[0].font.color.rgb = RGBColor.from_string("526475")
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    footer.add_run("Educational effort summary  |  Page ").font.size = Pt(9)
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    footer._p.append(field)
    doc.core_properties.title = "Pediatric clerkship educational effort summary"
    doc.core_properties.author = "Pediatric Clerkship"
    doc.core_properties.subject = "Scheduled third-year student teaching by academic year"

    def note(text, *, warning=False):
        p = doc.add_paragraph(text)
        for run in p.runs:
            run.font.size = Pt(9)
            run.font.color.rgb = RGBColor.from_string("8C3B25" if warning else "526475")
        return p

    def add_effort_table(entries, *, pending=False):
        table = doc.add_table(rows=1, cols=4)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        # 6.9 inches = the printable width; also set the table grid, not only cells.
        widths = tuple(Inches(value) for value in (2.45, 2.15, 1.05, 1.25))
        for col, width in zip(table.columns, widths):
            col.width = width
        titles = ("Provider label" if pending else "Preceptor", "Months with assignments",
                  "Student-shifts", "Educational hours*")
        repeat_header = OxmlElement("w:tblHeader")
        table.rows[0]._tr.get_or_add_trPr().append(repeat_header)

        def fill_row(row, values, *, heading=False, total=False, band=False):
            row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
            for index, (cell, value, width) in enumerate(zip(row.cells, values, widths)):
                cell.width = width
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                cell.text = str(value)
                properties = cell._tc.get_or_add_tcPr()
                margin = OxmlElement("w:tcMar")
                for edge, amount in (("top", 40), ("bottom", 40), ("left", 90), ("right", 90)):
                    item = OxmlElement("w:" + edge)
                    item.set(qn("w:w"), str(amount))
                    item.set(qn("w:type"), "dxa")
                    margin.append(item)
                properties.append(margin)
                if heading or total or band:
                    shade = OxmlElement("w:shd")
                    shade.set(qn("w:fill"), "24466B" if heading else "E8EEF5" if total else "F5F7FA")
                    properties.append(shade)
                p = cell.paragraphs[0]
                p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(1)
                p.paragraph_format.line_spacing = 1.0
                p.paragraph_format.keep_with_next = heading
                if index >= 2:
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                for run in p.runs:
                    run.font.name, run.font.size = "Calibri", Pt(10)
                    run.bold = heading or total
                    if heading:
                        run.font.color.rgb = RGBColor(255, 255, 255)

        fill_row(table.rows[0], titles, heading=True)
        for index, row in enumerate(entries):
            fill_row(table.add_row(), (row["preceptor_name"], row["months_brief"],
                                      f"{row['no_of_shifts']:,}", f"{row['educational_hours']:,}"),
                     band=index % 2 == 1)
        subtotal = "Awaiting attribution" if pending else "Named preceptors total"
        fill_row(table.add_row(), (subtotal, "", f"{sum(row['no_of_shifts'] for row in entries):,}",
                                   f"{sum(row['educational_hours'] for row in entries):,}"), total=True)
        # Avoid leaving the subtotal by itself at the top of a new page.
        if len(table.rows) > 2:
            for cell in table.rows[-2].cells:
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.keep_with_next = True
        return table

    for index, item in enumerate(summaries):
        year = item["academic_start_year"]
        if index:
            doc.add_page_break()
        doc.add_paragraph("Preceptor educational effort", style="Title")
        doc.add_paragraph("Third-year medical student teaching", style="Subtitle")
        doc.add_heading(f"Academic year {item['academic_year']}", level=1)
        p = doc.add_paragraph(f"July 1, {year} - June 30, {year + 1}")
        p.paragraph_format.space_after = Pt(9)
        if not item["no_of_shifts"]:
            doc.add_paragraph("No assigned student-shifts were found in the archived schedules for this academic year. "
                              "This does not establish that no teaching occurred.")
            continue

        p = doc.add_paragraph()
        p.add_run("Archived OPD schedules record ")
        p.add_run(f"{item['no_of_shifts']:,} student-shifts").bold = True
        p.add_run(", representing ")
        p.add_run(f"{item['educational_hours']:,} educational hours").bold = True
        p.add_run(" of scheduled teaching.")
        p = doc.add_paragraph()
        p.add_run("Named preceptors: ").bold = True
        p.add_run(str(item["named_preceptor_count"]))
        p.add_run("   |   Months with assignments: ").bold = True
        p.add_run(item["months_brief"])
        if item["unresolved_labels"]:
            pending_shifts = sum(row["no_of_shifts"] for row in item["unresolved_labels"])
            pending_hours = sum(row["educational_hours"] for row in item["unresolved_labels"])
            note(f"The totals include {pending_shifts:,} student-shifts ({pending_hours:,} hours) recorded under "
                 "site, slot, or combined provider labels. These are listed separately below and are not credited to an individual.",
                 warning=True)

        note(f"*One student assigned to one AM or PM shift = one student-shift and "
             f"{TEACHING_HOURS_PER_STUDENT_SHIFT} educational hours. Two students in the same shift count twice. "
             "These are student-weighted hours, not distinct clock hours or verified attendance.")
        note(f"Coverage: {item['source_count']} saved rotation schedule(s) overlap this academic year. "
             "Only archived assignments are represented; missing rotations are not assumed to have no teaching. "
             "Future scheduled assignments are included.")
        if item["has_missing_provider"]:
            note("Data review: at least one source rotation overlapping this year contains assignments without "
                 "an identifiable preceptor. Those assignments are excluded from provider totals; review Report_Notes.txt.",
                 warning=True)

        if item["named_preceptors"]:
            doc.add_heading("Educational effort by preceptor", level=2)
            add_effort_table(item["named_preceptors"])
        else:
            doc.add_paragraph("No assignments in this selection were recorded under an individual preceptor name.")
        if item["unresolved_labels"]:
            doc.add_heading("Assignments awaiting an individual preceptor name", level=2)
            add_effort_table(item["unresolved_labels"], pending=True)
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            p.add_run("Combined recorded total: ").bold = True
            p.add_run(f"{item['no_of_shifts']:,} student-shifts | {item['educational_hours']:,} educational hours")
        source = note("Source: current saved OPD schedules. "
                      f"Archive retrieved: {scan['generated_at']}. "
                      "Alphabetical listing; student names omitted. File-level source details accompany this report in the ZIP.")
        source.paragraph_format.space_before = Pt(6)

    output = BytesIO()
    doc.save(output)
    return output.getvalue()


def teaching_build_zip(scan, selected_years):
    """Return a ZIP and annual preview. Does not call GitHub or write plaintext there."""
    years = sorted({int(year) for year in selected_years})
    annual = teaching_annual_rows(scan, years)
    if not annual:
        raise OPDArchiveError("No student assignments were found for the selected academic year(s); no empty report was generated.")
    monthly_by_name = defaultdict(list)
    for item in scan["monthly"]:
        if item["academic_start_year"] in years:
            monthly_by_name[item["preceptor_name"]].append(item)
    labels = ", ".join(teaching_academic_label(year) for year in years)
    output = BytesIO()
    notes = [
        "PRECEPTOR TEACHING SUMMARY", f"Selected academic year(s): {labels}",
        f"Archive retrieved: {scan['generated_at']}", f"Repository snapshot: {scan['commit']}",
        f"Current OPD files read: {len(scan['sources'])}", "",
        "Source: every current OPD_YYYY-MM-DD.xlsx.enc in the configured archive folder.",
        "Superseded Git history is not counted. The archived files are not changed.",
        "All recognized OPD site worksheets are included, not just HOPE_DRIVE, NYES and ETOWN.",
        "This report counts scheduled student assignments, not patient encounters or confirmed attendance.",
        "One row in preceptor_teaching_summary.csv = one preceptor in one academic year.",
        f"{TEACHING_CHAIR_SUMMARY_FILENAME} = one combined Word summary for the chair.",
        "The chair summary lists named preceptors alphabetically and keeps unresolved provider labels separate.",
        "Academic year is July 1 through June 30 of the following calendar year.",
        "The actual session date, not the rotation's start date, determines the month and academic year.",
        "no_of_shifts counts student-shifts, not distinct half-days or distinct students.",
        f"educational_hours = no_of_shifts x {TEACHING_HOURS_PER_STUDENT_SHIFT}.",
        "Two different students in the same AM/PM session count as two assignments and eight hours.",
        "An identical preceptor/student/date/AM-or-PM duplicate counts once, even across overlapping OPDs.",
        "Provider availability with no student does not count. All filled student assignments are treated equally.",
        "Names are matched case-insensitively with normalized whitespace and comma spacing; no fuzzy identity matching.",
        "Student names and decrypted workbooks are not included in this ZIP.",
        "Future scheduled assignments in the selected academic year(s) are included.",
        "Providers with no assignments in the selected year(s) do not receive a report.",
        "The summary remains a snapshot until Load / refresh archived OPDs is clicked again.", "",
        "DATA QUALITY (the following diagnostics cover ALL scanned academic years)",
        f"Exact duplicate student-shifts removed: {scan['duplicate_assignments_removed']}",
        f"Future student-shifts in the entire archive when scanned: {scan['future_assignments_in_archive']}",
    ]
    if scan["unresolved_preceptor_labels"]:
        notes += ["Provider/site/slot labels needing review (retained literally, not assigned to a guessed individual):"]
        notes += ["  " + name for name in scan["unresolved_preceptor_labels"]]
    for item in scan["warnings"]:
        notes.append(f"Rotation {item['rotation_start']}: {item['issue']}: {item['details']}")
    source_columns = (
        "rotation_start", "last_scheduled_date", "archive_file", "github_blob_sha", "name_order",
        "assigned_student_shifts_read", "assigned_student_shifts_counted",
        "duplicate_student_shifts_removed", "missing_provider_cells",
    )
    with ZipFile(output, "w", compression=ZIP_DEFLATED) as zf:
        zf.writestr(TEACHING_CHAIR_SUMMARY_FILENAME, teaching_make_chair_summary(scan, years))
        zf.writestr("preceptor_teaching_summary.csv", teaching_csv_bytes(annual))
        used = set()
        for name in sorted(monthly_by_name, key=teaching_name_key):
            base = re.sub(r"[^A-Za-z0-9._-]+", "_", name).strip("._")[:110] or "preceptor"
            safe = base
            number = 1
            while safe.casefold() in used:
                number += 1
                safe = f"{base}_{number}"
            used.add(safe.casefold())
            zf.writestr(f"Preceptor_Reports/{safe}_Teaching_Report.docx",
                        teaching_make_docx(name, monthly_by_name[name], scan))
        zf.writestr("Report_Notes.txt", "\n".join(notes).encode("utf-8"))
        zf.writestr("Archive_Sources.csv", teaching_csv_bytes(scan["sources"], source_columns))
    return output.getvalue(), annual


def render_preceptor_teaching_summary():
    st.subheader("Preceptor Teaching Summary")
    st.write("Read current encrypted OPDs from GitHub, then generate a chair-friendly Word summary, "
             "the summary CSV, and one Word teaching report per preceptor.")
    st.caption("All OPD sites. One student assigned to one AM/PM session = one student-shift = four educational hours. "
               "Two students at once count twice. Hours are student-weighted scheduled hours, not unique clock hours.")
    current = teaching_academic_start(teaching_local_today())
    st.info(f"Current academic year: {teaching_academic_label(current)} (July 1, {current} through June 30, {current + 1}).")
    try:
        client = GitHubOPDArchive(get_opd_archive_config())
    except OPDArchiveError as exc:
        st.error(str(exc))
        return
    order = st.selectbox("Names around '~' in archived OPDs", TEACHING_NAME_ORDERS,
                          key="teaching_name_order",
                          help="The sample OPD uses Preceptor ~ Student. No rotation list is required. "
                               "Choose the actual name order in your archive. Multiple students may be on separate rows "
                               "or separated by semicolons/newlines on the student side; commas inside names are preserved.")
    options_signature = hashlib.sha256(json.dumps(
        [TEACHING_REPORT_VERSION, client.config.signature(), order, TEACHING_PRECEPTOR_NAME_MAP,
         TEACHING_OPD_NAME_ORDER_OVERRIDES], sort_keys=True).encode()).hexdigest()
    if st.session_state.get("teaching_options_signature") != options_signature:
        for key in ("teaching_scan", "teaching_zip", "teaching_zip_signature", "teaching_selected_years"):
            st.session_state.pop(key, None)
        st.session_state["teaching_options_signature"] = options_signature
    if st.button("Load / refresh archived OPDs", key="teaching_load_archives", type="primary"):
        for key in ("teaching_scan", "teaching_zip", "teaching_zip_signature", "teaching_selected_years"):
            st.session_state.pop(key, None)
        bar = st.progress(0, text="Reading the current encrypted archive...")
        try:
            with st.spinner("Decrypting OPDs and counting student assignments..."):
                scan = teaching_scan_archives(
                    client, order,
                    progress=lambda n, total: bar.progress(n / total, text=f"Read {n} of {total} current OPD files"),
                )
            st.session_state["teaching_scan"] = scan
        except OPDArchiveError as exc:
            st.error(str(exc))
        except Exception:
            # Do not echo workbook contents, credentials or learner data in the UI.
            st.error("The teaching summary could not be completed. No partial ZIP was created. "
                     "Check the workbook layout and installed requirements, then retry.")
        finally:
            bar.empty()
    scan = st.session_state.get("teaching_scan")
    if scan is None:
        st.caption("Nothing is downloaded or decrypted until you click Load / refresh archived OPDs. "
                   "This page never saves a report or decrypted OPD to GitHub.")
        return
    if not scan["sources"]:
        st.warning("No current encrypted OPDs were found in the configured GitHub archive folder.")
        return
    st.success(f"Read and decrypted {len(scan['sources'])} current OPD files. Archive snapshot: {scan['generated_at']}.")
    st.caption("Upload revisions in Create Student Schedule as usual. Click Load / refresh again to include newer saves. "
               "Only current files are counted, not their older Git versions.")
    years = sorted({row["academic_start_year"] for row in scan["monthly"]} | {current}, reverse=True)
    selected = st.multiselect("Academic year(s) to include", years, default=[current],
                              format_func=teaching_academic_label, key="teaching_selected_years",
                              help="The current year is selected by default. Select other available years to add "
                                   "a separate section for each year in the chair summary and every preceptor's Word report.")
    signature = (options_signature, scan["commit"], tuple(sorted(selected)))
    if st.session_state.get("teaching_zip_signature") != signature:
        st.session_state.pop("teaching_zip", None)
        st.session_state["teaching_zip_signature"] = signature
    annual = teaching_annual_rows(scan, selected)
    if scan["duplicate_assignments_removed"]:
        st.warning(f"{scan['duplicate_assignments_removed']} identical student-shift duplicates were counted once. "
                   "Different students assigned in the same shift still count separately.")
    if scan["unresolved_preceptor_labels"]:
        st.warning("Some assigned providers are recorded as site/slot/combined labels rather than individual names. "
                   "They are retained literally and marked for review, not attributed to a guessed person.")
        with st.expander("Provider labels to review"):
            st.write(", ".join(scan["unresolved_preceptor_labels"]))
    if scan["warnings"]:
        st.warning("Some filled assignment cells have no identifiable provider and cannot be attributed. "
                   "See the data-quality details and Report_Notes.txt.")
    with st.expander("Archive coverage and data-quality details"):
        st.dataframe(pd.DataFrame(scan["sources"]), hide_index=True, use_container_width=True)
        if scan["warnings"]:
            st.dataframe(pd.DataFrame(scan["warnings"]), hide_index=True, use_container_width=True)
    if not annual:
        st.info("No assigned student-shifts were found for the selected academic year(s). "
                "Choose another available year, check the '~' name order, or archive OPDs with completed student assignments. "
                "Provider availability without a student does not count.")
        return
    preview = pd.DataFrame(annual, columns=TEACHING_CSV_COLUMNS)
    a, b, c = st.columns(3)
    a.metric("Preceptors / provider labels", preview["preceptor_name"].nunique())
    b.metric("Assigned student-shifts", f"{preview['no_of_shifts'].sum():,}")
    c.metric("Student-weighted educational hours", f"{preview['educational_hours'].sum():,}")
    st.dataframe(preview, hide_index=True, use_container_width=True)
    st.caption("The CSV contains one row per preceptor per academic year. Months follow the actual assignment dates. "
               "Student names are not included in the CSV, Word reports, or source notes.")
    if st.button("Create teaching reports ZIP", key="teaching_build_zip"):
        st.session_state.pop("teaching_zip", None)
        try:
            with st.spinner("Creating chair summary, CSV and individual Word reports..."):
                zip_bytes, _ = teaching_build_zip(scan, selected)
                st.session_state["teaching_zip"] = zip_bytes
                st.session_state["teaching_zip_signature"] = signature
        except OPDArchiveError as exc:
            st.error(str(exc))
        except Exception:
            st.error("The Word/CSV export could not be completed. No partial ZIP was retained. "
                     "Check that python-docx is installed and retry.")
    if st.session_state.get("teaching_zip"):
        year_part = teaching_academic_label(selected[0]) if len(selected) == 1 else "Multiple_Academic_Years"
        st.download_button("Download CSV + Word reports (ZIP)",
                           data=st.session_state["teaching_zip"], file_name=f"Preceptor_Teaching_{year_part}.zip",
                           mime="application/zip", key="teaching_download_zip")
        # Reuse the Word bytes already packaged in the ZIP; no second scan or decryption.
        with ZipFile(BytesIO(st.session_state["teaching_zip"])) as generated_zip:
            chair_bytes = generated_zip.read(TEACHING_CHAIR_SUMMARY_FILENAME)
        st.download_button("Download chair summary only (Word)", data=chair_bytes,
                           file_name=TEACHING_CHAIR_SUMMARY_FILENAME,
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                           key="teaching_download_chair_summary")
        st.caption("The ZIP includes one combined chair summary, the unchanged summary CSV, individual Word reports "
                   "in Preceptor_Reports, Report_Notes.txt and Archive_Sources.csv. "
                   "Treat these downloads as unencrypted staff reports; do not commit them to the public repository.")


st.set_page_config(page_title="PSUCOM PEDIATRIC CLERKSHIP SCHEDULE CREATOR", layout="wide")
st.title("PSUCOM PEDIATRIC CLERKSHIP SCHEDULE CREATOR")

# ─── Sidebar mode selector ─────────────────────────────────────────────────────
mode = st.sidebar.radio("What do you want to do?",("Instructions", "Format OPD + Summary", "Create Student Schedule", "OPD Check", "Create Individual Schedules", "OPD Archive", "Preceptor Teaching Summary", "OPD MD PA Conflict Detector", "Shift Availability Tracker"), key="schedule_app_mode")
# ─── Sidebar mode selector ─────────────────────────────────────────────────────

if mode == "OPD Check":
    DAYS = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']
    
    def detect_am_pm_blocks(df):
        """
        Scan down column A and group runs whose cell text begins with 'AM' or 'PM',
        ignoring everything else.
        Returns a list of (label, start_row, end_row) tuples.
        """
        runs = []
        current_label = None
        run_start = None
        prev_row = None
    
        # Excel row 6 is df index 5
        for idx in range(5, len(df)):
            raw = df.iat[idx, 0]
            if isinstance(raw, str):
                ru = raw.upper()
                if ru.startswith("AM"):
                    label = "AM"
                elif ru.startswith("PM"):
                    label = "PM"
                else:
                    continue
            else:
                continue
    
            row = idx + 1  # convert back to 1‑based Excel row
            if current_label is None:
                # start new run
                current_label = label
                run_start = row
            elif label != current_label or row != prev_row + 1:
                # close out previous run
                runs.append((current_label, run_start, prev_row))
                current_label = label
                run_start = row
    
            prev_row = row
    
        # finish last run
        if current_label is not None:
            runs.append((current_label, run_start, prev_row))
    
        return runs
    
    st.title("OPD PRECEPTOR CHECK")
    
    baseline_file = st.file_uploader(
        "1) Upload Latest Updated OPD", 
        type=["xlsx"], key="baseline"
    )
    assigned_file = st.file_uploader(
        "2) Upload OPD with Student Assignments", 
        type=["xlsx"], key="assigned"
    )
    
    if baseline_file and assigned_file:
        SHEETS = [
            'HOPE_DRIVE','ETOWN','NYES','COMPLEX','WARD A',
            'PSHCH_NURSERY','HAMPDEN_NURSERY','SJR_HOSP','AAC',
            'AHOLOUKPE','ADOLMED'
        ]
    
        # Read all relevant sheets at once
        base_sheets = pd.read_excel(baseline_file, sheet_name=SHEETS, header=None)
        assn_sheets = pd.read_excel(assigned_file, sheet_name=SHEETS, header=None)

        # 1) Remove student suffix from every baseline sheet
        for sheet_name, df in base_sheets.items():
            for col in df.columns[1:]:
                df[col] = (
                    df[col]
                      .where(df[col].notna(), np.nan)      # keep NaN
                      .astype(str)
                      .str.partition('~')[0]               # take text before "~"
                      .replace({'nan': np.nan})            # restore NaN
                )
            base_sheets[sheet_name] = df
    
        results = {}
        
        for sheet in SHEETS:
            df_base = base_sheets[sheet]
            df_assn = assn_sheets[sheet]
        
            runs = detect_am_pm_blocks(df_base)
            week_pairs = [(runs[i], runs[i+1]) for i in range(0, len(runs), 2)]
        
            base_map = {}    # (w,period,day,pre) → (cell, student)
            assn_map = {}
        
            for w_idx, ((am_lbl, am_s, am_e), (pm_lbl, pm_s, pm_e)) in enumerate(week_pairs, start=1):
                for period, (lbl, start, end) in [('AM',(am_lbl,am_s,am_e)), ('PM',(pm_lbl,pm_s,pm_e))]:
                    for col, day in enumerate(DAYS, start=1):
                        for row in range(start, end+1):
                            cell = f"{chr(ord('A')+col)}{row}"
        
                            # baseline
                            vb = df_base.iat[row-1, col]
                            if pd.notna(vb) and isinstance(vb, str):
                                parts = str(vb).split('~', 1)
                                pre = parts[0].strip()
                                stu = parts[1].strip() if len(parts)==2 else None
                                base_map.setdefault((w_idx,period,day,pre), (cell, stu))
        
                            # assigned
                            va = df_assn.iat[row-1, col]
                            if pd.notna(va) and isinstance(va, str):
                                parts = str(va).split('~', 1)
                                pre = parts[0].strip()
                                stu = parts[1].strip() if len(parts)==2 else None
                                assn_map.setdefault((w_idx,period,day,pre), (cell, stu))
        
            dropped = []
            added   = []
        
            # drops: in base not in assigned
            for key, (cell, stu) in base_map.items():
                if key not in assn_map:
                    w,p,d,pre = key
                    dropped.append((w,p,d,pre,cell,stu))
        
            # adds: in assigned not in base
            for key, (cell, stu) in assn_map.items():
                if key not in base_map:
                    w,p,d,pre = key
                    added.append((w,p,d,pre,cell,stu))
        
            # sort exactly as before...
            dropped_sorted = sorted(
                dropped,
                key=lambda x: (x[0], {'AM':0,'PM':1}[x[1]], DAYS.index(x[2]), x[4])
            )
            added_sorted = sorted(
                added,
                key=lambda x: (x[0], {'AM':0,'PM':1}[x[1]], DAYS.index(x[2]), x[4])
            )
        
            results[sheet] = {"dropped": dropped_sorted, "added": added_sorted}



        
        #
    
    doc = Document()
    doc.add_heading('Change Report', level=1)
    
    for sheet, change in (locals().get('results') or {}).items():
        doc.add_heading(sheet, level=2)
    
        # build week→day map (collect both AM & PM under each day)
        week_map = defaultdict(lambda: defaultdict(lambda: {'dropped': [], 'added': []}))
        for w,p,d,pre,cell,stu in change['dropped']:
            week_map[w][d]['dropped'].append((p, pre, cell, stu))
        for w,p,d,pre,cell,stu in change['added']:
            week_map[w][d]['added'].append((p, pre, cell, stu))
    
        # emit in week order
        for w in sorted(week_map):
            doc.add_heading(f'Week {w}', level=3)
    
            for day in DAYS:
                slot = week_map[w].get(day)
                if not slot or (not slot['dropped'] and not slot['added']):
                    continue
    
                doc.add_heading(day, level=4)
    
                # DROPS
                for p, pre, cell, stu in slot['dropped']:
                    line = f"- Dropped: {pre} — was at {cell}"
                    if stu:
                        line += f"  (Student impacted: {stu})"
                    doc.add_paragraph(line, style='List Bullet')
    
                # ADDS
                for p, pre, cell, stu in slot['added']:
                    line = f"- Added: {pre} — now at {cell}"
                    if stu:
                        line += f"  (Student impacted: {stu})"
                    doc.add_paragraph(line, style='List Bullet')
    
            doc.add_paragraph()  # blank line between weeks




    # Save to in-memory buffer
    word_file = io.BytesIO()
    doc.save(word_file)
    word_file.seek(0)
    
    # Download button
    st.download_button(
        label="📄 Download Word Report",
        data=word_file,
        file_name="change_report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


elif mode == "Instructions":
    d = st.text_input('Start date (m/d/yyyy)')
    if d:
        try:
            s = datetime.strptime(d, '%m/%d/%Y')
            e = s + timedelta(days=34)
            st.write(f"{s:%B %d, %Y} → {e:%B %d, %Y}")
    
            st.write('Please go to https://login.qgenda.com/')
    
            # Display on‐screen instructions
            st.markdown(f"""Download four files and create reports based on **{s:%B %d, %Y}** → **{e:%B %d, %Y}**.""")
            st.write("Download instructions here:")


        except ValueError:
            st.error('Invalid format – use m/d/yyyy (e.g. 7/6/2021)')

        # --- Generate a Word document with the same instructions ---
        doc = Document()
        doc.add_heading('Qgenda Report Instructions', level=1)
        doc.styles['Normal'].font.size = Pt(8)
        
        # Bold the date range
        p = doc.add_paragraph()
        p.add_run('Date range: ')
        run_start = p.add_run(f'{s:%B %d, %Y}')
        run_start.bold = True
        p.add_run(' → ')
        run_end = p.add_run(f'{e:%B %d, %Y}')
        run_end.bold = True
        
        
        doc.add_paragraph('1. Go to https://login.qgenda.com/')
        
        # Helper to add each report block
        def add_report(title, steps):
            doc.add_heading(title, level=2)
            for step in steps:
                # If the step contains dates, bold them
                if 'Enter Start Date:' in step:
                    p = doc.add_paragraph(style='List Bullet')
                    prefix, dates = step.split(':', 1)
                    p.add_run(prefix + ': ')
                    # split the two dates on " and End Date:"
                    start_part, end_part = dates.strip().split(' and End Date:')
                    r1 = p.add_run(start_part.strip())
                    r1.bold = True
                    p.add_run(' and End Date: ')
                    r2 = p.add_run(end_part.strip())
                    r2.bold = True
                else:
                    doc.add_paragraph(step, style='List Bullet')
        
        add_report(
            'Report 1 – Penn State Health Hershey Medical Center - Academic General Pediatrics',
            [
                'Click Penn State Health Hershey Medical Center - Academic General Pediatrics → Schedule → Reports',
                'Set Report Type to Calendar by Task',
                'Set Format to Excel',
                f'Enter Start Date: {s:%m/%d/%Y} and End Date: {e:%m/%d/%Y}',
                'Ensure Calendar starts on Monday',
                'Show Staff by Last Name, First Name',
                'Show Tasks by Short Name',
                'Click Run Report'
            ]
        )
        add_report(
            "Report 2 – Penn State Health Children's Hospital – Hospitalists",
            [
                'Click Penn State Health Children\'s Hospital → Schedule → Reports',
                'Set Report Type to Calendar by Task',
                'Set Format to Excel',
                f'Enter Start Date: {s:%m/%d/%Y} and End Date: {e:%m/%d/%Y}',
                'Ensure Calendar starts on Monday',
                'Show Staff by Last Name, First Name',
                'Show Tasks by Long Name',
                'Click Run Report'
            ]
        )
        add_report(
            'Report 3 – Department of Pediatrics (Admin - Adolescent Med)',
            [
                'Click Department of Pediatrics → Schedule → Reports',
                'Select Admin - Adolescent Med in top-right corner',
                'Set Report Type to Calendar by Task',
                'Set Format to Excel',
                f'Enter Start Date: {s:%m/%d/%Y} and End Date: {e:%m/%d/%Y}',
                'Ensure Calendar starts on Monday',
                'Show Staff by Last Name, First Name',
                'Show Tasks by Long Name',
                'Click Run Report'
            ]
        )
        add_report(
            'Report 4 – Department of Pediatrics (Complex Care)',
            [
                'Click Department of Pediatrics → Schedule → Reports',
                'Select Complex Care in top-right corner',
                'Set Report Type to Calendar by Task',
                'Set Format to Excel',
                f'Enter Start Date: {s:%m/%d/%Y} and End Date: {e:%m/%d/%Y}',
                'Ensure Calendar starts on Monday',
                'Show Staff by Last Name, First Name',
                'Show Tasks by Long Name',
                'Click Run Report'
            ]
        )


        # Save to bytes and offer download
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        st.download_button(label="📄 Download Instructions (Word)",data=buf.getvalue(),file_name="Qgenda_Report_Instructions.docx",mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


elif mode == "Format OPD + Summary":
    # ─── Inputs ────────────────────────────────────────────────────────────────────
    # Required keywords to look for in the content
    #required_keywords = ["academic general pediatrics", "hospitalists", "complex care", "adol med"]
    required_keywords = ["academic general pediatrics", "hospitalists", "complex care"]
    found_keywords = set()
    
    schedule_files = st.file_uploader("1) Upload one or more QGenda calendar Excel(s)",type=["xlsx", "xls"],accept_multiple_files=True)
    
    #if schedule_files:
    #    for file in schedule_files:
    #        try:
    #            # Read the first sheet
    #            df = pd.read_excel(file, sheet_name=0, header=None)
    
    #            # Flatten all string values to a list of lowercase strings
    #            cell_values = df.astype(str).apply(lambda x: x.str.lower()).values.flatten().tolist()
    
    #            # Check if any keyword is found in cell values
    #            for keyword in required_keywords:
    #                if any(keyword in val for val in cell_values):
    #                    found_keywords.add(keyword)
    
    #       except Exception as e:
    #           st.error(f"Error reading {file.name}: {e}")

    if schedule_files:
        for file in schedule_files:
            try:
                df = pd.read_excel(file, sheet_name=0, header=None)
    
                cell_values = [
                    str(val).strip().lower()
                    for row in df.values
                    for val in row
                    if pd.notna(val)
                ]
    
                for keyword in required_keywords:
                    if any(keyword in val for val in cell_values):
                        found_keywords.add(keyword)
    
            except Exception as e:
                st.error(f"Error reading {file.name}: {e}")
    
        # Identify missing calendars
        missing_keywords = [k for k in required_keywords if k not in found_keywords]
    
        if missing_keywords:
            st.warning(f"Missing required calendar(s): {', '.join(missing_keywords)}. Please upload all required calendars.")
        else:
            st.success("All required calendars uploaded and verified by content.")
    
    student_file = st.file_uploader("2) Upload Redcap Rotation list CSV (must have a 'legal_name' and 'start_date' column)",type=["csv"])
    
    #record_id = st.text_input("3) Enter the REDCap record_id for this batch", "")
    
    record_id = "peds_clerkship"
    
    # ─── Guard ─────────────────────────────────────────────────────────────────────
    if not schedule_files or not student_file or not record_id:
        st.info("Please upload schedule Excel(s), student CSV")
        st.stop()
    
    # ─── Prep: Date regex & maps ───────────────────────────────────────────────────
    
    date_pat = re.compile(r'^[A-Za-z]+ \d{1,2}, \d{4}$')
    base_map = {
        "hope drive am continuity":    "hd_am_",
        "hope drive pm continuity":    "hd_pm_",
        
        "hope drive am acute precept": "hd_am_acute_",
        "hope drive pm acute precept": "hd_pm_acute_",
    
        "hope drive weekend acute 1": "hd_wknd_acute_1_", # Changed prefix
        "hope drive weekend acute 2": "hd_wknd_acute_2_", # Changed prefix
    
        "hope drive weekend continuity": "hd_wknd_am_",
        
        "etown am continuity":         "etown_am_",
        "etown pm continuity":         "etown_pm_",
        
        "nyes rd am continuity":       "nyes_am_",
        "nyes rd pm continuity":       "nyes_pm_",
        
        "nursery weekday 8a-6p":       ["nursery_am_", "nursery_pm_"],
        
        "rounder 1 7a-7p":             ["ward_a_am_","ward_a_pm_"],
        "rounder 2 7a-7p":             ["ward_a_am_","ward_a_pm_"],
        "rounder 3 7a-7p":             ["ward_a_am_","ward_a_pm_"],
    
        "hope drive clinic am":        "complex_am_",
        "hope drive clinic pm":        "complex_pm_",
        
        "briarcrest clinic am":       "adol_med_am_",
        "briarcrest clinic pm":       "adol_med_pm_",

        "lancaster am":       "lancaster_am_",
        "lancaster pm":       "lancaster_pm_",
    
    }
    
    # Which groups need at least 2 providers?
    min_required = {
        "hope drive am acute precept": 2,
        "hope drive pm acute precept": 2,
        
        "nursery weekday 8a-6p":       2,
        
        "rounder 1 7a-7p":             2,
        "rounder 2 7a-7p":             2,
        "rounder 3 7a-7p":             2,
    }
    
    file_configs = {"HAMPDEN_NURSERY.xlsx": {"title": "HAMPDEN_NURSERY","custom_text": "CUSTOM_PRINT","names": ["Folaranmi, Oluwamayoda","Alur, Pradeep","Nanda, Sharmilarani","HAMPDEN_NURSERY"]},
                    "SJR_HOSP.xlsx": {"title": "SJR_HOSPITALIST","custom_text": "CUSTOM_PRINT","names": ["Spangola, Haley","Gubitosi, Terry","SJR_1","SJR_2"]},
                    "AAC.xlsx": {"title": "AAC","custom_text": "CUSTOM_PRINT","names": ["Vaishnavi Harding","Abimbola Ajayi","Shilu Joshi","Desiree Webb","Amy Zisa","Abdullah Sakarcan","Anna Karasik","AAC_1","AAC_2","AAC_3",]},
                    "LANCASTER_CMG.xlsx": {"title": "LANCASTER_CMG","custom_text": "CUSTOM_PRINT","names": ["Ashleigh Sobotka","Susannah Christman"]},
                    "MAHOUSSI_AHOLOUKPE.xlsx": {"title": "MAHOUSSI_AHOLOUKPE","custom_text": "CUSTOM_PRINT","names": ["Mahoussi Aholoukpe"]},
                    #"REPLACE.xlsx": {"title": "REPLACE","custom_text": "CUSTOM_PRINT","names": ["ReplaceFirstName ReplaceLastName"]},
                   }
    
    # ─── HERE: generate sheet‐specific custom_print entries for the configss...  ────────────────────
    for cfg in file_configs.values():
        sheet = cfg["title"]              # e.g. "HAMPDEN_NURSERY"
        key   = sheet.lower() + "_print"  # e.g. "hampden_nursery_print"
        prefix = f"{cfg['custom_text'].lower()}_{sheet.lower()}_"
        base_map[key] = prefix
        
    # ─── 1. Aggregate schedule assignments by date ────────────────────────────────
    assignments_by_date = {}
    for file in schedule_files:
        df = pd.read_excel(file, header=None, dtype=str)
    
        # find all date cells
        date_positions = []
        for r in range(df.shape[0]):
            for c in range(df.shape[1]):
                val = str(df.iat[r,c]).replace("\xa0"," ").strip()
                if date_pat.match(val):
                    try:
                        d = pd.to_datetime(val).date()
                        date_positions.append((d,r,c))
                    except:
                        pass
    
        # dedupe to the topmost row per date
        unique = {}
        for d,r,c in date_positions:
            if d not in unique or r < unique[d][0]:
                unique[d] = (r,c)
    
        # before the loop, define:
        day_names = {"monday","tuesday","wednesday","thursday","friday","saturday","sunday"}
        
        # collect providers under each date
        for d, (row0,col0) in unique.items():
            grp = assignments_by_date.setdefault(d, {des:[] for des in base_map})
            
            for r in range(row0+1, df.shape[0]):
                raw = str(df.iat[r, col0]).replace("\xa0", " ").strip()
                # stop if we hit a blank row
                if raw == "":
                    break
                # stop if we hit another date header
                if date_pat.match(raw):
                    break
    
                desc = raw.lower()
                prov = str(df.iat[r, col0+1]).strip()
                if desc in grp and prov:
                    grp[desc].append(prov)
    
    # ─── Provider filter UI ──────────────────────────────────────────────────────
    all_providers = sorted({
        p.strip()
        for day in assignments_by_date.values()
        for provs in day.values()
        for p in provs
        if isinstance(p, str) and p.strip()
    })
    
    # Multiselect persists in session; start empty by design
    if "provider_filter" not in st.session_state:
        st.session_state["provider_filter"] = []
    
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("Select All Providers", key="prov_select_all"):
            st.session_state["provider_filter"] = all_providers
    with col2:
        if st.button("Clear Providers", key="prov_clear_all"):
            st.session_state["provider_filter"] = []
    with col3:
        # Switch to actually apply the filter. Off = treat as 'All'
        apply_provider_filter = st.checkbox(
            "Apply provider filter",
            value=False,
            key="prov_apply_filter",
            help="When OFF, everyone is included even if the multiselect is blank."
        )
    
    allowed_providers = st.multiselect(
        "Limit providers included in OPD",
        options=all_providers,
        key="provider_filter",
        help="Only selected providers will be written when 'Apply provider filter' is ON.",
    )
    
    # Effective allow-list:
    effective_allowed = (
        set(allowed_providers) if (apply_provider_filter and allowed_providers) else set(all_providers)
    )

    # ─── 2. Read student list and prepare s1, s2, … ───────────────────────────────
    students_df = pd.read_csv(student_file, dtype=str)
    legal_names = students_df["legal_name"].dropna().tolist()
    
    # ─── 3. Build the single REDCap row ───────────────────────────────────────────
    redcap_row = {"record_id": record_id}
    sorted_dates = sorted(assignments_by_date.keys())
    
    for idx, date in enumerate(sorted_dates, start=1):
        redcap_row[f"hd_day_date{idx}"] = date
        suffix = f"d{idx}_"
    
        # build day‑specific prefixes
        des_map = {
            des: ([p + suffix for p in prefs] if isinstance(prefs, list)
                  else [prefs + suffix])
            for des, prefs in base_map.items()
        }
    
        # 3a) schedule providers (respect provider filter)
        for des, provs in assignments_by_date[date].items():
            # Do not mutate the original list
            filtered = [p for p in provs if p in effective_allowed]
        
            # If the group has a minimum requirement, pad by repeating the first allowed provider
            req = min_required.get(des, len(filtered))
            if filtered and len(filtered) < req:
                filtered = filtered + [filtered[0]] * (req - len(filtered))
        
            # If nothing allowed and no minimum → skip write
            if not filtered:
                continue
        
            if des.startswith("rounder"):
                # rounder N 7a-7p → slot math
                # NOTE: req here is the number of providers per team (usually 2)
                team_idx = int(des.split()[1]) - 1  # 0-based team index
                for i, name in enumerate(filtered, start=1):
                    slot = team_idx * req + i  # team1→1..req, team2→req+1..2*req, etc.
                    for prefix in des_map[des]:
                        redcap_row[f"{prefix}{i if prefix.endswith('_am_') or prefix.endswith('_pm_') else slot}"] = name
                        # ^ If your rounder prefixes are lists like ["ward_a_am_","ward_a_pm_"],
                        #   they'll be in des_map[des] already; the index logic above preserves slots.
            else:
                for i, name in enumerate(filtered, start=1):
                    for prefix in des_map[des]:
                        redcap_row[f"{prefix}{i}"] = name

        # 3b) custom_print names — once per date, using the SAME suffix
        for fname, cfg in file_configs.items():
            sheet = cfg["title"]          
            key   = sheet.lower() + "_print"
            prefix = base_map[key]       # e.g. "custom_print_hampden_nursery_"
            
            for i, person in enumerate(cfg["names"], start=1):
                # note the suffix goes BEFORE the slot index
                redcap_row[f"{prefix}{suffix}{i}"] = person
                
    # append student slots s1,s2,...
    for i,name in enumerate(legal_names, start=1):
        redcap_row[f"s{i}"] = name
    
    # ─── 4. Display & slice out dates/am/acute and students ─────────────────────
    out_df = pd.DataFrame([redcap_row])
    
    # 1) shuffle
    students = legal_names.copy()
    random.shuffle(students)
    
    # 2) define slot sequence
    slot_seq = [1, 3, 5, 2, 4, 6]
    
    # 3) assign
    ward_a_assignment = {}
    
    for idx, student in enumerate(students):
        slot_group = idx // 4                  # every 4 students move to next slot
        slot       = slot_seq[slot_group % len(slot_seq)]
        week_idx   = idx % 4                   # 0→week1,1→week2,2→week3,3→week4
    
        ward_a_assignment[student] = week_idx
        
        # for their week, each Mon–Fri (days 1–5 + 7*week_idx)
        for day in range(1, 6):
            day_num = day + 7 * week_idx
            for shift in ("am", "pm"):
                key  = f"ward_a_{shift}_d{day_num}_{slot}"
                orig = redcap_row.get(key, "")
                redcap_row[key] = f"{orig} ~ {student}" if orig else f"~ {student}"
                
    # ─── track who’s already grabbed a nursery slot ─────────────────────────────
    nursery_assigned = set()
    
    # ─── HAMPDEN_NURSERY: max 1 student for week1 and 1 for week3, into slot _4 ##FOCUSES ON SLOT 4!!! ─────
    for week_idx in (0, 2):  # 0→week1, 2→week3
        pool = [
            s for s in legal_names
            if s not in nursery_assigned
            and ward_a_assignment.get(s, -1) != week_idx
        ]
        if not pool:
            continue
        student = random.choice(pool)
        nursery_assigned.add(student)    # ← mark them as “used”!
    
        for day in range(1, 6):
            d   = day + 7 * week_idx
            key = f"custom_print_hampden_nursery_d{d}_4"        
            orig = redcap_row.get(key, "")
            redcap_row[key] = f"{orig} ~ {student}" if orig else f"~ {student}"
    
    # ─── 2) SJR_HOSPITALIST (max 2 students, any weeks ≠ their Ward A week) ─────
    for week_idx in range(4):  # 0→wk1,1→wk2,2→wk3,3→wk4
        # build pool excluding Hampden and anyone on Ward A that week
        pool = [
            s for s in legal_names
            if s not in nursery_assigned
            and ward_a_assignment.get(s, -1) != week_idx
        ]
        random.shuffle(pool)
        # assign up to two students: first to slot 3, next to slot 4
        for slot_idx in (3, 4):
            if not pool:
                break
            student = pool.pop()
            nursery_assigned.add(student)
            # Mon–Fri of this week
            for day in range(1, 6):
                d   = day + 7 * week_idx
                key = f"custom_print_sjr_hospitalist_d{d}_{slot_idx}"
                orig = redcap_row.get(key, "")
                redcap_row[key] = f"{orig} ~ {student}" if orig else f"~ {student}"
    
    
    # ─── 3) PSHCH_NURSERY (everyone else, up to 8 slots: slot1 weeks1–4, then slot2 wks1–4) ─────────
    leftovers = [s for s in legal_names if s not in nursery_assigned]
    # build (week_idx, slot) in the desired order
    psch_slots = [(wk,1) for wk in range(4)] + [(wk,2) for wk in range(4)]
    for student in leftovers:
        for wk, slot in psch_slots:
            # skip if conflicts with Ward A week
            if ward_a_assignment.get(student, -1) == wk:
                continue
            # build key once (AM & PM) to test existence and avoid duping
            key_am = f"nursery_am_d{day}_ {slot}"
            # assign across Mon–Fri
            for day in range(1, 6):
                d = day + wk * 7
                for prefix in ("nursery_am_","nursery_pm_"):
                    key  = f"{prefix}d{d}_{slot}"
                    orig = redcap_row.get(key, "")
                    redcap_row[key] = f"{orig} ~ {student}" if orig else f"~ {student}"
            # remove this slot so no one else uses it
            psch_slots.remove((wk,slot))
            nursery_assigned.add(student)
            break
        # if no slot left, the student remains unassigned in PSHCH_NURSERY
    
    # format date columns
    for c in out_df.columns:
        if c.startswith("hd_day_date"):
            out_df[c] = pd.to_datetime(out_df[c]).dt.strftime("%m-%d-%Y")
    
    
    out_df = pd.DataFrame([redcap_row])
    csv_full = out_df.to_csv(index=False).encode("utf-8")
    
    def generate_opd_workbook(full_df: pd.DataFrame) -> bytes:
        import io
        import xlsxwriter
    
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    
        # ─── Formats ─────────────────────────────────────────────────────────────────
        format1     = workbook.add_format({'font_size':18,'bold':1,'align':'center','valign':'vcenter','font_color':'black','bg_color':'#FEFFCC','border':1})
        format4     = workbook.add_format({'font_size':12,'bold':1,'align':'center','valign':'vcenter','font_color':'black','bg_color':'#8ccf6f','border':1})
        format4a    = workbook.add_format({'font_size':12,'bold':1,'align':'center','valign':'vcenter','font_color':'black','bg_color':'#9fc5e8','border':1})
        format5     = workbook.add_format({'font_size':12,'bold':1,'align':'center','valign':'vcenter','font_color':'black','bg_color':'#FEFFCC','border':1})
        format5a    = workbook.add_format({'font_size':12,'bold':1,'align':'center','valign':'vcenter','font_color':'black','bg_color':'#d0e9ff','border':1})
        format11    = workbook.add_format({'font_size':18,'bold':1,'align':'center','valign':'vcenter','font_color':'black','bg_color':'#FEFFCC','border':1})
        formate     = workbook.add_format({'font_size':12,'bold':0,'align':'center','valign':'vcenter','font_color':'white','border':0})
        format3     = workbook.add_format({'font_size':12,'bold':1,'align':'center','valign':'vcenter','font_color':'black','bg_color':'#FFC7CE','border':1})
        format2     = workbook.add_format({'bg_color':'black'})
        format_date = workbook.add_format({'num_format':'m/d/yyyy','font_size':12,'bold':1,'align':'center','valign':'vcenter','font_color':'black','bg_color':'#FFC7CE','border':1})
        format_label= workbook.add_format({'font_size':12,'bold':1,'align':'center','valign':'vcenter','font_color':'black','bg_color':'#FFC7CE','border':1})
        merge_format= workbook.add_format({'bold':1,'align':'center','valign':'vcenter','text_wrap':True,'font_color':'red','bg_color':'#FEFFCC','border':1})
    
        # ─── Worksheets ─────────────────────────────────────────────────────────────
        worksheet_names = ['HOPE_DRIVE','ETOWN','NYES','LANCASTER','LANCASTER_CMG','COMPLEX','WARD A','PSHCH_NURSERY','HAMPDEN_NURSERY','SJR_HOSP','AAC','AHOLOUKPE','ADOLMED']
        
        sheets = {name: workbook.add_worksheet(name) for name in worksheet_names}
    
        # ─── Site headers ────────────────────────────────────────────────────────────
        site_list = ['Hope Drive','Elizabethtown','Nyes Road','Lancaster','Lancaster CMG','Complex Care','WARD A','PSHCH NURSERY','HAMPDEN NURSERY','SJR HOSPITALIST','AAC','AHOLOUKPE','ADOLMED']
        
        for ws, site in zip(sheets.values(), site_list):
            ws.write(0, 0, 'Site:', format1)
            ws.write(0, 1, site,   format1)
    
        # ─── HOPE_DRIVE specific ────────────────────────────────────────────────────
        hd = sheets['HOPE_DRIVE']
        for cr in ['A8:H15','A32:H39','A56:H63','A80:H87']:
            hd.conditional_format(cr, {'type':'cell','criteria':'>=','value':0,'format':format1})
        for cr in ['A18:H25','A42:H49','A66:H73','A90:H97']:
            hd.conditional_format(cr, {'type':'cell','criteria':'>=','value':0,'format':format5a})
        for cr in ['A6:H6','A7:H7','A30:H30','A31:H31','A54:H54','A55:H55','A78:H78','A79:H79']:
            hd.conditional_format(cr, {'type':'cell','criteria':'>=','value':0,'format':format4})
        for cr in ['A16:H16','A17:H17','A40:H40','A41:H41','A64:H64','A65:H65','A88:H88','A89:H89']:
            hd.conditional_format(cr, {'type':'cell','criteria':'>=','value':0,'format':format4a})

        # how many acute vs continuity rows per block
        ACUTE_COUNT      = 2
        CONTINUITY_COUNT = 8
        BLOCK_SIZE       = ACUTE_COUNT + CONTINUITY_COUNT  # should be 10
        AM_COUNT         = BLOCK_SIZE
        PM_COUNT         = BLOCK_SIZE
        
        # e.g. [6,16,30,40,54,64,78,88]
        BLOCK_STARTS = [6, 30, 54, 78]
        
        for start in BLOCK_STARTS:
            zero_row = start - 1
            # — AM half of the block —
            for i in range(AM_COUNT):
                # first 2 → ACUTES, rest → Continuity
                if i < ACUTE_COUNT:
                    label = 'AM - ACUTES'
                else:
                    label = 'AM - Continuity'
                hd.write(zero_row + i, 0, label, format5a)
            
            # — PM half of the block —
            for i in range(PM_COUNT):
                if i < ACUTE_COUNT:
                    label = 'PM - ACUTES'
                else:
                    label = 'PM - Continuity'
                hd.write(zero_row + AM_COUNT + i, 0, label, format5a)
            
        # ─── GENERIC SHEETS ─────────────────────────────────────────────────────────
        others       = [ws for name, ws in sheets.items() if name != 'HOPE_DRIVE']
        AM_COUNT     = 10
        PM_COUNT     = 10
        BLOCK_STARTS = [6, 30, 54, 78]
    
        for ws in others:
            # 1) conditional formats
            for cr in ['A6:H15','A30:H39','A54:H63','A78:H87']:
                ws.conditional_format(cr, {
                    'type':'cell','criteria':'>=','value':0,'format':format1
                })
            for cr in ['A16:H25','A40:H49','A64:H73','A88:H97']:
                ws.conditional_format(cr, {
                    'type':'cell','criteria':'>=','value':0,'format':format5a
                })
            for cr, fmt in [
                ('B6:H6',   format4), ('B16:H16', format4a),
                ('B30:H30', format4), ('B40:H40', format4a),
                ('B54:H54', format4), ('B64:H64', format4a),
                ('B78:H78', format4), ('B88:H88', format4a)
            ]:
                ws.conditional_format(cr, {
                    'type':'cell','criteria':'>=','value':0,'format':fmt
                })
    
            # 2) Write exactly 10 AM then 10 PM in column A
            for start in BLOCK_STARTS:
                zero_row = start - 1
                for i in range(AM_COUNT):
                    ws.write(zero_row + i, 0, 'AM', format5a)
                for i in range(PM_COUNT):
                    ws.write(zero_row + AM_COUNT + i, 0, 'PM', format5a)
    
        # ─── Universal formatting & dates ────────────────────────────────────────────
        date_cols = [f"hd_day_date{i}" for i in range(1,29)]
        dates     = pd.to_datetime(full_df[date_cols].iloc[0]).tolist()
        weeks     = [dates[i*7:(i+1)*7] for i in range(4)]
        days      = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']
    
        for ws in workbook.worksheets():
            ws.set_zoom(80)
            ws.set_column('A:A', 10)
            ws.set_column('B:H', 65)
            ws.set_row(0, 37.25)
    
            for idx, start in enumerate([2,26,50,74]):
                        # day names
                        for c, d in enumerate(days):
                            ws.write(start, 1+c, d, format3)
                        # dates
                        for c, val in enumerate(weeks[idx]):
                            ws.write(start+1, 1+c, val, format_date)
                            
                        # padding formula bars
                        ws.write_formula(f'A{start}',   '""', format_label)
                        
                        ws.conditional_format(
                            f'A{start+3}:H{start+3}',
                            {'type':'cell','criteria':'>=','value':0,'format':format_label}
                        )
    
    
            # black bars every 24 rows
            step = 24
            for row in range(2, 98, step):
                ws.merge_range(f'A{row}:H{row}', ' ', format2)
    
            # merge CRTS message on every sheet
            text1 = (
                'Students are to alert their preceptors when they have a Clinical '
                'Reasoning Teaching Session (CRTS).  Please allow the students to '
                'leave approximately 15 minutes prior to the start of their session '
                'so they can be prepared to actively participate.  - Thank you!'
            )
            ws.merge_range('C1:F1', text1, merge_format)
            ws.write('G1', '', merge_format)
            ws.write('H1', '', merge_format)
    
            #PAINT Empty White Spaces
            ws.write('A3', '', format_date)
            ws.write('A4', '', format_date)
            ws.write('A27', '', format_date)
            ws.write('A28', '', format_date)
            ws.write('A51', '', format_date)
            ws.write('A52', '', format_date)
            ws.write('A75', '', format_date)
            ws.write('A76', '', format_date)
    
    
        workbook.close()
        output.seek(0)
        return output.read()
    
    excel_bytes = generate_opd_workbook(out_df)
    #st.download_button(label="⬇️ Download OPD.xlsx",data=excel_bytes,file_name="OPD.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    
    import pandas as pd
    import io
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment # <--- NEW IMPORT
    
    # --- MODIFIED update_excel_from_csv function to work with bytes ---
    def update_excel_from_csv(excel_template_bytes: bytes, csv_data_bytes: bytes, mappings: list) -> bytes | None:
        """
        Updates an Excel file (from bytes) with values from a CSV (from bytes)
        based on provided mappings, and returns the updated Excel as bytes.
    
        Args:
            excel_template_bytes (bytes): The bytes content of the Excel file to be updated.
            csv_data_bytes (bytes): The bytes content of the CSV file containing the data.
            mappings (list of dict): A list of dictionaries, where each dictionary
                                     defines a mapping:
                                     {'csv_column': 'name_of_csv_column',
                                      'excel_sheet': 'Sheet Name',
                                      'excel_cell': 'Cell Address (e.g., B8)'}
        Returns:
            bytes: The bytes content of the updated Excel workbook, or None if an error occurs.
        """
        try:
            # Load the CSV data from bytes using BytesIO
            df_csv = pd.read_csv(io.BytesIO(csv_data_bytes))
    
            if df_csv.empty:
                # Assuming st is available in the Streamlit environment
                # If running outside Streamlit, you might use print() or logging
                # st.error("Error: The CSV data is empty. Cannot update Excel.")
                return None
    
            # Load the Excel workbook from bytes using BytesIO
            wb = load_workbook(io.BytesIO(excel_template_bytes))
    
            # Iterate through the mappings and update the Excel file
            for mapping in mappings:
                csv_column = mapping['csv_column']
                excel_sheet_name = mapping['excel_sheet']
                excel_cell = mapping['excel_cell']
    
                if csv_column not in df_csv.columns:
                    # st.warning(f"Warning: CSV column '{csv_column}' not found in CSV data. Skipping this mapping.")
                    continue
    
                # Get the value from the first row of the specified CSV column
                value_to_transfer = df_csv.loc[0, csv_column]
    
                if excel_sheet_name not in wb.sheetnames:
                    # st.warning(f"Warning: Excel sheet '{excel_sheet_name}' not found in the Excel template. Skipping this mapping.")
                    continue
    
                ws = wb[excel_sheet_name]
    
                # --- APPLY THE REQUESTED FORMATTING HERE ---
                # 1. Convert to string and append " ~ "
                # This ensures that even if value_to_transfer is a number, it can be concatenated
                #formatted_value = str(value_to_transfer) + ' ~ '

                orig = str(value_to_transfer).strip()
                # if it already contains a student‐delimiter, don’t add another
                if ' ~ ' in orig:
                    formatted_value = orig
                else:
                    formatted_value = orig + ' ~ '
    
                # 2. Write the formatted value to the cell
                ws[excel_cell] = formatted_value
    
                # 3. Set the alignment for the cell using openpyxl's Alignment
                cell = ws[excel_cell] # Get the cell object
                cell.alignment = Alignment(horizontal='center', vertical='center') # Set horizontal and vertical centering
    
                # st.info(f"Successfully wrote '{formatted_value}' from CSV column '{csv_column}' to '{excel_sheet_name}'!{excel_cell}")
    
            # Save the modified Excel workbook to a BytesIO object
            output_excel_bytes_io = io.BytesIO()
            wb.save(output_excel_bytes_io)
            output_excel_bytes_io.seek(0) # Rewind the buffer to the beginning
    
            return output_excel_bytes_io.getvalue()
    
        except Exception as e:
            # st.error(f"An error occurred during Excel update: {e}")
            return None
            
    # --- Configuration for update_excel_from_csv (your mappings) ---
    data_mappings        = []
    excel_column_letters = ['B','C','D','E','F','G','H']
    num_weeks            = 4
    
    # HOPE_DRIVE acute + continuity row offsets
    hd_row_defs = {
        'AM': {'acute_start': 6,  'cont_start': 8},
        'PM': {'acute_start': 16, 'cont_start': 18},
    }
    
    # Other sheets only need continuity (rows 6–13 for AM, 16–23 for PM)
    cont_row_defs = {
        'AM':  6,
        'PM': 16,
    }
    
    # your prefix map
    base_map = {
        "hope drive am continuity":      "hd_am_",
        "hope drive pm continuity":      "hd_pm_",
        "hope drive am acute precept":   "hd_am_acute_",
        "hope drive pm acute precept":   "hd_pm_acute_",
        "hope drive weekend acute 1":    "hd_wknd_acute_1_",
        "hope drive weekend acute 2":    "hd_wknd_acute_2_",
        "hope drive weekend continuity": "hd_wknd_am_",
        
        "etown am continuity":           "etown_am_",
        "etown pm continuity":           "etown_pm_",
        
        "nyes rd am continuity":         "nyes_am_",
        "nyes rd pm continuity":         "nyes_pm_",
        
        "nursery weekday 8a-6p":         ["nursery_am_","nursery_pm_"],
        
        "rounder 1 7a-7p":               ["ward_a_am_","ward_a_pm_"],
        "rounder 2 7a-7p":               ["ward_a_am_","ward_a_pm_"],
        "rounder 3 7a-7p":               ["ward_a_am_","ward_a_pm_"],
        
        "hope drive clinic am":          "complex_am_",
        "hope drive clinic pm":          "complex_pm_",
        
        "briarcrest clinic am":          "adol_med_am_",
        "briarcrest clinic pm":          "adol_med_pm_",

        "lancaster am":          "lancaster_am_",
        "lancaster pm":          "lancaster_pm_",
    
        'hampden_nursery_print':    'custom_print_hampden_nursery_',
        'sjr_hospitalist_print':    'custom_print_sjr_hospitalist_',
        'aac_print':                'custom_print_aac_',
        'lancaster_cmg_print':      'custom_print_lancaster_cmg_',
    
        'mahoussi_aholoukpe_print': 'custom_print_mahoussi_aholoukpe_',
        
    }
    
    # which keys from base_map for each sheet
    sheet_map = {
        'ETOWN':           ('etown am continuity','etown pm continuity'),
        'NYES':            ('nyes rd am continuity','nyes rd pm continuity'),
        'LANCASTER':            ('lancaster am','lancaster pm'),
        'LANCASTER_CMG':        ('lancaster_cmg_print',),
        
        'COMPLEX':         ('hope drive clinic am','hope drive clinic pm'),
        'WARD A':             ('rounder 1 7a-7p','rounder 2 7a-7p','rounder 3 7a-7p'),
        'PSHCH_NURSERY':    ("nursery weekday 8a-6p","nursery weekday 8a-6p"),
        
        'HAMPDEN_NURSERY': ('hampden_nursery_print',),
        'SJR_HOSP':        ('sjr_hospitalist_print',),
        'AAC':             ('aac_print',),
        'AHOLOUKPE':        ('mahoussi_aholoukpe_print',),
        
        'ADOLMED':             ('briarcrest clinic am','briarcrest clinic pm'),
    }
    
    worksheet_names = ['HOPE_DRIVE','ETOWN','NYES','LANCASTER', 'LANCASTER_CMG', 'COMPLEX','WARD A','PSHCH_NURSERY','HAMPDEN_NURSERY','SJR_HOSP','AAC','AHOLOUKPE','ADOLMED']
    
    for ws in worksheet_names:
        # ─── HOPE_DRIVE ───────────────────────────────────────────
        if ws == 'HOPE_DRIVE':
                    # ─── HOPE_DRIVE: exact same 4‑week AM/PM acute+cont logic ───
            for week_idx in range(1, num_weeks + 1):
                week_base  = (week_idx - 1) * 24
                day_offset = (week_idx - 1) * 7
    
                for day_idx, col in enumerate(excel_column_letters, start=1):
                    is_weekday = day_idx <= 5
                    day_num    = day_idx + day_offset
    
                    # AM acute + continuity
                    if is_weekday:
                        # acute (_1–2)
                        for prov in range(1, 3):
                            row = week_base + hd_row_defs['AM']['acute_start'] + (prov - 1)
                            data_mappings.append({
                                'csv_column':  f'hd_am_acute_d{day_num}_{prov}',
                                'excel_sheet': 'HOPE_DRIVE',
                                'excel_cell':  f'{col}{row}',
                            })
                        # continuity (_1–8)
                        for prov in range(1, 9):
                            row = week_base + hd_row_defs['AM']['cont_start'] + (prov - 1)
                            data_mappings.append({
                                'csv_column':  f'hd_am_d{day_num}_{prov}',
                                'excel_sheet': 'HOPE_DRIVE',
                                'excel_cell':  f'{col}{row}',
                            })
                    else:
                        # weekend acute 1 & 2
                        for acute_type in (1, 2):
                            row = week_base + hd_row_defs['AM']['acute_start'] + (acute_type - 1)
                            data_mappings.append({
                                'csv_column':  f'hd_wknd_acute_{acute_type}_d{day_num}_1',
                                'excel_sheet': 'HOPE_DRIVE',
                                'excel_cell':  f'{col}{row}',
                            })
                        # weekend continuity
                        for prov in range(1, 9):
                            row = week_base + hd_row_defs['AM']['cont_start'] + (prov - 1)
                            data_mappings.append({
                                'csv_column':  f'hd_wknd_am_d{day_num}_{prov}',
                                'excel_sheet': 'HOPE_DRIVE',
                                'excel_cell':  f'{col}{row}',
                            })
    
                    # PM acute + continuity
                    if is_weekday:
                        for prov in range(1, 3):
                            row = week_base + hd_row_defs['PM']['acute_start'] + (prov - 1)
                            data_mappings.append({
                                'csv_column':  f'hd_pm_acute_d{day_num}_{prov}',
                                'excel_sheet': 'HOPE_DRIVE',
                                'excel_cell':  f'{col}{row}',
                            })
                        for prov in range(1, 9):
                            row = week_base + hd_row_defs['PM']['cont_start'] + (prov - 1)
                            data_mappings.append({
                                'csv_column':  f'hd_pm_d{day_num}_{prov}',
                                'excel_sheet': 'HOPE_DRIVE',
                                'excel_cell':  f'{col}{row}',
                            })
                    else:
                        for acute_type in (1, 2):
                            row = week_base + hd_row_defs['PM']['acute_start'] + (acute_type - 1)
                            data_mappings.append({
                                'csv_column':  f'hd_wknd_pm_acute_{acute_type}_d{day_num}_1',
                                'excel_sheet': 'HOPE_DRIVE',
                                'excel_cell':  f'{col}{row}',
                            })
                        for prov in range(1, 9):
                            row = week_base + hd_row_defs['PM']['cont_start'] + (prov - 1)
                            data_mappings.append({
                                'csv_column':  f'hd_wknd_pm_d{day_num}_{prov}',
                                'excel_sheet': 'HOPE_DRIVE',
                                'excel_cell':  f'{col}{row}',
                            })
            # done with HOPE_DRIVE
            continue
    
    
        # ─── W_A (rounders) ───────────────────────────────────────
        if ws == 'W_A':
            mapping_keys = sheet_map[ws]  # ('rounder 1…','rounder 2…','rounder 3…')
            for week_idx in range(1, num_weeks+1):
                week_base  = (week_idx - 1) * 24
                day_offset = (week_idx - 1) * 7
    
                for day_idx, col in enumerate(excel_column_letters, start=1):
                    day_num = day_idx + day_offset
    
                    # AM block → rows 6–…
                    row = week_base + cont_row_defs['AM']
                    for team_idx, key in enumerate(mapping_keys):
                        am_pref = base_map[key][0]  # e.g. "ward_a_am_"
                        provs   = assignments_by_date[date][key]
                        req     = min_required.get(key, len(provs))
                        # pad to exactly 2 providers
                        while len(provs) < req:
                            provs.append(provs[0])
                        offset = team_idx * req
                        for i, name in enumerate(provs, start=1):
                            slot = offset + i     # team1→1,2; team2→3,4; team3→5,6
                            data_mappings.append({
                                'csv_column': f"{am_pref}d{day_num}_{slot}",
                                'excel_sheet': ws,
                                'excel_cell': f"{col}{row}",
                            })
                            row += 1
    
                    # PM block → rows 16–…
                    row = week_base + cont_row_defs['PM']
                    for team_idx, key in enumerate(mapping_keys):
                        pm_pref = base_map[key][1]  # e.g. "ward_a_pm_"
                        provs   = assignments_by_date[date][key]
                        req     = min_required.get(key, len(provs))
                        while len(provs) < req:
                            provs.append(provs[0])
                        offset = team_idx * req
                        for i, name in enumerate(provs, start=1):
                            slot = offset + i
                            data_mappings.append({
                                'csv_column': f"{pm_pref}d{day_num}_{slot}",
                                'excel_sheet': ws,
                                'excel_cell': f"{col}{row}",
                            })
                            row += 1
    
            continue  # skip the generic logic below
    
        # ─── ALL OTHER SHEETS ──────────────────────────────────────
        mapping_keys = sheet_map.get(ws, ())
        if not mapping_keys:
            continue
    
        for key in mapping_keys:
            val = base_map[key]
        
            # Decide which side(s) this key applies to
            am_prefix = pm_prefix = None
            if isinstance(val, list):
                am_prefix, pm_prefix = val
            else:
                k = key.lower()
                if " am " in k:
                    am_prefix = val
                elif " pm " in k:
                    pm_prefix = val
                else:
                    # keys that don't encode AM/PM (rare) write to both
                    am_prefix = pm_prefix = val
    
            for week_idx in range(1, num_weeks + 1):
                week_base  = (week_idx - 1) * 24
                day_offset = (week_idx - 1) * 7
    
                for day_idx, col in enumerate(excel_column_letters, start=1):
                    day_num = day_idx + day_offset
    
                    # AM continuity (_1–10)
                    for prov in range(1, 11):
                        row = week_base + cont_row_defs['AM'] + (prov - 1)
                        data_mappings.append({
                            'csv_column': f"{am_prefix}d{day_num}_{prov}",
                            'excel_sheet': ws,
                            'excel_cell': f"{col}{row}",
                        })
    
                    # PM continuity (_1-10)
                    for prov in range(1, 11):
                        row = week_base + cont_row_defs['PM'] + (prov - 1)
                        data_mappings.append({
                            'csv_column': f"{pm_prefix}d{day_num}_{prov}",
                            'excel_sheet': ws,
                            'excel_cell': f"{col}{row}",
                        })
    
        def hide_blank_rows_all_sheets(excel_bytes: bytes):
            """
            Hide rows where col A starts with AM/PM and ALL of B..H are empty.
            Works for every sheet. Preserves row indices so conditional formats stay aligned.
        
            Returns (new_excel_bytes, per_sheet_hidden_counts, total_hidden)
            """
            import io, re
            from openpyxl import load_workbook
        
            def _empty(v):
                return v is None or (isinstance(v, str) and v.strip() == "")
        
            wb = load_workbook(io.BytesIO(excel_bytes))
            per_sheet = {}
            total = 0
        
            for ws in wb.worksheets:
                hidden = 0
                for r in range(1, ws.max_row + 1):
                    a1 = ws.cell(row=r, column=1).value
                    if not (isinstance(a1, str) and re.match(r"^\s*(AM|PM)\b", a1, re.IGNORECASE)):
                        continue
                    if all(_empty(ws.cell(row=r, column=c).value) for c in range(2, 9)):
                        ws.row_dimensions[r].hidden = True
                        ws.row_dimensions[r].height = 0
                        hidden += 1
                per_sheet[ws.title] = hidden
                total += hidden
        
            out = io.BytesIO()
            wb.save(out)
            out.seek(0)
            return out.getvalue(), per_sheet, total
                
    # --- Main execution flow for generating and then updating the workbook ---
    st.subheader("Generate & Update OPD.xlsx + Summary")
    
    if st.button("Generate OPD File For Sarah to Load Students"):
        # 1) Generate the initial OPD workbook
        excel_template_bytes = generate_opd_workbook(out_df)
        if not excel_template_bytes:
            st.error("Failed to generate OPD template.")
            st.stop()
    
        # 2) Update it with your CSV data
        updated_excel_bytes = update_excel_from_csv(excel_template_bytes, csv_full, data_mappings)
        if not updated_excel_bytes:
            st.error("Failed to update OPD.xlsx with data.")
            st.stop()

        cleaned_bytes, hidden_map, hidden_total = hide_blank_rows_all_sheets(updated_excel_bytes)
        st.success("✅ OPD.xlsx updated successfully!")
    
        # 3) Build your summary DataFrame (reuse your df_summary logic)
        summary = []
        for student in legal_names:
            entry = {"Student": student}
            for w in range(4):
                days = [d + w*7 for d in range(1,6)]
                assigns = []
                # Ward A
                ward_found = False
                for shift in ("am","pm"):
                    for slot in range(1,7):
                        for d in days:
                            key = f"ward_a_{shift}_d{d}_{slot}"
                            if student in redcap_row.get(key,""):
                                assigns.append("Ward A")
                                ward_found = True
                                break
                        if ward_found: break
                    if ward_found: break
                # Hampden
                if not ward_found:
                    for d in days:
                        key = f"custom_print_hampden_nursery_d{d}_4"
                        if student in redcap_row.get(key,""):
                            assigns.append("Hampden")
                            break
                # SJR
                sjr_found = False
                for slot in (3,4):
                    for d in days:
                        key = f"custom_print_sjr_hospitalist_d{d}_{slot}"
                        if student in redcap_row.get(key,""):
                            assigns.append("SJR")
                            sjr_found = True
                            break
                    if sjr_found: break
                # PSHCH
                pshch_found = False
                for slot in (1,2):
                    for d in days:
                        for pref in ("nursery_am_","nursery_pm_"):
                            key = f"{pref}d{d}_{slot}"
                            if student in redcap_row.get(key,""):
                                assigns.append("PSHCH")
                                pshch_found = True
                                break
                        if pshch_found: break
                    if pshch_found: break
    
                entry[f"Week {w+1}"] = ", ".join(assigns) or ""
            summary.append(entry)
        df_summary = pd.DataFrame(summary)
    
        # 4) Build a Word doc with the summary table
        doc = Document()
        # make landscape
        section = doc.sections[0]
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width, section.page_height = section.page_height, section.page_width
        
        doc.add_heading("Assignment Summary by Week", level=1)
        
        cols  = df_summary.columns.tolist()
        table = doc.add_table(rows=1, cols=len(cols), style="Table Grid")
        hdr_cells = table.rows[0].cells
        for i, c in enumerate(cols):
            hdr_cells[i].text = c
        
        for _, row in df_summary.iterrows():
            row_cells = table.add_row().cells
            for i, c in enumerate(cols):
                row_cells[i].text = str(row[c])
        
        # **Save** into bytes
        word_io = io.BytesIO()
        doc.save(word_io)
        word_io.seek(0)
        word_bytes = word_io.read()
        
        # 5) Package into a ZIP
        zip_io = io.BytesIO()
        with zipfile.ZipFile(zip_io, "w") as z:
            z.writestr("Updated_OPD.xlsx", cleaned_bytes)
            z.writestr("Assignment_Summary.docx", word_bytes)
        zip_io.seek(0)
        
        # 6) Single download
        st.download_button(label="⬇️ Download OPD.xlsx + Summary (zip)",data=zip_io.read(),file_name="Batch_Output.zip",mime="application/zip")

elif mode == "Create Student Schedule":
    st.subheader("Create Student Schedule")

    def create_ms_schedule_template(students, dates):
        buf = io.BytesIO()
        wb = xlsxwriter.Workbook(buf, {'in_memory': True, 'strings_to_formulas': False, 'strings_to_urls': False})
    
        # — Formats —
        f1 = wb.add_format({'font_size':14,'bold':1,'align':'center','valign':'vcenter',
                            'font_color':'black','text_wrap':True,'bg_color':'#FEFFCC','border':1})
        f2 = wb.add_format({'font_size':10,'bold':1,'align':'center','valign':'vcenter',
                            'font_color':'yellow','bg_color':'black','border':1,'text_wrap':True})
        f3 = wb.add_format({'font_size':12,'bold':1,'align':'center','valign':'vcenter',
                            'font_color':'black','bg_color':'#FFC7CE','border':1})
        f4 = wb.add_format({'num_format':'mm/dd/yyyy','font_size':12,'bold':1,'align':'center',
                            'valign':'vcenter','font_color':'black','bg_color':'#F4F6F7','border':1})
        f5 = wb.add_format({'font_size':12,'bold':1,'align':'center','valign':'vcenter',
                            'font_color':'black','bg_color':'#F4F6F7','border':1})
        f6 = wb.add_format({'bg_color':'black','border':1})
        f7 = wb.add_format({'font_size':12,'bold':1,'align':'center','valign':'vcenter',
                            'font_color':'black','bg_color':'#90EE90','border':1})
        f8 = wb.add_format({'font_size':12,'bold':1,'align':'center','valign':'vcenter',
                            'font_color':'black','bg_color':'#89CFF0','border':1})
    
        days = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']
        start_rows = [2, 10, 18, 26]
        weeks = ['Week 1','Week 2','Week 3','Week 4']
        due_texts = [
            '',
            '',
            '',
            'All Clinical Encounter Logs are Due, Solicitation of Clinical Assessments, Observed H&Ps and Observed Handoff Due']
    
        used_titles = set()
        for name in students:
            safe = re.sub(r"[\[\]:*?/\\]", "-", str(name)).strip().strip("'") or "Student"
            title, suffix_number = safe[:31], 1
            while title.casefold() in used_titles:
                suffix_number += 1
                suffix = f"_{suffix_number}"
                title = safe[:31-len(suffix)] + suffix
            used_titles.add(title.casefold())
            ws = wb.add_worksheet(title)
            ws.set_zoom(70)
    
            # Header
            ws.merge_range('A1:A2','Student Name:', f1)
            ws.merge_range('B1:B2',      str(name),      f1)
            #note = ("*Note* Protected Self-Study Time is for coursework only. During this time period, "
            #        "we expect students to do coursework, be available for any additional educational "
            #        "activities, and any extra clinical time that may be available. If the student is not "
            #        "available during this time period and has not made an absence request, the student "
            #        "will be cited for unprofessionalism and will risk failing the course.")
            
            note = ("*Note* Protected Self-Study Time is reserved for independent learning and completion of required "
                    "coursework. Students are encouraged to use this time to complete assignments, review course materials, "
                    "prepare for patient care, and reinforce concepts encountered during the clerkship.")
            
            ws.merge_range('C1:H2', note, f2)
    
            # Column widths & row height
            ws.set_column('A:A', 20)
            ws.set_column('B:B', 30)
            ws.set_column('C:G', 40)
            ws.set_column('H:H',155)
            ws.set_row(0, 37.25)
    
            # Days headers and dates
            date_idx = 0
            for block, row in enumerate(start_rows):
                # 1) write the days on row `row`, cols B–H
                for col_offset, day in enumerate(days, start=1):
                    ws.write(row, col_offset, day, f3)
        
            # 2) write the dates directly beneath in B–H (row+1)
                for col_offset in range(7):
                    if date_idx < len(dates):
                        # note the +1 here instead of +2
                        ws.write(row+1, col_offset+1, dates[date_idx], f4)
                        date_idx += 1
    
            # Week labels
            for i, week in enumerate(weeks):
                row = 4 + (i * 8)
                ws.write(f'A{row}', week, f3)
    
            # AM / PM labels
            for i in range(4):
                ws.write(f'A{6 + i*8}', 'AM', f3)
                ws.write(f'A{7 + i*8}', 'PM', f3)
    
            # Fill AM/PM blocks with Asynchronous Time (cols C–J)
            for block in range(4):
                am_row = 5 + block*8
                pm_row = 6 + block*8
                for col in range(1, 8):
                    ws.write(am_row, col, "Protected Self-Study Time", f5)
                    ws.write(pm_row, col, "Protected Self-Study Time", f5)
    
            # Separators
            for sep in [10, 18, 26, 34]:
                ws.merge_range(f'A{sep}:H{sep}', '', f6)
    
            # Green filler rows
            for filler in [8, 16, 24, 32]:
                for col in range(8):
                    ws.write(filler, col, ' ', f7)
    
            # Assignment‑due rows
            for i, base in enumerate([8, 16, 24, 32]):
                ws.write(f'A{base}', 'ASSIGNMENT DUE:', f8)
                for col in range(1, 8):
                    if col == 5:
                        ws.write(base-1, col, 'Ask for Feedback!', f8)
                    elif col == 7:
                        ws.write(base-1, col, due_texts[i], f8)
                    else:
                        ws.write(base-1, col, ' ', f8)
    
        wb.close()
        buf.seek(0)
        return buf
    

    # Archive the original OPD at this upload step, before producing schedules.
    try:
        archive_client = GitHubOPDArchive(get_opd_archive_config())
        _opd_scope_archive_session(archive_client.config)
    except OPDArchiveError as exc:
        st.error(str(exc))
        st.info("Complete the GitHub/Streamlit Secrets setup first. OPDs will not be processed without a verified archive.")
        st.stop()

    source_choice = st.radio("OPD source", ("Upload new/revised OPD", "Reload archived OPD"),
                             key="opd_source_choice", horizontal=True)
    raw_opd, opd_details, archive_ok = None, None, False
    if source_choice == "Upload new/revised OPD":
        opd_upload = st.file_uploader("Upload original OPD.xlsx", type=["xlsx"], key="opd_main",
                                      on_change=_opd_upload_changed)
        if opd_upload is not None:
            try:
                raw_opd, opd_details, archive_ok = _opd_archive_upload_ui(opd_upload, archive_client)
            except OPDArchiveError as exc:
                st.error(str(exc))
    else:
        loaded_opd = _opd_archive_picker(archive_client, "schedule_archive")
        if loaded_opd:
            raw_opd, opd_details, archive_ok = loaded_opd["raw"], loaded_opd["details"], True

    rot_upload = st.file_uploader("Upload Rotation Schedule (.xlsx or .csv)", type=["xlsx", "csv"],
                                  key="rot_main")
    df_rot = None
    if rot_upload is not None:
        try:
            rotation_bytes = rot_upload.getvalue()
            df_rot = (pd.read_csv(BytesIO(rotation_bytes)) if rot_upload.name.lower().endswith(".csv")
                      else pd.read_excel(BytesIO(rotation_bytes)))
        except Exception:
            st.error("The rotation list could not be read. Check the uploaded CSV or Excel file.")

    if raw_opd is None or df_rot is None:
        st.info("Upload or reload an OPD and upload its rotation list to create student schedules.")
        st.session_state.pop("opd_generated_master", None)
    else:
        if not {"legal_name", "start_date"}.issubset(df_rot.columns):
            st.error("The rotation list must contain legal_name and start_date columns.")
            st.stop()
        roster_df = df_rot.loc[df_rot["legal_name"].notna()].copy()
        roster_df["legal_name"] = roster_df["legal_name"].astype(str).str.strip()
        roster_df = roster_df.loc[roster_df["legal_name"].ne("")]
        roster_df["start_date"] = pd.to_datetime(roster_df["start_date"], errors="coerce")
        if roster_df.empty or roster_df["start_date"].isna().any():
            st.error("Each student in the rotation list must have a valid start_date.")
            st.stop()
        roster_mondays = {
            value.date() - timedelta(days=value.weekday()) for value in roster_df["start_date"]
        }
        if roster_mondays != {opd_details["rotation_start"]}:
            st.error(f"The rotation list does not match this OPD's first Monday "
                     f"({opd_details['rotation_start']:%m/%d/%Y}). Upload the matching rotation list. "
                     "The original OPD is archived under its own dates, not the roster's dates.")
            st.stop()
        if len(opd_details["week_mondays"]) != 4:
            st.error("The original has been archived, but the existing MS_Schedule template requires exactly four weeks.")
            st.stop()

        students = list({_opd_name_key(name): name for name in roster_df["legal_name"]}.values())
        name_order = st.selectbox("Names around '~' in the OPD", OPD_NAME_ORDERS,
                                  key="opd_name_order",
                                  help="Auto-detect matches the student to legal_name in the rotation list. "
                                       "It accepts either name order and ignores spaces around '~'.")
        signature = (hashlib.sha256(raw_opd).hexdigest(), hashlib.sha256(rotation_bytes).hexdigest(), name_order)
        if st.session_state.get("opd_master_signature") != signature:
            st.session_state.pop("opd_generated_master", None)
            st.session_state["opd_master_signature"] = signature
        try:
            assignments, unmatched = collect_opd_assignments(raw_opd, students, name_order)
        except OPDArchiveError as exc:
            st.error(str(exc))
            st.stop()
        if unmatched:
            st.warning(f"{len(unmatched)} filled assignment cell(s) do not match this rotation list/name order "
                       "and will not populate a student schedule. This can include students from another course.")
            with st.expander("Unmatched assignment cells"):
                st.write(", ".join(unmatched))
        assigned_students = {item["student"] for item in assignments}
        if any(name not in assigned_students for name in students):
            st.warning("Some students have no matching OPD assignments. Review the roster and name-order setting before building.")
        conflicts = defaultdict(list)
        for item in assignments:
            conflicts[(item["student"], item["week"], item["shift"], item["column"])].append(item)
        has_conflicts = False
        days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        for (student, week, shift, col), items in conflicts.items():
            if len(items) > 1:
                has_conflicts = True
                locations = "; ".join(f"{item['site']}@{item['coordinate']}" for item in items)
                st.warning(f"Week {week+1} {days[col-2]} {shift}: {student} double-booked ({locations}). "
                           "As in the previous version, the last assignment read fills that schedule cell.")
        if not has_conflicts:
            st.success("No AM/PM shift conflicts detected among matched students.")
        st.caption(f"OPD rotation: {opd_details['rotation_start']:%B %d, %Y}; "
                   f"{len(students)} student(s); {len(assignments)} matched assignment cell(s).")
        if not archive_ok:
            st.error("Schedule generation is disabled until the original OPD is successfully archived.")
            st.session_state.pop("opd_generated_master", None)
        if st.button("Create & Download Fully-Populated MS_Schedule", disabled=not archive_ok,
                     key="opd_build_master"):
            try:
                # Do not silently reuse or overwrite an older session's source.
                current_archive = archive_client.load(opd_details["rotation_start"])
                if not hmac.compare_digest(current_archive["raw"], raw_opd):
                    raise OPDArchiveError("A different OPD is now current in the archive for this rotation. "
                                          "Reload the latest archived OPD, or deliberately upload your revised file again.")
                dates = pd.date_range(start=opd_details["rotation_start"], periods=28, freq="D").tolist()
                blank_buf = create_ms_schedule_template(students, dates)
                st.session_state["opd_generated_master"] = populate_ms_schedule(blank_buf.getvalue(), assignments)
            except OPDArchiveError as exc:
                st.session_state.pop("opd_generated_master", None)
                st.error(str(exc))
        if archive_ok and st.session_state.get("opd_generated_master"):
            st.download_button("Download MS_Schedule.xlsx", data=st.session_state["opd_generated_master"],
                               file_name="MS_Schedule.xlsx", mime=OPD_XLSX_MIME, key="opd_master_download")
            st.caption("Next: open Create Individual Schedules and upload this MS_Schedule.xlsx. "
                       "The individual student ZIP and Power Automate preceptor workbook remain available there.")


elif mode == "Create Individual Schedules":
    st.subheader("Individual Schedule Creator")

    # -------------------------------------------------------------------------
    # HARD-CODED PRECEPTOR EMAIL MAP
    # -------------------------------------------------------------------------
    # Add each preceptor exactly as the name appears in the OPD/student schedule.
    # Matching ignores capitalization and extra spaces.
    #
    # Example:
    # PRECEPTOR_EMAIL_MAP = {
    #     "Smith, Jane": "jsmith@pennstatehealth.psu.edu",
    #     "Jones, Robert": "rjones@pennstatehealth.psu.edu",
    # }
    PRECEPTOR_EMAIL_MAP = {
        # "Preceptor Name": "preceptor_email@pennstatehealth.psu.edu",
    }

    FOCUS_SITES = {"HOPE_DRIVE", "NYES", "ETOWN"}
    REPORT_COLUMNS = [
        "preceptor_name",
        "student_name",
        "no_of_sessions",
        "monday_date",
        "primary_preceptor",
        "fragmented_preceptor",
        "primary_preceptor_flag",
        "primary_preceptor_flag_reason",
        "email",
    ]

    # Increment this whenever the generated report columns or output logic change.
    # Streamlit keeps session_state across app reruns/deployments, so an older
    # preview DataFrame may otherwise remain cached without newly added columns.
    INDIVIDUAL_REPORT_SCHEMA_VERSION = 3
    INDIVIDUAL_OUTPUT_STATE_KEYS = (
        "individual_schedule_zip",
        "individual_preceptor_report",
        "individual_preceptor_preview",
        "individual_missing_emails",
    )

    if (
        st.session_state.get("individual_report_schema_version")
        != INDIVIDUAL_REPORT_SCHEMA_VERSION
    ):
        for state_key in INDIVIDUAL_OUTPUT_STATE_KEYS:
            st.session_state.pop(state_key, None)
        st.session_state["individual_report_schema_version"] = (
            INDIVIDUAL_REPORT_SCHEMA_VERSION
        )

    uploaded = st.file_uploader(
        "Upload the master Excel (.xlsx) with one tab per person",
        type=["xlsx"],
        key="individual_schedule_master",
    )

    def _normalize_name(value):
        """Normalize names for reliable hard-coded email matching."""
        return re.sub(r"\s+", " ", str(value or "").strip()).casefold()

    def _normalize_site(value):
        """Convert site labels such as 'Hope Drive' to 'HOPE_DRIVE'."""
        cleaned = re.sub(r"[^A-Za-z0-9]+", "_", str(value or "").strip().upper())
        return cleaned.strip("_")

    NORMALIZED_EMAIL_MAP = {
        _normalize_name(preceptor): str(email).strip()
        for preceptor, email in PRECEPTOR_EMAIL_MAP.items()
        if str(preceptor).strip()
    }

    def _parse_excel_date(value):
        """Return a Python date from a typical openpyxl/pandas date value."""
        if value is None or value == "":
            return None

        if isinstance(value, datetime):
            return value.date()

        if hasattr(value, "date") and not isinstance(value, str):
            try:
                return value.date()
            except Exception:
                pass

        try:
            parsed = pd.to_datetime(value, errors="coerce")
            if pd.notna(parsed):
                return parsed.date()
        except Exception:
            pass

        return None

    def _week_monday_from_date_row(ws, date_row):
        """Find/infer Monday from the B:H date cells for a weekly block."""
        for day_offset, col in enumerate(range(2, 9)):
            parsed = _parse_excel_date(ws.cell(row=date_row, column=col).value)
            if parsed is not None:
                return parsed - timedelta(days=day_offset)
        return None

    def _student_name_from_sheet(ws):
        """Prefer the displayed student name; fall back to the worksheet title."""
        for coordinate in ("B1", "B2"):
            value = ws[coordinate].value
            if value is not None and str(value).strip():
                return str(value).strip()
        return str(ws.title).strip()

    def _parse_schedule_assignment(value):
        """Parse a student-schedule cell formatted as 'Preceptor - [SITE]'."""
        if not isinstance(value, str) or not value.strip():
            return None, None

        match = re.match(r"^\s*(.*?)\s*-\s*\[\s*([^\]]+)\s*\]\s*$", value)
        if not match:
            return None, None

        preceptor = re.sub(r"\s+", " ", match.group(1).strip())
        site = _normalize_site(match.group(2))
        return preceptor, site

    def build_preceptor_assignment_report(master_wb):
        """
        Build one row per student/preceptor/week for HOPE_DRIVE, NYES, and ETOWN.

        Primary logic:
          - Every student/week represented in this report receives exactly one
            primary preceptor.
          - Prefer the preceptor with the largest session count among those with
            at least 3 sessions.
          - If nobody reaches 3 sessions, select the highest-session preceptor
            anyway and flag that primary assignment for review.
          - The same preceptor may be primary for more than one student. Repeated
            primary assignments within the same week are also flagged.
          - Ties are resolved alphabetically for stable output.

        Fragmentation logic:
          - YES when that preceptor has <3 sessions with the student that week.
          - A count of exactly 3 is not fragmented.
        """
        weekly_rows = []

        # Rows created by create_ms_schedule_template:
        # week 1 dates/AM/PM = 4/6/7; then every 8 rows for weeks 2-4.
        week_layout = [
            {"date_row": 4, "session_rows": (6, 7)},
            {"date_row": 12, "session_rows": (14, 15)},
            {"date_row": 20, "session_rows": (22, 23)},
            {"date_row": 28, "session_rows": (30, 31)},
        ]

        for ws in master_wb.worksheets:
            student_name = _student_name_from_sheet(ws)

            for layout in week_layout:
                monday_date = _week_monday_from_date_row(ws, layout["date_row"])
                if monday_date is None:
                    continue

                preceptor_counts = Counter()

                for session_row in layout["session_rows"]:
                    for col in range(2, 9):  # Monday-Sunday, B:H
                        preceptor, site = _parse_schedule_assignment(
                            ws.cell(row=session_row, column=col).value
                        )
                        if not preceptor or site not in FOCUS_SITES:
                            continue
                        preceptor_counts[preceptor] += 1

                if not preceptor_counts:
                    continue

                # Every student/week must have one primary. Prefer >=3 sessions;
                # otherwise choose the best available preceptor and flag it.
                eligible_primary = [
                    (preceptor, count)
                    for preceptor, count in preceptor_counts.items()
                    if count >= 3
                ]
                primary_pool = eligible_primary or list(preceptor_counts.items())
                primary_pool.sort(
                    key=lambda item: (-item[1], _normalize_name(item[0]))
                )
                primary_name, primary_count = primary_pool[0]
                below_threshold = primary_count < 3

                for preceptor, count in preceptor_counts.items():
                    is_primary = preceptor == primary_name
                    flag_reasons = []
                    if is_primary and below_threshold:
                        flag_reasons.append("SELECTED PRIMARY HAS FEWER THAN 3 SESSIONS")

                    weekly_rows.append(
                        {
                            "preceptor_name": preceptor,
                            "student_name": student_name,
                            "no_of_sessions": int(count),
                            "monday_date": monday_date,
                            "primary_preceptor": "YES" if is_primary else "NO",
                            "fragmented_preceptor": "YES" if count < 3 else "NO",
                            "primary_preceptor_flag": "YES" if flag_reasons else "NO",
                            "primary_preceptor_flag_reason": "; ".join(flag_reasons),
                            "email": NORMALIZED_EMAIL_MAP.get(
                                _normalize_name(preceptor), ""
                            ),
                        }
                    )

        report_df = pd.DataFrame(weekly_rows, columns=REPORT_COLUMNS)

        if not report_df.empty:
            # Flag a provider who is serving as primary for multiple students in
            # the same week. Reuse is allowed; the flag simply makes it visible.
            primary_mask = report_df["primary_preceptor"].eq("YES")
            primary_rows = report_df.loc[
                primary_mask,
                ["monday_date", "preceptor_name", "student_name"],
            ].copy()
            primary_rows["_normalized_preceptor"] = primary_rows[
                "preceptor_name"
            ].map(_normalize_name)

            repeated_keys = set(
                primary_rows.groupby(
                    ["monday_date", "_normalized_preceptor"], dropna=False
                )["student_name"]
                .nunique()
                .loc[lambda counts: counts > 1]
                .index.tolist()
            )

            for row_idx in report_df.index[primary_mask]:
                key = (
                    report_df.at[row_idx, "monday_date"],
                    _normalize_name(report_df.at[row_idx, "preceptor_name"]),
                )
                if key not in repeated_keys:
                    continue

                reason = str(
                    report_df.at[row_idx, "primary_preceptor_flag_reason"] or ""
                ).strip()
                repeat_reason = "PRECEPTOR IS PRIMARY FOR MULTIPLE STUDENTS THIS WEEK"
                report_df.at[row_idx, "primary_preceptor_flag"] = "YES"
                report_df.at[row_idx, "primary_preceptor_flag_reason"] = (
                    f"{reason}; {repeat_reason}" if reason else repeat_reason
                )

            report_df["_primary_sort"] = report_df["primary_preceptor"].map(
                {"YES": 0, "NO": 1}
            )
            report_df = (
                report_df.sort_values(
                    by=[
                        "monday_date",
                        "student_name",
                        "_primary_sort",
                        "no_of_sessions",
                        "preceptor_name",
                    ],
                    ascending=[True, True, True, False, True],
                    kind="stable",
                )
                .drop(columns=["_primary_sort"])
                .reset_index(drop=True)
            )

        return report_df

    def build_preceptor_report_workbook(report_df):
        """
        Create a Power Automate-ready .xlsx workbook using XlsxWriter.

        The report is written as a genuine Excel table without adding a second,
        overlapping worksheet AutoFilter. Avoiding that overlap prevents Excel's
        "We found a problem with some content" repair warning.
        """
        output = BytesIO()
        workbook = xlsxwriter.Workbook(
            output,
            {
                "in_memory": True,
                "strings_to_formulas": False,
                "strings_to_urls": False,
            },
        )

        worksheet = workbook.add_worksheet("Preceptor Assignments")
        worksheet.freeze_panes(1, 0)
        worksheet.set_zoom(90)
        worksheet.set_column("A:A", 28)
        worksheet.set_column("B:B", 28)
        worksheet.set_column("C:C", 16)
        worksheet.set_column("D:D", 15)
        worksheet.set_column("E:E", 20)
        worksheet.set_column("F:F", 23)
        worksheet.set_column("G:G", 23)
        worksheet.set_column("H:H", 58)
        worksheet.set_column("I:I", 38)

        header_format = workbook.add_format(
            {
                "bold": True,
                "font_color": "#FFFFFF",
                "bg_color": "#1F4E78",
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "border_color": "#D9E2F3",
            }
        )
        body_format = workbook.add_format(
            {
                "valign": "top",
                "bottom": 1,
                "bottom_color": "#D9E2F3",
            }
        )
        integer_format = workbook.add_format(
            {
                "valign": "top",
                "align": "center",
                "num_format": "0",
                "bottom": 1,
                "bottom_color": "#D9E2F3",
            }
        )
        date_format = workbook.add_format(
            {
                "valign": "top",
                "align": "center",
                "num_format": "mm/dd/yyyy",
                "bottom": 1,
                "bottom_color": "#D9E2F3",
            }
        )
        primary_formats = [
            workbook.add_format(
                {
                    "valign": "top",
                    "bg_color": "#E2F0D9",
                    "bottom": 1,
                    "bottom_color": "#D9E2F3",
                    **({"align": "center", "num_format": "0"} if col == 2 else {}),
                    **({"align": "center", "num_format": "mm/dd/yyyy"} if col == 3 else {}),
                }
            )
            for col in range(len(REPORT_COLUMNS))
        ]
        fragmented_formats = [
            workbook.add_format(
                {
                    "valign": "top",
                    "bg_color": "#FFF2CC",
                    "bottom": 1,
                    "bottom_color": "#D9E2F3",
                    **({"align": "center", "num_format": "0"} if col == 2 else {}),
                    **({"align": "center", "num_format": "mm/dd/yyyy"} if col == 3 else {}),
                }
            )
            for col in range(len(REPORT_COLUMNS))
        ]
        flagged_primary_formats = [
            workbook.add_format(
                {
                    "valign": "top",
                    "bg_color": "#FCE4D6",
                    "font_color": "#9C0006",
                    "bottom": 1,
                    "bottom_color": "#D9E2F3",
                    **({"align": "center", "num_format": "0"} if col == 2 else {}),
                    **({"align": "center", "num_format": "mm/dd/yyyy"} if col == 3 else {}),
                }
            )
            for col in range(len(REPORT_COLUMNS))
        ]
        missing_email_format = workbook.add_format(
            {
                "valign": "top",
                "bg_color": "#FCE4D6",
                "bottom": 1,
                "bottom_color": "#D9E2F3",
            }
        )

        # Write headers explicitly. The table is added after the data is written.
        for col_idx, header in enumerate(REPORT_COLUMNS):
            worksheet.write(0, col_idx, header, header_format)

        for row_offset, row in enumerate(
            report_df.itertuples(index=False, name=None), start=1
        ):
            is_primary = str(row[4]).strip().upper() == "YES"
            is_fragmented = str(row[5]).strip().upper() == "YES"
            is_flagged_primary = str(row[6]).strip().upper() == "YES"

            for col_idx, value in enumerate(row):
                if is_flagged_primary:
                    cell_format = flagged_primary_formats[col_idx]
                elif is_primary:
                    cell_format = primary_formats[col_idx]
                elif is_fragmented:
                    cell_format = fragmented_formats[col_idx]
                elif col_idx == 2:
                    cell_format = integer_format
                elif col_idx == 3:
                    cell_format = date_format
                else:
                    cell_format = body_format

                if col_idx == len(REPORT_COLUMNS) - 1 and not str(value or "").strip():
                    cell_format = missing_email_format

                if col_idx == 3:
                    parsed_date = _parse_excel_date(value)
                    if parsed_date is not None:
                        worksheet.write_datetime(
                            row_offset,
                            col_idx,
                            datetime.combine(parsed_date, datetime.min.time()),
                            cell_format,
                        )
                    else:
                        worksheet.write_blank(row_offset, col_idx, None, cell_format)
                elif value is None or (isinstance(value, float) and pd.isna(value)):
                    worksheet.write_blank(row_offset, col_idx, None, cell_format)
                else:
                    worksheet.write(row_offset, col_idx, value, cell_format)

        # Power Automate needs a named Excel table. Only the table owns the filter;
        # do not also call worksheet.autofilter() on the same range.
        if not report_df.empty:
            worksheet.add_table(
                0,
                0,
                len(report_df),
                len(REPORT_COLUMNS) - 1,
                {
                    "name": "PreceptorAssignmentTable",
                    "style": "Table Style Medium 2",
                    "columns": [{"header": header} for header in REPORT_COLUMNS],
                },
            )

        notes = workbook.add_worksheet("Definitions")
        notes.set_column("A:A", 24)
        notes.set_column("B:B", 90)
        notes_format = workbook.add_format({"valign": "top", "text_wrap": True})
        notes_label_format = workbook.add_format(
            {"bold": True, "valign": "top", "text_wrap": True}
        )
        definitions = [
            ("Report scope", "HOPE_DRIVE, NYES, and ETOWN only"),
            (
                "Primary preceptor",
                "Exactly one per student/week. The app prefers the highest-session "
                "preceptor with at least 3 sessions. If nobody reaches 3, the "
                "highest-session preceptor is still selected and flagged.",
            ),
            (
                "Repeated primary",
                "The same preceptor may be primary for multiple students. Those "
                "primary rows are flagged for visibility.",
            ),
            ("Fragmented preceptor", "YES when no_of_sessions < 3."),
            (
                "Primary preceptor flag",
                "YES when the selected primary has fewer than 3 sessions or the "
                "same preceptor is primary for multiple students that week.",
            ),
            (
                "Email mapping",
                "Emails come from PRECEPTOR_EMAIL_MAP in the Streamlit source code.",
            ),
        ]
        for row_idx, (label, definition) in enumerate(definitions):
            notes.write(row_idx, 0, label, notes_label_format)
            notes.write(row_idx, 1, definition, notes_format)

        missing_names = (
            sorted(
                report_df.loc[
                    report_df["email"].fillna("").astype(str).str.strip().eq(""),
                    "preceptor_name",
                ]
                .dropna()
                .unique()
                .tolist(),
                key=_normalize_name,
            )
            if not report_df.empty
            else []
        )

        if missing_names:
            missing_ws = workbook.add_worksheet("Missing Emails")
            missing_ws.freeze_panes(1, 0)
            missing_ws.set_column("A:A", 32)
            missing_ws.set_column("B:B", 70)
            missing_ws.write(0, 0, "preceptor_name", header_format)
            missing_ws.write(0, 1, "action_needed", header_format)
            for row_idx, name in enumerate(missing_names, start=1):
                missing_ws.write(row_idx, 0, name, body_format)
                missing_ws.write(
                    row_idx,
                    1,
                    "Add this name and email to PRECEPTOR_EMAIL_MAP in app_sch_2026.py",
                    body_format,
                )

        workbook.close()
        output.seek(0)
        return output, missing_names

    def copy_sheet_to_new_wb(src_ws):
        """Return a BytesIO of a new .xlsx containing src_ws with formatting."""
        from openpyxl import Workbook
        from io import BytesIO

        wb_new = Workbook()
        ws_new = wb_new.active
        ws_new.title = src_ws.title[:31]

        # Column widths & visibility
        for col_letter, dim in src_ws.column_dimensions.items():
            if dim.width is not None:
                ws_new.column_dimensions[col_letter].width = dim.width
            ws_new.column_dimensions[col_letter].hidden = dim.hidden

        # Row heights & visibility
        for idx, dim in src_ws.row_dimensions.items():
            if dim.height is not None:
                ws_new.row_dimensions[idx].height = dim.height
            ws_new.row_dimensions[idx].hidden = dim.hidden

        # Sheet settings (best-effort)
        try:
            ws_new.sheet_format.defaultColWidth = src_ws.sheet_format.defaultColWidth
            ws_new.sheet_format.defaultRowHeight = src_ws.sheet_format.defaultRowHeight
        except Exception:
            pass
        ws_new.freeze_panes = src_ws.freeze_panes
        try:
            ws_new.page_setup.orientation = src_ws.page_setup.orientation
            ws_new.page_setup.fitToWidth = src_ws.page_setup.fitToWidth
            ws_new.page_setup.fitToHeight = src_ws.page_setup.fitToHeight
            ws_new.page_margins = copy(src_ws.page_margins)
            ws_new.print_options.horizontalCentered = src_ws.print_options.horizontalCentered
            ws_new.print_options.verticalCentered = src_ws.print_options.verticalCentered
            ws_new.print_area = src_ws.print_area
        except Exception:
            pass

        try:
            # Set zoom to 70%
            ws_new.sheet_view.zoomScale = 70

            # Set column widths
            ws_new.column_dimensions["A"].width = 20
            ws_new.column_dimensions["B"].width = 30
            for col in ["C", "D", "E", "F", "G"]:
                ws_new.column_dimensions[col].width = 40
            ws_new.column_dimensions["H"].width = 155
        except Exception:
            pass

        # Copy cells: values + (copied) styles
        for row in src_ws.iter_rows():
            for cell in row:
                # Skip non-master cells from merged ranges
                if isinstance(cell, MergedCell):
                    continue

                # Get a reliable numeric column index
                col_idx = getattr(cell, "col_idx", None)
                if col_idx is None:
                    col = cell.column  # may be int or letter depending on version
                    col_idx = col if isinstance(col, int) else column_index_from_string(col)

                # Create target cell with value (formula preserved if present)
                tgt = ws_new.cell(row=cell.row, column=col_idx, value=cell.value)

                # Copy style safely
                if getattr(cell, "has_style", False):
                    try:
                        if cell.font:
                            tgt.font = copy(cell.font)
                        if cell.fill:
                            tgt.fill = copy(cell.fill)
                        if cell.border:
                            tgt.border = copy(cell.border)
                        if cell.alignment:
                            tgt.alignment = copy(cell.alignment)
                        if cell.protection:
                            tgt.protection = copy(cell.protection)
                        tgt.number_format = cell.number_format
                    except Exception:
                        pass

        # Copy merged cell ranges (after values)
        for merged in list(src_ws.merged_cells.ranges):
            try:
                ws_new.merge_cells(str(merged))
            except Exception:
                pass

        # Copy data validations (best-effort)
        try:
            if src_ws.data_validations and src_ws.data_validations.dataValidation:
                from openpyxl.worksheet.datavalidation import DataValidation

                for dv in src_ws.data_validations.dataValidation:
                    dv_new = DataValidation(
                        type=dv.type,
                        formula1=dv.formula1,
                        formula2=dv.formula2,
                        allow_blank=dv.allow_blank,
                        operator=dv.operator,
                        showDropDown=dv.showDropDown,
                        showErrorMessage=dv.showErrorMessage,
                        errorTitle=dv.errorTitle,
                        error=dv.error,
                        promptTitle=dv.promptTitle,
                        prompt=dv.prompt,
                    )
                    for sqref in getattr(dv, "sqref", []):
                        dv_new.add(sqref)
                    ws_new.add_data_validation(dv_new)
        except Exception:
            pass

        # Filters
        try:
            ws_new.auto_filter.ref = getattr(src_ws.auto_filter, "ref", None)
        except Exception:
            pass

        # Save to buffer
        buf = BytesIO()
        wb_new.save(buf)
        buf.seek(0)
        return buf

    if uploaded is not None:
        # Clear previously generated output when a different source file is uploaded.
        source_signature = (uploaded.name, hashlib.sha256(uploaded.getvalue()).hexdigest())
        if st.session_state.get("individual_schedule_source") != source_signature:
            st.session_state["individual_schedule_source"] = source_signature
            st.session_state.pop("individual_schedule_zip", None)
            st.session_state.pop("individual_preceptor_report", None)
            st.session_state.pop("individual_preceptor_preview", None)
            st.session_state.pop("individual_missing_emails", None)

        # Keep formulas/formatting -> data_only=False
        wb = load_workbook(uploaded, data_only=False)
        st.write(f"Found **{len(wb.sheetnames)}** tabs.")
        st.caption(
            "The preceptor report includes only HOPE_DRIVE, NYES, and ETOWN. "
            "Each student/week with focus-site assignments receives one primary preceptor. The app prefers "
            ">=3 sessions; fallback and repeated primary assignments are flagged. "
            "Fragmented = <3 sessions."
        )

        if st.button("Build individual schedules + preceptor report"):
            report_df = build_preceptor_assignment_report(wb)
            report_buf, missing_emails = build_preceptor_report_workbook(report_df)

            zip_buf = BytesIO()
            with ZipFile(zip_buf, mode="w", compression=ZIP_DEFLATED) as zf:
                used_names = defaultdict(int)

                for sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]

                    # Skip truly empty sheets (no cells with value)
                    has_any_value = any(
                        cell.value is not None
                        for row in ws.iter_rows()
                        for cell in row
                    )
                    if not has_any_value:
                        continue

                    # Safe file name
                    base = re.sub(r"[^A-Za-z0-9._-]+", "_", sheet_name).strip("_") or "sheet"
                    used_names[base] += 1
                    safe_name = (
                        base
                        if used_names[base] == 1
                        else f"{base}_{used_names[base]}"
                    )

                    # Copy this sheet into its own new workbook (preserving formatting)
                    out_buf = copy_sheet_to_new_wb(ws)
                    zf.writestr(f"{safe_name}.xlsx", out_buf.getvalue())

                # Include the new report in the same ZIP.
                zf.writestr(
                    "Preceptor_Assignment_Report.xlsx", report_buf.getvalue()
                )

            zip_buf.seek(0)
            st.session_state["individual_schedule_zip"] = zip_buf.getvalue()
            st.session_state["individual_preceptor_report"] = report_buf.getvalue()
            st.session_state["individual_preceptor_preview"] = report_df
            st.session_state["individual_missing_emails"] = missing_emails

        report_preview = st.session_state.get("individual_preceptor_preview")

        # Defensive protection for stale output created by an older app version.
        # A prior preview may be a valid DataFrame but lack newly added columns.
        required_preview_columns = {
            "primary_preceptor",
            "primary_preceptor_flag",
            "primary_preceptor_flag_reason",
        }
        preview_is_current = (
            isinstance(report_preview, pd.DataFrame)
            and required_preview_columns.issubset(report_preview.columns)
        )
        if report_preview is not None and not preview_is_current:
            for state_key in INDIVIDUAL_OUTPUT_STATE_KEYS:
                st.session_state.pop(state_key, None)
            report_preview = None
            st.info(
                "The preceptor report format was updated. Please click "
                "'Build individual schedules + preceptor report' to regenerate both outputs."
            )

        if report_preview is not None:
            st.markdown("**Preceptor assignment report preview**")
            if report_preview.empty:
                st.warning(
                    "No HOPE_DRIVE, NYES, or ETOWN assignments were found in the uploaded workbook."
                )
            else:
                st.dataframe(report_preview, use_container_width=True, hide_index=True)

            flagged_primary_rows = report_preview.loc[
                report_preview["primary_preceptor"].eq("YES")
                & report_preview["primary_preceptor_flag"].eq("YES")
            ]
            if not flagged_primary_rows.empty:
                st.warning(
                    f"{len(flagged_primary_rows)} primary assignment(s) require review. "
                    "See primary_preceptor_flag_reason in the preview or Excel report."
                )
            elif not report_preview.empty:
                st.success("Every primary assignment met the preferred criteria without reuse flags.")

            missing_emails = st.session_state.get("individual_missing_emails", [])
            if missing_emails:
                st.warning(
                    f"{len(missing_emails)} preceptor(s) are missing from "
                    "PRECEPTOR_EMAIL_MAP. Their email cells are blank, and their names "
                    "are listed on the 'Missing Emails' tab."
                )
            elif not report_preview.empty:
                st.success("All preceptors in this report have a mapped email address.")

        if st.session_state.get("individual_preceptor_report") is not None:
            st.download_button(
                label="Download preceptor assignment report (Excel)",
                data=st.session_state["individual_preceptor_report"],
                file_name="Preceptor_Assignment_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        if st.session_state.get("individual_schedule_zip") is not None:
            st.download_button(
                label="Download individual schedules + report (ZIP)",
                data=st.session_state["individual_schedule_zip"],
                file_name="individual_schedules_with_preceptor_report.zip",
                mime="application/zip",
            )


elif mode == "OPD MD PA Conflict Detector":
    import streamlit as st
    import pandas as pd
    from collections import defaultdict

    st.title("OPD MD/PA Double-Booking & Availability")
    st.write(
        "Upload the MD and PA OPD Excel files to scan for double-booked preceptors and to list availability "
        "by site/date/period (including other sites)."
    )

    # -----------------------------
    # Config & Constants
    # -----------------------------
    DAYS = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']
    DEFAULT_FOCUS = ['HOPE_DRIVE','NYES','ETOWN','LANCASTER']

    TRUST_ONLY_AM_PM = True          # Only parse rows with Column A starting AM/PM
    REQUIRE_VALID_DATE = True        # Bookings require a valid date anchor

    # -----------------------------
    # Helpers
    # -----------------------------
    @st.cache_data(show_spinner=False)
    def read_sheet_names(file):
        try:
            xl = pd.ExcelFile(file)
            return xl.sheet_names
        except Exception as e:
            st.error(f"Failed to read sheet names: {e}")
            return []

    @st.cache_data(show_spinner=False)
    def load_sheet(file, sheet_name):
        return pd.read_excel(file, sheet_name=sheet_name, header=None)

    def _try_parse_date(x):
        try:
            d = pd.to_datetime(x, errors='coerce')
            if pd.notna(d):
                return d
            return None
        except Exception:
            return None

    def find_week_headers(df: pd.DataFrame):
        """
        Robust week header finder.
        - Find 'Monday' (case-insensitive) in Column B.
        - Within next 10 rows, pick first row where B..H has >=2 parseable dates -> dates row.
        - monday_date is B's parsed date; if missing, infer from any parsed day in that row.
        Returns [(monday_row, date_row, monday_date)]
        """
        col = 1
        s = df.iloc[:, col].astype(str).str.strip().str.lower()
        day_rows = df.index[s.eq('monday')].tolist()

        headers = []
        for dr in day_rows:
            date_r, monday_date = None, None
            for look_ahead in range(1, 11):
                r = dr + look_ahead
                if r >= len(df):
                    break
                parsed_dates = {}
                for i in range(7):
                    c = 1 + i
                    if c >= df.shape[1]:
                        continue
                    parsed = _try_parse_date(df.iat[r, c])
                    if parsed is not None:
                        parsed_dates[i] = parsed
                if len(parsed_dates) >= 2:
                    date_r = r
                    if 0 in parsed_dates:
                        monday_date = parsed_dates[0].date()
                    else:
                        i0 = sorted(parsed_dates.keys())[0]
                        monday_date = (parsed_dates[i0] - pd.Timedelta(days=i0)).date()
                    break
            headers.append((dr, date_r, monday_date))
        return headers

    def row_to_week_monday(row_idx: int, headers):
        prev = [h for h in headers if h[0] <= row_idx]
        if not prev:
            return None
        prev.sort(key=lambda x: x[0])
        return prev[-1][2]

    def detect_am_pm_runs(df: pd.DataFrame, start_row: int = 0):
        """Scan Col A for AM/PM rows and group consecutive runs."""
        runs, current, run_start, prev_idx = [], None, None, None
        for idx in range(start_row, len(df)):
            label = None
            raw = df.iat[idx, 0]
            if isinstance(raw, str):
                ru = raw.strip().upper()
                if ru.startswith('AM'):
                    label = 'AM'
                elif ru.startswith('PM'):
                    label = 'PM'
            if TRUST_ONLY_AM_PM and label is None:
                continue
            if label is None:
                continue
            if current is None:
                current, run_start = label, idx
            elif label != current or (prev_idx is not None and idx != prev_idx + 1):
                runs.append((current, run_start, prev_idx))
                current, run_start = label, idx
            prev_idx = idx
        if current is not None:
            runs.append((current, run_start, prev_idx))
        return runs

    def build_maps_and_roster(df: pd.DataFrame):
        """
        Returns:
          mapping_by_week: {(monday_date, period, day, preceptor) -> {'student','cell','date'}}
                           (only when a real student exists)
          index_by_date:   {(date, period, preceptor) -> {'student','cell','day'}}  # for date-based matching
          roster_week:     {(monday_date, period) -> set(preceptors)}               # week-level roster
          day_roster:      {(monday_date, period, day) -> set(preceptors)}          # day-level roster
          week_dates:      {monday_date -> {day -> date}}
          occupied:        set((monday_date, period, day, preceptor)) even w/o student
          diag_weeks:      diagnostics list
        """
        import re, unicodedata

        def is_placeholder_preceptor(text: str) -> bool:
            if not text: return True
            t = text.strip().upper()
            EXCLUDE_PREFIXES = [
                'CLOSED','CLOSE','BLOCK','VACATION','ADMIN','MEETING','NO CLINIC',
                'CLINIC CANCELLED','CANCELLED','HOLIDAY','OFF','PTO','SICK',
                'NOTE','NOTES','REFERENCE','INFO','FYI','ORIENTATION'
            ]
            return any(t.startswith(pfx) for pfx in EXCLUDE_PREFIXES)

        def _norm(s: str) -> str:
            s = unicodedata.normalize("NFKC", s)
            s = s.replace("\u00A0", " ")                    # NBSP -> space
            s = s.replace("\u2013", "-").replace("\u2014", "-")  # en/em dash -> '-'
            s = re.sub(r"\s+", " ", s.strip())
            return s

        def parse_cell(val: str):
            """Robust split on '~'. RHS counts as student only if alphanumeric & not a placeholder."""
            if not isinstance(val, str): return None, None
            raw = _norm(val)
            if "~" not in raw:
                pre = _norm(raw)
                return (pre if pre else None), None
            pre, rhs = re.split(r"\s*~\s*", raw, maxsplit=1)
            pre = _norm(pre)
            rhs = _norm(rhs)

            BLANK_TOKENS = {"", "nan", "n/a", "na", "-", "--", "—", "none", "null"}
            PLACEHOLDER_HINTS = {"note", "notes", "ref", "reference", "info", "fyi"}

            rhs_l = rhs.lower()
            # treat as empty unless it has at least one letter/number and is not a placeholder
            if (rhs_l in BLANK_TOKENS) or (not re.search(r"[a-z0-9]", rhs_l)) or any(h in rhs_l for h in PLACEHOLDER_HINTS):
                rhs = None
            return (pre if pre else None), rhs

        headers = find_week_headers(df)
        runs = detect_am_pm_runs(df, start_row=0)

        mapping_by_week = {}
        index_by_date = {}
        roster_week = defaultdict(set)
        day_roster = defaultdict(set)
        week_dates = defaultdict(dict)
        occupied = set()
        diag_weeks = []

        # Build date anchors with fallback
        for (day_row, date_row, monday_date) in headers:
            if monday_date is None or date_row is None:
                continue
            inferred_days = []
            for i, day in enumerate(DAYS):
                col_idx = 1 + i  # B..H
                val = df.iat[date_row, col_idx] if col_idx < df.shape[1] else None
                parsed = _try_parse_date(val)
                if parsed is not None:
                    week_dates[monday_date][day] = parsed.date()
                else:
                    # fallback Monday + i days
                    try:
                        fallback = (pd.to_datetime(monday_date) + pd.Timedelta(days=i)).date()
                        week_dates[monday_date][day] = fallback
                        inferred_days.append(day)
                    except Exception:
                        week_dates[monday_date][day] = None
                        inferred_days.append(day)
            diag_weeks.append({
                'monday_date': monday_date,
                'date_row': date_row,
                'inferred_days': inferred_days
            })

        # Parse inside AM/PM runs
        for period, rstart, rend in runs:
            monday_date = row_to_week_monday(rstart, headers)
            if monday_date is None:
                continue
            for col_idx, day in enumerate(DAYS, start=1):
                date_anchor = week_dates[monday_date].get(day)
                for row in range(rstart, rend+1):
                    if col_idx >= df.shape[1]:
                        continue
                    val = df.iat[row, col_idx]
                    if pd.isna(val) or not isinstance(val, str):
                        continue
                    pre, stu = parse_cell(val)
                    if not pre or is_placeholder_preceptor(pre):
                        continue

                    # Present in week & specific day
                    roster_week[(monday_date, period)].add(pre)
                    day_roster[(monday_date, period, day)].add(pre)
                    occupied.add((monday_date, period, day, pre))  # presence marker

                    # Keep booking only if real student + valid date
                    if stu is None:
                        continue
                    if REQUIRE_VALID_DATE and _try_parse_date(date_anchor) is None:
                        continue

                    cell = f"{chr(ord('A')+col_idx)}{row+1}"
                    wk_key = (monday_date, period, day, pre)
                    mapping_by_week.setdefault(wk_key, {'student': stu, 'cell': cell, 'date': date_anchor})
                    # date-based index for cross-file matching
                    dt_key = (pd.to_datetime(date_anchor).date(), period, pre)
                    # prefer first seen student for stability
                    index_by_date.setdefault(dt_key, {'student': stu, 'cell': cell, 'day': day})

        return mapping_by_week, index_by_date, roster_week, day_roster, week_dates, occupied, diag_weeks

    # -----------------------------
    # UI - File uploads
    # -----------------------------
    col1, col2 = st.columns(2)
    with col1:
        md_file = st.file_uploader("Upload MD OPD (xlsx)", type=["xlsx"], key="md")
    with col2:
        pa_file = st.file_uploader("Upload PA OPD (xlsx)", type=["xlsx"], key="pa")

    if md_file and pa_file:
        md_sheets = read_sheet_names(md_file)
        pa_sheets = read_sheet_names(pa_file)

        common_sheets = sorted([s for s in DEFAULT_FOCUS if s in md_sheets and s in pa_sheets])
        selected_sheets = st.multiselect(
            "Sites (tabs) to compare",
            options=sorted(list(set(md_sheets) & set(pa_sheets))),
            default=common_sheets or sorted(list(set(md_sheets) & set(pa_sheets)))
        )

        if not selected_sheets:
            st.warning("No common sheets selected.")
            st.stop()

        # Keep per-site context so we can search "other sites"
        site_ctx = {}

        conflict_rows = []
        diagnostics = []

        for sheet in selected_sheets:
            df_md = load_sheet(md_file, sheet)
            df_pa = load_sheet(pa_file, sheet)

            (md_map_wk, md_idx_date, md_roster_wk, md_day_roster, md_week_dates,
             md_occupied, md_diag) = build_maps_and_roster(df_md)
            (pa_map_wk, pa_idx_date, pa_roster_wk, pa_day_roster, pa_week_dates,
             pa_occupied, pa_diag) = build_maps_and_roster(df_pa)
            diagnostics.append({'site': sheet, 'md': md_diag, 'pa': pa_diag})

            # Save for cross-site availability
            site_ctx[sheet] = dict(
                md_idx_date=md_idx_date,
                pa_idx_date=pa_idx_date,
                md_week_dates=md_week_dates,
                pa_week_dates=pa_week_dates,
                md_day_roster=md_day_roster,
                pa_day_roster=pa_day_roster
            )

            # --------- CONFLICTS by actual date ---------
            md_keys = set(md_idx_date.keys())
            pa_keys = set(pa_idx_date.keys())
            for (date_obj, period, pre) in sorted(md_keys & pa_keys):
                md_entry = md_idx_date[(date_obj, period, pre)]
                pa_entry = pa_idx_date[(date_obj, period, pre)]
                conflict_rows.append({
                    'site': sheet,
                    'date': date_obj,
                    'day': pd.to_datetime(date_obj).strftime('%A'),
                    'period': period,
                    'preceptor': pre,
                    'md_student': md_entry['student'],
                    'pa_student': pa_entry['student']
                })

        # Conflicts dataframe
        conflicts_df = pd.DataFrame(conflict_rows)

        # --------- helpers to build pools ---------
        def pool_for_site_day(site, day_name, period, date_obj):
            """Union of preceptors present in THIS site for (day, period, date)."""
            ctx = site_ctx[site]
            pool = set()
            md_week_dates, pa_week_dates = ctx['md_week_dates'], ctx['pa_week_dates']
            md_day_roster, pa_day_roster = ctx['md_day_roster'], ctx['pa_day_roster']
            # MD
            for m in md_week_dates.keys():
                if md_week_dates[m].get(day_name) == date_obj:
                    pool |= (md_day_roster.get((m, period, day_name)) or set())
            # PA
            for m in pa_week_dates.keys():
                if pa_week_dates[m].get(day_name) == date_obj:
                    pool |= (pa_day_roster.get((m, period, day_name)) or set())
            return pool

        def pool_for_other_sites(current_site, day_name, period, date_obj):
            """Union of preceptors present in ALL OTHER sites for (day, period, date)."""
            pool = set()
            for site in site_ctx.keys():
                if site == current_site:
                    continue
                pool |= pool_for_site_day(site, day_name, period, date_obj)
            return pool

        def count_assigned_any_site(pre, date_obj, period):
            """How many students (MD+PA) does preceptor have across all sites at this date/period?"""
            total = 0
            for s, ctx in site_ctx.items():
                if (date_obj, period, pre) in ctx['md_idx_date']:
                    total += 1
                if (date_obj, period, pre) in ctx['pa_idx_date']:
                    total += 1
            return total

        # --------- AVAILABILITY (same-site & other-sites) for conflict slots ---------
        availability_same_rows = []
        availability_other_rows = []
        suggestions_rows = []

        if not conflicts_df.empty:
            for _, r in conflicts_df.iterrows():
                site   = r['site']
                date_o = r['date']
                day_nm = r['day']         # 'Monday'...'Sunday'
                period = r['period']

                # Pools
                same_pool  = pool_for_site_day(site, day_nm, period, date_o)
                other_pool = pool_for_other_sites(site, day_nm, period, date_o)

                # Build availability function
                def add_pool(pool, dest_list, pool_site_label):
                    for pre in sorted(pool):
                        total_assigned = count_assigned_any_site(pre, date_o, period)
                        is_acute = ("ACUTE" in str(pre).upper())
                        capacity = 2 if is_acute else 1
                        seats_left = max(0, capacity - total_assigned)
                        if seats_left > 0:
                            dest_list.append({
                                'site_of_conflict': site,
                                'candidate_site': pool_site_label,
                                'date': date_o,
                                'day': day_nm,
                                'period': period,
                                'conflict_preceptor': r['preceptor'],
                                'preceptor': pre,
                                'is_acute': is_acute,
                                'current_students': total_assigned,
                                'capacity': capacity,
                                'seats_left': seats_left,
                                'status': 'available'
                            })

                add_pool(same_pool, availability_same_rows, site)
                # For other sites, keep which site each candidate belongs to.
                for other_site in site_ctx.keys():
                    if other_site == site:
                        continue
                    pool = pool_for_site_day(other_site, day_nm, period, date_o)
                    add_pool(pool, availability_other_rows, other_site)

        avail_same_df = pd.DataFrame(availability_same_rows)
        avail_other_df = pd.DataFrame(availability_other_rows)

        # --------- SUGGESTIONS (top-3) prefer same-site, then other-sites ---------
        if not conflicts_df.empty:
            for _, r in conflicts_df.iterrows():
                in_slot_same  = avail_same_df[
                    (avail_same_df['site_of_conflict'] == r['site']) &
                    (avail_same_df['date'] == r['date']) &
                    (avail_same_df['period'] == r['period'])
                ].copy()

                in_slot_other = avail_other_df[
                    (avail_other_df['site_of_conflict'] == r['site']) &
                    (avail_other_df['date'] == r['date']) &
                    (avail_other_df['period'] == r['period'])
                ].copy()

                # Put same preceptor first if eligible (Acute w/ 1 student)
                def order(df):
                    df['_self'] = (df['preceptor'] == r['preceptor'])
                    return df.sort_values(['_self','candidate_site','preceptor'], ascending=[False, True, True]).drop(columns=['_self'])

                ordered = pd.concat([order(in_slot_same), order(in_slot_other)], ignore_index=True)

                if not ordered.empty:
                    for _, a in ordered.head(3).iterrows():
                        label = a['preceptor']
                        if a['preceptor'] == r['preceptor']:
                            label = f"{a['preceptor']} (currently assigned)"
                        suggestions_rows.append({
                            'conflict_site': r['site'],
                            'date': r['date'],
                            'day': r['day'],
                            'period': r['period'],
                            'conflict_preceptor': r['preceptor'],
                            'md_student': r['md_student'],
                            'pa_student': r['pa_student'],
                            'suggested_preceptor': label,
                            'suggested_site': a['candidate_site'],
                            'suggested_is_acute': bool(a['is_acute']),
                            'suggested_current_students': int(a['current_students']),
                            'suggested_capacity': int(a['capacity']),
                            'suggested_seats_left': int(a['seats_left'])
                        })
                else:
                    suggestions_rows.append({
                        'conflict_site': r['site'],
                        'date': r['date'],
                        'day': r['day'],
                        'period': r['period'],
                        'conflict_preceptor': r['preceptor'],
                        'md_student': r['md_student'],
                        'pa_student': r['pa_student'],
                        'suggested_preceptor': '⚠️ No alternative preceptors available',
                        'suggested_site': None,
                        'suggested_is_acute': None,
                        'suggested_current_students': None,
                        'suggested_capacity': None,
                        'suggested_seats_left': None
                    })

        suggestions_df = pd.DataFrame(suggestions_rows)

        # -----------------------------
        # Results UI (conflict-focused)
        # -----------------------------
        st.subheader("Results (conflict-focused)")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.metric("Sites compared", len(selected_sheets))
        with c2:
            st.metric("Double bookings found", 0 if conflicts_df.empty else len(conflicts_df))
        with c3:
            st.metric("Avail. (same-site)", 0 if avail_same_df.empty else len(avail_same_df))
        with c4:
            st.metric("Avail. (other-sites)", 0 if avail_other_df.empty else len(avail_other_df))

        st.markdown("**Double-booked preceptors (MD & PA in same slot)**")
        if conflicts_df.empty:
            st.info("No double-bookings found for the selected sites.")
        else:
            st.dataframe(conflicts_df[['site','date','day','period','preceptor','md_student','pa_student']], use_container_width=True)
            st.download_button(
                label="Download double-bookings CSV",
                data=conflicts_df[['site','date','day','period','preceptor','md_student','pa_student']].to_csv(index=False).encode('utf-8'),
                file_name="opd_double_bookings.csv",
                mime="text/csv"
            )

        # Availability (same site)
        show_same  = st.toggle("Show available preceptors in the SAME site (Acutes can take 2)",value=False, key="tog_same_site")
        if show_same:
            if avail_same_df.empty:
                st.info("No same-site availability for the conflicted slots.")
            else:
                st.markdown("**Available preceptors (same site as conflict)** — Acutes shown if <2 students; others only if unbooked.")
                st.dataframe(
                    avail_same_df[['site_of_conflict','candidate_site','date','day','period',
                                   'conflict_preceptor','preceptor','is_acute',
                                   'current_students','capacity','seats_left','status']],
                    use_container_width=True
                )
                st.download_button(
                    label="Download same-site availability CSV",
                    data=avail_same_df.to_csv(index=False).encode('utf-8'),
                    file_name="opd_availability_same_site.csv",
                    mime="text/csv"
                )

        # Availability (other sites)
        show_other = st.toggle("Show available preceptors in OTHER sites (Acutes can take 2)",value=False, key="tog_other_site")
        if show_other:
            if avail_other_df.empty:
                st.info("No other-site availability for the conflicted slots.")
            else:
                st.markdown("**Available preceptors (other sites)** — same date & AM/PM, different site.")
                st.dataframe(
                    avail_other_df[['site_of_conflict','candidate_site','date','day','period',
                                    'conflict_preceptor','preceptor','is_acute',
                                    'current_students','capacity','seats_left','status']],
                    use_container_width=True
                )
                st.download_button(
                    label="Download other-site availability CSV",
                    data=avail_other_df.to_csv(index=False).encode('utf-8'),
                    file_name="opd_availability_other_sites.csv",
                    mime="text/csv"
                )

        # Suggestions
        show_sugg  = st.toggle("Show suggestions to resolve each conflict (prefers same site, then other sites)",
                       value=False, key="tog_suggestions")
        if show_sugg:
            if suggestions_df.empty:
                st.info("No suggestions available.")
            else:
                st.markdown("**Targeted suggestions** — same site first; if none, suggests from other sites on the same date & AM/PM.")
                st.dataframe(
                    suggestions_df[['conflict_site','date','day','period','conflict_preceptor',
                                    'md_student','pa_student','suggested_preceptor','suggested_site',
                                    'suggested_is_acute','suggested_current_students',
                                    'suggested_capacity','suggested_seats_left']],
                    use_container_width=True
                )
                st.download_button(
                    label="Download suggestions CSV",
                    data=suggestions_df.to_csv(index=False).encode('utf-8'),
                    file_name="opd_targeted_suggestions_cross_site.csv",
                    mime="text/csv"
                )

        # -----------------------------
        # Optional Annotated Downloads Toggle
        # -----------------------------
        show_annotated = st.toggle("Generate annotated OPD files (highlight conflicts in RED)",
                           value=False, key="tog_annotated_downloads")
        
        if show_annotated:
            st.markdown("---")
            st.subheader("Download annotated OPDs (conflicts highlighted in RED)")
            st.caption("Cells are red when the *other* OPD already has that preceptor booked for the same site, date, and AM/PM.")
        
            from io import BytesIO
            from openpyxl import load_workbook
            from openpyxl.styles import Font, Color, Border, Side, PatternFill
            from openpyxl.comments import Comment
            
            def _annot_make_copy(uploaded_file, other_idx_by_site: dict, selected_sheets: list) -> bytes:
                """
                Annotate: red font (if visible), THICK RED BORDER, and a small note so conflicts
                are obvious even when Conditional Formatting overrides font color.
                """
                raw = uploaded_file.getvalue()
                wb = load_workbook(BytesIO(raw))
            
                # --- helpers matching your main parser ---
                import re, unicodedata, pandas as pd
                def _norm(s: str) -> str:
                    s = unicodedata.normalize("NFKC", s).replace("\u00A0", " ")
                    s = s.replace("\u2013", "-").replace("\u2014", "-")
                    return re.sub(r"\s+", " ", s.strip())
            
                def _parse_cell(val: str):
                    if not isinstance(val, str): return None, None
                    raw = _norm(val)
                    if "~" not in raw:
                        pre = _norm(raw); return (pre if pre else None), None
                    pre, rhs = re.split(r"\s*~\s*", raw, maxsplit=1)
                    pre = _norm(pre); rhs = _norm(rhs)
                    if rhs.lower() in {"", "nan", "n/a", "na", "-", "--", "—", "none", "null"}:
                        rhs = None
                    return (pre if pre else None), rhs
            
                def _is_placeholder_preceptor(text: str) -> bool:
                    if not text: return True
                    t = text.strip().upper()
                    return any(t.startswith(pfx) for pfx in [
                        'CLOSED','CLOSE','BLOCK','VACATION','ADMIN','MEETING','NO CLINIC',
                        'CLINIC CANCELLED','CANCELLED','HOLIDAY','OFF','PTO','SICK',
                        'NOTE','NOTES','REFERENCE','INFO','FYI','ORIENTATION'
                    ])
            
                # opaque ARGB
                OPAQUE_RED = Color(rgb="FFFF0000")
                RED_SIDE   = Side(style="thick", color="FFFF0000")
                RED_BORDER = Border(left=RED_SIDE, right=RED_SIDE, top=RED_SIDE, bottom=RED_SIDE)
            
                for sheet in selected_sheets:
                    if sheet not in wb.sheetnames:
                        continue
                    ws = wb[sheet]
            
                    # rebuild date map from the uploaded file (aligns weeks/days to dates)
                    df = load_sheet(uploaded_file, sheet)
                    headers = find_week_headers(df)
                    runs = detect_am_pm_runs(df, start_row=0)
            
                    week_dates = {}
                    for (_day_row, date_row, monday_date) in headers:
                        if monday_date is None or date_row is None:
                            continue
                        week_dates.setdefault(monday_date, {})
                        for i, day in enumerate(DAYS):
                            c = 1 + i  # B..H
                            if c >= df.shape[1]: continue
                            parsed = pd.to_datetime(df.iat[date_row, c], errors='coerce')
                            if pd.notna(parsed):
                                week_dates[monday_date][day] = parsed.date()
            
                    other_idx = other_idx_by_site.get(sheet, {})  # keys: (date, period, preceptor)
            
                    for period, rstart, rend in runs:
                        monday_date = row_to_week_monday(rstart, headers)
                        if monday_date is None:
                            continue
                        for c_idx, day in enumerate(DAYS, start=1):  # B..H
                            dt = week_dates.get(monday_date, {}).get(day)
                            if dt is None:
                                continue
                            for row in range(rstart, rend+1):
                                if c_idx >= df.shape[1]: continue
                                val = df.iat[row, c_idx]
                                if pd.isna(val) or not isinstance(val, str): continue
                                pre, _stu = _parse_cell(val)
                                if not pre or _is_placeholder_preceptor(pre): continue
            
                                if (dt, period, pre) in other_idx:
                                    addr = f"{chr(ord('A')+c_idx)}{row+1}"
                                    cell = ws[addr]
            
                                    # Try to ensure red font (CF may still override)
                                    f = cell.font or Font()
                                    try:
                                        cell.font = f.copy(color="FFFF0000")
                                    except Exception:
                                        cell.font = Font(
                                            name=f.name, size=f.size or 11, bold=f.bold,
                                            italic=f.italic, underline=f.underline, color=OPAQUE_RED
                                        )
            
                                    # Add thick red border (highly visible even with CF)
                                    cell.border = RED_BORDER
            
                                    # Add a small note/comment (red triangle)
                                    if cell.comment is None:
                                        txt = f"Booked in other OPD\n{sheet} — {day} {dt} — {period}\nPreceptor: {pre}"
                                        try:
                                            cell.comment = Comment(txt, "MD↔PA conflict")
                                        except Exception:
                                            pass
            
                out = BytesIO()
                wb.save(out)
                out.seek(0)
                return out.getvalue()

        
            # Compare across files
            md_compare_against_pa = {s: site_ctx[s]['pa_idx_date'] for s in site_ctx}
            pa_compare_against_md = {s: site_ctx[s]['md_idx_date'] for s in site_ctx}
            
            col_md, col_pa = st.columns(2)
            with col_md:
                md_bytes = _annot_make_copy(md_file, md_compare_against_pa, selected_sheets)
                st.download_button("⬇️ MD annotated (RED = booked in PA)", md_bytes,
                                   "md_opd_annotated.xlsx",
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with col_pa:
                pa_bytes = _annot_make_copy(pa_file, pa_compare_against_md, selected_sheets)
                st.download_button("⬇️ PA annotated (RED = booked in MD)", pa_bytes,
                                   "pa_opd_annotated.xlsx",
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    else:
        st.info("Upload both the MD and PA OPD files to begin.")

# --------------------
# UI: title + upload
# --------------------
elif mode == "Shift Availability Tracker":
    # --------------------
    # Helpers: parsing
    # --------------------
    def is_date_header_row(series, min_dates=3):
        """A row is a 'date header' if it has >= min_dates parsable dates in columns 2+."""
        parsed = pd.to_datetime(series, errors="coerce")
        return parsed.notna().sum() >= min_dates
    
    def extract_names(cell: object) -> set[str]:
        """
        From a single Excel cell, return unique preceptor names.
        - Ignores any 'Closed' (case-insens., 'Closed?')
        - Requires '~' marker
        - Splits multiple entries on ';', '/', or 2+ spaces
        """
        if not isinstance(cell, str):
            return set()
        s = cell.replace("\r", " ").replace("\n", " ").strip()
        if re.search(r"\bclosed\b", s, flags=re.IGNORECASE):
            return set()
        if "~" not in s:
            return set()
        parts = re.split(r"[;/]| {2,}", s)
        return {p.replace("~", "").strip() for p in parts if p.replace("~", "").strip()}
    
    def build_segmented_name_map(excel: pd.ExcelFile) -> dict:
        """
        Return dict: (site, date, shift_label) -> set(names), computed SEGMENT-BY-SEGMENT.
        A 'segment' begins at each date-header row and ends before the next date-header row.
        """
        bucket = {}
        for sheet in excel.sheet_names:
            df = pd.read_excel(excel, sheet_name=sheet, header=None)
            header_rows = [r for r in range(len(df)) if is_date_header_row(df.iloc[r, 1:])]
            if not header_rows:
                continue
            header_rows.append(len(df))  # sentinel end
    
            valid = (
                {"AM - ACUTES", "AM - CONTINUITY", "PM - ACUTES", "PM - CONTINUITY"}
                if sheet == "HOPE_DRIVE" else {"AM", "PM"}
            )
    
            for h in range(len(header_rows) - 1):
                date_row, end_row = header_rows[h], header_rows[h+1]
                dates = pd.to_datetime(df.iloc[date_row, 1:], errors="coerce")
    
                seg_bucket = {}
                for i in range(date_row + 1, end_row):
                    label = str(df.iat[i, 0]).strip().upper()
                    if label in valid:
                        for j, d in enumerate(dates, start=1):
                            if pd.isna(d):
                                continue
                            names = extract_names(df.iat[i, j])
                            if not names:
                                continue
                            key = (sheet, pd.Timestamp(d).date(), label)
                            seg_bucket.setdefault(key, set()).update(names)
    
                for k, s in seg_bucket.items():
                    bucket.setdefault(k, set()).update(s)
        return bucket
    
    def fold_hope_drive_rows(sub_df: pd.DataFrame):
        """
        For HOPE_DRIVE on a given date, combine:
          - AM = (AM - ACUTES) ∪ (AM - CONTINUITY)
          - PM = (PM - ACUTES) ∪ (PM - CONTINUITY)
        Return list of dict rows with Shift in {'AM','PM'}, Names list, Count.
        """
        am, pm = set(), set()
        for _, r in sub_df.iterrows():
            if r["Shift"].startswith("AM"):
                am |= set(r["Names"])
            elif r["Shift"].startswith("PM"):
                pm |= set(r["Names"])
        out = []
        if am:
            out.append({"Shift": "AM", "Names": sorted(am), "Count": len(am)})
        if pm:
            out.append({"Shift": "PM", "Names": sorted(pm), "Count": len(pm)})
        return out
    
    # --------------------
    # UI: title + upload
    # --------------------
    st.title("Shift Availability Tracker")
    
    opd_file = st.file_uploader("Upload md_opd.xlsx", type=["xlsx"])
    if not opd_file:
        st.stop()
    
    excel = pd.ExcelFile(opd_file)
    
    # --------------------
    # Build name-level map + daily counts (segment-aware)
    # --------------------
    name_map = build_segmented_name_map(excel)
    
    rows = []
    for (site, dt, shift), names in name_map.items():
        rows.append({"Site": site, "Date": pd.to_datetime(dt), "Shift": shift, "Names": sorted(names)})
    
    raw = pd.DataFrame(rows)
    
    # Merge HOPE_DRIVE detailed labels into AM/PM for counts and names
    collapsed = []
    for (site, dt), sub in raw.groupby(["Site", "Date"]):
        if site == "HOPE_DRIVE":
            for r in fold_hope_drive_rows(sub):
                collapsed.append({"Site": site, "Date": dt, **r})
        else:
            for _, r in sub.iterrows():
                collapsed.append({
                    "Site": site,
                    "Date": r["Date"],
                    "Shift": r["Shift"],
                    "Names": r["Names"],
                    "Count": len(r["Names"]),
                })
    
    daily = pd.DataFrame(collapsed)
    if daily.empty:
        st.warning("No preceptors with '~' found. Check file/layout.")
        st.stop()
    
    daily["Weekday"] = daily["Date"].dt.weekday              # Mon=0..Sun=6
    daily["DayName"] = daily["Date"].dt.day_name()
    daily["WeekStart"] = daily["Date"] - pd.to_timedelta(daily["Date"].dt.weekday, unit="D")
    
    # --------------------
    # Weekly Grid (single table per site)
    # --------------------
    st.subheader("Weekly Grid (single table per site)")
    
    site_list = sorted(daily["Site"].unique().tolist())
    site_sel = st.selectbox("Site", site_list, index=(site_list.index("HOPE_DRIVE") if "HOPE_DRIVE" in site_list else 0))
    
    day_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    def shift_order_for(site):
        return ["AM - ACUTES", "AM - CONTINUITY", "PM - ACUTES", "PM - CONTINUITY"] if site == "HOPE_DRIVE" else ["AM", "PM"]
    
    if site_sel == "HOPE_DRIVE":
        raw_site = raw[raw["Site"] == "HOPE_DRIVE"].copy()
        raw_site["WeekStart"] = raw_site["Date"] - pd.to_timedelta(raw_site["Date"].dt.weekday, unit="D")
        raw_site["DayName"] = raw_site["Date"].dt.day_name()
        raw_site["DayCat"] = pd.Categorical(raw_site["DayName"], categories=day_order, ordered=True)
        raw_site["ShiftCat"] = pd.Categorical(raw_site["Shift"], categories=shift_order_for("HOPE_DRIVE"), ordered=True)
    
        blocks = []
        for wk, wkdf in raw_site.sort_values(["WeekStart", "ShiftCat", "DayCat"]).groupby("WeekStart"):
            grid = (
                wkdf.assign(Count=wkdf["Names"].apply(len))
                    .pivot_table(index="ShiftCat", columns="DayCat", values="Count", aggfunc="max")
                    .reindex(index=shift_order_for("HOPE_DRIVE"), columns=day_order)
                    .fillna(0).astype(int)
            )
            grid.index.name = "Shift"
            grid.insert(0, "Week of", f"Week of {wk:%Y-%m-%d}")
            blocks.append(grid.reset_index())
        weekly_single_table = pd.concat(blocks, axis=0, ignore_index=True) if blocks else pd.DataFrame()
    else:
        site_df = daily[daily["Site"] == site_sel].copy()
        site_df["DayCat"] = pd.Categorical(site_df["DayName"], categories=day_order, ordered=True)
        site_df["ShiftCat"] = pd.Categorical(site_df["Shift"], categories=shift_order_for(site_sel), ordered=True)
    
        blocks = []
        for wk, wkdf in site_df.sort_values(["WeekStart", "ShiftCat", "DayCat"]).groupby("WeekStart"):
            grid = (
                wkdf.pivot_table(index="ShiftCat", columns="DayCat", values="Count", aggfunc="max")
                    .reindex(index=shift_order_for(site_sel), columns=day_order)
                    .fillna(0).astype(int)
            )
            grid.index.name = "Shift"
            grid.insert(0, "Week of", f"Week of {wk:%Y-%m-%d}")
            blocks.append(grid.reset_index())
        weekly_single_table = pd.concat(blocks, axis=0, ignore_index=True) if blocks else pd.DataFrame()
    
    if weekly_single_table.empty:
        st.info("No data for this site.")
    else:
        st.dataframe(weekly_single_table, use_container_width=True)
        c1, c2 = st.columns(2)
        csv_bytes = weekly_single_table.to_csv(index=False).encode("utf-8")
        c1.download_button("Download Weekly Grid (CSV)", data=csv_bytes,
                           file_name=f"{site_sel}_weekly_grid.csv", mime="text/csv")
        xbuf = io.BytesIO()
        with pd.ExcelWriter(xbuf, engine="xlsxwriter") as writer:
            weekly_single_table.to_excel(writer, sheet_name="Weekly Grid", index=False)
        c2.download_button("Download Weekly Grid (Excel)", data=xbuf.getvalue(),
                           file_name=f"{site_sel}_weekly_grid.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    # --------------------
    # Weekly capacity across ETOWN + HOPE_DRIVE + NYES
    # AM = all 5 days (min across Mon–Fri)
    # PM = drop EXACTLY one day (second-smallest across Mon–Fri)
    # --------------------
    st.subheader("Weekly Max Students (ETOWN + HOPE_DRIVE + NYES)")
    
    def daily_caps_three_sites(day_df: pd.DataFrame) -> pd.DataFrame:
        df3 = day_df[day_df["Site"].isin({"ETOWN", "HOPE_DRIVE", "NYES"})].copy()
        df3["Weekday"] = df3["Date"].dt.weekday
        df3 = df3[df3["Weekday"] <= 4]  # Mon-Fri
        def ampm_bucketize(group):
            am_cap = group.loc[group["Shift"].str.startswith("AM"), "Count"].sum()
            pm_cap = group.loc[group["Shift"].str.startswith("PM"), "Count"].sum()
            return pd.Series({"AM_Capacity": am_cap, "PM_Capacity": pm_cap})
        dc = df3.groupby("Date").apply(ampm_bucketize).reset_index()
        dc["WeekStart"] = dc["Date"] - pd.to_timedelta(dc["Date"].dt.weekday, unit="D")
        return dc
    
    daily_caps = daily_caps_three_sites(daily)
    
    def weekly_student_capacity(g):
        am_vals = sorted(g["AM_Capacity"].tolist())      # Mon..Fri
        pm_vals = sorted(g["PM_Capacity"].tolist())      # Mon..Fri
        S_am = am_vals[0] if am_vals else 0              # AM: all 5 days → min
        S_pm = pm_vals[1] if len(pm_vals) >= 2 else (pm_vals[0] if pm_vals else 0)  # PM: drop 1 → second-smallest
        return pd.Series({"AM_students_max": S_am, "PM_students_max": S_pm, "Total_students_max": S_am + S_pm})
    
    weekly_capacity = daily_caps.groupby("WeekStart").apply(weekly_student_capacity).reset_index()
    st.dataframe(weekly_capacity, use_container_width=True)
    


elif mode == "OPD Archive":
    render_opd_archive_page()


elif mode == "Preceptor Teaching Summary":
    render_preceptor_teaching_summary()
