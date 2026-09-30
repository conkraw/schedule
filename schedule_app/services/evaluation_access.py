"""Shared section-only password gate for OER and PTS.

No browser cookie, query parameter, or process-global flag grants access.
Both entrypoints check the gate before rendering data or doing network work.
Date-setting callbacks also check it, because callbacks run before render().
Other scheduling sections and their session data are not protected by this gate.
"""
from __future__ import annotations
from functools import wraps
import hashlib
import hmac
import math
import time
import streamlit as st

P = "_evaluation_access_"
MIN_PASSWORD_LENGTH = 16
IDLE_SECONDS = 30 * 60
MAX_SESSION_SECONDS = 8 * 60 * 60
# New boundary: do not inherit a login issued by the older OER-only gate.
ACCESS_VERSION = "oer-pts-v2"
PROTECTED_STATE_PREFIXES = (
    "oasis_", "_oasis_", "oer_", "evaluation_privacy_",
    "teaching_", "assessment_completion_", "pts_", "_pts_",
)


def _configured_password():
    try:
        config = dict(st.secrets.get("evaluation_access", {}))
        password = config.get("password", "")
        archive = dict(st.secrets.get("opd_archive", {}))
    except Exception:
        return ""
    if (not isinstance(password, str) or len(password) < MIN_PASSWORD_LENGTH
            or len(password) > 1024
            or password.startswith(("REPLACE_", "YOUR_", "GENERATE_"))
            or password in (archive.get("encryption_key"), archive.get("github_token"))):
        return ""
    return password


def _fingerprint(password):
    return hashlib.sha256((ACCESS_VERSION + "\0" + password).encode("utf-8")).hexdigest()


def _clear_evaluation_data():
    """Discard OER/PTS session data, including decrypted content and report bytes.

    This only removes in-session state. It never deletes files or saved mappings
    in GitHub, and never clears unrelated OPD or individual-schedule state.
    """
    for key in list(st.session_state):
        if str(key).startswith(PROTECTED_STATE_PREFIXES):
            st.session_state.pop(key, None)


def lock_evaluation_records():
    """Lock both protected sections; retain this function name for compatibility."""
    _clear_evaluation_data()
    for suffix in ("auth", "issued", "last_seen", "password"):
        st.session_state.pop(P + suffix, None)


def _login_callback():
    # Callbacks execute before widgets are recreated. Remove the supplied value
    # here so the plaintext password is not retained in widget/session state.
    supplied = st.session_state.pop(P + "password", "")
    expected = _configured_password()
    now = time.time()
    if now < st.session_state.get(P + "retry_at", 0):
        return
    valid = (isinstance(supplied, str) and len(supplied) <= 1024 and bool(expected)
             and hmac.compare_digest(supplied.encode("utf-8"), expected.encode("utf-8")))
    if valid:
        st.session_state[P + "auth"] = _fingerprint(expected)
        st.session_state[P + "issued"] = now
        st.session_state[P + "last_seen"] = now
        for suffix in ("failed", "failures", "retry_at"):
            st.session_state.pop(P + suffix, None)
    else:
        lock_evaluation_records()
        failures = min(int(st.session_state.get(P + "failures", 0)) + 1, 10)
        st.session_state[P + "failures"] = failures
        st.session_state[P + "failed"] = True
        # Per-session delay, not a cross-session/brute-force prevention service.
        st.session_state[P + "retry_at"] = now + (5 if failures < 5 else 60)


def evaluation_access_is_valid(*, touch=False):
    password = _configured_password()
    if not password:
        return False
    now = time.time()
    try:
        issued = float(st.session_state.get(P + "issued", 0))
        seen = float(st.session_state.get(P + "last_seen", 0))
        valid = (st.session_state.get(P + "auth") == _fingerprint(password)
                 and math.isfinite(issued) and math.isfinite(seen)
                 and 0 <= now - seen < IDLE_SECONDS
                 and 0 <= now - issued < MAX_SESSION_SECONDS)
    except (TypeError, ValueError, OverflowError):
        return False
    if valid and touch:
        st.session_state[P + "last_seen"] = now
    return bool(valid)


def protected_evaluation_callback(callback):
    """Prevent a stale/locked date-widget callback from reading or writing GitHub.

    A callback can fire before the next section render checks the password.
    A failed check clears protected caches and lets render show the login form.
    """
    @wraps(callback)
    def guarded(*args, **kwargs):
        if not evaluation_access_is_valid(touch=True):
            lock_evaluation_records()
            return None
        return callback(*args, **kwargs)
    return guarded


def require_evaluation_access(*, show_lock=True, lock_key="evaluation_lock_main", section_name="OER"):
    """Return False before protected widgets/data/network work if not authorized.

    The same secret and session authorization cover OER and PTS; no second
    password or separate encryption key is introduced.
    """
    if section_name not in ("OER", "PTS"):
        raise ValueError("Unknown protected section.")
    if evaluation_access_is_valid(touch=True):
        if show_lock:
            st.sidebar.button("Lock OER / PTS", key=lock_key, on_click=lock_evaluation_records)
        return True
    lock_evaluation_records()
    st.subheader(f"{section_name} — locked")
    if not _configured_password():
        st.error("Setup required: add [evaluation_access] password in Streamlit Secrets. "
                 "Use a unique password or passphrase of at least 16 characters, "
                 "different from the encryption key and GitHub token. Other app sections remain available.")
        return False
    if st.session_state.get(P + "failed"):
        st.error("The password was not accepted.")
    remaining = max(0, int(math.ceil(st.session_state.get(P + "retry_at", 0) - time.time())))
    if remaining:
        st.warning(f"Wait {remaining} seconds before trying again.")
    st.caption("This password unlocks OER and PTS in this session. Access expires after 30 minutes "
               "without an interaction in either protected section, or after eight hours. "
               "Other scheduling sections remain available without this password.")
    with st.form(P + "login_form", clear_on_submit=True):
        st.text_input("OER / PTS password", type="password", key=P + "password", max_chars=1024)
        st.form_submit_button(f"Unlock {section_name}", on_click=_login_callback,
                              disabled=bool(remaining))
    return False
