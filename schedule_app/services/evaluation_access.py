"""A section-only shared-password gate for evaluation administration.

This does NOT gate Preceptor Teaching Summary or its linked report contents.
No browser cookie, query parameter, or process-global flag grants access.
Password/config changes invalidate authorization on the next section interaction.
"""
from __future__ import annotations
import hashlib
import hmac
import math
import time
import streamlit as st

P = "_evaluation_access_"
MIN_PASSWORD_LENGTH = 16
IDLE_SECONDS = 30 * 60
MAX_SESSION_SECONDS = 8 * 60 * 60


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
    return hashlib.sha256(("evaluation-section-v1\0" + password).encode("utf-8")).hexdigest()


def _clear_evaluation_data():
    # These are the evaluation ADMIN screens' state keys, including cached
    # downloads and upload widgets. OPD/teaching report state is independent.
    for key in list(st.session_state):
        if str(key).startswith(("oasis_", "_oasis_", "oer_", "evaluation_privacy_")):
            st.session_state.pop(key, None)


def lock_evaluation_records():
    """Callback: clear admin caches and authorization, not the rest of the app."""
    _clear_evaluation_data()
    for suffix in ("auth", "issued", "last_seen", "password"):
        st.session_state.pop(P + suffix, None)


def _login_callback():
    # Streamlit callbacks run before widgets on the next script execution, so
    # removing the submitted widget value here avoids retaining the password.
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


def require_evaluation_access(*, show_lock=True, lock_key="evaluation_lock_main"):
    """Return False before ANY evaluation widgets/data/network work if locked."""
    if evaluation_access_is_valid(touch=True):
        if show_lock:
            st.sidebar.button("Lock Evaluation Records", key=lock_key, on_click=lock_evaluation_records)
        return True
    lock_evaluation_records()
    st.subheader("Evaluation Records — locked")
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
    st.caption("Only Evaluation Records requires this password. Access expires after 30 minutes "
               "without an interaction in this section, or after eight hours. Other sections are not locked.")
    with st.form(P + "login_form", clear_on_submit=True):
        st.text_input("Evaluation Records password", type="password", key=P + "password",
                      max_chars=1024)
        st.form_submit_button("Unlock Evaluation Records", on_click=_login_callback,
                              disabled=bool(remaining))
    return False
