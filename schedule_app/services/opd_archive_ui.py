"""Shared Streamlit archive upload, verification and reload widgets.

Extracted from the supplied app; this module performs no page rendering on import.
"""

from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.opd_archive import OPD_XLSX_MIME
from schedule_app.services.opd_archive import inspect_opd_rotation
import hashlib
import streamlit as st


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
