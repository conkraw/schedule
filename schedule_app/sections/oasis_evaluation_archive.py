"""Main-menu OASIS upload, encrypted save, and original-CSV recovery."""
import hashlib

import streamlit as st

from schedule_app.services.opd_archive import (
    GitHubOPDArchive, OPDArchiveError, get_opd_archive_config,
)
from schedule_app.services.oasis_evaluations import (
    GitHubOASISEvaluations, OASIS_ARCHIVE_VERSION, OASISArchiveError, oasis_export_label,
)

SCOPE = "_oasis_archive_scope"
SAVE = "_oasis_archive_save"
LIST = "_oasis_archive_list"
LOADED = "_oasis_archive_loaded"
CHOICE = "oasis_saved_export"
UPLOAD = "oasis_evaluation_upload"


def _upload_changed():
    st.session_state.pop(SAVE, None)
    st.session_state.pop(LOADED, None)


def _refresh_exports():
    for key in (LIST, LOADED):
        st.session_state.pop(key, None)


def _selection_changed():
    st.session_state.pop(LOADED, None)


def _show_details(details):
    a, b, c = st.columns(3)
    a.metric("Response rows", f"{details['row_count']:,}")
    b.metric("Columns preserved", details["column_count"])
    c.metric("Original size", f"{details['byte_count'] / (1024 * 1024):.2f} MiB")
    if details["date_range_complete"]:
        st.caption(f"Course-date coverage: {details['course_start']} to {details['course_end']}. "
                   "These are the Start Date / End Date fields, not the submission-date range. "
                   "Response rows are question-level rows, not a count of evaluations.")
    else:
        st.warning("Some course dates are blank, unreadable, or reversed. The original CSV is preserved "
                   "under an undated identifier; no course dates or responses were changed.")
    st.caption("All original columns and values are retained. Student names, evaluation answers and "
               "comments are not displayed on this page.")


def render():
    st.subheader("OASIS Evaluation Archive")
    st.write("Upload the original OASIS evaluation CSV. The app encrypts it, saves it next to your "
             "OPD archive in GitHub, then reloads and decrypts it to verify an exact match.")
    st.info("Each different export is kept as a separate saved snapshot. Re-uploading an identical "
            "CSV does not add another copy, even if its filename changes. Exports are not merged, "
            "and OASIS data does not change OPD teaching reports.")
    try:
        config = get_opd_archive_config()
        client = GitHubOASISEvaluations(GitHubOPDArchive(config))
    except OPDArchiveError as exc:
        st.error(str(exc))
        st.caption("This section uses the existing [opd_archive] Streamlit Secrets; no new key is needed.")
        return
    signature = (OASIS_ARCHIVE_VERSION, config.signature())
    if st.session_state.get(SCOPE) != signature:
        for key in (SAVE, LIST, LOADED, CHOICE):
            st.session_state.pop(key, None)
        st.session_state[SCOPE] = signature
    st.caption(f"Storage: {config.owner}/{config.repo} | branch {config.branch} | {client.folder}/")
    st.caption("This page has no app-password gate. Anyone who can access the running app can use "
               "its archive upload and original-file download functions. Follow your institution's "
               "approved access and storage requirements for evaluations.")

    uploaded = st.file_uploader(
        "Upload OASIS evaluation export (.csv) — automatically encrypted and saved",
        type=["csv"], key=UPLOAD, on_change=_upload_changed,
        help="Up to 10 MiB. Use the original CSV export; do not first open and resave it in Excel.",
    )
    if uploaded is not None:
        raw = uploaded.getvalue()
        upload_signature = (signature, getattr(uploaded, "file_id", None),
                            hashlib.sha256(raw).hexdigest())
        state = st.session_state.get(SAVE)
        if not isinstance(state, dict) or state.get("signature") != upload_signature:
            state = None
        if state is None:
            with st.spinner("Encrypting, saving and verifying the OASIS export..."):
                try:
                    receipt = client.save(raw)
                    state = {"signature": upload_signature, "receipt": receipt}
                    st.session_state.pop(LIST, None)
                    st.session_state.pop(LOADED, None)
                    # The dropdown has not been rendered yet in this run.
                    st.session_state[CHOICE] = receipt["filename"]
                except OPDArchiveError as exc:
                    state = {"signature": upload_signature, "error": str(exc)}
                except Exception:
                    # Never echo CSV values, raw server responses, or credentials.
                    state = {"signature": upload_signature, "error":
                             "The OASIS upload could not be completed or verified. Check the CSV "
                             "and archive settings, then retry. No successful save is confirmed."}
            st.session_state[SAVE] = state
        if state.get("error"):
            st.error("OASIS save is NOT confirmed. " + state["error"])
            st.button("Retry OASIS encrypted save", key="oasis_retry_save", on_click=_upload_changed)
        else:
            receipt = state["receipt"]
            if receipt["action"] == "created":
                st.success("OASIS export archived and verified. The original CSV can be reloaded below.")
            else:
                st.success("This identical OASIS export is already archived and verified. No additional copy was saved.")
            _show_details(receipt["details"])
            st.caption("Verified saved file: " + receipt["path"])
    else:
        # Do not let a removed uploader leave a stale success receipt visible.
        st.session_state.pop(SAVE, None)

    st.markdown("### Reload a saved OASIS export")
    st.button("Refresh saved OASIS exports", key="oasis_refresh_exports", on_click=_refresh_exports)
    if LIST not in st.session_state:
        try:
            st.session_state[LIST] = client.list_exports()
        except OPDArchiveError as exc:
            st.error(str(exc))
            return
        except Exception:
            st.error("The OASIS export list could not be loaded. Refresh the list to retry.")
            return
    snapshot = st.session_state[LIST]
    filenames = snapshot["filenames"]
    if not filenames:
        st.info("No saved OASIS exports yet. Upload your CSV above to create the first encrypted copy.")
        return
    if st.session_state.get(CHOICE) not in filenames:
        st.session_state.pop(CHOICE, None)
    selected = st.selectbox("Saved OASIS exports", filenames, format_func=oasis_export_label,
                            key=CHOICE, on_change=_selection_changed,
                            help="Sorted by course-date coverage, not upload date. The export ID "
                                 "distinguishes files covering the same dates.")
    loaded = st.session_state.get(LOADED)
    if loaded and loaded.get("filename") != selected:
        st.session_state.pop(LOADED, None)
        loaded = None
    if st.button("Load / decrypt selected OASIS export", key="oasis_load_export"):
        st.session_state.pop(LOADED, None)
        try:
            with st.spinner("Retrieving and decrypting the original CSV..."):
                # Fetch current HEAD, so a removed/altered file is not silently
                # downloaded from a stale list's historical commit.
                loaded = client.load(selected)
                st.session_state[LOADED] = loaded
        except OPDArchiveError as exc:
            loaded = None
            st.error(str(exc))
        except Exception:
            loaded = None
            st.error("This OASIS export could not be loaded or verified. No file is available to download.")
    loaded = st.session_state.get(LOADED)
    if loaded and loaded["filename"] == selected:
        st.success("Original OASIS CSV loaded and verified. Reloading does not write to GitHub.")
        _show_details(loaded["details"])
        # Preserve exact bytes; use a neutral download filename, not patient/
        # learner/provider information in the uploaded local filename.
        download_name = selected.removesuffix(".enc")
        st.download_button("Download original OASIS CSV", data=loaded["raw"],
                           file_name=download_name, mime="text/csv", key="oasis_download_original")
        st.caption("The downloaded CSV is unencrypted. It has the original contents and a neutral archive filename. "
                   "Do not commit it to a public repository.")
