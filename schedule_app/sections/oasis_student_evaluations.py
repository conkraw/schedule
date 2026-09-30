"""Student-evaluation upload/recovery inside OASIS Evaluations; no report mixing."""
from __future__ import annotations

import hashlib
import streamlit as st

from schedule_app.services.opd_archive import (
    GitHubOPDArchive, OPDArchiveError, get_opd_archive_config,
)
from schedule_app.services.oasis_student_evaluations import (
    GitHubOASISStudentEvaluations, STUDENT_ARCHIVE_VERSION, student_export_label,
)

P = "oasis_student_archive_"


def _clear_loaded():
    st.session_state.pop(P + "loaded", None)


def _upload_changed():
    st.session_state.pop(P + "save", None)
    _clear_loaded()


def _refresh():
    st.session_state.pop(P + "list", None)
    _clear_loaded()


def _scope(config):
    signature = (STUDENT_ARCHIVE_VERSION, config.signature())
    if st.session_state.get(P + "scope") != signature:
        # Never clear educator sources, outputs, usernames, dates, or teaching data.
        # The uploader has not been instantiated yet in this run.
        for key in list(st.session_state):
            if key.startswith(P):
                st.session_state.pop(key, None)
        st.session_state[P + "scope"] = signature
    return signature


def _show_details(details):
    a, b, c = st.columns(3)
    a.metric("Forms in this file", f"{details['form_count']:,}")
    b.metric("Question-response rows", f"{details['row_count']:,}")
    c.metric("Original size", f"{details['byte_count'] / (1024 * 1024):.2f} MiB")
    st.caption("Includes: " + "; ".join(details["form_types"]) + ".")
    if details["date_range_complete"]:
        st.caption(f"Course-date coverage: {details['course_start']} to {details['course_end']}. "
                   "These dates organize the original file; no Submit Date or academic-year filter is applied.")
    else:
        st.warning("Some course dates are blank, unreadable, or reversed. The unchanged original is "
                   "stored with an undated identifier. No source dates were guessed or corrected.")
    if details["missing_form_record_rows"]:
        st.warning("Some rows have no Form Record identifier. They are preserved in the original, "
                   "but cannot contribute to the file's distinct-form count.")
    if details["blank_submit_date_rows"]:
        st.caption("Rows without a Submit Date are retained as well. Forms in this file is a source-file "
                   "count, not a claim that every form is a completed evaluation.")
    st.caption("All original columns, identifiers, grades, comments, and line breaks are preserved. "
               "Individual student names, answers, and comments are not displayed on this page.")


def _save_upload(uploaded, client, scope):
    if uploaded is None:
        st.session_state.pop(P + "save", None)
        return
    raw = uploaded.getvalue()
    signature = (scope, getattr(uploaded, "file_id", None), hashlib.sha256(raw).hexdigest())
    state = st.session_state.get(P + "save")
    if not isinstance(state, dict) or state.get("signature") != signature:
        _clear_loaded()
        state = None
    if state is None:
        try:
            with st.spinner("Encrypting, saving, and verifying the original student-evaluation CSV..."):
                receipt = client.save(raw)
            state = {"signature": signature, "receipt": receipt}
            _refresh()
            # Set before the dropdown is instantiated; no stale student file remains.
            st.session_state[P + "choice"] = receipt["filename"]
        except OPDArchiveError as exc:
            state = {"signature": signature, "error": str(exc)}
        except Exception:
            state = {"signature": signature, "error":
                     "The student-evaluation CSV could not be saved or verified. Check the export "
                     "and archive connection, then retry. No successful save is confirmed."}
        st.session_state[P + "save"] = state
    if state.get("error"):
        st.error("Student-evaluation save is NOT confirmed. " + state["error"])
        if st.button("Retry student-evaluation encrypted save", key=P + "retry_save"):
            _upload_changed()
            st.rerun()
        st.caption("A save or verification can fail after a write; retry to verify. "
                   "No educator summary is generated from this upload.")
        return
    receipt = state["receipt"]
    if receipt["action"] == "created":
        st.success("Student-evaluation CSV encrypted, archived, and verified in GitHub.")
    else:
        st.success("This identical student-evaluation CSV is already archived and verified. No duplicate was saved.")
    _show_details(receipt["details"])
    st.caption("Verified saved file: " + receipt["path"])


def _reload(client):
    st.markdown("### Reload a saved student-evaluation CSV")
    if st.button("Refresh saved student-evaluation exports", key=P + "refresh"):
        _refresh()
    if P + "list" not in st.session_state:
        try:
            st.session_state[P + "list"] = client.list_exports()
        except OPDArchiveError as exc:
            st.error(str(exc))
            return
        except Exception:
            st.error("The student-evaluation file list could not be loaded. Refresh to retry.")
            return
    filenames = st.session_state[P + "list"]["filenames"]
    if not filenames:
        _clear_loaded()
        st.info("No student-evaluation CSVs have been archived yet. Upload the original export above.")
        return
    if st.session_state.get(P + "choice") not in filenames:
        st.session_state.pop(P + "choice", None)
    selected = st.selectbox("Saved student-evaluation exports", filenames,
                            key=P + "choice", format_func=student_export_label,
                            on_change=_clear_loaded,
                            help="Ordered by course-date coverage, not upload date. The export identifier "
                                 "distinguishes different files covering the same dates.")
    loaded = st.session_state.get(P + "loaded")
    if loaded and loaded.get("filename") != selected:
        _clear_loaded()
    if st.button("Load / decrypt selected student-evaluation CSV", key=P + "load"):
        _clear_loaded()
        try:
            with st.spinner("Loading and decrypting the original student-evaluation CSV..."):
                # Read current HEAD, not a file resurrected from an old list snapshot.
                st.session_state[P + "loaded"] = client.load(selected)
        except OPDArchiveError as exc:
            st.error(str(exc))
        except Exception:
            st.error("The student-evaluation file could not be decrypted or verified. No download was prepared.")
    loaded = st.session_state.get(P + "loaded")
    if loaded and loaded.get("filename") == selected:
        st.success("Original student-evaluation CSV loaded and verified. Reloading does not change GitHub.")
        _show_details(loaded["details"])
        st.download_button("Download original student-evaluation CSV (optional)", loaded["raw"],
                           file_name=selected.removesuffix(".enc"), mime="text/csv", key=P + "download")
        st.caption("This download is the unencrypted original with a neutral archive filename. "
                   "It contains identifiable student assessments; keep it out of the public repository.")


def render():
    st.markdown("### Preceptor evaluations of students")
    st.write("Upload the original OASIS student-evaluation CSV. It is encrypted, saved in GitHub, "
             "then retrieved and decrypted to verify a byte-for-byte match.")
    st.info("Archive only: these assessments stay separate from feedback about educators. "
            "No student grades or comments are added to educator summaries, chair reports, or preceptor Word reports. "
            "No reporting dates or username corrections are needed to save the original.")
    try:
        config = get_opd_archive_config()
        client = GitHubOASISStudentEvaluations(GitHubOPDArchive(config))
    except OPDArchiveError as exc:
        st.error(str(exc))
        return
    signature = _scope(config)
    st.caption(f"Storage: {config.owner}/{config.repo} | branch {config.branch} | {client.folder}/")
    uploaded = st.file_uploader(
        "Upload preceptor evaluations of students (.csv) — automatically encrypted and saved",
        type=["csv"], key=P + "upload", on_change=_upload_changed,
        help="Up to 10 MiB. Supports Clinical Assessment of Student, PEDS Handoff, and "
             "PEDS History Taking & Physical Exam in the same original CSV.",
    )
    _save_upload(uploaded, client, signature)
    st.caption("Each changed export is retained separately. Re-uploading identical contents—even with a "
               "different filename—does not add a copy. This archive does not merge or score evaluations.")
    _reload(client)
    st.caption("Uses the existing GitHub token and encryption key. No app password has been added. "
               "Anyone able to access the running app can use this upload/recovery feature unless access is "
               "restricted elsewhere. Use institution-approved access and storage for student assessments.")
