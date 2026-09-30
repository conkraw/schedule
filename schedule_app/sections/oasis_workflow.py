"""One OASIS menu: educator feedback/reporting or separate student-assessment archive."""
from __future__ import annotations
import hashlib
import json

import streamlit as st

from schedule_app.services.opd_archive import GitHubOPDArchive, OPDArchiveError, get_opd_archive_config
from schedule_app.services.oasis_evaluations import GitHubOASISEvaluations, oasis_export_label
from schedule_app.services.oasis_educator_reports import (
    OASISReportError, educator_summary, question_key_rows, validate_username,
)
from schedule_app.services.oasis_educator_usernames import GitHubOASISUsernames
from schedule_app.services.oasis_workflow import (
    WORKFLOW_VERSION, OASISOutputScope, GitHubOASISSummaries,
    load_cumulative_evaluations, make_period_summary, summary_csv,
)
from schedule_app.sections import oasis_date_controls
from schedule_app.services.oasis_student_evaluations import validate_educator_upload_kind

OASIS_EVALUATION_KINDS = ("Evaluations of educators", "Evaluations of students")

P = "oasis_combined_"


def _clear_output():
    for suffix in ("receipt", "publish_signature", "publish_error", "attempt_signature"):
        st.session_state.pop(P+suffix, None)


def _clear_loaded():
    for suffix in ("prepared", "read_error", "username_catalog", "username_error"):
        st.session_state.pop(P+suffix, None)
    _clear_output()


def _error(exc):
    st.error(str(exc))
    if getattr(exc, "issues", None):
        # Keep diagnosis on screen: the only generated report/download is the
        # summary CSV. Do not make extra issue/question-key CSV/ZIP outputs.
        st.dataframe(exc.issues, hide_index=True, use_container_width=True)


def _scope_state(archive):
    scope = (WORKFLOW_VERSION, archive.config.signature())
    if st.session_state.get(P+"scope") != scope:
        for key in list(st.session_state):
            if key.startswith(P):
                st.session_state.pop(key, None)
        st.session_state[P+"scope"] = scope


def _upload_original(upload, client) -> bool:
    """One write attempt per upload event; user explicitly retries a failed write."""
    if upload is None:
        st.session_state.pop(P+"active_upload", None)
        st.session_state.pop(P+"upload_status", None)
        return True
    raw = upload.getvalue()
    # Changing the upload signature revalidates an already-open older session.
    signature = ("minimized-evaluation-source-v2", getattr(upload, "file_id", None), hashlib.sha256(raw).hexdigest())
    if st.session_state.get(P+"active_upload") != signature:
        st.session_state[P+"active_upload"] = signature
        st.session_state.pop(P+"upload_status", None)
        _clear_loaded()
    if st.button("Retry reduced CSV save", key=P+"retry_upload",
                 disabled=not bool(st.session_state.get(P+"upload_status", {}).get("error"))):
        st.session_state.pop(P+"upload_status", None)
        _clear_loaded()
    if P+"upload_status" not in st.session_state:
        try:
            with st.spinner("Encrypting and verifying the reduced OASIS CSV in GitHub..."):
                # Do not mistake student grades for an educator's teaching feedback.
                validate_educator_upload_kind(raw)
                receipt = client.save(raw)
            st.session_state[P+"upload_status"] = {"receipt": receipt}
            # Never reuse a report calculated before this upload was archived.
            _clear_loaded()
        except OPDArchiveError as exc:
            st.session_state[P+"upload_status"] = {"error": exc}
        except Exception:
            st.session_state[P+"upload_status"] = {"error": OASISReportError(
                "The reduced CSV save was not confirmed. Retry after checking the archive settings/connection.")}
    status = st.session_state[P+"upload_status"]
    if status.get("error"):
        _clear_output()
        _error(status["error"])
        st.info("The reduced upload is not confirmed archived. No updated summary will be published for it.")
        return False
    receipt = status["receipt"]
    if receipt["action"] == "unchanged":
        st.success("Reduced OASIS CSV already archived and verified; the identical file was not saved twice.")
    else:
        st.success("Reduced OASIS CSV archived, encrypted, and verified in GitHub.")
    st.caption(f"Source rows: {receipt['details']['row_count']:,}. Previous exports remain available for cumulative evaluation counts.")
    privacy = receipt["details"].get("privacy", {})
    st.caption(f"Retained {privacy.get('retained_column_count', receipt['details']['column_count'])} reporting columns; "
               f"removed {privacy.get('removed_column_count', 0)} unneeded columns BEFORE encryption. "
               "Structured student identifiers, gender and other unused demographics are not stored in new educator-feedback files.")
    return True


def _load_data(archive, client):
    if P+"prepared" in st.session_state or P+"read_error" in st.session_state:
        return
    try:
        with st.spinner("Reading the cumulative archive and removing repeated evaluation/question rows..."):
            prepared = load_cumulative_evaluations(client)
            # Both the source data and initial username map refer to this snapshot.
            names = GitHubOASISUsernames(archive).load(commit=prepared["archive_commit"])
        st.session_state[P+"prepared"] = prepared
        st.session_state[P+"username_catalog"] = names
    except OPDArchiveError as exc:
        st.session_state[P+"read_error"] = exc
    except Exception:
        st.session_state[P+"read_error"] = OASISReportError(
            "Cumulative evaluations could not be read completely. No partial output was created. Refresh and retry.")


def _selected_filters(prepared):
    courses = sorted({f["key"][0] for f in prepared["forms"]})
    types = sorted({f["key"][1] for f in prepared["forms"]})
    for suffix, available, defaults in (("courses", courses, courses),
                                       ("types", types, [t for t in types if t == "*Clinical Teaching Eval"] or types)):
        key = P+suffix
        if key not in st.session_state:
            st.session_state[key] = list(defaults)
        elif any(v not in available for v in st.session_state[key]):
            st.session_state[key] = [v for v in st.session_state[key] if v in available]
    with st.expander("Courses and evaluation forms"):
        selected_courses = st.multiselect("Course(s) included in the output", courses, key=P+"courses")
        selected_types = st.multiselect("Evaluation form(s) included in the output", types, key=P+"types")
    return selected_courses, selected_types


def _username_controls(summary, prepared, filters, service):
    notice = st.session_state.pop(P+"username_notice", None)
    if notice:
        st.success(notice)
    unresolved = {e["educator_key"] for e in summary["issues"]}
    if unresolved:
        _clear_output()
        st.warning(f"Action needed: {len(unresolved)} educator(s) need a missing/duplicate username corrected. "
                   "The reduced source is stored; the output CSV has NOT been created for this selection yet.")
        st.dataframe([{k: e[k] for k in ("educator_name", "record_id", "issue")} for e in summary["issues"]],
                     hide_index=True, use_container_width=True)
    educators = sorted(summary["educators"], key=lambda e: (e["educator_key"] not in unresolved, e["educator_name"].casefold()))
    if not educators:
        return summary
    with st.expander("Fix / maintain educator usernames", expanded=bool(unresolved)):
        lookup = {e["educator_key"]: e for e in educators}
        if st.session_state.get(P+"educator_choice") not in lookup:
            st.session_state.pop(P+"educator_choice", None)
        selected = st.selectbox("Educator", list(lookup), key=P+"educator_choice",
                                format_func=lambda k: lookup[k]["educator_name"] + (" — needs correction" if k in unresolved else ""))
        item = lookup[selected]
        st.write("Source email: " + (item["evaluator_email"] or "Not provided"))
        st.caption("Exported username (reference only): " + (item["exported_username"] or "Not provided"))
        suffix = hashlib.sha256(selected.encode()).hexdigest()[:16]
        username = st.text_input("Username to use as record_id (not a full email)",
                                 value=item["record_id"], key=P+"username_"+suffix)
        if st.button("Save username and continue", key=P+"apply_username"):
            try:
                username = validate_username(username)
                catalog = st.session_state[P+"username_catalog"]
                proposed = {**catalog["entries"], selected: {"record_id": username}}
                check = educator_summary(prepared, proposed, **filters)
                if any(i["educator_key"] == selected and "Duplicate record_id" in i["issue"] for i in check["issues"]):
                    raise OASISReportError("That record_id already belongs to another educator. Enter a different verified username.")
                saved = service.save(selected, item["educator_name"], username, expected=catalog)
                st.session_state[P+"username_catalog"] = saved
                _clear_output()
                st.session_state[P+"username_notice"] = "Username saved encrypted in GitHub and verified. Continuing automatically."
                st.rerun()
            except OPDArchiveError as exc:
                _clear_output()
                _error(exc)
            except Exception:
                _clear_output()
                st.error("The username update was not confirmed. Refresh the cumulative data/usernames and retry.")
        catalog = st.session_state[P+"username_catalog"]
        if selected in catalog["entries"]:
            # Consent is tied to the exact educator and catalog revision.
            confirm = st.checkbox("Remove this username override and use the source email again",
                                  key=P+"remove_confirm_"+suffix+str(catalog["sha"])[:12])
            if st.button("Remove selected username override", disabled=not confirm, key=P+"remove_username"):
                try:
                    st.session_state[P+"username_catalog"] = service.remove(selected, expected=catalog)
                    _clear_output()
                    st.session_state[P+"username_notice"] = "Override removed; source email rules apply again."
                    st.rerun()
                except OPDArchiveError as exc:
                    _error(exc)
        st.caption("record_id is the username before @ in Evaluator Email, unless you supply an override. "
                   "The override persists in GitHub; the reduced source CSV is unchanged and no email is invented.")
    return educator_summary(prepared, st.session_state[P+"username_catalog"]["entries"], **filters)


def _publish(archive, prepared, scope, summary):
    raw = summary_csv(summary)
    names = st.session_state[P+"username_catalog"]
    signature = hashlib.sha256(json.dumps([
        archive.config.signature(), prepared["archive_filenames"],
        scope.filename, scope.period.signature(), names["sha"], hashlib.sha256(raw).hexdigest(),
    ], sort_keys=True).encode()).hexdigest()
    if st.session_state.get(P+"attempt_signature") != signature:
        _clear_output()
    if st.session_state.get(P+"publish_error"):
        if st.button("Retry encrypted output CSV save", key=P+"retry_output"):
            _clear_output()
    if P+"attempt_signature" not in st.session_state:
        st.session_state[P+"attempt_signature"] = signature
        try:
            with st.spinner("Creating, encrypting, and verifying the cumulative summary CSV in GitHub..."):
                receipt = GitHubOASISSummaries(archive).save(scope, raw, prepared=prepared, username_catalog=names)
            st.session_state[P+"receipt"] = receipt
            st.session_state[P+"publish_signature"] = signature
            # Refresh any old saved-output list without losing the active report.
            st.session_state.pop(P+"saved_output_list", None)
            st.session_state.pop(P+"saved_output_loaded", None)
        except OPDArchiveError as exc:
            st.session_state[P+"publish_error"] = exc
        except Exception:
            st.session_state[P+"publish_error"] = OASISReportError(
                "Output CSV save or verification failed. No success is confirmed. Retry or refresh cumulative evaluations.")
    receipt = st.session_state.get(P+"receipt")
    if receipt and st.session_state.get(P+"publish_signature") == signature:
        action = {"created": "created", "updated": "updated", "unchanged": "already current"}[receipt["action"]]
        st.success(f"Output CSV {action}, encrypted, and verified in GitHub — "
                   f"{receipt['details']['educator_count']:,} educators; {receipt['details']['evaluation_count']:,} evaluations.")
        st.code(receipt["path"], language=None)
        st.caption("Only the summary CSV is generated for this reporting period. No Word, ZIP, question-key, or detail output is created. "
                   "No download is required. Other periods are rebuilt when you select/load and apply those dates.")
        with st.expander("Optional: review or download this output CSV"):
            st.dataframe([{k: v for k,v in row.items() if not k.endswith("_comments")} for row in summary["rows"]],
                         hide_index=True, use_container_width=True)
            st.download_button("Download output CSV for review", receipt["raw"],
                               file_name=scope.filename.removesuffix(".enc"), mime="text/csv", key=P+"download_current")
            st.caption("The review download is unencrypted; comments can contain identifying details.")
    elif st.session_state.get(P+"publish_error"):
        _error(st.session_state[P+"publish_error"])
        st.caption("No new output is confirmed. Any older file in GitHub remains a previous report, not a refreshed result.")


def _process(archive, period):
    client = GitHubOASISEvaluations(archive)
    st.markdown("### Upload learner feedback about educators")
    uploaded = st.file_uploader("Upload the original or updated educator-feedback CSV", type=["csv"], key=P+"upload")
    upload_ok = _upload_original(uploaded, client)
    if st.button("Refresh cumulative evaluations and usernames", key=P+"refresh_data"):
        _clear_loaded()
    if not upload_ok:
        return
    _load_data(archive, client)
    if st.session_state.get(P+"read_error"):
        _clear_output()
        _error(st.session_state[P+"read_error"])
        st.caption("All prior archived sources are preserved. Conflicting copies of an existing answer stop the report, "
                   "rather than silently choosing or averaging different versions.")
        return
    prepared = st.session_state.get(P+"prepared")
    if not prepared:
        return
    st.caption("This is a snapshot of this session's source data. Use Refresh cumulative evaluations and usernames to include uploads made from another session.")
    st.caption(f"Cumulative source data: {len(prepared['sources'])} archived export(s), "
               f"{len(prepared['forms']):,} distinct submitted evaluations, "
               f"{prepared['duplicates_removed']:,} repeated question rows counted once.")
    if prepared["unsubmitted_forms_excluded"]:
        st.warning(f"{prepared['unsubmitted_forms_excluded']} evaluation(s) without Submit Date are excluded as unsubmitted.")
    courses, types = _selected_filters(prepared)
    if not courses or not types:
        _clear_output()
        st.info("Select at least one course and evaluation form to prepare a summary.")
        return
    filters = {"courses": courses, "evaluation_types": types}
    if period:
        filters.update(start_date=period.start, end_date=period.end, date_field="Submit Date")
    try:
        summary = educator_summary(prepared, st.session_state[P+"username_catalog"]["entries"], **filters)
        summary = _username_controls(summary, prepared, filters, GitHubOASISUsernames(archive))
        if not period:
            _clear_output()
            st.info("Your source is stored. Apply the Submit Date reporting period above to create the encrypted output CSV.")
            return
        scope = OASISOutputScope(period, tuple(courses), tuple(types))
        summary = make_period_summary(prepared, scope, st.session_state[P+"username_catalog"]["entries"])
        if summary["issues"]:
            _clear_output()
            return
        if not summary["rows"]:
            _clear_output()
            st.info("No submitted evaluations match these dates/course/forms. No empty output was saved, and previous outputs were not deleted.")
            return
        st.markdown("### Encrypted output CSV")
        _publish(archive, prepared, scope, summary)
        with st.expander("How averages and comments are reported"):
            st.write("One row per educator. Evaluation counts use distinct submitted Form Records, not question rows. "
                     "Means use Multiple Choice Value with a separate response count for each Question ID. "
                     "Blank/N/A ratings are excluded. Strengths and improvement comments are combined separately, without rewriting them.")
            st.dataframe(question_key_rows(prepared), hide_index=True, use_container_width=True)
            st.caption("q1286_mean is an average of duration-category codes, not weeks or a teaching-quality score.")
    except OPDArchiveError as exc:
        _clear_output()
        _error(exc)


def _review_saved(archive):
    service = GitHubOASISSummaries(archive)
    with st.expander("Optional: retrieve a previously saved output CSV"):
        st.caption("These are saved snapshots. To update a period with newer sources, load its dates above and apply the reporting period.")
        if st.button("List / refresh saved output CSVs", key=P+"list_outputs"):
            st.session_state.pop(P+"saved_output_loaded", None)
            st.session_state.pop(P+"saved_output_list", None)
            try:
                st.session_state[P+"saved_output_list"] = service.list_outputs()
            except OPDArchiveError as exc:
                _error(exc)
        listing = st.session_state.get(P+"saved_output_list")
        if listing is None:
            return
        names = listing["filenames"]
        if not names:
            st.info("No output CSVs have been saved yet.")
            return
        if st.session_state.get(P+"saved_output_choice") not in names:
            st.session_state.pop(P+"saved_output_choice", None)
        chosen = st.selectbox("Saved output CSV", names, key=P+"saved_output_choice")
        if st.button("Load / decrypt selected output CSV", key=P+"load_output"):
            st.session_state.pop(P+"saved_output_loaded", None)
            try:
                st.session_state[P+"saved_output_loaded"] = service.load(chosen)
            except OPDArchiveError as exc:
                _error(exc)
        loaded = st.session_state.get(P+"saved_output_loaded")
        if loaded and loaded["filename"] == chosen:
            d = loaded["details"]
            st.write(f"{d['label']} | {d['start_date']} to {d['end_date']} | "
                     f"{d['educator_count']} educators | {d['evaluation_count']} evaluations")
            st.download_button("Download saved output CSV for review", loaded["raw"],
                               file_name=chosen.removesuffix(".enc"), mime="text/csv", key=P+"download_saved")


def _review_originals(archive):
    """Recovery of already stored originals; not an additional generated report."""
    service = GitHubOASISEvaluations(archive)
    with st.expander("Optional: retrieve a reduced source CSV"):
        st.caption("Only required reporting columns are returned. Even older full exports are reduced before download. This does not change the saved archive.")
        if st.button("List / refresh source exports", key=P+"list_originals"):
            st.session_state.pop(P+"original_loaded", None)
            st.session_state.pop(P+"original_list", None)
            try:
                st.session_state[P+"original_list"] = service.list_exports()
            except OPDArchiveError as exc:
                _error(exc)
        listing = st.session_state.get(P+"original_list")
        if not listing:
            return
        names = listing["filenames"]
        if not names:
            st.info("No source exports are archived yet.")
            return
        if st.session_state.get(P+"original_choice") not in names:
            st.session_state.pop(P+"original_choice", None)
        chosen = st.selectbox("Saved source CSV", names, format_func=oasis_export_label, key=P+"original_choice")
        if st.button("Load / decrypt reduced source CSV", key=P+"load_original"):
            st.session_state.pop(P+"original_loaded", None)
            try:
                st.session_state[P+"original_loaded"] = service.load(chosen)
            except OPDArchiveError as exc:
                _error(exc)
        loaded = st.session_state.get(P+"original_loaded")
        if loaded and loaded["filename"] == chosen:
            st.download_button("Download reduced source CSV", loaded["raw"],
                               file_name=chosen.removesuffix(".enc"), mime="text/csv", key=P+"download_original")
            st.caption("This reduced CSV omits structured student identifiers. Educator-feedback comments remain and may identify people. The download is unencrypted.")


def render():
    from schedule_app.services.evaluation_access import require_evaluation_access
    if not require_evaluation_access(lock_key="evaluation_lock_oasis_workflow"):
        return
    st.subheader("OER")
    try:
        privacy_archive = GitHubOPDArchive(get_opd_archive_config())
    except OPDArchiveError as exc:
        st.error(str(exc))
        return
    from schedule_app.sections.evaluation_privacy import render as render_privacy_review
    render_privacy_review(privacy_archive)
    kind = st.radio("Evaluation type", OASIS_EVALUATION_KINDS,
                    key="oasis_evaluation_kind", horizontal=True,
                    help="Keep learners' feedback about educators separate from preceptors' assessments of students.")
    if kind == "Evaluations of students":
        from schedule_app.sections.oasis_student_evaluations import render as render_student_archive
        render_student_archive()
        return
    st.write("Upload → unnecessary columns removed → reduced CSV encrypted and saved → correct missing usernames → encrypted educator-summary CSV saved automatically.")
    st.caption("All archived source exports contribute. New evaluations are added; repeated copies count once. "
               "Averages and comments are recalculated from responses, never by appending or averaging prior summary rows.")
    try:
        archive = GitHubOPDArchive(get_opd_archive_config())
    except OPDArchiveError as exc:
        st.error(str(exc))
        return
    _scope_state(archive)
    period = oasis_date_controls.render(archive)
    _process(archive, period)
    _review_saved(archive)
    _review_originals(archive)
    st.caption("Uses your existing GitHub token and encryption key. OER and PTS share the protected-section password. "
               "Linked evaluation content and report downloads in PTS also require an unlocked session. "
               "The optional review CSV is unencrypted; do not upload it to a public repository.")
