"""Educator evaluation CSV reports with explicit missing-username correction."""
from __future__ import annotations

import hashlib
import json
from datetime import date

import streamlit as st

from schedule_app.services.opd_archive import GitHubOPDArchive, OPDArchiveError, get_opd_archive_config
from schedule_app.services.oasis_evaluations import GitHubOASISEvaluations, oasis_export_label
from schedule_app.services.oasis_educator_reports import (
    OASIS_REPORT_VERSION, OASISReportError, ISSUE_COLUMNS, SUMMARY_FILENAME,
    prepare_reports, educator_summary, build_report_downloads, csv_bytes,
    validate_username, question_key_rows, _date,
)
from schedule_app.services.oasis_educator_usernames import GitHubOASISUsernames

PREFIX = "oer_"


def _drop_outputs():
    st.session_state.pop(PREFIX + "downloads", None)
    st.session_state.pop(PREFIX + "report_signature", None)


def _show_error(exc):
    st.error(str(exc))
    issues = getattr(exc, "issues", [])
    if issues:
        st.dataframe(issues, hide_index=True, use_container_width=True)
        st.download_button("Download OASIS source issues", csv_bytes(issues, ISSUE_COLUMNS),
                           file_name="OASIS_Source_Issues.csv", mime="text/csv", key=PREFIX+"source_issues")


def _effective_overrides():
    saved = st.session_state.get(PREFIX + "username_catalog", {}).get("entries", {})
    local = st.session_state.get(PREFIX + "local_usernames", {})
    return {**saved, **local}


def _load_prepared(exports, signature):
    st.session_state.pop(PREFIX + "prepared", None)
    st.session_state.pop(PREFIX + "read_error", None)
    _drop_outputs()
    try:
        with st.spinner("Checking evaluation IDs, matching questions, and removing duplicate responses..."):
            st.session_state[PREFIX + "prepared"] = prepare_reports(exports)
        st.session_state[PREFIX + "loaded_source_signature"] = signature
    except OPDArchiveError as exc:
        st.session_state[PREFIX + "read_error"] = exc
    except Exception:
        st.session_state[PREFIX + "read_error"] = OASISReportError(
            "The selected evaluations could not be processed. No partial CSV was created. Check the export format and retry.")


def _username_controls(summary, prepared, filters, service):
    educators = summary["educators"]
    if not educators:
        return summary
    unresolved = {e["educator_key"] for e in summary["issues"]}
    if unresolved:
        st.warning(f"{len(unresolved)} educator(s) need a username or identity correction before the final CSV can be downloaded.")
        st.dataframe([{k: e[k] for k in ("educator_name", "record_id", "issue")} for e in summary["issues"]],
                     hide_index=True, use_container_width=True)
        st.download_button("Download username issues", csv_bytes(summary["issues"], ("educator_name", "record_id", "issue")),
                           file_name="OASIS_Username_Issues.csv", mime="text/csv", key=PREFIX+"username_issues")
    with st.expander("Add / correct an educator username", expanded=bool(unresolved)):
        st.caption("The username is the part before @ in Evaluator Email. Exported Evaluator Username is shown only as a reference; "
                   "it is not substituted automatically. Enter the username you want as record_id. Do not enter a full email address.")
        order = sorted(educators, key=lambda e: (e["educator_key"] not in unresolved, e["educator_name"].casefold()))
        lookup = {e["educator_key"]: e for e in order}
        if st.session_state.get(PREFIX+"educator_choice") not in lookup:
            st.session_state.pop(PREFIX+"educator_choice", None)
        key = st.selectbox("Educator to update", list(lookup), key=PREFIX+"educator_choice",
                           format_func=lambda k: lookup[k]["educator_name"] + (" — needs username" if k in unresolved else ""))
        e = lookup[key]
        st.write("Exported email: " + (e["evaluator_email"] or "Not provided"))
        st.write("Exported Evaluator Username (reference only): " + (e["exported_username"] or "Not provided"))
        shortkey = hashlib.sha256(key.encode()).hexdigest()[:16]
        entry = st.text_input("Username to use as record_id", value=e["record_id"], key=PREFIX+"username_"+shortkey)
        can_save = service is not None and PREFIX+"username_catalog" in st.session_state
        persist = st.checkbox("Save this username encrypted in GitHub for future reports", value=can_save,
                              disabled=not can_save, key=PREFIX+"save_username_github")
        if st.button("Apply username", key=PREFIX+"apply_username"):
            try:
                username = validate_username(entry)
                proposed = {**_effective_overrides(), key: {"record_id": username}}
                preview = educator_summary(prepared, proposed, **filters)
                if any(issue["educator_key"] == key and "Duplicate record_id" in issue["issue"] for issue in preview["issues"]):
                    raise OASISReportError("That record_id is already used by another educator. No username change was applied.")
                if persist and can_save:
                    saved = service.save(key, e["educator_name"], username,
                                         expected=st.session_state[PREFIX+"username_catalog"])
                    st.session_state[PREFIX+"username_catalog"] = saved
                    local = dict(st.session_state.get(PREFIX+"local_usernames", {}))
                    local.pop(key, None)
                    st.session_state[PREFIX+"local_usernames"] = local
                    st.success("Username saved encrypted in GitHub and verified.")
                else:
                    st.session_state.setdefault(PREFIX+"local_usernames", {})[key] = {"record_id": username}
                    st.success("Username applied for this session. It has not been saved to GitHub.")
                _drop_outputs()
            except OPDArchiveError as exc:
                _show_error(exc)
            except Exception:
                st.error("The username change was not confirmed. Refresh saved usernames and retry; no export was created.")
        overrides = _effective_overrides()
        if key in overrides:
            confirmed = st.checkbox("Remove this saved/session override and return to the source email", key=PREFIX+"confirm_remove_"+shortkey)
            if st.button("Remove username override", disabled=not confirmed, key=PREFIX+"remove_username"):
                try:
                    catalog = st.session_state.get(PREFIX+"username_catalog", {})
                    if key in catalog.get("entries", {}):
                        st.session_state[PREFIX+"username_catalog"] = service.remove(key, expected=catalog)
                    st.session_state.get(PREFIX+"local_usernames", {}).pop(key, None)
                    _drop_outputs()
                    st.success("Override removed. Source email-based identification is used again; a missing email will be flagged.")
                except OPDArchiveError as exc:
                    _show_error(exc)
        st.caption("Saving an override never changes the original CSV or invents an email. A missing email remains flagged in the final CSV.")
    return educator_summary(prepared, _effective_overrides(), **filters)


def render():
    st.subheader("OASIS Educator Reports")
    st.write("Create one CSV row per educator: number of evaluations, an average and response count for each multiple-choice question, "
             "and combined strengths / areas-for-improvement comments.")
    st.caption("Uses Evaluator as the educator. Counts distinct submitted Form Records, not question rows. "
               "Question IDs and wording align old and new forms; changing question numbers do not mix the results.")
    archive = None
    config_error = None
    try:
        archive = GitHubOPDArchive(get_opd_archive_config())
    except OPDArchiveError as exc:
        config_error = str(exc)
    scope = (OASIS_REPORT_VERSION, archive.config.signature() if archive else "local_upload_only")
    if st.session_state.get(PREFIX+"scope") != scope:
        for key in list(st.session_state):
            if key.startswith(PREFIX):
                st.session_state.pop(key, None)
        st.session_state[PREFIX+"scope"] = scope
    source = st.radio("Evaluation source", ("Saved OASIS exports in GitHub", "Upload CSV for this report only"),
                      key=PREFIX+"source", horizontal=True)
    exports = None
    signature = None
    if source == "Saved OASIS exports in GitHub":
        if archive is None:
            st.error(config_error or "Configure the existing OPD archive Secrets first.")
            return
        client = GitHubOASISEvaluations(archive)
        if st.button("Refresh saved OASIS exports", key=PREFIX+"refresh_exports"):
            for key in ("export_list", "prepared", "read_error"):
                st.session_state.pop(PREFIX+key, None)
            _drop_outputs()
        if PREFIX+"export_list" not in st.session_state:
            try:
                st.session_state[PREFIX+"export_list"] = client.list_exports()
            except OPDArchiveError as exc:
                _show_error(exc)
                return
            except Exception:
                st.error("Saved exports could not be listed. Refresh and retry.")
                return
        snapshot = st.session_state[PREFIX+"export_list"]
        filenames = snapshot["filenames"]
        if not filenames:
            st.info("No saved exports yet. Use OASIS Evaluation Archive to encrypt/save your CSV, then return here.")
            return
        selection_key = PREFIX+"selected_exports"
        if selection_key in st.session_state:
            st.session_state[selection_key] = [f for f in st.session_state[selection_key] if f in filenames]
        chosen = st.multiselect("OASIS exports to include", filenames, default=filenames[:100],
                                format_func=oasis_export_label, key=selection_key)
        st.caption("You may combine snapshots: identical Form Record / Question ID responses count once. "
                   "Different copies of the same response stop processing; select the intended snapshot rather than averaging conflicting copies.")
        signature = (scope, source, snapshot["commit"], tuple(sorted(chosen)))
        ready = bool(chosen) and len(chosen) <= 100
    else:
        uploaded = st.file_uploader("Upload OASIS evaluation CSV(s)", type=["csv"], accept_multiple_files=True,
                                    key=PREFIX+"uploads")
        st.caption("These uploads are used only for this report. They are not saved here; use OASIS Evaluation Archive to preserve an encrypted original.")
        exports = [(f.name, f.getvalue()) for f in uploaded]
        signature = (scope, source, tuple((name, hashlib.sha256(raw).hexdigest()) for name, raw in exports))
        ready = bool(exports)
    if st.session_state.get(PREFIX+"active_signature") != signature:
        for key in ("prepared", "read_error"):
            st.session_state.pop(PREFIX+key, None)
        _drop_outputs()
        st.session_state[PREFIX+"active_signature"] = signature
    if st.button("Read selected evaluations", disabled=not ready, key=PREFIX+"read", type="primary"):
        if exports is None:
            st.session_state.pop(PREFIX+"prepared", None)
            st.session_state.pop(PREFIX+"read_error", None)
            _drop_outputs()
            try:
                with st.spinner("Downloading and decrypting the selected OASIS exports..."):
                    exports = []
                    size = 0
                    for filename in chosen:
                        item = client.load(filename, commit=snapshot["commit"])
                        size += len(item["raw"])
                        if size > 64 * 1024 * 1024:
                            raise OASISReportError("Selected exports exceed 64 MiB. Choose fewer exports.")
                        exports.append((filename, item["raw"]))
                _load_prepared(exports, signature)
            except OPDArchiveError as exc:
                st.session_state[PREFIX+"read_error"] = exc
            except Exception:
                st.session_state[PREFIX+"read_error"] = OASISReportError("A selected export could not be retrieved. No partial report was created. Refresh the export list and retry.")
        else:
            _load_prepared(exports, signature)
    if st.session_state.get(PREFIX+"read_error"):
        _show_error(st.session_state[PREFIX+"read_error"])
        return
    prepared = st.session_state.get(PREFIX+"prepared")
    if prepared is None:
        return
    if not prepared["forms"]:
        st.warning("No submitted evaluations were found. Forms without Submit Date are not counted.")
        return
    st.success(f"Read {len(prepared['sources'])} export(s): {len(prepared['forms']):,} distinct submitted evaluations, "
               f"{prepared['duplicates_removed']:,} duplicate question rows removed.")
    if prepared["unsubmitted_forms_excluded"]:
        st.warning(f"{prepared['unsubmitted_forms_excluded']} form(s) have no Submit Date and were excluded as unsubmitted.")
    courses = sorted({f["key"][0] for f in prepared["forms"]})
    types = sorted({f["key"][1] for f in prepared["forms"]})
    # Widget keys include the source identity so a new file cannot leave stale filters.
    skey = hashlib.sha256(repr(signature).encode()).hexdigest()[:12]
    selected_courses = st.multiselect("Course(s)", courses, default=courses, key=PREFIX+"courses_"+skey)
    default_types = [t for t in types if t == "*Clinical Teaching Eval"] or types
    selected_types = st.multiselect("Evaluation form(s)", types, default=default_types, key=PREFIX+"types_"+skey)
    filters = {"courses": selected_courses, "evaluation_types": selected_types}
    filter_description = "All submitted evaluations in selected courses/forms and exports."
    limited = st.checkbox("Limit to a reporting date range", value=False, key=PREFIX+"limit_dates")
    if limited:
        field = st.selectbox("Date field to filter", ("Submit Date", "Start Date", "End Date"), key=PREFIX+"date_field")
        days = [_date(f["metadata"][field]) for f in prepared["forms"]]
        days = [d for d in days if d]
        a, b = st.columns(2)
        start = a.date_input("Start date (included)", value=min(days) if days else date.today(), key=PREFIX+"start_"+skey+field)
        end = b.date_input("End date (included)", value=max(days) if days else date.today(), key=PREFIX+"end_"+skey+field)
        filters.update(start_date=start, end_date=end, date_field=field)
        filter_description = f"Date filter: {field}, {start} through {end}, both included."
    service = GitHubOASISUsernames(archive) if archive else None
    if service:
        if st.button("Refresh saved usernames", key=PREFIX+"refresh_usernames"):
            st.session_state.pop(PREFIX+"username_catalog", None)
            st.session_state.pop(PREFIX+"username_load_error", None)
            _drop_outputs()
        if PREFIX+"username_catalog" not in st.session_state and PREFIX+"username_load_error" not in st.session_state:
            try:
                st.session_state[PREFIX+"username_catalog"] = service.load()
            except OPDArchiveError as exc:
                st.session_state[PREFIX+"username_load_error"] = str(exc)
            except Exception:
                st.session_state[PREFIX+"username_load_error"] = "Saved username corrections could not be loaded."
        if st.session_state.get(PREFIX+"username_load_error"):
            st.warning(st.session_state[PREFIX+"username_load_error"])
            if not st.checkbox("Continue without saved usernames for this session", key=PREFIX+"ignore_username_load_error"):
                _drop_outputs()
                return
    try:
        summary = educator_summary(prepared, _effective_overrides(), **filters)
        summary = _username_controls(summary, prepared, filters, service)
    except OPDArchiveError as exc:
        _drop_outputs()
        _show_error(exc)
        return
    if not summary["rows"]:
        _drop_outputs()
        st.info("No evaluations match the chosen courses, forms, and dates.")
        return
    a, b = st.columns(2)
    a.metric("Educators", len(summary["rows"]))
    b.metric("Evaluations", summary["evaluation_count"])
    st.markdown("**Educator summary preview**")
    st.dataframe([{k: v for k, v in row.items() if not k.endswith("_comments")} for row in summary["rows"]],
                 hide_index=True, use_container_width=True)
    st.caption("Each q<ID>_mean has its own q<ID>_n response count. Blank/N/A scores are excluded, not counted as zero. "
               "q1286_mean is a mean of time-category codes, NOT weeks or a teaching-quality score. No composite average is calculated.")
    with st.expander("Question IDs, exact wording, and CSV columns"):
        st.dataframe(question_key_rows(prepared), hide_index=True, use_container_width=True)
    with st.expander("Review combined comments"):
        st.dataframe([{k: row[k] for k in ("educator_name", "strengths_comments", "areas_for_improvement_comments")} for row in summary["rows"]],
                     hide_index=True, use_container_width=True)
    if any(row["email_missing"] == "YES" for row in summary["rows"]):
        st.caption("Some source emails are missing. After you supply a username, email_missing remains YES for transparency; no email is invented.")
    report_signature = hashlib.sha256(json.dumps([signature, filters, _effective_overrides()], sort_keys=True, default=str).encode()).hexdigest()
    if st.session_state.get(PREFIX+"report_signature") != report_signature:
        _drop_outputs()
    if summary["issues"]:
        _drop_outputs()
        st.info("Complete the username corrections above to enable the final CSV. No educator is silently dropped.")
    if st.button("Create educator report CSV", disabled=bool(summary["issues"]), key=PREFIX+"build_csv"):
        _drop_outputs()
        try:
            downloads = build_report_downloads(prepared, summary, filter_description=filter_description)
            st.session_state[PREFIX+"downloads"] = downloads
            st.session_state[PREFIX+"report_signature"] = report_signature
        except OPDArchiveError as exc:
            _show_error(exc)
        except Exception:
            st.error("The report could not be completed. No partial CSV was retained.")
    if st.session_state.get(PREFIX+"downloads"):
        downloads = st.session_state[PREFIX+"downloads"]
        st.download_button("Download educator summary CSV", downloads["csv"], file_name=SUMMARY_FILENAME,
                           mime="text/csv", key=PREFIX+"download_csv")
        st.download_button("Download CSV + question key + detail (ZIP)", downloads["zip"],
                           file_name="OASIS_Educator_Reports.zip", mime="application/zip", key=PREFIX+"download_zip")
    st.caption("The report omits structured student identifiers, but verbatim comments may still identify someone. "
               "CSV/ZIP downloads are unencrypted and are not automatically saved to GitHub. This section has no app-password gate; "
               "anyone with access to the running app can use it.")
