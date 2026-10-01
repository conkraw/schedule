"""Date and OPD source controls shared by protected PTS and PTS Matching."""
from datetime import date
import hashlib
import json
import pandas as pd
import streamlit as st

from schedule_app.reports.teaching_export import teaching_csv_bytes
from schedule_app.reports.chair_summary import CHAIR_STUDENT_CONTINUITY_REPORT_VERSION
from schedule_app.services.report_diagnostics import REPORT_ISSUE_COLUMNS, REPORT_OUTPUT_VERSION, REPORT_BUILD_ID
from schedule_app.services.student_continuity import STUDENT_CONTINUITY_SCHEMA_VERSION, require_student_continuity_data
from schedule_app.services.teaching_validation import (STRICT_CONFLICT_SCHEMA_VERSION, STRICT_REPORT_VERSION,
    CONFLICT_CSV_COLUMNS, require_conflict_source_data)
from schedule_app.services.learner_reach import LEARNER_REACH_SCHEMA_VERSION, PARTICIPATION_REPORT_VERSION, require_learner_reach_data
from schedule_app.services.opd_archive import GitHubOPDArchive, OPDArchiveError, get_opd_archive_config
from schedule_app.services.reporting_periods import DATE_RANGE_SCHEMA_VERSION, teaching_report_date_text
from schedule_app.services.teaching_analysis import (teaching_academic_label, teaching_academic_start,
    teaching_filter_date_range, teaching_local_today, teaching_require_date_range_data, teaching_scan_archives)
from schedule_app.settings import (TEACHING_NAME_ORDERS, TEACHING_OPD_NAME_ORDER_OVERRIDES,
    TEACHING_PRECEPTOR_NAME_MAP, TEACHING_REPORT_VERSION, TEACHING_WORK_TYPE_MAP, TEACHING_WORK_TYPE_ORDER)
from schedule_app.services.teaching_priority import OUTPATIENT_PRIORITY_VERSION
from schedule_app.sections.reporting_date_controls import (REPORTING_MODES,
    render_period_controls as _render_period_controls, clear_teaching_downloads as _clear_teaching_downloads)
from schedule_app.sections.ignored_student_entries import load_ignored_student_entries_ui, P as IGNORE_P
from schedule_app.services.ignored_student_entries import (GitHubIgnoredStudentEntries, exclusion_signature,
    EXCLUSION_VERSION, require_matching_exclusions)


def _render_conflicts(exc):
    _clear_teaching_downloads()
    st.error(str(exc))
    st.write("Each conflict ID groups all source cells for one preceptor/date/AM-or-PM. "
             "A YES or NO indicates whether a non-ignored student remains in that cell; student names are omitted.")
    columns = ["conflict_id", "preceptor_name", "date", "shift", "conflicting_work_types",
               "rotation_start", "archive_file", "worksheet", "cell", "listed_work_type", "student_assigned"]
    if st.checkbox("Show OPD conflict details", key="pts_conflict_details", value=False):
        st.dataframe(pd.DataFrame(exc.rows).reindex(columns=columns), hide_index=True, use_container_width=True)
    st.download_button("Download OPD conflicts to correct (CSV)",
        data=teaching_csv_bytes(exc.rows, CONFLICT_CSV_COLUMNS),
        file_name="OPD_Conflict_Review.csv", mime="text/csv", key="teaching_download_conflicts")
    st.info("Open OPD Archive, select the listed rotation and download its original workbook. "
            "Review the listed worksheets/cells and correct which clinical experience owns the shift. "
            "Re-upload the corrected OPD in Create Student Schedule so it replaces that rotation's current archive. "
            "Then return here and click Load / refresh archived OPDs. No teaching report or pie chart will run until conflicts in the selected dates are resolved.")


def _render_report_issues(exc):
    """Explain the failing report entry without exporting any learner data."""
    _clear_teaching_downloads()
    st.error(str(exc))
    if st.checkbox("Show report issue details", key="pts_report_issue_details", value=False):
        st.dataframe(pd.DataFrame(exc.rows).reindex(columns=REPORT_ISSUE_COLUMNS),
                     hide_index=True, use_container_width=True)
    st.download_button("Download report issue details (CSV)",
        data=teaching_csv_bytes(exc.rows, REPORT_ISSUE_COLUMNS),
        file_name="Learner_Reach_Report_Issues.csv", mime="text/csv",
        key="teaching_download_report_issues")
    st.info("This identifies a report-calculation or report-writing problem, not necessarily an incorrect OPD. "
            "Restart after installing every update file and refresh archived OPDs. If it still fails, "
            "share the issue CSV. Student names, raw OPDs, tokens and encryption keys are not included.")
    st.caption("Report builder: " + REPORT_BUILD_ID)


def render_pts_workspace():
    """Shared date/source controls, called only after the page password gate."""
    mode, period, issue = _render_period_controls()
    if mode == REPORTING_MODES[0] and period is None:
        _clear_teaching_downloads()
        st.info(issue or "Select both reporting dates and enter a report label.")
    try:
        client = GitHubOPDArchive(get_opd_archive_config())
    except OPDArchiveError as exc:
        _clear_teaching_downloads()
        st.error(str(exc))
        return
    with st.expander("OPD settings (optional)", expanded=False):
        order = st.selectbox("Names around '~' in archived OPDs", TEACHING_NAME_ORDERS,
                         key="teaching_name_order",
                         help="Choose the actual name order in your OPDs. No rotation list is required. "
                              "Commas within names are preserved; use separate rows or semicolons for multiple students.")
    ignored_catalog = load_ignored_student_entries_ui(client, order)
    if ignored_catalog is None:
        _clear_teaching_downloads()
        return
    ignored_signature = exclusion_signature(ignored_catalog["entries"])
    options_signature = hashlib.sha256(json.dumps(
        [EXCLUSION_VERSION, ignored_signature, TEACHING_REPORT_VERSION, DATE_RANGE_SCHEMA_VERSION, LEARNER_REACH_SCHEMA_VERSION, STRICT_CONFLICT_SCHEMA_VERSION, OUTPATIENT_PRIORITY_VERSION, STUDENT_CONTINUITY_SCHEMA_VERSION, client.config.signature(), order,
         TEACHING_PRECEPTOR_NAME_MAP, TEACHING_OPD_NAME_ORDER_OVERRIDES,
         TEACHING_WORK_TYPE_MAP, TEACHING_WORK_TYPE_ORDER], sort_keys=True).encode()).hexdigest()
    if st.session_state.get("teaching_options_signature") != options_signature:
        # Keep reporting-year selections; the available-year check below drops
        # only values that no longer exist in the loaded source data.
        for key in ("teaching_scan", "teaching_zip", "teaching_zip_signature"):
            st.session_state.pop(key, None)
        st.session_state["teaching_options_signature"] = options_signature
    reload_clicked = st.button("Load / refresh archived OPDs", key="teaching_load_archives", type="primary")
    pending = st.session_state.pop(IGNORE_P + "rescan", None)
    auto_rescan = bool(pending and pending.get("scope") == client.config.signature() and pending.get("order") == order)
    if reload_clicked or auto_rescan:
        st.session_state.pop("teaching_scan", None)
        st.session_state.pop("assessment_completion_inputs", None)
        st.session_state.pop("assessment_completion_scope", None)
        _clear_teaching_downloads()
        bar = st.progress(0, text="Reading the current encrypted archive...")
        try:
            with st.spinner("Decrypting OPDs and counting retained student assignments..."):
                # A new explicit scan also refreshes other users' exclusion edits.
                current = GitHubIgnoredStudentEntries(client).load()
                if exclusion_signature(current["entries"]) != ignored_signature:
                    st.session_state[IGNORE_P + "catalog"] = current
                    st.session_state[IGNORE_P + "rescan"] = {
                        "scope": client.config.signature(), "order": order,
                        "commit": pending.get("commit") if auto_rescan and not reload_clicked else None,
                    }
                    st.rerun()
                inventory = []
                scan = teaching_scan_archives(
                    client, order,
                    commit=pending.get("commit") if auto_rescan and not reload_clicked else None,
                    ignored_student_entries=tuple(ignored_catalog["entries"]),
                    student_entry_collector=inventory.extend,
                    progress=lambda n, total: bar.progress(n / total, text=f"Read {n} of {total} current OPD files"))
            st.session_state[IGNORE_P + "inventory"] = inventory
            st.session_state["teaching_scan"] = scan
        except OPDArchiveError as exc:
            st.error(str(exc))
        except Exception:
            st.error("The teaching summary could not be completed. No partial ZIP was created. "
                     "Check the workbook layout and installed requirements, then retry.")
        finally:
            bar.empty()
    scan = st.session_state.get("teaching_scan")
    if scan is None:
        st.info("Load the archived OPDs to continue. Saved usernames and student matches are managed in PTS Matching.")
        return {"client": client, "order": order, "period": period, "mode": mode,
                "scan": None, "report_scan": None, "selected": []}
    try:
        require_matching_exclusions(scan, ignored_catalog)
        teaching_require_date_range_data(scan)
        require_learner_reach_data(scan)
        require_conflict_source_data(scan)
        require_student_continuity_data(scan)
    except OPDArchiveError as exc:
        st.session_state.pop("teaching_scan", None)
        _clear_teaching_downloads()
        st.warning(str(exc))
        return
    if not scan["sources"]:
        _clear_teaching_downloads()
        st.warning("No current encrypted OPDs were found in the configured GitHub archive folder.")
        return
    st.success(f"Read and decrypted {len(scan['sources'])} current OPD files. Archive snapshot: {scan['generated_at']}.")
    st.caption("One loaded snapshot is shared with PTS Matching. Refresh only to include newer OPD uploads.")
    if mode == REPORTING_MODES[0]:
        if period is None:
            return
        try:
            report_scan = teaching_filter_date_range(scan, period)
        except OPDArchiveError as exc:
            _clear_teaching_downloads()
            st.error(str(exc))
            return
        selected = [period.start.year]  # one custom period, not a calendar-year split
        selection_signature = (mode, *period.signature())
        file_part = period.filename_part()
        st.info(f"Reporting period {period.label}: {teaching_report_date_text(report_scan, selected[0])} (inclusive).")
    else:
        current = teaching_academic_start(teaching_local_today())
        years = sorted({row["academic_start_year"] for row in scan["monthly"]}
                       | {teaching_academic_start(date.fromisoformat(row["date"])) for row in scan["clinical_daily"]}
                       | {current}, reverse=True)
        if any(year not in years for year in st.session_state.get("teaching_selected_years", [])):
            st.session_state.pop("teaching_selected_years", None)
        selected = st.multiselect("Academic year(s) to include", years, default=[current],
                                  format_func=teaching_academic_label, key="teaching_selected_years")
        report_scan = scan
        selection_signature = (mode, *sorted(selected))
        file_part = teaching_academic_label(selected[0]) if len(selected) == 1 else "Multiple_Academic_Years"
    signature = (options_signature, scan["commit"], selection_signature, PARTICIPATION_REPORT_VERSION,
                 STRICT_REPORT_VERSION, CHAIR_STUDENT_CONTINUITY_REPORT_VERSION, REPORT_OUTPUT_VERSION)
    return {"client": client, "scan": scan, "report_scan": report_scan,
            "selected": selected, "signature": signature, "file_part": file_part,
            "period": period if mode == REPORTING_MODES[0] else None, "mode": mode,
            "order": order}
