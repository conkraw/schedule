"""Sidebar: teaching effort with exact, editable reporting periods.

Only this page's date controls and teaching reports change. Archive writes,
student schedules and the Power Automate workbook are unaffected.
"""
from datetime import date
from io import BytesIO
from zipfile import ZipFile
import hashlib
import json
import pandas as pd
import streamlit as st

from schedule_app.reports.teaching_export import teaching_build_zip
from schedule_app.services.learner_reach import (
    LEARNER_REACH_SCHEMA_VERSION, LEARNER_REACH_COLUMNS, REACH_DEFINITION,
    REACH_SCOPE_NOTE, require_learner_reach_data, reach_percent, reach_totals,
)
from schedule_app.services.opd_archive import GitHubOPDArchive, OPDArchiveError, get_opd_archive_config
from schedule_app.services.reporting_periods import (
    DATE_RANGE_SCHEMA_VERSION, ReportingPeriod, read_reporting_period_json,
    reporting_period_json, teaching_report_date_text,
)
from schedule_app.services.teaching_analysis import (
    teaching_academic_label, teaching_academic_start, teaching_annual_rows,
    teaching_filter_date_range, teaching_local_today, teaching_require_date_range_data,
    teaching_scan_archives, teaching_work_type_rows,
)
from schedule_app.settings import (
    TEACHING_CHAIR_SUMMARY_FILENAME, TEACHING_CSV_COLUMNS, TEACHING_NAME_ORDERS,
    TEACHING_OPD_NAME_ORDER_OVERRIDES, TEACHING_PRECEPTOR_NAME_MAP, TEACHING_REPORT_VERSION,
    TEACHING_WORK_TYPE_CSV_COLUMNS, TEACHING_WORK_TYPE_MAP, TEACHING_WORK_TYPE_ORDER,
)

REPORTING_MODES = ("Custom dates", "Standard July-June academic years")
PERIOD_FIELDS = {"label": "teaching_period_label", "start_date": "teaching_period_start",
                 "end_date": "teaching_period_end", "mode": "teaching_reporting_mode"}


def _remember_period_inputs():
    """Non-widget state survives switching to another section within this session."""
    preferences = dict(st.session_state.get("teaching_period_preferences", {}))
    for name, key in PERIOD_FIELDS.items():
        if key in st.session_state:
            preferences[name] = st.session_state[key]
    st.session_state["teaching_period_preferences"] = preferences


def _clear_teaching_downloads():
    st.session_state.pop("teaching_zip", None)
    st.session_state.pop("teaching_zip_signature", None)


def _load_period_settings():
    """Button callback: populate controls *before* the subsequent page rerun."""
    upload = st.session_state.get("teaching_period_json_upload")
    try:
        if upload is None:
            raise OPDArchiveError("Choose a saved reporting-date JSON file first.")
        period = read_reporting_period_json(upload.getvalue())
    except OPDArchiveError as exc:
        st.session_state["teaching_period_import_error"] = str(exc)
        return
    values = {"label": period.label, "start_date": period.start, "end_date": period.end,
              "mode": REPORTING_MODES[0]}
    for name, key in PERIOD_FIELDS.items():
        st.session_state[key] = values[name]
    st.session_state["teaching_period_preferences"] = values
    st.session_state.pop("teaching_period_import_error", None)
    st.session_state["teaching_period_import_ok"] = True
    _clear_teaching_downloads()


def _render_period_controls():
    preferences = st.session_state.get("teaching_period_preferences", {})
    # Initialize values only after Streamlit has removed obsolete widget state,
    # for example when returning from another sidebar section.
    for name, key in PERIOD_FIELDS.items():
        if key not in st.session_state and name in preferences:
            st.session_state[key] = preferences[name]
    if st.session_state.get("teaching_reporting_mode") not in REPORTING_MODES:
        st.session_state["teaching_reporting_mode"] = REPORTING_MODES[0]
    st.markdown("**Reporting period**")
    mode = st.radio("Choose reporting dates", REPORTING_MODES, key="teaching_reporting_mode",
                    horizontal=True, on_change=_remember_period_inputs)
    period, issue = None, None
    if mode == REPORTING_MODES[0]:
        left, right = st.columns(2)
        start = left.date_input("Start date (included)", value=None, min_value=date(1970, 1, 1),
                                max_value=date(2100, 12, 31), format="MM/DD/YYYY",
                                key="teaching_period_start", on_change=_remember_period_inputs)
        end = right.date_input("End date (included)", value=None, min_value=date(1970, 1, 1),
                               max_value=date(2100, 12, 31), format="MM/DD/YYYY",
                               key="teaching_period_end", on_change=_remember_period_inputs)
        label = st.text_input("Report label / academic year", value="", placeholder="e.g., 26-27",
                              max_chars=60, key="teaching_period_label", on_change=_remember_period_inputs,
                              help="This label goes in the academic_year CSV column and every Word report. "
                                   "It does not determine the dates or force a July boundary.")
        try:
            period = ReportingPeriod(label, start, end)
        except OPDArchiveError as exc:
            issue = str(exc)
        st.caption("Choose the exact dates you need, even for a period longer than 12 months. "
                   "Both dates are included. Only assignments on those dates count; July 1 does not split a custom report.")
    else:
        st.caption("Optional original mode: July 1 through June 30, with a separate section for each selected year.")
    _remember_period_inputs()
    with st.expander("Save / reload these date settings (optional)"):
        st.caption("Dates and the label stay selected while you use this session. To reuse them after closing "
                   "the session, download the small JSON settings file and load it here next time. "
                   "This saves no OPDs, student names or credentials and does not write to GitHub.")
        if period is not None:
            st.download_button("Download reporting-date settings", data=reporting_period_json(period),
                               file_name=f"Reporting_Dates_{period.filename_part()}.json", mime="application/json",
                               key="teaching_save_period")
        upload = st.file_uploader("Load saved reporting-date settings (.json)", type=["json"],
                                   key="teaching_period_json_upload")
        st.button("Apply saved reporting dates", key="teaching_apply_period", disabled=upload is None,
                  on_click=_load_period_settings)
        if st.session_state.get("teaching_period_import_error"):
            st.error(st.session_state["teaching_period_import_error"])
        if st.session_state.pop("teaching_period_import_ok", False):
            st.success("Reporting dates and label loaded. Generate the reports again for this period.")
    return mode, period, issue


def render():
    st.subheader("Preceptor Teaching Summary")
    st.write("Read current encrypted OPDs from GitHub, then generate a chair-friendly Word summary, "
             "overall and work-type CSVs, and one Word teaching report per preceptor.")
    st.caption("HOPE_DRIVE + ETOWN + NYES = Academic Pediatrics. Ward A, PSHCH Nursery, Complex Care "
               "and other services stay separate. One student-shift = four educational hours; two students "
               "at once count twice. These are scheduled student-weighted hours, not distinct clock hours.")
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
    order = st.selectbox("Names around '~' in archived OPDs", TEACHING_NAME_ORDERS,
                         key="teaching_name_order",
                         help="Choose the actual name order in your OPDs. No rotation list is required. "
                              "Commas within names are preserved; use separate rows or semicolons for multiple students.")
    options_signature = hashlib.sha256(json.dumps(
        [TEACHING_REPORT_VERSION, DATE_RANGE_SCHEMA_VERSION, LEARNER_REACH_SCHEMA_VERSION, client.config.signature(), order,
         TEACHING_PRECEPTOR_NAME_MAP, TEACHING_OPD_NAME_ORDER_OVERRIDES,
         TEACHING_WORK_TYPE_MAP, TEACHING_WORK_TYPE_ORDER], sort_keys=True).encode()).hexdigest()
    if st.session_state.get("teaching_options_signature") != options_signature:
        for key in ("teaching_scan", "teaching_zip", "teaching_zip_signature", "teaching_selected_years"):
            st.session_state.pop(key, None)
        st.session_state["teaching_options_signature"] = options_signature
    if st.button("Load / refresh archived OPDs", key="teaching_load_archives", type="primary"):
        st.session_state.pop("teaching_scan", None)
        _clear_teaching_downloads()
        bar = st.progress(0, text="Reading the current encrypted archive...")
        try:
            with st.spinner("Decrypting OPDs and counting student assignments..."):
                scan = teaching_scan_archives(
                    client, order,
                    progress=lambda n, total: bar.progress(n / total, text=f"Read {n} of {total} current OPD files"))
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
        st.caption("Nothing is downloaded or decrypted until you click Load / refresh archived OPDs. "
                   "This section never uploads a report or decrypted OPD to GitHub.")
        return
    try:
        teaching_require_date_range_data(scan)
        require_learner_reach_data(scan)
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
    st.caption("After loading, changing dates recalculates from this snapshot without another GitHub download. "
               "Click Load / refresh again to include later OPD uploads. Older Git history is not counted.")
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
    signature = (options_signature, scan["commit"], selection_signature)
    if st.session_state.get("teaching_zip_signature") != signature:
        _clear_teaching_downloads()
        st.session_state["teaching_zip_signature"] = signature
    annual = teaching_annual_rows(report_scan, selected)
    if report_scan["unresolved_preceptor_labels"]:
        st.warning("Some provider labels identify a site/slot/combined entry rather than a person. "
                   "They remain literal and are separated from named preceptors in the Word reports.")
        with st.expander("Provider labels to review"):
            st.write(", ".join(report_scan["unresolved_preceptor_labels"]))
    if report_scan.get("work_type_conflicts"):
        st.warning("Some identical assignments appear in different work types. Each counts once under Work type needs review.")
    if report_scan.get("clinical_shift_conflicts"):
        st.warning("Some clinical half-days appear in multiple work types. Overall Learner Reach counts each shift once; "
                   "affected work-type percentages are N/A pending review. See Clinical_Shift_Review.csv in the ZIP.")
    with st.expander("Archive coverage and data-quality details"):
        st.caption("These file-level counts describe whole archived rotations, not just the selected date range. "
                   "The teaching totals below are date-filtered. Missing rotations are not assumed to have zero teaching.")
        if scan["duplicate_assignments_removed"]:
            st.write(f"Exact duplicate student-shifts removed across the whole archive: {scan['duplicate_assignments_removed']:,}.")
        st.dataframe(pd.DataFrame(scan["sources"]), hide_index=True, use_container_width=True)
        if scan["warnings"]:
            st.warning("Some cells were excluded: missing/nonclinical provider labels or missing '~' markers. "
                       "These archive-wide warnings may refer to dates outside your selected period.")
            st.dataframe(pd.DataFrame(scan["warnings"]), hide_index=True, use_container_width=True)
        if report_scan.get("clinical_shift_conflicts"):
            st.dataframe(pd.DataFrame(report_scan["clinical_shift_conflicts"]), hide_index=True, use_container_width=True)
        st.write(f"Repeated clinical listings combined across all archived rotations: {scan.get('duplicate_clinical_listings_removed', 0):,}.")
        if report_scan.get("work_type_conflicts"):
            st.dataframe(pd.DataFrame(report_scan["work_type_conflicts"]), hide_index=True, use_container_width=True)
        st.dataframe(pd.DataFrame([{"opd_site": site, "work_type": work_type}
                                  for site, work_type in report_scan["site_work_type_mapping"].items()]),
                     hide_index=True, use_container_width=True)
    if not annual:
        _clear_teaching_downloads()
        st.info("No recorded clinical shifts were found for the selected reporting period. "
                "Adjust the dates/year selection, check the '~' name order or archive the missing OPDs. "
                "Named providers with blank student fields are included when present.")
        return
    preview = pd.DataFrame(annual, columns=tuple(TEACHING_CSV_COLUMNS) + LEARNER_REACH_COLUMNS)
    a, b, c = st.columns(3)
    a.metric("Preceptors / provider labels", preview["preceptor_name"].nunique())
    b.metric("Assigned student-shifts", f"{preview['no_of_shifts'].sum():,}")
    c.metric("Student-weighted educational hours", f"{preview['educational_hours'].sum():,}")
    totals = reach_totals(annual)
    a, b, c, d = st.columns(4)
    a.metric("Recorded OPD hours", f"{totals['recorded_clinical_hours']:,}")
    b.metric("Hours with students", f"{totals['hours_with_students']:,}")
    c.metric("Hours without students", f"{totals['hours_without_students']:,}")
    d.metric("Learner Reach", reach_percent(totals["learner_reach_pct"]))
    st.caption(REACH_DEFINITION)
    with st.expander("What Learner Reach does and does not measure"):
        st.write(REACH_SCOPE_NOTE)
        st.write("Educational hours remain student-weighted; simultaneous students count twice only for that measure. "
                 "Clinical hours and Learner Reach count the half-day once. People with recorded shifts but no assignments appear with 0%.")
    st.markdown("**Teaching by type of work**")
    st.dataframe(pd.DataFrame(teaching_work_type_rows(report_scan, selected), columns=tuple(TEACHING_WORK_TYPE_CSV_COLUMNS) + LEARNER_REACH_COLUMNS),
                 hide_index=True, use_container_width=True)
    with st.expander("Overall totals and Learner Reach"):
        st.dataframe(preview, hide_index=True, use_container_width=True)
    st.caption("In Custom dates mode, academic_year contains your report label. All dates in that range stay in one "
               "report section, even across July. Boundary months include only the chosen days. "
               "Student names are not exported. Future scheduled assignments within the selected dates are included.")
    if st.button("Create teaching reports ZIP", key="teaching_build_zip"):
        st.session_state.pop("teaching_zip", None)
        try:
            with st.spinner("Creating the chair summary, CSVs and individual Word reports..."):
                zip_bytes, _ = teaching_build_zip(report_scan, selected)
                st.session_state["teaching_zip"] = zip_bytes
                st.session_state["teaching_zip_signature"] = signature
        except OPDArchiveError as exc:
            st.error(str(exc))
        except Exception:
            st.error("The Word/CSV export could not be completed. No partial ZIP was retained. "
                     "Check that python-docx is installed and retry.")
    if st.session_state.get("teaching_zip"):
        st.download_button("Download CSV + Word reports (ZIP)", data=st.session_state["teaching_zip"],
                           file_name=f"Preceptor_Teaching_{file_part}.zip", mime="application/zip", key="teaching_download_zip")
        with ZipFile(BytesIO(st.session_state["teaching_zip"])) as generated_zip:
            chair_bytes = generated_zip.read(TEACHING_CHAIR_SUMMARY_FILENAME)
        st.download_button("Download chair summary only (Word)", data=chair_bytes,
                           file_name=TEACHING_CHAIR_SUMMARY_FILENAME,
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                           key="teaching_download_chair_summary")
        st.caption("The ZIP includes the chair summary, overall/work-type CSVs with Learner Reach, monthly Learner Reach CSV, individual Word reports, source notes and "
                   "(for custom dates) the exact reporting-date settings. These are unencrypted staff reports; "
                   "do not commit the downloads to the public repository.")
