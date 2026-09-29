"""Sidebar: teaching effort with exact, editable reporting periods.

Date presets are saved by a separate controls module. OPD archive writes,
student schedules, report calculations and the Power Automate workbook are unchanged.
"""
from datetime import date
from io import BytesIO
from zipfile import ZipFile
import hashlib
import json
import pandas as pd
import streamlit as st

from schedule_app.reports.teaching_export import teaching_build_zip, teaching_csv_bytes
from schedule_app.services.teaching_validation import (
    STRICT_CONFLICT_SCHEMA_VERSION, STRICT_REPORT_VERSION, CONFLICT_CSV_COLUMNS,
    TeachingConflictError, validate_teaching_report, require_conflict_source_data,
)
from schedule_app.services.learner_reach import (
    LEARNER_REACH_SCHEMA_VERSION, LEARNER_REACH_COLUMNS, REACH_DEFINITION,
    REACH_SCOPE_NOTE, require_learner_reach_data, reach_percent, reach_totals,
    PARTICIPATION_REPORT_VERSION, PARTICIPATION_SCOPE_NOTE, REACH_DETAIL_TOTAL_NOTE,
    teaching_participation_keys, reach_group_year,
)
from schedule_app.services.opd_archive import GitHubOPDArchive, OPDArchiveError, get_opd_archive_config
from schedule_app.services.reporting_periods import (
    DATE_RANGE_SCHEMA_VERSION, teaching_report_date_text,
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

from schedule_app.services.teaching_priority import (
    OUTPATIENT_PRIORITY_VERSION, OUTPATIENT_PRIORITY_NOTE, PRIORITY_AUDIT_COLUMNS,
    selected_priority_adjustments, outpatient_priority_audit_rows,
)

from schedule_app.sections.reporting_date_controls import (
    REPORTING_MODES, PERIOD_FIELDS,
    render_period_controls as _render_period_controls,
    remember_period_inputs as _remember_period_inputs,
    clear_teaching_downloads as _clear_teaching_downloads,
)


def _render_conflicts(exc):
    _clear_teaching_downloads()
    st.error(str(exc))
    st.write("Each conflict ID groups all source cells for one preceptor/date/AM-or-PM. "
             "A YES or NO only indicates whether a student is recorded in that cell; student names are omitted.")
    columns = ["conflict_id", "preceptor_name", "date", "shift", "conflicting_work_types",
               "rotation_start", "archive_file", "worksheet", "cell", "listed_work_type", "student_assigned"]
    st.dataframe(pd.DataFrame(exc.rows).reindex(columns=columns), hide_index=True, use_container_width=True)
    st.download_button("Download OPD conflicts to correct (CSV)",
        data=teaching_csv_bytes(exc.rows, CONFLICT_CSV_COLUMNS),
        file_name="OPD_Conflict_Review.csv", mime="text/csv", key="teaching_download_conflicts")
    st.info("Open OPD Archive, select the listed rotation and download its original workbook. "
            "Review the listed worksheets/cells and correct which clinical experience owns the shift. "
            "Re-upload the corrected OPD in Create Student Schedule so it replaces that rotation's current archive. "
            "Then return here and click Load / refresh archived OPDs. No teaching report or pie chart will run until conflicts in the selected dates are resolved.")


def render():
    st.subheader("Preceptor Teaching Summary")
    st.write("Read current encrypted OPDs from GitHub, then generate a chair-friendly Word summary, "
             "overall and work-type CSVs, and one Word teaching report per preceptor.")
    st.caption(OUTPATIENT_PRIORITY_NOTE)
    st.caption("After that exception, reports are blocked if a preceptor/date/AM-or-PM appears in different clinical experiences in the selected dates. "
               "The issue table identifies the exact archived OPDs and source cells. Every valid report includes numeric Learner Reach percentages and clinical-experience pie charts.")
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
        [TEACHING_REPORT_VERSION, DATE_RANGE_SCHEMA_VERSION, LEARNER_REACH_SCHEMA_VERSION, STRICT_CONFLICT_SCHEMA_VERSION, OUTPATIENT_PRIORITY_VERSION, client.config.signature(), order,
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
        st.caption("No OPD is downloaded or decrypted until you click Load / refresh archived OPDs. "
                   "Saved date presets load separately from GitHub. "
                   "This section never uploads a report or decrypted OPD to GitHub.")
        return
    try:
        teaching_require_date_range_data(scan)
        require_learner_reach_data(scan)
        require_conflict_source_data(scan)
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
    signature = (options_signature, scan["commit"], selection_signature, PARTICIPATION_REPORT_VERSION, STRICT_REPORT_VERSION)
    if st.session_state.get("teaching_zip_signature") != signature:
        _clear_teaching_downloads()
        st.session_state["teaching_zip_signature"] = signature
    adjustments = selected_priority_adjustments(report_scan, selected)
    if adjustments:
        st.info(f"Academic Pediatrics took priority over PSHCH Nursery for {len(adjustments):,} "
                "overlapping clinical half-day(s) in these dates. Nursery hours and nursery student "
                "assignments were excluded only for those half-days; the source OPDs were not changed.")
        with st.expander("Outpatient priority adjustments (not conflicts)"):
            audit_rows = outpatient_priority_audit_rows(report_scan, selected)
            st.caption("One adjustment ID groups all relevant cells for one preceptor/date/AM-or-PM. "
                       "Student-assigned YES/NO describes the original cell, not teaching credit. "
                       "An unassigned clinic shift stays unassigned even if a nursery student was listed.")
            st.dataframe(pd.DataFrame(audit_rows).reindex(columns=PRIORITY_AUDIT_COLUMNS),
                         hide_index=True, use_container_width=True)
            st.download_button("Download outpatient priority adjustments (CSV)",
                               teaching_csv_bytes(audit_rows, PRIORITY_AUDIT_COLUMNS),
                               file_name="Outpatient_Priority_Adjustments.csv", mime="text/csv",
                               key="teaching_priority_adjustments")
    try:
        validate_teaching_report(report_scan, selected)
    except TeachingConflictError as exc:
        _render_conflicts(exc)
        return
    except OPDArchiveError as exc:
        _clear_teaching_downloads()
        st.error(str(exc))
        return
    annual = teaching_annual_rows(report_scan, selected)
    eligible = teaching_participation_keys(report_scan, selected)
    active_names = {name for name, _, _ in eligible}
    typed = teaching_work_type_rows(report_scan, selected)
    active_sites = {site for row in typed for site in row.get("source_sites", "").split("; ") if site}
    review_names = [name for name in report_scan["unresolved_preceptor_labels"] if name in active_names]
    if review_names:
        st.warning("Some provider labels identify a site/slot/combined entry rather than a person. "
                   "They remain literal and are separated from named preceptors in the Word reports.")
        with st.expander("Provider labels to review"):
            st.write(", ".join(review_names))
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
        st.write(f"Repeated clinical listings combined across all archived rotations: {scan.get('duplicate_clinical_listings_removed', 0):,}.")
        if report_scan.get("work_type_conflicts"):
            st.dataframe(pd.DataFrame(report_scan["work_type_conflicts"]), hide_index=True, use_container_width=True)
        st.dataframe(pd.DataFrame([{"opd_site": site, "work_type": work_type}
                                  for site, work_type in report_scan["site_work_type_mapping"].items() if site in active_sites]),
                     hide_index=True, use_container_width=True)
    if not annual:
        _clear_teaching_downloads()
        st.info("No student assignments were found for the selected reporting period. "
                "Preceptors and services with only unassigned shifts are omitted. "
                "Adjust the dates/year selection, check the '~' name order or archive the missing OPDs.")
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
    st.caption(PARTICIPATION_SCOPE_NOTE)
    if sum(row["recorded_clinical_hours"] for row in typed) != totals["recorded_clinical_hours"]:
        st.caption(REACH_DETAIL_TOTAL_NOTE)
    st.caption(REACH_DEFINITION)
    with st.expander("What Learner Reach does and does not measure"):
        st.write(REACH_SCOPE_NOTE)
        st.write("Educational hours remain student-weighted; simultaneous students count twice only for that measure. "
                 "Clinical hours and Learner Reach count the half-day once. Wholly unassigned preceptors/categories are not listed; included preceptors retain their non-teaching shifts in the denominator.")
    st.markdown("**Teaching by type of work**")
    st.dataframe(pd.DataFrame(typed, columns=tuple(TEACHING_WORK_TYPE_CSV_COLUMNS) + LEARNER_REACH_COLUMNS),
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
        except TeachingConflictError as exc:
            _render_conflicts(exc)
            return
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
            chart_files = sorted(name for name in generated_zip.namelist()
                                 if name.startswith("Learner_Reach_Charts/") and name.endswith(".png"))
            with st.expander("Learner Reach pies by clinical experience"):
                st.caption("Each pie compares recorded hours with students versus without students in that experience. "
                           "These are the same category percentages shown in the chair summary, not student-weighted hours.")
                for chart_file in chart_files:
                    st.image(generated_zip.read(chart_file), width=680)
        st.download_button("Download chair summary only (Word)", data=chair_bytes,
                           file_name=TEACHING_CHAIR_SUMMARY_FILENAME,
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                           key="teaching_download_chair_summary")
        st.caption("The ZIP includes the chair summary with clinical-experience pies, chart PNGs and their CSV, overall/work-type CSVs with Learner Reach, monthly Learner Reach CSV, individual Word reports, source notes and "
                   "(for custom dates) the exact reporting-date settings. These are unencrypted staff reports; "
                   "do not commit the downloads to the public repository.")
