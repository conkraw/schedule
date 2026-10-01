"""PTS is the reporting page; matching and exclusions live in PTS Matching.

Routine tables and chart images are opt-in. All calculation, diagnostic and
report outputs are retained. Data reuse is confined to this authenticated
session or a single ZIP build; no public/global cache of evaluations is used.
"""
from io import BytesIO
from time import perf_counter
from zipfile import ZipFile
import pandas as pd
import streamlit as st

from schedule_app.sections.pts_workspace import render_pts_workspace, _render_conflicts, _render_report_issues
from schedule_app.sections.pts_navigation import open_pts_matching
from schedule_app.sections.reporting_date_controls import clear_teaching_downloads as _clear_teaching_downloads
from schedule_app.sections.preceptor_oasis_links import render_teaching_oasis_links
from schedule_app.sections.assessment_completion import render_assessment_completion
from schedule_app.services.evaluation_access import require_evaluation_access, evaluation_access_is_valid, lock_evaluation_records
from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.teaching_analysis import teaching_annual_rows, teaching_work_type_rows
from schedule_app.services.teaching_validation import validate_teaching_report, TeachingConflictError
from schedule_app.services.learner_reach import reach_totals, reach_percent
from schedule_app.services.educational_time import teaching_time_rows, TIME_PREVIEW_LABELS
from schedule_app.services.teaching_priority import selected_priority_adjustments, outpatient_priority_audit_rows, PRIORITY_AUDIT_COLUMNS
from schedule_app.services.ignored_student_entries import GitHubIgnoredStudentEntries, require_matching_exclusions
from schedule_app.services.report_diagnostics import ReportDataError, REPORT_BUILD_ID
from schedule_app.services.assessment_completion import completion_signature
from schedule_app.services.teaching_evaluations import TEACHING_OASIS_REPORT_VERSION, feedback_signature, load_feedback_bundle
from schedule_app.reports.teaching_export import teaching_build_zip, teaching_csv_bytes
from schedule_app.settings import TEACHING_CHAIR_SUMMARY_FILENAME

<<<<<<< HEAD
PTS_REPORT_SCREEN_VERSION = "2026-10-01-feedback-shifts-assessed-totals-1"
=======
PTS_REPORT_SCREEN_VERSION = "2026-10-01-student-cohort-consistency-1"
>>>>>>> 706c8e4168ede05bba26315c4418e919cab7d537


def _optional_details(scan, report_scan, selected):
    """Only run after an explicit checkbox: an expander alone is not lazy."""
    with st.expander("Diagnostic tables and previews", expanded=True):
        st.caption("These tables are optional. The same detailed reports and source notes remain in the ZIP.")
        st.caption("Report builder: " + REPORT_BUILD_ID)
        st.dataframe(pd.DataFrame(teaching_time_rows(report_scan, selected)).rename(columns=TIME_PREVIEW_LABELS),
                     hide_index=True, use_container_width=True)
        st.dataframe(pd.DataFrame(teaching_time_rows(report_scan, selected, by_work_type=True)).rename(columns=TIME_PREVIEW_LABELS),
                     hide_index=True, use_container_width=True)
        st.dataframe(scan["sources"], hide_index=True, use_container_width=True)
        if scan["warnings"]:
            st.dataframe(scan["warnings"], hide_index=True, use_container_width=True)
        audit = outpatient_priority_audit_rows(report_scan, selected)
        if audit:
            st.dataframe(audit, hide_index=True, use_container_width=True)
            st.download_button("Download outpatient priority adjustments (CSV)",
                teaching_csv_bytes(audit, PRIORITY_AUDIT_COLUMNS),
                file_name="Outpatient_Priority_Adjustments.csv", mime="text/csv", key="teaching_priority_adjustments")


def _downloads(file_part):
    raw = st.session_state.get("teaching_zip")
    if not raw:
        st.session_state.pop("teaching_zip_payload", None)
        return
    # Extract the chair only once; do not read every chart on each widget rerun.
    payload = st.session_state.get("teaching_zip_payload")
    if payload is None or payload["zip"] is not raw:
        with ZipFile(BytesIO(raw)) as zf:
            payload = {"zip": raw, "chair": zf.read(TEACHING_CHAIR_SUMMARY_FILENAME),
                       "charts": sorted(name for name in zf.namelist()
                                        if name.startswith("Learner_Reach_Charts/") and name.endswith(".png"))}
        st.session_state["teaching_zip_payload"] = payload
    st.success("Reports ready. The ZIP includes the chair summary, individual reports, CSVs and clinical-experience pie charts.")
    st.download_button("Download CSV + Word reports (ZIP)", data=raw,
                       file_name=f"Preceptor_Teaching_{file_part}.zip", mime="application/zip", key="teaching_download_zip")
    st.download_button("Download chair summary only (Word)", data=payload["chair"],
                       file_name=TEACHING_CHAIR_SUMMARY_FILENAME,
                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                       key="teaching_download_chair_summary")
    if st.checkbox("Preview clinical-experience pie charts (optional)", value=False, key="pts_preview_charts"):
        with ZipFile(BytesIO(raw)) as zf:
            for chart in payload["charts"]:
                st.image(zf.read(chart), width=680)
    st.caption("Downloads are unencrypted staff reports and may contain evaluation comments. Do not upload them to the public repository.")


def render():
    if not require_evaluation_access(section_name="PTS", lock_key="evaluation_lock_pts"):
        return
    st.subheader("PTS")
    st.caption("Choose dates, load your archived OPDs, and create the reports. Usernames, student-name corrections and ignored entries are managed in PTS Matching.")
    st.button("Open PTS Matching", key="pts_open_matching", on_click=open_pts_matching)
    context = render_pts_workspace()
    if not context or context.get("report_scan") is None:
        return
    client, scan = context["client"], context["scan"]
    report_scan, selected = context["report_scan"], context["selected"]
    show_details = st.checkbox("Show diagnostic tables and report previews (optional)", value=False,
                               key="pts_show_diagnostics")
    try:
        validate_teaching_report(report_scan, selected)
        annual = teaching_annual_rows(report_scan, selected)
        if not annual:
            _clear_teaching_downloads()
            st.info("No student assignments were found for these dates after exclusions. Change the dates or use PTS Matching to restore an ignored entry.")
            return
        totals = reach_totals(annual)
    except TeachingConflictError as exc:
        _render_conflicts(exc)
        return
    except ReportDataError as exc:
        _render_report_issues(exc)
        return
    except OPDArchiveError as exc:
        _clear_teaching_downloads()
        st.error(str(exc))
        return
    a, b, c = st.columns(3)
    a.metric("Total scheduled availability (hours)", f"{totals['recorded_clinical_hours']:,}")
    b.metric("Educational hours", f"{totals['hours_with_students']:,}")
    c.metric("Learner Reach", reach_percent(totals["learner_reach_pct"]))
    st.caption("Includes weekends. One AM/PM shift with one or more students = four educational hours; simultaneous students do not multiply the time.")
    review = set(report_scan["unresolved_preceptor_labels"]) & {r["preceptor_name"] for r in annual}
    if review:
        st.warning(f"{len(review):,} provider label(s) need review. They remain separate from named preceptors in the reports.")
    if scan["warnings"]:
        st.warning(f"{len(scan['warnings']):,} source-data notice(s) are available in optional diagnostics and ZIP source notes. These cover the whole archive.")
    adjustments = selected_priority_adjustments(report_scan, selected)
    if adjustments:
        st.caption(f"Outpatient priority applied to {len(adjustments):,} Academic Pediatrics / PSHCH Nursery half-day overlap(s).")
    # PTS uses saved keys. It never renders the preceptor or student correction editors.
    plan, links_ready = render_teaching_oasis_links(client, report_scan, selected,
                                                   allow_missing_summaries=True,
                                                   manage_usernames=False, show_tables=show_details)
    completion = render_assessment_completion(client, report_scan, selected,
                                              manage_students=False, show_tables=show_details)
    report_scan = dict(report_scan, assessment_completion=completion)
    if show_details:
        _optional_details(scan, report_scan, selected)
    signature = (context["signature"], PTS_REPORT_SCREEN_VERSION, completion_signature(completion),
                 TEACHING_OASIS_REPORT_VERSION, feedback_signature(plan["bundle"]) if plan else None)
    if st.session_state.get("teaching_zip_signature") != signature:
        _clear_teaching_downloads()
        st.session_state["teaching_zip_signature"] = signature
    if not links_ready:
        _clear_teaching_downloads()
        st.info("Choose/refresh the OASIS summary link, or turn off linked evaluations for teaching-only reports. Name corrections belong in PTS Matching.")
    if st.button("Create teaching reports ZIP", key="teaching_build_zip", disabled=not links_ready):
        # Existing freshness checks still run. Reuse only an exact current build,
        # not a prior file for different dates, matches, thresholds or sources.
        old_zip = st.session_state.get("teaching_zip")
        bar = st.progress(0, text="Verifying saved settings and evaluation links...")
        started = perf_counter()
        try:
            require_matching_exclusions(report_scan, GitHubIgnoredStudentEntries(client).load())
            feedback = None if plan is None else load_feedback_bundle(
                client, report_scan, selected, plan["catalog"], plan["summaries"], allow_missing_summaries=True)
            if old_zip:
                zip_bytes = old_zip
            else:
                zip_bytes, _ = teaching_build_zip(
                    report_scan, selected, oasis_feedback=feedback,
                    progress=lambda done, total, text: bar.progress(min(1.0, done / max(total, 1)), text=text))
            if not evaluation_access_is_valid(touch=True):
                lock_evaluation_records()
                st.warning("Access expired during report generation. Unlock PTS and create the reports again.")
                return
            st.session_state["teaching_zip"] = zip_bytes
            st.session_state["teaching_zip_signature"] = signature
            st.session_state["teaching_build_seconds"] = perf_counter() - started
        except TeachingConflictError as exc:
            _render_conflicts(exc)
            return
        except ReportDataError as exc:
            _render_report_issues(exc)
            return
        except OPDArchiveError as exc:
            _clear_teaching_downloads()
            st.error(str(exc))
        except Exception:
            _clear_teaching_downloads()
            st.error("The report ZIP could not be completed. No partial ZIP was retained. Check installed update files and retry.")
        finally:
            bar.empty()
    _downloads(context["file_part"])
