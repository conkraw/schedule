"""Optional, non-blocking assessment/feedback completeness check."""
from __future__ import annotations
import hashlib
import json
import streamlit as st
from schedule_app.services.assessment_completion import (
    ASSESSMENT_VERSION, load_completion_inputs, build_completion_bundle,
    completion_context, completion_display, unverified_bundle,
)
from datetime import date
from schedule_app.services.assessment_progress import assessment_as_of, REVIEW, ABSENT
from schedule_app.services.teaching_analysis import teaching_local_today
from schedule_app.sections.reporting_date_controls import clear_teaching_downloads
from schedule_app.sections.assessment_settings import render_assessment_threshold
from schedule_app.sections.student_name_matches import render_student_name_matches
from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.assessment_diagnostics import DIAGNOSTICS_UI_VERSION
from schedule_app.sections.assessment_diagnostics import render_completion_diagnostics

P = "assessment_completion_"


def _clear_downloads():
    for key in ("teaching_zip", "teaching_zip_signature"):
        st.session_state.pop(key, None)


def render_assessment_completion(archive, scan, years, *, manage_students=True,
                                 show_tables=True, matching_only=False):
    st.markdown("**Student assessment completion and evaluation-record checks**")
    enabled = True if matching_only else st.checkbox("Include assessment completion and missing-evaluation alerts", value=True, key=P + "enabled")
    if not enabled:
        return None
    today = teaching_local_today()
    date_key = P + "as_of"
    if date_key not in st.session_state:
        st.session_state[date_key] = today
    as_of = st.date_input("Assessments as of (included)", key=date_key,
                         min_value=date(1970, 1, 1), max_value=today, format="MM/DD/YYYY",
                         on_change=clear_teaching_downloads,
                         help="Only scheduled shifts and submitted student assessments on or before this date count toward completion. Teaching hours and linked educator-feedback periods do not change.")
    if as_of is None:
        clear_teaching_downloads()
        st.info("Select an Assessments as of date to calculate completion. No date has been assumed.")
        return None
    as_of = assessment_as_of(as_of)
    minimum_shifts = render_assessment_threshold(archive)
    if minimum_shifts is None:
        return unverified_bundle(scan, years, "Not checked: minimum-shifts setting not verified", minimum_shifts=None, as_of=as_of)
    # The loaded source records do not depend on the threshold. Keep them when
    # this setting changes; recompute denominators and percentages below.
    signature = hashlib.sha256(json.dumps([ASSESSMENT_VERSION, DIAGNOSTICS_UI_VERSION, archive.config.signature(),
                                          {k: v for k, v in completion_context(scan, years, as_of=as_of).items()
                                           if k != "assessments_as_of"}], sort_keys=True).encode()).hexdigest()
    if st.session_state.get(P + "scope") != signature:
        for key in ("inputs", "failure", "course_choice", "student_choice"):
            st.session_state.pop(P + key, None)
        st.session_state[P + "scope"] = signature
        _clear_downloads()
    st.caption("Two separate measures: Clinical Assessment of Student and History Taking & Physical Exam. "
               f"Eligible = the same student assigned to the preceptor for {minimum_shifts}+ distinct AM/PM shifts, not days. "
               "Completed forms use Submit Date through the cutoff above. No assessment on file is a progress notice, not overdue; it does not require a name confirmation. Teaching-hour dates are unchanged.")
    if st.button("Load / refresh evaluation completeness", key=P + "refresh"):
        st.session_state.pop(P + "inputs", None)
        st.session_state.pop(P + "failure", None)
        _clear_downloads()
        try:
            with st.spinner("Checking archived student assessments, retained OPD assignments, usernames and educator feedback..."):
                st.session_state[P + "inputs"] = load_completion_inputs(archive, scan, years)
        except OPDArchiveError as exc:
            st.session_state[P + "failure"] = str(exc)
        except Exception:
            st.session_state[P + "failure"] = "The assessment check could not be verified. Refresh or review installed update files; no partial percentages are used."
    inputs = st.session_state.get(P + "inputs")
    if inputs is None:
        failure = st.session_state.get(P + "failure")
        if failure:
            st.warning(failure + " Teaching reports remain available.")
        else:
            st.info("Click Load / refresh evaluation completeness to check both directions. Until then, the reports mark these measures as not checked.")
        bundle = unverified_bundle(scan, years, "Not checked: assessment data could not be verified" if failure else "Not checked: load evaluation completeness", minimum_shifts=minimum_shifts, as_of=as_of)
        if show_tables:
            with st.expander("Preceptors with evaluation records not yet checked"):
                st.dataframe(bundle["warnings"], hide_index=True, use_container_width=True)
        return bundle
    courses = inputs["prepared"]["courses"]
    if len(courses) > 1:
        if P + "course_choice" not in st.session_state:
            st.session_state[P + "course_choice"] = []
        selected = st.multiselect("Student-assessment course(s) for these OPD assignments", courses, key=P + "course_choice")
        if not selected:
            st.warning("Choose the matching clerkship course(s). No course is guessed and no percentages are calculated from all courses by default.")
    else:
        selected = courses
    # A username edited above invalidates the cached match; no network request on
    # every widget interaction. The next explicit refresh reads the new catalog.
    current_links = st.session_state.get("teaching_oasis_catalog")
    # The link editor uses its own prefix; read it without changing that module.
    from schedule_app.sections.preceptor_oasis_links import P as LINKS_P
    current_links = st.session_state.get(LINKS_P + "catalog", current_links)
    if (current_links is not None and current_links.get("scope") == inputs["catalog"].get("scope")
            and current_links.get("sha") != inputs["catalog"].get("sha")):
        st.warning("Preceptor usernames/summary links changed in PTS Matching. Click Load / refresh evaluation completeness to use the verified new links. Teaching reports remain available.")
        return unverified_bundle(scan, years, "Not checked: refresh after username/summary-link changes", minimum_shifts=minimum_shifts, as_of=as_of)
    try:
        bundle, unmatched = build_completion_bundle(inputs, scan, years, courses=selected, minimum_shifts=minimum_shifts, as_of=as_of)
    except OPDArchiveError as exc:
        st.warning(str(exc) + " Teaching reports remain available.")
        return unverified_bundle(scan, years, "Not checked: assessment selection needs review", minimum_shifts=minimum_shifts, as_of=as_of)
    st.caption(f"Student-assessment files checked: {bundle['source_count']} | Retrieved: {bundle['retrieved_at']}. "
               "Refresh this check after new OASIS uploads or saved username changes. It never modifies originals.")
    if manage_students:
        render_student_name_matches(archive, inputs, scan, years, unmatched, minimum_shifts=minimum_shifts,
                                    show_tables=show_tables, as_of=as_of)
    else:
        from schedule_app.services.student_name_review import student_name_review
        review = student_name_review(inputs, scan, years, unmatched, as_of=as_of)
        if review["missing"]:
            st.warning(f"Possible student-name discrepancies: {len(review['missing']):,}. Open PTS Matching → Student names when needed. "
                       f"{review['eligible_missing_count']:,} name(s) affect completion; those numeric results are labeled provisional.")
        if review["absent"]:
            st.info(f"{len(review['absent']):,} OPD student name(s) are not in the loaded OASIS records. No name confirmation is required. Eligible students remain in the denominator with no documented assessment.")
    if matching_only:
        return bundle
    no_record_rows = [r for r in bundle["rows"] if r.get("either_students_without_assessment")]
    if no_record_rows:
        st.info(f"No assessment on file for some eligible students in {len(no_record_rows):,} preceptor/period result(s). Percentages still calculate; these assessments are not necessarily overdue.")
    render_completion_diagnostics(archive, inputs, bundle, unmatched)
    for direction, title in (("Student → educator", "Educators without verified student feedback"),
                             ("Preceptor → student", "Preceptors without verified completed student assessments")):
        issues = [r for r in bundle["warnings"] if r["direction"] == direction]
        if issues:
            st.warning(f"{title}: {len(issues)} record(s) to review. This is an alert, not a report-generation block.")
            if show_tables:
                with st.expander(title, expanded=False):
                    st.dataframe(issues, hide_index=True, use_container_width=True)
        else:
            st.success(title.replace("without", "with") + ": no missing-record alerts.")
    if show_tables:
        st.dataframe([{"Preceptor": r["preceptor_name"], "Period": r["academic_year"],
                       f"Students {minimum_shifts}+ shifts": r["eligible_students"],
                       "Clinical Assessment": completion_display(r, "clinical"),
                       "History & Physical": completion_display(r, "hp"),
                       "Either form": completion_display(r, "either"), "Status": r["assessment_status"]}
                      for r in bundle["rows"]], hide_index=True, use_container_width=True)
    if inputs["prepared"]["issues"]:
        st.warning(f"Student-assessment source issues: {len(inputs['prepared']['issues']):,}. "
                   "Details are available in the optional diagnostics; only affected results are unverified.")
    if inputs["prepared"]["issues"] and show_tables:
        with st.expander("Student-assessment source metadata to review"):
            st.caption("Question answers are ignored. These issues concern form identity, Student External ID, Evaluator Email or Submit Date. "
                       "No partial or guessed percentages are published for affected preceptors.")
            st.dataframe(inputs["prepared"]["issues"], hide_index=True, use_container_width=True)
    return bundle
