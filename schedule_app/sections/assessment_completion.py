"""Optional, non-blocking assessment/feedback completeness check."""
from __future__ import annotations
import hashlib
import json
import streamlit as st
from schedule_app.services.assessment_completion import (
    ASSESSMENT_VERSION, load_completion_inputs, build_completion_bundle,
    completion_context, completion_display, unverified_bundle,
)
from schedule_app.services.student_assessment_links import GitHubStudentAssessmentLinks, student_name_key
from schedule_app.services.opd_archive import OPDArchiveError

P = "assessment_completion_"


def _clear_downloads():
    for key in ("teaching_zip", "teaching_zip_signature"):
        st.session_state.pop(key, None)


def render_assessment_completion(archive, scan, years):
    st.markdown("**Student assessment completion and evaluation-record checks**")
    enabled = st.checkbox("Include assessment completion and missing-evaluation alerts", value=True, key=P + "enabled")
    if not enabled:
        return None
    signature = hashlib.sha256(json.dumps([ASSESSMENT_VERSION, archive.config.signature(),
                                          completion_context(scan, years)], sort_keys=True).encode()).hexdigest()
    if st.session_state.get(P + "scope") != signature:
        for key in ("inputs", "failure", "course_choice", "student_choice"):
            st.session_state.pop(P + key, None)
        st.session_state[P + "scope"] = signature
        _clear_downloads()
    st.caption("Two separate measures: Clinical Assessment of Student and History Taking & Physical Exam. "
               "Eligible = the same student assigned to the preceptor for 3+ distinct AM/PM shifts, not days. "
               "Completed forms use Submit Date within the same reporting period. Missing records only warn; they do not stop teaching reports.")
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
        bundle = unverified_bundle(scan, years, "Not checked: assessment data could not be verified" if failure else "Not checked: load evaluation completeness")
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
    if (st.session_state.get(LINKS_P + "include", False) and current_links is not None
            and current_links.get("sha") != inputs["catalog"].get("sha")):
        st.warning("Preceptor usernames/summary links changed. Refresh links and OASIS summaries above, then refresh evaluation completeness. Teaching reports remain available.")
        return unverified_bundle(scan, years, "Not checked: refresh after username/summary-link changes")
    try:
        bundle, unmatched = build_completion_bundle(inputs, scan, years, courses=selected)
    except OPDArchiveError as exc:
        st.warning(str(exc) + " Teaching reports remain available.")
        return unverified_bundle(scan, years, "Not checked: assessment selection needs review")
    st.caption(f"Student-assessment files checked: {bundle['source_count']} | Retrieved: {bundle['retrieved_at']}. "
               "Refresh this check after new OASIS uploads or saved username changes. It never modifies originals.")
    for direction, title in (("Student → educator", "Educators without verified student feedback"),
                             ("Preceptor → student", "Preceptors without verified completed student assessments")):
        issues = [r for r in bundle["warnings"] if r["direction"] == direction]
        if issues:
            st.warning(f"{title}: {len(issues)} record(s) to review. This is an alert, not a report-generation block.")
            with st.expander(title, expanded=True):
                st.dataframe(issues, hide_index=True, use_container_width=True)
        else:
            st.success(title.replace("without", "with") + ": no missing-record alerts.")
    st.dataframe([{"Preceptor": r["preceptor_name"], "Period": r["academic_year"],
                   "Students 3+ shifts": r["eligible_students_3plus_shifts"],
                   "Clinical Assessment": completion_display(r, "clinical"),
                   "History & Physical": completion_display(r, "hp"),
                   "Either form": completion_display(r, "either"), "Status": r["assessment_status"]}
                  for r in bundle["rows"]], hide_index=True, use_container_width=True)
    if unmatched:
        st.warning("Some eligible OPD students cannot be matched uniquely to Student External ID. "
                   "Their denominator has not been reduced. Only affected completion percentages are marked Not verified.")
    with st.expander("Resolve student external-ID matches (optional)"):
        st.caption("OPDs contain student names, not Student External ID. Exact name matches ignore case/spacing and "
                   "the OASIS '; MD2028'-style suffix. No fuzzy matching. Corrections are encrypted in GitHub. "
                   "Student names/IDs shown here are not added to the chair report, individual reports or report ZIP.")
        if unmatched:
            st.dataframe(unmatched, hide_index=True, use_container_width=True)
        choices = {r["name_key"]: r["student_name"] for r in unmatched}
        editing = st.checkbox("Review or correct saved student ID links", key=P + "edit_student_links")
        if editing:
            for key, row in inputs["student_links"]["entries"].items():
                choices[key] = row["student_name"]
        if choices:
            keys = sorted(choices)
            if st.session_state.get(P + "student_choice") not in keys:
                st.session_state[P + "student_choice"] = keys[0]
            key = st.selectbox("OPD student name needing an ID match", keys, key=P + "student_choice", format_func=lambda k: choices[k])
            old = inputs["student_links"]["entries"].get(key, {})
            suffix = hashlib.sha256((key + "|" + str(inputs["student_links"].get("sha"))).encode()).hexdigest()[:16]
            with st.form(P + "student_form_" + suffix):
                sid = st.text_input("Student External ID (not email or username unless it is the actual external ID)",
                                    value=old.get("external_id", ""), key=P + "external_id_" + suffix)
                confirm = st.checkbox("I verified that this Student External ID belongs to the selected OPD student", key=P + "confirm_id_" + suffix)
                save = st.form_submit_button("Save student ID link encrypted in GitHub")
            service = GitHubStudentAssessmentLinks(archive)
            if save:
                if not confirm:
                    st.warning("Confirm the student's external ID before saving.")
                else:
                    try:
                        inputs["student_links"] = service.save_link(choices[key], sid, expected=inputs["student_links"])
                        st.session_state[P + "inputs"] = inputs
                        _clear_downloads()
                        st.rerun()
                    except OPDArchiveError as exc:
                        st.warning(str(exc))
            if old:
                remove = st.checkbox("Confirm removal of this saved student ID link", key=P + "confirm_remove_" + suffix)
                if st.button("Remove saved student ID link", key=P + "remove", disabled=not remove):
                    try:
                        inputs["student_links"] = service.remove_link(choices[key], expected=inputs["student_links"])
                        _clear_downloads()
                        st.rerun()
                    except OPDArchiveError as exc:
                        st.warning(str(exc))
        else:
            st.info("No unresolved student names in the 3+ shift denominator.")
    if inputs["prepared"]["issues"]:
        with st.expander("Student-assessment source metadata to review"):
            st.caption("Question answers are ignored. These issues concern form identity, Student External ID, Evaluator Email or Submit Date. "
                       "No partial or guessed percentages are published for affected preceptors.")
            st.dataframe(inputs["prepared"]["issues"], hide_index=True, use_container_width=True)
    return bundle
