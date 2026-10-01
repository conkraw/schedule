"""Opt-in diagnostic for a missing percentage; all identity editing stays in PTS Matching."""
from __future__ import annotations
import streamlit as st

from schedule_app.services.assessment_diagnostics import (
    DIAGNOSTICS_UI_VERSION, FOCUS_KEY, unavailable_completion_choices,
    completion_review_detail, make_student_review_focus,
)
from schedule_app.services.evaluation_access import (
    evaluation_access_is_valid, protected_evaluation_callback,
)
from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.student_assessment_links import GitHubStudentAssessmentLinks
from schedule_app.sections.reporting_date_controls import clear_teaching_downloads

P = "assessment_completion_diagnostics_"


@protected_evaluation_callback
def open_targeted_student_review(focus):
    # Navigation is a callback, before the sidebar widget is constructed again.
    st.session_state[FOCUS_KEY] = focus
    st.session_state["pts_matching_task"] = "Student names"
    st.session_state["schedule_app_mode"] = "PTS Matching"


def render_completion_diagnostics(archive, inputs, bundle, unmatched):
    if not evaluation_access_is_valid(touch=True):
        return
    choices = unavailable_completion_choices(bundle)
    notice = st.session_state.pop(P + "notice", None)
    if notice:
        st.success(notice)
    if not choices:
        return
    st.caption(f"Optional assessment review is available for {len(choices):,} preceptor/period result(s). "
               "No assessment on file does not require identity confirmation. Provisional results flag possible name discrepancies.")
    if not st.checkbox("Review documented completion or missing records (optional)",
                       value=False, key=P + "show"):
        return
    with st.expander("Check an assessment-completion result", expanded=True):
        options = list(choices)
        key = P + "preceptor"
        if st.session_state.get(key) not in choices:
            st.session_state[key] = options[0]
        selected = st.selectbox("Preceptor / reporting period to check", options, key=key,
            format_func=lambda token: (
                f"{choices[token]['preceptor_name']} | {choices[token]['academic_year']} | "
                f"{choices[token]['report_start_date']} to {choices[token]['report_end_date']}"))
        row = choices[selected]
        detail = completion_review_detail(row, unmatched)
        st.write("Result: " + str(detail["assessment_status"]))
        st.caption("Assessments as of: " + str(detail["assessments_as_of"]))
        if detail.get("either_students_without_assessment"):
            st.info(f"No assessment on file for {detail['either_students_without_assessment']} eligible student(s). They stay in the denominator. This is not an overdue judgment.")
        st.caption("Saved preceptor username: " + (detail["record_id"] or "Not assigned"))
        denominator = detail["eligible_students"]
        st.write("Eligible students: " + (str(denominator) if denominator is not None else "Not checked")
                 + (f" (minimum {detail['minimum_shifts']} AM/PM shifts)" if detail["minimum_shifts"] is not None else ""))
        clinical, hp = detail["clinical_forms_submitted"], detail["hp_forms_submitted"]
        st.write("Submitted forms found in the selected dates: Clinical Assessment — "
                 + (str(clinical) if clinical is not None else "Not checked / not verified")
                 + "; History & Physical — "
                 + (str(hp) if hp is not None else "Not checked / not verified"))
        st.caption("These are form counts, not the percentage numerator. The numerator counts only unique "
                   "eligible students assessed by this preceptor. Multiple forms or students below the "
                   "minimum cannot be divided directly by the eligible-student count.")
        missing = detail["unresolved_students"]
        if missing:
            st.write(f"These {len(missing)} eligible OPD name(s) still need a student match:")
            st.dataframe([{"OPD student name": item["student_name"],
                           "Assigned AM/PM shifts": item["assigned_shifts"],
                           "Issue": item["issue"]} for item in missing],
                         hide_index=True, use_container_width=True)
            st.button("Fix these student names in PTS Matching", key=P + "open_matching",
                      on_click=open_targeted_student_review,
                      args=(make_student_review_focus(bundle, row),))
            st.caption("These are possible name discrepancies, not proof of an identity. Confirm only a real match. "
                       "An absent OASIS student with no similar-name issue requires no action and remains in the denominator.")
        else:
            st.info("No eligible student-name corrections are required for this result. "
                    "Students who have not been evaluated do not need to be matched to another person. "
                    "Refresh after new OER uploads; review source/username issues only when the result is not checked.")
        if st.button("Recheck saved student matches from GitHub", key=P + "recheck"):
            try:
                saved = GitHubStudentAssessmentLinks(archive).load()
            except OPDArchiveError as exc:
                st.error(str(exc))
                st.caption("Existing inputs and saved matches were kept. A failed read is not an empty catalog.")
            else:
                st.session_state["assessment_completion_inputs"] = {**inputs, "student_links": saved}
                clear_teaching_downloads()
                st.session_state[P + "notice"] = (
                    "Saved student matches reloaded and verified. Completion results have been recalculated; "
                    "new OPD or OER source uploads still require Load / refresh evaluation completeness.")
                st.rerun()
        st.caption("Completion review: " + DIAGNOSTICS_UI_VERSION + ". This check does not change any source file.")
