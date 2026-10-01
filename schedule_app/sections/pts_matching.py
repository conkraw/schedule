"""Password-protected preceptor/student matching, separate from PTS reporting."""
import streamlit as st
from schedule_app.services.evaluation_access import require_evaluation_access
from schedule_app.sections.pts_workspace import render_pts_workspace
from schedule_app.sections.pts_navigation import return_to_pts
from schedule_app.sections.preceptor_oasis_links import render_preceptor_username_management
from schedule_app.sections.assessment_completion import render_assessment_completion
from schedule_app.sections.ignored_student_entries import render_ignored_student_entries


def render():
    # Gate before controls, archive reads, student names or downloads.
    if not require_evaluation_access(section_name="PTS Matching", lock_key="evaluation_lock_pts_matching"):
        return
    st.subheader("PTS Matching")
    st.caption("Correct only missing usernames or unmatched student names. Saved changes are encrypted in GitHub and reused in PTS.")
    st.button("Return to PTS reports", key="pts_return_to_reports", on_click=return_to_pts)
    task = st.selectbox("What needs updating?", ("Preceptor usernames", "Student names", "Ignored student entries"),
                        key="pts_matching_task")
    context = render_pts_workspace()
    if not context:
        return
    if task == "Ignored student entries":
        render_ignored_student_entries(context["client"], context["order"], period=context.get("period"))
        return
    scan = context.get("report_scan")
    if scan is None:
        return
    if task == "Preceptor usernames":
        render_preceptor_username_management(context["client"], scan, context["selected"])
    else:
        st.caption("Select the OPD name and corresponding OASIS name only. (MD), (PA), (DO) and recognized class labels are handled automatically when the remaining name is unambiguous.")
        render_assessment_completion(context["client"], scan, context["selected"],
                                     manage_students=True, show_tables=False, matching_only=True)
    st.caption("After correcting names, return to PTS, refresh evaluation completeness when prompted, and create a fresh report ZIP. Original OPDs and OASIS files are unchanged.")
