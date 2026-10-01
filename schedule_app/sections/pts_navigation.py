"""Lightweight cross-page preferences and password-checked navigation."""
import streamlit as st
from schedule_app.services.evaluation_access import protected_evaluation_callback


def preserve_pts_preferences():
    # Only widget values; never buttons, upload widgets or credentials. Assigning
    # a retained key to itself interrupts Streamlit's hidden-widget cleanup.
    for key in ("teaching_name_order", "teaching_selected_years", "teaching_oasis_include",
                "assessment_completion_enabled", "assessment_completion_course_choice",
                "pts_matching_task"):
        if key in st.session_state:
            st.session_state[key] = st.session_state[key]


@protected_evaluation_callback
def open_pts_matching():
    st.session_state["schedule_app_mode"] = "PTS Matching"


@protected_evaluation_callback
def return_to_pts():
    st.session_state["schedule_app_mode"] = "PTS"
