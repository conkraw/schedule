"""Pediatric clerkship schedule app entrypoint.

Run: streamlit run app_sch_2026.py
Editable mappings are in schedule_app/settings.py.
"""

from importlib import import_module
import streamlit as st

st.set_page_config(page_title="PSUCOM PEDIATRIC CLERKSHIP SCHEDULE CREATOR", layout="wide")
st.title("PSUCOM PEDIATRIC CLERKSHIP SCHEDULE CREATOR")

# Keep OER and PTS as the last two choices.
# Preserve the session key used by the existing archive navigation callbacks.
SECTIONS = {
    "Instructions": "instructions",
    "Format OPD + Summary": "format_opd_summary",
    "Create Student Schedule": "create_student_schedule",
    "OPD Check": "opd_check",
    "Create Individual Schedules": "create_individual_schedules",
    "OPD Archive": "opd_archive",
    "OPD MD PA Conflict Detector": "opd_md_pa_conflict_detector",
    "Shift Availability Tracker": "shift_availability_tracker",
    "PTS Matching": "pts_matching",
    "OER": "oasis_workflow",
    "PTS": "preceptor_teaching_summary",
}

# Migrate old labels before creating the sidebar widget (including open sessions).
LEGACY_SECTIONS = {"Evaluation Records": "OER", "OASIS Evaluation Archive": "OER",
                   "OASIS Educator Reports": "OER", "OASIS Evaluations": "OER",
                   "Preceptor Teaching Summary": "PTS"}
previous_mode = st.session_state.get("schedule_app_mode")
if previous_mode in LEGACY_SECTIONS:
    st.session_state["schedule_app_mode"] = LEGACY_SECTIONS[previous_mode]

# Preserve only non-secret PTS widget choices when navigating away. Streamlit
# normally removes widget state for controls not rendered on the next page.
from schedule_app.sections.pts_navigation import preserve_pts_preferences
preserve_pts_preferences()

mode = st.sidebar.radio(
    "What do you want to do?", tuple(SECTIONS), key="schedule_app_mode"
)

# Import only the chosen section. render() runs on every normal Streamlit rerun;
# uploading a file or clicking a button does not rely on re-importing a module.
section = import_module(f"schedule_app.sections.{SECTIONS[mode]}")
section.render()
