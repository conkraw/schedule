"""Pediatric clerkship schedule app entrypoint.

Run: streamlit run app_sch_2026.py
Edit a section in schedule_app/sections/, not this launcher.
Editable mappings are in schedule_app/settings.py.
"""

from importlib import import_module
import streamlit as st

st.set_page_config(page_title="PSUCOM PEDIATRIC CLERKSHIP SCHEDULE CREATOR", layout="wide")
st.title("PSUCOM PEDIATRIC CLERKSHIP SCHEDULE CREATOR")

# Keep non-OASIS labels/order and the original session key. Archive navigation callbacks
# use schedule_app_mode to select Create Student Schedule.
SECTIONS = {
    "Instructions": "instructions",
    "Format OPD + Summary": "format_opd_summary",
    "Create Student Schedule": "create_student_schedule",
    "OPD Check": "opd_check",
    "Create Individual Schedules": "create_individual_schedules",
    "OPD Archive": "opd_archive",
    "Evaluation Records": "oasis_workflow",
    "Preceptor Teaching Summary": "preceptor_teaching_summary",
    "OPD MD PA Conflict Detector": "opd_md_pa_conflict_detector",
    "Shift Availability Tracker": "shift_availability_tracker",
}

# Migrate an open session from either former OASIS screen before creating the widget.
if st.session_state.get("schedule_app_mode") in ("OASIS Evaluation Archive", "OASIS Educator Reports", "OASIS Evaluations"):
    st.session_state["schedule_app_mode"] = "Evaluation Records"

mode = st.sidebar.radio(
    "What do you want to do?", tuple(SECTIONS), key="schedule_app_mode"
)

# Import only the chosen section. render() runs on every normal Streamlit rerun;
# uploading a file or clicking a button does not rely on re-importing a module.
section = import_module(f"schedule_app.sections.{SECTIONS[mode]}")
section.render()
