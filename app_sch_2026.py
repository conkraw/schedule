"""Pediatric clerkship schedule app entrypoint.

Run: streamlit run app_sch_2026.py
Edit a section in schedule_app/sections/, not this launcher.
Editable mappings are in schedule_app/settings.py.
"""

from importlib import import_module
import streamlit as st

st.set_page_config(page_title="PSUCOM PEDIATRIC CLERKSHIP SCHEDULE CREATOR", layout="wide")
st.title("PSUCOM PEDIATRIC CLERKSHIP SCHEDULE CREATOR")

# Keep the original labels, order and session key. Archive navigation callbacks
# use schedule_app_mode to select Create Student Schedule.
SECTIONS = {
    "Instructions": "instructions",
    "Format OPD + Summary": "format_opd_summary",
    "Create Student Schedule": "create_student_schedule",
    "OPD Check": "opd_check",
    "Create Individual Schedules": "create_individual_schedules",
    "OPD Archive": "opd_archive",
    "OASIS Evaluation Archive": "oasis_evaluation_archive",
    "OASIS Educator Reports": "oasis_educator_reports",
    "Preceptor Teaching Summary": "preceptor_teaching_summary",
    "OPD MD PA Conflict Detector": "opd_md_pa_conflict_detector",
    "Shift Availability Tracker": "shift_availability_tracker",
}

mode = st.sidebar.radio(
    "What do you want to do?", tuple(SECTIONS), key="schedule_app_mode"
)

# Import only the chosen section. render() runs on every normal Streamlit rerun;
# uploading a file or clicking a button does not rely on re-importing a module.
section = import_module(f"schedule_app.sections.{SECTIONS[mode]}")
section.render()
