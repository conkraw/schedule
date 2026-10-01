"""Test-only entrypoint for retained legacy OASIS render functions.
The deployed launcher exposes only the new combined workflow.
"""
import streamlit as st
from importlib import import_module
st.set_page_config(page_title="Legacy OASIS module regression", layout="wide")
section = {
    "OASIS Evaluation Archive": "oasis_evaluation_archive",
    "OASIS Educator Reports": "oasis_educator_reports",
}[st.radio("Legacy section", ("OASIS Evaluation Archive", "OASIS Educator Reports"), key="schedule_app_mode")]
import_module("schedule_app.sections." + section).render()
