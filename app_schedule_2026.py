import streamlit as st
import csv
import datetime
import pandas as pd
import numpy as np
from datetime import timedelta
import xlsxwriter
import openpyxl
from openpyxl import Workbook
import io
from io import BytesIO, StringIO
import os
import time 
import random
from openpyxl.styles import Font, Alignment

# Configure page
st.set_page_config(
    page_title="OPD Creator",      # shown in browser tab
    layout="wide",                 # full‑width
    initial_sidebar_state="expanded"
)

# Main title
st.title("Outpatient Department (OPD) Schedule Creator")

# Sidebar
st.sidebar.title("Navigation")
page = st.sidebar.radio("Go to:", [
    "Home",
    "Create OPD",
    "Upload Files",
    "Generate Schedule",
    "Download OPD"
])

file_configs = {
    "HAMPDEN_NURSERY.xlsx": {"title": "HAMPDEN NURSERY","custom_text": "CUSTOM_PRINT","names": ["Folaranmi, Oluwamayoda", "Alur, Pradeep", "Nanda, Sharmilarani", "HAMPDEN_NURSERY"]},
    "SJR_HOSP.xlsx": {"title": "SJR HOSPITALIST","custom_text": "CUSTOM_PRINT","names": ["Spangola, Haley", "Gubitosi, Terry", "SJR_1", "SJR_2"]}, 
    "AAC.xlsx": {"title": "AAC","custom_text": "CUSTOM_PRINT","names": ["Vaishnavi Harding", "Abimbola Ajayi", "Shilu Joshi", "Desiree Webb", "Amy Zisa", "Abdullah Sakarcan", "Anna Karasik", "AAC_1", "AAC_2", "AAC_3"]},
    "AL.xlsx": {"title": "AL","custom_text": "CUSTOM_PRINT","names": ["Aholoukpe, Mahoussi"]}, 
}

# Example of reacting to the sidebar choice
if page == "Home":
    st.write("👋 Welcome! Use the sidebar to navigate through the app.")
# … after your sidebar radio and imports …

elif page == "Create OPD":
    st.header("Date Input for OPD")
    st.write("Enter start date in **m/d/yyyy** (no leading zeros, e.g. 7/6/2021):")

    date_input = st.text_input("Start Date")

    if st.button("Submit Date") and date_input:
        try:
            # parse and store in session
            start_date = datetime.datetime.strptime(date_input, "%m/%d/%Y")
            end_date   = start_date + datetime.timedelta(days=34)
            st.session_state.start_date = start_date
            st.session_state.end_date   = end_date

            st.success(
                f"✅ Valid date: {start_date:%B %d, %Y}  |  "
                f"Range: {start_date:%B %d, %Y} ➝ {end_date:%B %d, %Y}"
            )

            # generate all clinic workbooks
            generated = {}
            for fname, cfg in file_configs.items():
                path = generate_excel_file(
                    start_date,
                    cfg["title"],
                    cfg["custom_text"],
                    fname,
                    cfg["names"]
                )
                generated[fname] = path

            st.session_state.generated_files = generated

            # move on
            st.experimental_set_query_params(page="Upload Files")
            st.experimental_rerun()

        except ValueError:
            st.error("❌ Invalid format. Please use m/d/yyyy.")

elif page == "Upload Files":
    st.write("⬆️ Upload your Excel/CSV files to be processed.")
elif page == "Generate Schedule":
    st.write("⚙️ Processing files and assigning students...")
elif page == "Download OPD":
    st.write("👇 Finally, download your completed OPD.xlsx here.")
