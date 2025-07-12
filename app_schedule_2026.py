import streamlit as st
import datetime
from datetime import timedelta
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
import io
from io import BytesIO, StringIO
import os
import random

# ── Page setup ──
st.set_page_config(
    page_title="OPD Creator",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Configuration ──
DEFAULT_CUSTOM_TEXT = "CUSTOM_PRINT"
file_configs = {
    "HAMPDEN_NURSERY.xlsx": {
        "title": "HAMPDEN NURSERY",
        "custom_text": DEFAULT_CUSTOM_TEXT,
        "names": [
            "Folaranmi, Oluwamayoda", "Alur, Pradeep",
            "Nanda, Sharmilarani", "HAMPDEN_NURSERY"
        ]
    },
    "SJR_HOSP.xlsx": {
        "title": "SJR HOSPITALIST",
        "custom_text": DEFAULT_CUSTOM_TEXT,
        "names": ["Spangola, Haley", "Gubitosi, Terry", "SJR_1", "SJR_2"]
    },
    "AAC.xlsx": {
        "title": "AAC",
        "custom_text": DEFAULT_CUSTOM_TEXT,
        "names": [
            "Vaishnavi Harding", "Abimbola Ajayi", "Shilu Joshi",
            "Desiree Webb", "Amy Zisa", "Abdullah Sakarcan",
            "Anna Karasik", "AAC_1", "AAC_2", "AAC_3"
        ]
    },
    "AL.xlsx": {
        "title": "AL",
        "custom_text": DEFAULT_CUSTOM_TEXT,
        "names": ["Aholoukpe, Mahoussi"]
    },
}

def generate_excel_file(start_date, title, custom_text, file_name, names):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"], ws["A2"] = title, custom_text

    custom_value_cols = ["A","C","E","G","I","K","M"]
    name_cols         = ["B","D","F","H","J","L","N"]
    days = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
    start_row = 4
    num_weeks  = 5
    week_h     = 13

    for w in range(num_weeks):
        base = start_date + datetime.timedelta(weeks=w)
        # write days & dates
        for i, day in enumerate(days):
            col = chr(65 + i*2)
            ws[f"{col}{start_row}"]     = day
            ws[f"{col}{start_row+1}"]   = (base + datetime.timedelta(days=i))\
                                          .strftime("%B %-d, %Y")
        # write names + custom_value
        name_start = start_row+2
        for i, col in enumerate(name_cols):
            for j, nm in enumerate(names or ["Default Name ~"]):
                r = name_start + j
                ws[f"{col}{r}"]               = nm
                ws[f"{custom_value_cols[i]}{r}"] = "custom_value"
        # fill leftover with custom_value
        next_start = start_row + week_h
        for col in custom_value_cols:
            for r in range(start_row+1, next_start):
                if r >= name_start + len(names):
                    ws[f"{col}{r}"] = "custom_value"
        start_row = next_start

    wb.save(file_name)
    return file_name

# ── Sidebar Navigation ──
st.title("Outpatient Department (OPD) Schedule Creator")
page = st.sidebar.radio(
    "Go to:",
    ["Home","Create OPD","Upload Files","Generate Schedule","Download OPD"],
    index=0
)

# ── Page Logic ──
if page == "Home":
    st.write("👋 Welcome! Use the sidebar to navigate through the app.")

elif page == "Create OPD":
    st.header("Date Input for OPD")
    st.write("Enter start date in **m/d/yyyy** (no leading zeros, e.g. 7/6/2021):")
    date_input = st.text_input("Start Date")

    if st.button("Submit Date") and date_input:
        try:
            sd = datetime.datetime.strptime(date_input, "%m/%d/%Y")
            ed = sd + timedelta(days=34)
            st.session_state.start_date = sd
            st.session_state.end_date   = ed

            st.success(f"✅ {sd:%B %d, %Y} → {ed:%B %d, %Y}")

            # generate workbooks
            generated = {}
            for fname, cfg in file_configs.items():
                path = generate_excel_file(
                    sd,
                    cfg["title"],
                    cfg["custom_text"],
                    fname,
                    cfg["names"]
                )
                generated[fname] = path
            st.session_state.generated_files = generated

            # move to next step
            st.session_state.page = "Upload Files"
            st.experimental_rerun()

        except ValueError:
            st.error("❌ Invalid format. Please use m/d/yyyy.")

elif page == "Upload Files":
    st.header("Upload Files")
    st.write("⬆️ Upload your Excel/CSV files to be processed here.")

elif page == "Generate Schedule":
    st.header("Generate Schedule")
    st.write("⚙️ Processing files and assigning students…")

elif page == "Download OPD":
    st.header("Download OPD")
    st.write("👇 Download your completed `OPD.xlsx` here.")
