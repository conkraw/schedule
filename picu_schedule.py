import io
import streamlit as st
import pandas as pd
import numpy as np 
import re
import xlsxwriter
import random
from openpyxl import load_workbook # Ensure load_workbook is imported
import io, zipfile
from docx import Document
from docx.shared import Pt
from docx.enum.section import WD_ORIENT
from datetime import timedelta
from xlsxwriter import Workbook as Workbook
from collections import defaultdict
from datetime import datetime, timedelta
from collections import Counter

st.set_page_config(page_title="Batch Preceptor → REDCap Import", layout="wide")
st.title("Batch Preceptor → REDCap Import Generator")

# ─── Sidebar mode selector ─────────────────────────────────────────────────────
mode = st.sidebar.radio("What do you want to do?", ("Format OPD + Summary",))
# ─── Sidebar mode selector ─────────────────────────────────────────────────────

if mode == "Format OPD + Summary":
    # ─── Inputs ────────────────────────────────────────────────────────────────────
    required_keywords = ["department of pediatrics"]
    found_keywords = set()
    
    schedule_files = st.file_uploader(
        "1) Upload one or more QGenda calendar Excel(s)",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )
    
    if schedule_files:
        for file in schedule_files:
            try:
                df = pd.read_excel(file, sheet_name=0, header=None)
                cell_values = df.astype(str)\
                                .apply(lambda x: x.str.lower())\
                                .values.flatten().tolist()
                for keyword in required_keywords:
                    if any(keyword in val for val in cell_values):
                        found_keywords.add(keyword)
            except Exception as e:
                st.error(f"Error reading {file.name}: {e}")
    
        missing = [k for k in required_keywords if k not in found_keywords]
        if missing:
            st.warning(f"Missing required calendar(s): {', '.join(missing)}")
        else:
            st.success("All required calendars uploaded and verified by content.")
    
    student_file = st.file_uploader(
        "2) Upload Redcap Rotation list CSV (must have 'legal_name' and 'start_date')",
        type=["csv"]
    )
    
    record_id = "peds_clerkship"
    
    if not schedule_files or not student_file or not record_id:
        st.info("Please upload schedule Excel(s) and student CSV to proceed.")
        st.stop()
    
    # ─── Prep: Date regex & Hope Drive maps ────────────────────────────────────────
    date_pat = re.compile(r'^[A-Za-z]+ \d{1,2}, \d{4}$')
    base_map = {
        "1st picu attending 7:30a-4p":        "d_att_",
        "1st picu attending 7:30a-2p":        "d_att_",
        "1st picu attending 7:30a-5p":        "d_att_",

        "2nd picu attending 7:45a-12p":       "d_att_",

        "picu attending pm call 2p-8a":       "n_att_",
        "picu attending pm call 4p-8a":       "n_att_",
        "picu attending pm call 5p-11:30a":    "n_att_",
        
        "app/fellow day 6:30a-6:30p":         "d_app_",
        "app/fellow night 5p-7a":             "n_app_"}

    

    FIRST_APP_FELLOW_DAY = "app/fellow day 6:30a-6:30p"  # <-- add

    FIRST_ATT_KEYS = {"1st picu attending 7:30a-4p", "1st picu attending 7:30a-2p", "1st picu attending 7:30a-5p"}
    SECOND_ATT_KEYS = {"2nd picu attending 7:45a-12p"}
    
    # ─── 1. Aggregate schedule assignments by date ────────────────────────────────
    assignments_by_date = {}
    for file in schedule_files:
        df = pd.read_excel(file, header=None, dtype=str)
        # find all date headers
        date_positions = []
        for r in range(df.shape[0]):
            for c in range(df.shape[1]):
                val = str(df.iat[r,c]).strip().replace("\xa0"," ")
                if date_pat.match(val):
                    try:
                        d = pd.to_datetime(val).date()
                        date_positions.append((d,r,c))
                    except:
                        pass
        # pick topmost row per date
        unique = {}
        for d,r,c in date_positions:
            if d not in unique or r < unique[d][0]:
                unique[d] = (r,c)
        
        for d, (row0,col0) in unique.items():
            grp = assignments_by_date.setdefault(d, {des: [] for des in base_map})
            for r in range(row0+1, df.shape[0]):
                raw = str(df.iat[r, col0]).strip().replace("\xa0", " ")
                if not raw or date_pat.match(raw):
                    break
                desc = raw.lower()
                prov = str(df.iat[r, col0+1]).strip()
                if desc in grp and prov:
                    #IF ONLY WANT TO THE FIRST APP/FELLOW THEN UNHASH
                    #if desc == FIRST_APP_FELLOW_DAY and grp[desc]:
                    #    continue  # skip any additional ones
                    grp[desc].append(prov)
    
    # ─── 2. Read student list ─────────────────────────────────────────────────────
    students_df = pd.read_csv(student_file, dtype=str)
    legal_names = students_df["legal_name"].dropna().tolist()
    
    # ─── 3. Build REDCap rows (one per student, record_id from CSV) ────────────────
    sorted_dates = sorted(assignments_by_date.keys())
    
    # Precompute provider fields once (same schedule for all students)
    provider_fields = {}
    for idx, date in enumerate(sorted_dates, start=0):  # d00, d01, ...
        day_suffix = f"{idx:02}"  # "00", "01", ...
        day_data = assignments_by_date[date]
    
        # First & second attending pinned to fixed slots
        first_att = next(
            (day_data[k][0] for k in FIRST_ATT_KEYS if k in day_data and day_data[k]),
            None
        )
        if first_att:
            provider_fields[f"d_att{day_suffix}_1"] = first_att
    
        second_att = next(
            (day_data[k][0] for k in SECOND_ATT_KEYS if k in day_data and day_data[k]),
            None
        )
        if second_att:
            provider_fields[f"d_att{day_suffix}_2"] = second_att
    
        # Everything else (skip keys already handled above)
        for des, provs in day_data.items():
            if des in FIRST_ATT_KEYS or des in SECOND_ATT_KEYS:
                continue
    
            # Keep only the first app/fellow day, if desired
            if des == "app/fellow day 6:30a-6:30p":
                provs = provs[:1]
    
            prefs = base_map.get(des)
            if not prefs:
                continue
            prefixes = [prefs + day_suffix + "_"] if isinstance(prefs, str) else [p + day_suffix + "_" for p in prefs]
    
            for i, name in enumerate(provs, start=1):
                for prefix in prefixes:
                    provider_fields[f"{prefix}{i}"] = name
    
    # Validate the student CSV has record_id
    if "record_id" not in students_df.columns:
        st.error("The student CSV must include a 'record_id' column.")
        st.stop()
    
    # Build one row per student
    rows = []
    for _, srow in students_df.iterrows():
        rid = str(srow["record_id"]).strip()
        start_date_value = str(srow.get("start_date", "")).strip()
        if not rid:
            continue  # or raise/warn
    
        row = {
            "record_id": rid,
            "start_date": start_date_value,
        }
        row.update(provider_fields)  # add all schedule/provider columns
        rows.append(row)
    
    # ─── 4. Display & download ────────────────────────────────────────────────────
    out_df = pd.DataFrame(rows)
    csv_full = out_df.to_csv(index=False).encode("utf-8")
    
    st.subheader("✅ Full REDCap Import Preview")
    st.dataframe(out_df)
    st.download_button("⬇️ Download Full CSV", csv_full, "batch_import_full.csv", "text/csv")
