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

def to_date_or_none(x):
    try:
        return pd.to_datetime(x).date()
    except Exception:
        return None

def window_dates(all_dates, start_date):
    """Return sorted dates in [start_date, start_date + 4 weeks)."""
    if not isinstance(start_date, datetime) and not isinstance(start_date, pd.Timestamp):
        # allow date or string
        sd = to_date_or_none(start_date)
    else:
        sd = start_date.date()
    if sd is None:
        return []
    end = sd + timedelta(weeks=4)
    return [d for d in sorted(all_dates) if sd <= d < end]

st.set_page_config(page_title="Batch Preceptor → REDCap Import", layout="wide")
st.title("Batch Preceptor → REDCap Import Generator")

# ─── Sidebar mode selector ─────────────────────────────────────────────────────
mode = st.sidebar.radio("What do you want to do?", ("Roster_HMC","Format OPD + Summary","Oasis_Eval_Redcap_Creator","OASIS Evaluation"))

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
        "1st picu attending 7:30a-4p":        "d_att",
        "1st picu attending 7:30a-2p":        "d_att",
        "1st picu attending 7:30a-5p":        "d_att",

        "2nd picu attending 7:45a-12p":       "d_att",

        "picu attending pm call 2p-8a":       "n_att",
        "picu attending pm call 4p-8a":       "n_att",
        "picu attending pm call 5p-11:30a":   "n_att",
        "picu attending pm call 5p-8a":       "n_att",
        
        "app/fellow day 6:30a-6:30p":         "d_app",
        "app/fellow night 5p-7a":             "n_app",
        "on-call 6:30a-6:30a":                "n_app",}

    

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
    
    # ─── 3. Build REDCap rows (one per student, with per-student 4-week window) ───
    all_dates = sorted(assignments_by_date.keys())
    
    if "record_id" not in students_df.columns:
        st.error("The student CSV must include a 'record_id' column.")
        st.stop()
    if "start_date" not in students_df.columns:
        st.error("The student CSV must include a 'start_date' column.")
        st.stop()
    
    rows = []
    for _, srow in students_df.iterrows():
        rid = str(srow["record_id"]).strip()
        sd_raw = str(srow["start_date"]).strip()
        sd = to_date_or_none(sd_raw)
        if not rid or sd is None:
            # Skip or warn if missing/invalid
            continue
    
        # Dates to include for this student: [start_date, start_date + 4 weeks)
        dates_for_student = window_dates(all_dates, sd)
        
        dates_for_student = dates_for_student[:27]
        
        if not dates_for_student:
            # If QGenda doesn't contain that start_date window, you can warn/skip
            # st.warning(f"No schedule dates found for {rid} from {sd} to {sd + timedelta(weeks=4)}")
            continue
    
        # Build provider fields for this student's window only
        provider_fields = {}
        for day_idx, date in enumerate(dates_for_student, start=1):  # 00, 01, ...
            day_suffix = f"{day_idx:02}"
            day_data = assignments_by_date.get(date, {})
    
            # Pin first & second attending
            first_att = next((day_data[k][0] for k in FIRST_ATT_KEYS if k in day_data and day_data[k]), None)
            if first_att:
                provider_fields[f"d_att{day_suffix}_1"] = first_att
    
            second_att = next((day_data[k][0] for k in SECOND_ATT_KEYS if k in day_data and day_data[k]), None)
            if second_att:
                provider_fields[f"d_att{day_suffix}_2"] = second_att
    
            # Everything else (skip the pinned attending keys)
            for des, provs in day_data.items():
                if des in FIRST_ATT_KEYS or des in SECOND_ATT_KEYS:
                    continue
                if des == "app/fellow day 6:30a-6:30p":
                    provs = provs[:2]  # cap at two

                if des == "on-call 6:30a-6:30a":
                    provs = provs[:1]  # cap at one
    
                prefs = base_map.get(des)
                if not prefs:
                    continue
                prefixes = [prefs + day_suffix + "_"] if isinstance(prefs, str) \
                           else [p + day_suffix + "_" for p in prefs]
                for i, name in enumerate(provs, start=1):
                    for prefix in prefixes:
                        provider_fields[f"{prefix}{i}"] = name
    
        # Build the student row
        row = {
            "record_id": rid,
            "start_date": sd.strftime("%Y-%m-%d"),  # keep the original string as provided
        }
        row.update(provider_fields)
        rows.append(row)
    
    # ─── 4. Display & download ────────────────────────────────────────────────────
    out_df = pd.DataFrame(rows)
    csv_full = out_df.to_csv(index=False).encode("utf-8")
    
    st.subheader("✅ Full REDCap Import Preview")
    st.dataframe(out_df)
    st.download_button("⬇️ Download Full CSV", csv_full, "batch_import_full.csv", "text/csv")


elif mode == "Roster_HMC":
    st.header("🔖 Roster_HMC")
    st.markdown("[🔗 Roster Website](https://oasis.pennstatehealth.net/admin/course/roster/)")

    # upload exactly one CSV
    roster_file = st.file_uploader("Upload exactly one Roster CSV",type=["csv"],accept_multiple_files=False,key="roster")
    
    if not roster_file:
        st.stop()

    # read as CSV
    df_roster = pd.read_csv(roster_file, dtype=str)

    df_roster.columns = df_roster.columns.str.strip()

    # map your columns to REDCap-friendly names
    rename_map = {
        "#":                              "row_number",
        "Student":                        "student",
        "Legal Name":                     "legal_name",
        "Previous Name":                  "previous_name",
        "Username":                       "username",
        "Confidential":                   "confidential",
        "External ID":                    "record_id",
        "Email Address":                  "email",
        "Phone":                          "phone",
        "Pager":                          "pager",
        "Mobile":                         "mobile",
        "Gender":                         "gender",
        "Pronouns":                       "pronouns",
        "Ethnicity":                      "ethnicity",
        "Designation":                    "designation",
        "AAMC ID":                        "aamc_id",
        "USMLE ID":                       "usmle_id",
        "Home School":                    "home_school",
        "Campus":                         "campus",
        "Date of Birth":                  "date_of_birth",
        "Emergency Contact":              "emergency_contact",
        "Emergency Phone":                "emergency_phone",
        "Primary Academic Department":    "primary_academic_department",
        "Secondary Academic Department":  "secondary_academic_department",
        "Academic Type":                  "academic_type",
        "Primary Site":                   "primary_site",
        "NBME":                           "nbme_score",
        "PSU ID":                         "psu_id",
        "Productivity Specialty":         "productivity_specialty",
        "Grade":                          "grade",
        "Status":                         "status",
        "Student Level":                  "student_level",
        "Track":                          "track",
        "Location":                       "location",
        "Start Date":                     "start_date",
        "End Date":                       "end_date",
        "Weeks":                          "weeks",
        "Credits":                        "credits",
        "Enrolled":                       "enrolled",
        "Actions":                        "actions",
        "Aprv By":                     "approved_by"
    }
    df_roster = df_roster.rename(columns=rename_map)

    # keep only those renamed columns (in this exact order)
    df_roster = df_roster[list(rename_map.values())]

    # move record_id to the front
    cols = ["record_id"] + [c for c in df_roster.columns if c != "record_id"]
    df_roster = df_roster[cols]

    # add REDCap repeater - dont need
    #df_roster["redcap_repeat_instrument"] = "roster"
    #df_roster["redcap_repeat_instance"]   = df_roster.groupby("record_id").cumcount() + 1

    # ─── split “student” into last_name / first_name ─────────────────────────
    # 1) drop everything after the semicolon
    name_only = df_roster["student"].str.split(";", n=1).str[0]
    # 2) split on comma into last / first
    parts = name_only.str.split(",", n=1, expand=True)
    df_roster["lastname"]  = parts[0].str.strip()
    df_roster["firstname"] = parts[1].str.strip()

    df_roster["name"] = df_roster["firstname"] + " " + df_roster["lastname"]

    df_roster["legal_name"] = df_roster["lastname"] + ", " + df_roster["firstname"] + " (MD)" 

    df_roster["email_2"] = df_roster["record_id"] + "@psu.edu"

    #legal name ... legal_name
    
    # 3) (optional) drop the original combined column
    renamed_cols_a = ["row_number","student","previous_name","username","confidential","phone","pager","mobile","gender","pronouns","ethnicity","designation","aamc_id","usmle_id","home_school"]
    renamed_cols_b = ["campus","date_of_birth","emergency_contact","emergency_phone","primary_academic_department","secondary_academic_department","academic_type","primary_site","nbme_score"]
    renamed_cols_c = ["productivity_specialty","grade","status","student_level","weeks","credits","enrolled","actions","approved_by"]

    renamed_cols = renamed_cols_a + renamed_cols_b + renamed_cols_c

    df_roster.drop(columns=renamed_cols, errors="ignore", inplace=True)

    #DUE DATES
    
    # ─── 1) Ensure start_date and end_date are datetime ─────────────────────────
    df_roster["start_date"] = pd.to_datetime(df_roster["start_date"], infer_datetime_format=True)
    df_roster["end_date"]   = pd.to_datetime(df_roster["end_date"], infer_datetime_format=True)
    
    # ─── 2) Compute first Sunday on/after start_date ────────────────────────────
    days_to_sunday = (6 - df_roster["start_date"].dt.weekday) % 7
    first_sunday   = df_roster["start_date"] + pd.to_timedelta(days_to_sunday, unit="D")
    
    # ─── 3) Create quiz_due_1 … quiz_due_4 ──────────────────────────────────────
    for n in range(1, 5):
        df_roster[f"quiz_due_{n}"] = first_sunday + pd.Timedelta(weeks=(n - 1))
    
    # ─── 4) Alias assignment & doc-assignment due dates ─────────────────────────
    df_roster["ass_middue_date"]   = df_roster["quiz_due_2"]
    df_roster["ass_due_date"]      = df_roster["quiz_due_4"]
    df_roster["docass_due_date_1"] = df_roster["quiz_due_2"]
    df_roster["docass_due_date_2"] = df_roster["quiz_due_4"]
    
    # ─── 5) Grade due date: 6 weeks after end_date ──────────────────────────────
    df_roster["grade_due_date"] = df_roster["end_date"] + pd.Timedelta(weeks=6)

    # ─── 6) Normalize all due dates to 23:59 with no seconds ─────────────────
    due_cols = [
        "quiz_due_1","quiz_due_2","quiz_due_3","quiz_due_4",
        "ass_middue_date","ass_due_date",
        "docass_due_date_1","docass_due_date_2",
        "grade_due_date"
    ]
    
    for col in due_cols:
        df_roster[col] = (df_roster[col].dt.normalize() + pd.Timedelta(hours=23, minutes=59)).dt.strftime("%m-%d-%Y 23:59")

    df_roster["start_date"] = df_roster["start_date"].dt.strftime('%Y-%m-%d')  
    df_roster["end_date"] = df_roster["end_date"].dt.strftime('%Y-%m-%d')

    
    df_roster["student_demographics_complete"] = 2 
    
    # preview + download
    st.dataframe(df_roster, height=400)
    
    st.download_button("📥 Download formatted Roster CSV",df_roster.to_csv(index=False).encode("utf-8"),file_name="roster_formatted.csv",mime="text/csv")

elif mode == "Oasis_Eval_Redcap_Creator":
    oasis_files = st.file_uploader("Upload Oasis Evaluation Sample",type=["csv"],accept_multiple_files=False)

elif mode == "OASIS Evaluation":
    st.header("📋 OASIS Evaluation Formatter")
    st.markdown("[Open OASIS Clinical Assessment of Student Setup](https://oasis.pennstatehealth.net/admin/course/e_manage/student_performance/setup_analysis_report.html)")

    uploaded = st.file_uploader("Upload your raw OASIS CSV", type="csv", key="oasis")
    if not uploaded:
        st.stop()

    df = pd.read_csv(uploaded, dtype=str)

    # 自动把 "Course ID"→"course_id", "1 Question Number"→"q1_question_number", …
    def rename_oasis(col: str) -> str:
        col = col.strip()
        m = re.match(r"^(\d+)\s+(.+)$", col)
        if m:
            num, rest = m.groups()
            return f"q{num}_{rest.lower().replace(' ', '_')}"
        return col.lower().replace(" ", "_")

    df.columns = [rename_oasis(c) for c in df.columns]

    # build master_cols
    front = [
        "record_id","course_id","department","course","location",
        "start_date","end_date","course_type","student","student_username",
        "student_external_id","student_designation","student_email",
        "student_aamc_id","student_usmle_id","student_gender","student_level",
        "student_default_classification","evaluator","evaluator_username",
        "evaluator_external_id","evaluator_email","evaluator_gender",
        "who_completed","evaluation","form_record","submit_date"
    ]
    q_sufs = [
        "question_number","question_id","question","answer_text",
        "multiple_choice_order","multiple_choice_value","multiple_choice_label"
    ]
    questions = [f"q{i}_{s}" for i in range(1,20) for s in q_sufs]
    tail = ["oasis_eval_complete"]
    master_cols = front + questions + tail

    # reorder (will KeyError if you missed any)
    df = df.reindex(columns=master_cols)

    # inject REDCap fields
    df["record_id"]                = df["student_external_id"]
    df["redcap_repeat_instrument"] = "oasis_eval"
    df["redcap_repeat_instance"]   = df.groupby("record_id").cumcount() + 1

    # final column order
    keep_front = ["record_id","redcap_repeat_instrument","redcap_repeat_instance"]
    rest       = [c for c in master_cols if c not in keep_front]
    df = df.reindex(columns=keep_front + rest)


    # 8) remove student & location
    df = df.drop(columns=["student","location","start_date","end_date","location"]) #Cannot have these columns in the repeating instrument. 

    df["oasis_eval_complete"] = 2 
    
    st.dataframe(df, height=400)
    st.download_button(
        "📥 Download formatted OASIS CSV",
        df.to_csv(index=False).encode("utf-8"),
        file_name="oasis_eval_formatted.csv",
        mime="text/csv",
    )

    
    


