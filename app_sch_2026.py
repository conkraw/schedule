import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="Batch Preceptor → REDCap Import", layout="wide")
st.title("Batch Preceptor → REDCap Import Generator")

# ─── Inputs ────────────────────────────────────────────────────────────────────
schedule_files = st.file_uploader(
    "1) Upload one or more AGP calendar Excel(s)",
    type=["xlsx","xls"],
    accept_multiple_files=True
)

student_file = st.file_uploader(
    "2) Upload student list CSV (must have a 'legal_name' column)",
    type=["csv"]
)

record_id = st.text_input("3) Enter the REDCap record_id for this batch", "")

# ─── Guard ─────────────────────────────────────────────────────────────────────
if not schedule_files or not student_file or not record_id:
    st.info("Please upload schedule Excel(s), student CSV, and enter a record_id.")
    st.stop()

# ─── Prep: Date regex & maps ───────────────────────────────────────────────────

date_pat = re.compile(r'^[A-Za-z]+ \d{1,2}, \d{4}$')
    base_map = {
        "hope drive am continuity":    "hd_am_",
        "hope drive pm continuity":    "hd_pm_",
        
        "hope drive am acute precept": "hd_am_acute_",
        "hope drive pm acute precept": "hd_pm_acute_",
        
        "etown am continuity":         "etown_am_",
        "etown pm continuity":         "etown_pm_",
        
        "nyes rd am continuity":       "nyes_am_",
        "nyes rd pm continuity":       "nyes_pm_",
        
        "nursery weekday 8a-6p":       ["nursery_am_", "nursery_pm_"],
        
        "rounder 1 7a-7p":             ["ward_a_am_team_1_","ward_a_pm_team_1_"],
        "rounder 2 7a-7p":             ["ward_a_am_team_2_","ward_a_pm_team_2_"],
        "rounder 3 7a-7p":             ["ward_a_am_team_3_","ward_a_pm_team_3_"],

        "hope drive clinic am":        "complex_am_1_",
        "hope drive clinic pm":        "complex_pm_1_",
        
        "briarcrest clinic am":       "adol_med_am_1_",
        "briarcrest clinic pm":       "adol_med_pm_1_",
        
    }

    # Which groups need at least 2 providers?
    min_required = {
        "hope drive am acute precept": 2,
        "hope drive pm acute precept": 2,
        "nursery weekday 8a-6p":       2,
        "rounder 1 7a-7p":             2,
        "rounder 2 7a-7p":             2,
        "rounder 3 7a-7p":             2,
    }

# ─── 1. Aggregate schedule assignments by date ────────────────────────────────
assignments_by_date = {}
for file in schedule_files:
    df = pd.read_excel(file, header=None, dtype=str)

    # find all date cells
    date_positions = []
    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            val = str(df.iat[r,c]).replace("\xa0"," ").strip()
            if date_pat.match(val):
                try:
                    d = pd.to_datetime(val).date()
                    date_positions.append((d,r,c))
                except:
                    pass

    # dedupe to the topmost row per date
    unique = {}
    for d,r,c in date_positions:
        if d not in unique or r < unique[d][0]:
            unique[d] = (r,c)

    # collect providers under each date
    for d, (row0,col0) in unique.items():
        grp = assignments_by_date.setdefault(d, {des:[] for des in base_map})
        for r in range(row0+1, df.shape[0]):
            cell = str(df.iat[r,col0]).replace("\xa0"," ").strip().lower()
            if "monday" in cell:
                break
            prov = str(df.iat[r,col0+1]).strip()
            if cell in grp and prov:
                grp[cell].append(prov)

# ─── 2. Read student list and prepare s1, s2, … ───────────────────────────────
students_df = pd.read_csv(student_file, dtype=str)
legal_names = students_df["legal_name"].dropna().tolist()

# ─── 3. Build the single REDCap row ───────────────────────────────────────────
redcap_row = {"record_id": record_id}

# sort dates chronologically
sorted_dates = sorted(assignments_by_date.keys())

# loop days for schedule fields
for idx, date in enumerate(sorted_dates, start=1):
    # date
    redcap_row[f"hd_day_date{idx}"] = date
    suffix = f"d{idx}_"
    # designation→ day-specific prefixes
    des_map = {
        des: ([p+suffix for p in prefs] if isinstance(prefs,list) else [prefs+suffix])
        for des,prefs in base_map.items()
    }
    # providers for this date
    for des, provs in assignments_by_date[date].items():
        req = min_required.get(des, len(provs))
        while len(provs) < req and provs:
            provs.append(provs[0])
        for i,name in enumerate(provs, start=1):
            for prefix in des_map[des]:
                redcap_row[f"{prefix}{i}"] = name

# append student slots s1,s2,...
for i,name in enumerate(legal_names, start=1):
    redcap_row[f"s{i}"] = name

# ─── 4. Display & slice out dates/am/acute and students ─────────────────────
out_df = pd.DataFrame([redcap_row])

# format date columns
for c in out_df.columns:
    if c.startswith("hd_day_date"):
        out_df[c] = pd.to_datetime(out_df[c]).dt.strftime("%m-%d-%Y")

st.subheader("✅ Full REDCap Import Preview")
st.dataframe(out_df)

# subset columns
date_cols     = [c for c in out_df.columns if c.startswith("hd_day_date")]
am_cont_cols  = [f"hd_am_d1_{i}" for i in range(1, 19)] + [f"hd_am_d2_{i}" for i in range(1, 19)]
am_acute_cols = [f"hd_am_acute_d1_{i}" for i in (1,2)] + [f"hd_am_acute_d2_{i}" for i in (1,2)]
student_cols  = [f"s{i}" for i in range(1, len(legal_names)+1)]

subset = ["record_id"] + date_cols + am_cont_cols + am_acute_cols + student_cols
subset = [c for c in subset if c in out_df.columns]

st.subheader("📅 Dates, AM Continuity/Acute & Students")
st.dataframe(out_df[subset])

# downloads
csv_full = out_df.to_csv(index=False).encode("utf-8")
st.download_button("⬇️ Download Full CSV", csv_full, "batch_import_full.csv", "text/csv")

csv_sub  = out_df[subset].to_csv(index=False).encode("utf-8")
st.download_button("⬇️ Download Dates+AM+Students CSV", csv_sub, "dates_am_students.csv", "text/csv")





