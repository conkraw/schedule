import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="Batch Preceptor → REDCap Import", layout="wide")
st.title("Batch Preceptor → REDCap Import Generator")

# 1️⃣ Multi‑file upload + single record_id
uploaded_files = st.file_uploader("Upload one or more AGP calendar Excels",type=["xlsx","xls"],accept_multiple_files=True)

record_id = st.text_input("Enter the REDCap record_id for this batch", "")

if uploaded_files and record_id:
    # Regex to detect “Month D, YYYY”
    date_pat = re.compile(r'^[A-Za-z]+ \d{1,2}, \d{4}$')

    # Base designation → prefix map
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

    # 2️⃣ Aggregate assignments by date → designation → [providers]
    assignments_by_date = {}

    for file in uploaded_files:
        df = pd.read_excel(file, header=None, dtype=str)

        # a) find all date cells anywhere
        date_positions = []
        for r in range(df.shape[0]):
            for c in range(df.shape[1]):
                val = str(df.iat[r, c]).replace("\xa0"," ").strip()
                if date_pat.match(val):
                    try:
                        d = pd.to_datetime(val).date()
                        date_positions.append((d, r, c))
                    except:
                        pass

        # b) dedupe lowest‐row position for each date
        unique = {}
        for d, r, c in date_positions:
            if d not in unique or r < unique[d][0]:
                unique[d] = (r, c)

        # c) process each date in this file
        for d, (row0, col0) in unique.items():
            # ensure we have a dict for this date
            assignments_by_date.setdefault(d, {des: [] for des in base_map})

            # scan downward until next "Monday"
            for r in range(row0 + 1, df.shape[0]):
                cell = str(df.iat[r, col0]).replace("\xa0"," ").strip().lower()
                if "monday" in cell:
                    break
                prov = str(df.iat[r, col0 + 1]).strip()
                if cell in assignments_by_date[d] and prov:
                    assignments_by_date[d][cell].append(prov)

    if not assignments_by_date:
        st.error("No valid session dates or assignments found across your files.")
        st.stop()

    # 3️⃣ Build the single REDCap row
    redcap_row = {"record_id": record_id}

    # sort dates chronologically
    sorted_dates = sorted(assignments_by_date.keys())

    for idx, date in enumerate(sorted_dates, start=1):
        # a) date field
        redcap_row[f"hd_day_date{idx}"] = date

        # b) day‐specific prefixes
        suffix = f"d{idx}_"
        desig_map = {
            des: ([p + suffix for p in prefs] if isinstance(prefs, list)
                  else [prefs + suffix])
            for des, prefs in base_map.items()
        }

        # c) get providers, pad as needed, populate fields
        for des, providers in assignments_by_date[date].items():
            req = min_required.get(des, len(providers))
            while len(providers) < req and providers:
                providers.append(providers[0])
            for i, name in enumerate(providers, start=1):
                for prefix in desig_map[des]:
                    redcap_row[f"{prefix}{i}"] = name
        
    # 4️⃣ Display & download
    out_df = pd.DataFrame([redcap_row])

    # Format all hd_day_dateN columns as MM-DD-YYYY
    for col in out_df.columns:
        if col.startswith("hd_day_date"):
            out_df[col] = pd.to_datetime(out_df[col]).dt.strftime("%m-%d-%Y")
            
    st.subheader("✅ REDCap Import Preview")
    st.dataframe(out_df)

    csv = out_df.to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Download Combined CSV", csv, "batch_import.csv", "text/csv")

    # After you've created and formatted out_df...

    # 1️⃣ Identify columns
    date_cols      = [c for c in out_df.columns if c.startswith("hd_day_date")]
    am_cont_cols   = [f"hd_am_d1_{i}" for i in range(1, 19)]
    am_acute_cols  = [f"hd_am_acute_d1_{i}" for i in (1, 2)]
    
    # 2️⃣ Subset
    subset_cols = date_cols + am_cont_cols + am_acute_cols
    dates_am_df = out_df.loc[:, [c for c in subset_cols if c in out_df.columns]]
    
    # 3️⃣ Display
    st.subheader("📅 Dates & AM Continuity/Acute Preview")
    st.dataframe(dates_am_df)
    
    # 4️⃣ Download
    csv = dates_am_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Download Dates + AM Continuity/Acute CSV",
        data=csv,
        file_name="dates_and_am_only.csv",
        mime="text/csv"
    )

elif uploaded_files and not record_id:
    st.info("Enter a record_id to generate the import row.")
else:
    st.info("Upload at least one Excel file and enter a record_id to begin.")





