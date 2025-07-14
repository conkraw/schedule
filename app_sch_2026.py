import streamlit as st
import pandas as pd
import re
from pathlib import Path

st.set_page_config(page_title="Batch Preceptor → REDCap Import", layout="wide")
st.title("Batch Preceptor → REDCap Import Generator")

# 1️⃣ Allow multiple Excel uploads
uploaded_files = st.file_uploader(
    "Upload one or more AGP calendar Excels",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)
record_prefix = st.text_input(
    "Enter a REDCap record_id prefix (e.g. 'B1R')", ""
)

if uploaded_files and record_prefix:
    all_rows = []

    # Regex for Month D, YYYY
    date_pat = re.compile(r'^[A-Za-z]+ \d{1,2}, \d{4}$')

    # Base designation → prefix map (no day suffix yet)
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

    for file in uploaded_files:
        df = pd.read_excel(file, header=None, dtype=str)

        # 2️⃣ Find all date cells anywhere
        date_positions = []
        for r in range(df.shape[0]):
            for c in range(df.shape[1]):
                val = str(df.iat[r, c]).replace("\xa0"," ").strip()
                if date_pat.match(val):
                    try:
                        d = pd.to_datetime(val).date()
                        date_positions.append((r, c, d))
                    except:
                        pass

        if not date_positions:
            st.warning(f"No session dates in {file.name}; skipping.")
            continue

        # 3️⃣ Deduplicate and sort by date
        unique = {}
        for r, c, d in date_positions:
            if d not in unique or r < unique[d][0]:
                unique[d] = (r, c)
        sessions = sorted((d, rc[0], rc[1]) for d, rc in unique.items())

        # 4️⃣ Prepare one row per file
        row = {}
        # Derive a record_id from prefix + filename
        stem = Path(file.name).stem.replace(" ", "_")
        row["record_id"] = f"{record_prefix}_{stem}"

        # 5️⃣ Loop through each detected session date
        for day_idx, (session_date, row0, col0) in enumerate(sessions, start=1):
            # a) date field
            row[f"hd_day_date{day_idx}"] = session_date

            # b) build day‑specific prefixes
            suffix = f"d{day_idx}_"
            desig_map = {}
            for des, pref in base_map.items():
                if isinstance(pref, list):
                    desig_map[des] = [p + suffix for p in pref]
                else:
                    desig_map[des] = [pref + suffix]

            # c) collect under that date until next "Monday"
            grouped = {des: [] for des in desig_map}
            for r in range(row0 + 1, df.shape[0]):
                cell = str(df.iat[r, col0]).replace("\xa0"," ").strip().lower()
                if cell == "":
                    continue
                if "monday" in cell:
                    break
                prov = str(df.iat[r, col0 + 1]).strip()
                if cell in grouped and prov:
                    grouped[cell].append(prov)

            # d) pad to min if needed
            for des, provs in grouped.items():
                req = min_required.get(des, len(provs))
                while len(provs) < req and provs:
                    provs.append(provs[0])
                # e) populate fields
                for idx, name in enumerate(provs, start=1):
                    for prefix in desig_map[des]:
                        row[f"{prefix}{idx}"] = name

        all_rows.append(row)

    # 6️⃣ Combine & display
    out_df = pd.DataFrame(all_rows)
    st.subheader("✅ Combined REDCap Import Preview")
    st.dataframe(out_df)

    csv = out_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Download CSV",
        csv,
        "batch_redcap_import.csv",
        "text/csv"
    )

elif uploaded_files and not record_prefix:
    st.info("Enter a record_id prefix so I can build rows for each file.")
else:
    st.info("Upload one or more Excel files and enter a record_id prefix to begin.")




