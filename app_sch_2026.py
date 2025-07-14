import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="Batch Preceptor → REDCap Import", layout="wide")
st.title("Batch Preceptor → REDCap Import Generator")

# 1️⃣ Multiple files + single record_id
uploaded_files = st.file_uploader(
    "Upload one or more AGP calendar Excels",
    type=["xlsx","xls"],
    accept_multiple_files=True
)
record_id = st.text_input("Enter the REDCap record_id for all these sessions", "")

if uploaded_files and record_id:
    # Regex for Month D, YYYY
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
    }

    # Which groups need at least 2 providers?
    min_required = {
        "hope drive am acute precept": 2,
        "hope drive pm acute precept": 2,
        "nursery weekday 8a-6p":       2,
    }

    # 2️⃣ Create a single output row dict
    redcap_row = {"record_id": record_id}

    # 3️⃣ Process each file, layering fields onto the same row
    session_counter = 0
    for file in uploaded_files:
        df = pd.read_excel(file, header=None, dtype=str)

        # a) Find all date cells anywhere
        positions = [
            (r, c, pd.to_datetime(df.iat[r, c]).date())
            for r in range(df.shape[0]) for c in range(df.shape[1])
            if isinstance(df.iat[r, c], str)
            and date_pat.match(df.iat[r, c].replace("\xa0"," ").strip())
        ]
        if not positions:
            st.warning(f"No dates in {file.name}; skipping.")
            continue

        # b) Dedupe/sort by date → this file’s sessions
        unique = {}
        for r, c, d in positions:
            if d not in unique or r < unique[d][0]:
                unique[d] = (r, c)
        sessions = sorted((d, rc[0], rc[1]) for d, rc in unique.items())

        # c) Loop through sessions in this file
        for session_date, row0, col0 in sessions:
            session_counter += 1
            suffix = f"d{session_counter}_"

            # store the date
            redcap_row[f"hd_day_date{session_counter}"] = session_date

            # build this day’s prefix map
            desig_map = {
                des: ([p + suffix for p in prefs] if isinstance(prefs, list)
                      else [prefs + suffix])
                for des, prefs in base_map.items()
            }

            # collect down until next Monday
            grouped = {des: [] for des in desig_map}
            for r in range(row0 + 1, df.shape[0]):
                cell = str(df.iat[r, col0]).replace("\xa0"," ").strip().lower()
                if "monday" in cell:
                    break
                prov = str(df.iat[r, col0 + 1]).strip()
                if cell in grouped and prov:
                    grouped[cell].append(prov)

            # pad & write fields
            for des, provs in grouped.items():
                req = min_required.get(des, len(provs))
                while len(provs) < req and provs:
                    provs.append(provs[0])
                for idx, name in enumerate(provs, start=1):
                    for prefix in desig_map[des]:
                        redcap_row[f"{prefix}{idx}"] = name

    # 4️⃣ Show & download
    out_df = pd.DataFrame([redcap_row])
    st.subheader("✅ Combined REDCap Import Preview")
    st.dataframe(out_df)

    csv = out_df.to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Download CSV", csv, "batch_import.csv", "text/csv")

elif uploaded_files and not record_id:
    st.info("Enter a record_id to use for all uploaded files.")
else:
    st.info("Upload AGP calendar Excel(s) and enter a REDCap record_id to begin.")




