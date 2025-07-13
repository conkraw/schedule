import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="Hope Drive → Multi‑Day Import", layout="wide")
st.title("Auto‑Detect Multi‑Day Preceptor → REDCap Import")

uploaded_file = st.file_uploader("Upload your AGP calendar (Excel)", type=["xlsx","xls"])
record_id     = st.text_input("Enter REDCap record_id for this session", "")

if uploaded_file and record_id:
    # 1️⃣ Read once
    df = pd.read_excel(uploaded_file, header=None, dtype=str)

    # 2️⃣ Find all “Month D, YYYY” cells anywhere
    date_pat = re.compile(r'^[A-Za-z]+ \d{1,2}, \d{4}$')
    date_positions = []
    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            val = str(df.iat[r, c]).replace("\xa0", " ").strip()
            if date_pat.match(val):
                d = pd.to_datetime(val).date()
                date_positions.append((r, c, d))

    if not date_positions:
        st.error("No Month‑Day‑Year headers found. Check your sheet.")
        st.stop()

    # 3️⃣ De‑dup & sort by date
    unique = {}
    for r, c, d in date_positions:
        unique.setdefault(d, (r, c))
    sessions = sorted((d, rc[0], rc[1]) for d, rc in unique.items())

    # 4️⃣ Base designation → prefix map
    base_map = {
        "hope drive am continuity":    "hd_am_",
        "hope drive pm continuity":    "hd_pm_",
        "hope drive am acute precept": "hd_am_acute_",
        "hope drive pm acute precept": "hd_pm_acute_",
    }
    # nursery gets two prefixes
    base_map["nursery weekday 8a-6p"] = ["nursery_am_", "nursery_pm_"]

    # 5️⃣ Which need at least 2 slots?
    min_required = {
        "hope drive am acute precept": 2,
        "hope drive pm acute precept": 2,
        "nursery weekday 8a-6p":       2,
    }

    # 6️⃣ Build one REDCap row
    redcap_row = {"record_id": record_id}

    for day_idx, (session_date, row0, col0) in enumerate(sessions, start=1):
        # a) Store the date
        redcap_row[f"hd_day_date{day_idx}"] = session_date

        # b) Build day‑specific prefixes
        suffix = f"d{day_idx}_"
        designation_map = {}
        for des, pref in base_map.items():
            if isinstance(pref, list):
                designation_map[des] = [p + suffix for p in pref]
            else:
                designation_map[des] = [pref + suffix]

        # c) Collect designations/providers down the column
        grouped = {des: [] for des in designation_map}
        for r in range(row0 + 1, df.shape[0]):
            cell = str(df.iat[r, col0]).replace("\xa0", " ").strip().lower()
            if "monday" in cell:  # hit next week
                break
            prov = str(df.iat[r, col0 + 1]).strip()
            if cell in grouped and prov:
                grouped[cell].append(prov)

        # d) Pad groups
        for des, provs in grouped.items():
            req = min_required.get(des, len(provs))
            while len(provs) < req and provs:
                provs.append(provs[0])

            # e) Populate fields
            for idx, name in enumerate(provs, start=1):
                for prefix in designation_map[des]:
                    redcap_row[f"{prefix}{idx}"] = name

    # 7️⃣ Output & download
    out_df = pd.DataFrame([redcap_row])
    st.subheader("✅ REDCap Import Preview")
    st.dataframe(out_df)

    st.download_button(
        "⬇️ Download full CSV",
        data=out_df.to_csv(index=False).encode("utf-8"),
        file_name="multi_day_import.csv",
        mime="text/csv"
    )

elif uploaded_file and not record_id:
    st.info("Please enter a record_id so I can build the import row for this session.")
else:
    st.info("Upload your AGP calendar and enter a record_id to begin.")



