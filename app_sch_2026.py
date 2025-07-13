import streamlit as st
import pandas as pd
from datetime import timedelta

st.set_page_config(page_title="Hope Drive → REDCap Import", layout="wide")
st.title("Hope Drive Preceptors → REDCap Import Template")

uploaded_file = st.file_uploader("Upload your AGP calendar (Excel)", type=["xlsx", "xls"])
record_id = st.text_input("Enter REDCap record_id for this session", "")

if uploaded_file and record_id:
    df = pd.read_excel(uploaded_file, header=None, dtype=str)

    # 1. Extract and format session date from A5
    try:
        raw_date = pd.to_datetime(df.iat[4, 0])
        hd_day_date = raw_date.date()
    except Exception:
        st.error("⚠️ Could not parse a valid date from cell A5.")
        st.stop()

    day0_str = hd_day_date.strftime('%B %-d, %Y').lower()
    day7_str = (hd_day_date + timedelta(days=7)).strftime('%B %-d, %Y').lower()

    # 2. Normalize and inspect column A
    col0_raw = df.iloc[:, 0].fillna("").astype(str)
    col0_clean = col0_raw.str.replace("\xa0", " ", regex=False).str.strip().str.lower()

    st.subheader("🔍 Preview of cleaned column A")
    st.write(col0_clean.head(50).to_frame())

    # 3. Locate start and end rows by substring match
    start_matches = col0_clean[col0_clean.str.contains(day0_str, na=False)]
    end_matches   = col0_clean[col0_clean.str.contains(day7_str, na=False)]

    if start_matches.empty or end_matches.empty:
        st.error(f"Could not locate rows containing '{day0_str}' or '{day7_str}'.")
        st.stop()

    start_row = start_matches.index[0]
    end_row   = end_matches.index[0]

    # 4. Scan rows in that window for "Hope Drive AM Continuity"
    providers = []
    for r in range(start_row + 1, end_row):
        for c in range(df.shape[1] - 1):
            cell = str(df.iat[r, c]).replace("\xa0", " ").strip()
            if cell == "Hope Drive AM Continuity":
                prov = str(df.iat[r, c + 1]).replace("\xa0", " ").strip()
                if prov:
                    providers.append(prov)

    if not providers:
        st.error("⚠️ No 'Hope Drive AM Continuity' entries found between date blocks.")
        st.stop()

    # 5. Build the REDCap import row
    data = {"record_id": record_id, "hd_day_date": hd_day_date}
    for i, name in enumerate(providers, start=1):
        data[f"hd_am_d1_{i}"] = name

    out_df = pd.DataFrame([data])

    # 6. Display and download
    st.subheader("📋 REDCap Import Preview")
    st.dataframe(out_df)

    csv = out_df.to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Download import CSV", csv, "hope_drive_import.csv", "text/csv")

elif uploaded_file and not record_id:
    st.info("Please enter a record_id so I can build the import row for you.")
else:
    st.info("Upload an Excel file and enter a record_id to get started.")


