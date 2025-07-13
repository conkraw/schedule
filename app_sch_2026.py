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

    # 2. Build formatted day0 and day7 strings (e.g., "July 7, 2025")
    day0_str = hd_day_date.strftime('%B %-d, %Y')
    day7_str = (hd_day_date + timedelta(days=7)).strftime('%B %-d, %Y')

    # 3. Find row indices of day0_str and day7_str in column 0
    col0 = df.iloc[:, 0].fillna("").str.strip()
    try:
        start_row = col0[col0 == day0_str].index[0]
        end_row = col0[col0 == day7_str].index[0]
    except IndexError:
        st.error(f"Could not find '{day0_str}' or '{day7_str}' in column A.")
        st.stop()

    # 4. Scan between start_row+1 and end_row for "Hope Drive AM Continuity"
    providers = []
    for r in range(start_row + 1, end_row):
        for c in range(df.shape[1] - 1):
            cell = str(df.iat[r, c]).strip()
            if cell == "Hope Drive AM Continuity":
                prov = str(df.iat[r, c + 1]).strip()
                if prov:
                    providers.append(prov)

    if not providers:
        st.error("⚠️ No 'Hope Drive AM Continuity' entries found between date blocks.")
        st.stop()

    # 5. Build REDCap import row
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


