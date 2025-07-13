import streamlit as st
import pandas as pd
from datetime import timedelta

st.set_page_config(page_title="Hope Drive → REDCap Import", layout="wide")
st.title("Hope Drive Preceptors → REDCap Import Template")

uploaded_file = st.file_uploader("Upload your AGP calendar (Excel)", type=["xlsx","xls"])
record_id    = st.text_input("Enter REDCap record_id for this session", "")

if uploaded_file and record_id:
    # 1. Read raw sheet as strings
    df = pd.read_excel(uploaded_file, header=None, dtype=str)

    # 2. Parse session date from A5
    try:
        raw_date     = pd.to_datetime(df.iat[4, 0])
        hd_day_date  = raw_date.date()
    except Exception:
        st.error("⚠️ Could not parse a valid date from cell A5.")
        st.stop()

    # 3. Format day0 and day7 exactly as they appear in column A
    day0 = hd_day_date.strftime('%B %-d, %Y')
    day7 = (hd_day_date + timedelta(days=7)).strftime('%B %-d, %Y')

    # 4. Find start/end indices by matching those formatted strings
    col0 = df.iloc[:, 0].str.strip()
    try:
        start_idx = col0[col0 == day0].index[0]
        end_idx   = col0[col0 == day7].index[0]
    except IndexError:
        st.error(f"⚠️ Couldn’t locate either '{day0}' or '{day7}' in column A.")
        st.stop()

    # 5. Scan only between start_idx+1 and end_idx for continuity rows
    providers = []
    for r in range(start_idx + 1, end_idx):
        for c in range(df.shape[1] - 1):
            cell_text = str(df.iat[r, c]).strip()
            if cell_text == "Hope Drive AM Continuity":
                prov_text = str(df.iat[r, c + 1]).strip()
                if prov_text:
                    providers.append(prov_text)

    if not providers:
        st.error("⚠️ No 'Hope Drive AM Continuity' entries found in that date block.")
        st.stop()

    # 6. Build the import row
    data = {"record_id": record_id, "hd_day_date": hd_day_date}
    for i, name in enumerate(providers, start=1):
        data[f"hd_am_d1_{i}"] = name

    out_df = pd.DataFrame([data])

    # 7. Display & download
    st.subheader("📋 REDCap Import Preview")
    st.dataframe(out_df)

    csv = out_df.to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Download import CSV", csv, "hope_drive_import.csv", "text/csv")

elif uploaded_file and not record_id:
    st.info("Please enter a record_id so I can build the import row for you.")
else:
    st.info("Upload an Excel file and enter a record_id to get started.")

