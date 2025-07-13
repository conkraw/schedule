import streamlit as st
import pandas as pd
from datetime import timedelta

st.set_page_config(page_title="Hope Drive → REDCap Import", layout="wide")
st.title("Hope Drive Preceptors → REDCap Import Template")

# Upload + record_id
uploaded_file = st.file_uploader("Upload your AGP calendar (Excel)", type=["xlsx","xls"])
record_id    = st.text_input("Enter REDCap record_id for this session", "")

if uploaded_file and record_id:
    # 1️⃣ Read raw sheet
    df = pd.read_excel(uploaded_file, header=None, dtype=str)

    # 2️⃣ Parse the session date from A5
    try:
        raw_date = pd.to_datetime(df.iat[4, 0])
        hd_day_date = raw_date.date()
    except Exception:
        st.error("⚠️ Could not parse a valid date from cell A5.")
        st.stop()

    # 3️⃣ Build formatted day‑0 and day‑7 strings just like your old code
    #     (e.g. "July 7, 2025")
    day0 = hd_day_date.strftime('%B %-d, %Y')
    day7 = (hd_day_date + timedelta(days=7)).strftime('%B %-d, %Y')

    # 4️⃣ Find those two rows in column 0
    col0 = df.iloc[:, 0].str.strip()
    try:
        start_idx = col0[col0 == day0].index[0]
        end_idx   = col0[col0 == day7].index[0]
    except IndexError:
        st.error(f"⚠️ Couldn’t locate either '{day0}' or '{day7}' in column A.")
        st.stop()

    # 5️⃣ Scan only between those rows for your continuity entries
    providers = []
    for r in range(start_idx + 1, end_idx):
        for c in range(df.shape[1] - 1):
            if df.iat[r, c].strip() == "Hope Drive AM Continuity":
                p = df.iat[r, c + 1].strip()
                if p:
                    providers.append(p)

    if not providers:
        st.error("⚠️ No 'Hope Drive AM Continuity' entries found in that date block.")
        st.stop()

    # 6️⃣ Build the REDCap import row
    data = {"record_id": record_id, "hd_day_date": hd_day_date}
    for i, name in enumerate(providers, start=1):
        data[f"hd_am_d1_{i}"] = name

    out_df = pd.DataFrame([data])

    # 7️⃣ Display & download
    st.subheader("📋 REDCap Import Preview")
    st.dataframe(out_df)

    csv = out_df.to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Download import CSV", csv, "hope_drive_import.csv", "text/csv")

elif uploaded_file and not record_id:
    st.info("Please enter a record_id so I can build the import row for you.")
else:
    st.info("Upload an Excel file and enter a record_id to get started.")

