import streamlit as st
import pandas as pd
from datetime import timedelta

st.set_page_config(page_title="Hope Drive → REDCap Import", layout="wide")
st.title("Hope Drive Preceptors → REDCap Import Template")

uploaded_file = st.file_uploader("Upload your AGP calendar (Excel)", type=["xlsx", "xls"])
record_id = st.text_input("Enter REDCap record_id for this session", "")

if uploaded_file and record_id:
    df = pd.read_excel(uploaded_file, header=None, dtype=str)

    # 1. Parse the session date from A5
    try:
        raw_date = pd.to_datetime(df.iat[4, 0])
        hd_day_date = raw_date.date()
    except Exception:
        st.error("⚠️ Could not parse a valid date from cell A5.")
        st.stop()

    # 2. Format dates to match cleaned column A
    day0_str = hd_day_date.strftime('%B %-d, %Y').lower()   # e.g. "july 7, 2025"
    day7_str = (hd_day_date + timedelta(days=7)).strftime('%B %-d, %Y').lower()

    # 3. Clean column A (normalize spaces, lowercase)
    col0 = (
        df.iloc[:, 0].fillna("")
        .astype(str)
        .str.replace("\xa0", " ", regex=False)
        .str.strip()
        .str.lower()
    )

    try:
        start_row = col0[col0 == day0_str].index[0]
        end_row   = col0[col0 == day7_str].index[0]
        
    except IndexError:
        st.error(f"❌ Could not find '{day0_str}' or '{day7_str}' in column A.")
        st.dataframe(col0.to_frame(name="column_A_cleaned").head(50))
        st.stop()

    # 4. Scan column pairs (0+1, 2+3, 4+5, ...)
    providers = set()  # use set to avoid duplicates
    for c in range(0, df.shape[1] - 1, 2):
        for r in range(start_row + 1, end_row):
            designation = str(df.iat[r, c]).replace("\xa0", " ").strip()
            if designation == "Hope Drive AM Continuity":
                provider = str(df.iat[r, c + 1]).replace("\xa0", " ").strip()
                if provider:
                    providers.add(provider)

    providers = sorted(providers)

    if not rows:
        st.error("⚠️ No 'Hope Drive AM Continuity' entries found in that date block.")
        st.stop()

    df_preview = pd.DataFrame(rows)
    st.subheader("🧾 Hope Drive Assignments")
    st.dataframe(df_preview)

    # 5. Build the REDCap import row
    data = {"record_id": record_id, "hd_day_date": hd_day_date}
    for i, name in enumerate(providers, start=1):
        data[f"hd_am_d1_{i}"] = name

    out_df = pd.DataFrame([data])

    # 6. Display & download
    st.subheader("📋 REDCap Import Preview")
    st.dataframe(out_df)

    csv = out_df.to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Download import CSV", csv, "hope_drive_import.csv", "text/csv")

elif uploaded_file and not record_id:
    st.info("Please enter a record_id so I can build the import row for you.")
else:
    st.info("Upload an Excel file and enter a record_id to get started.")


