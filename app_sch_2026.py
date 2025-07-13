import streamlit as st
import pandas as pd
from datetime import timedelta

st.set_page_config(page_title="Hope Drive → REDCap Import", layout="wide")
st.title("Hope Drive Preceptors → REDCap Import Template")

uploaded_file = st.file_uploader("Upload your AGP calendar (Excel)", type=["xlsx","xls"])
record_id    = st.text_input("Enter REDCap record_id for this session", "")

if uploaded_file and record_id:
    # 1️⃣ Read raw sheet
    df = pd.read_excel(uploaded_file, header=None)

    # 2️⃣ Parse your session date from A5
    try:
        hd_day_date = pd.to_datetime(df.iat[4, 0]).date()
    except Exception:
        st.error("⚠️ Could not parse a valid date from cell A5.")
        st.stop()

    # 3️⃣ Compute the 7-day cutoff
    hd_end_date = hd_day_date + timedelta(days=7)

    # 4️⃣ Vectorize column A → dates (NaT on non-dates)
    col_dates = pd.to_datetime(df.iloc[:, 0], errors="coerce").dt.date

    # 5️⃣ Find all row indices where date is between start and end (inclusive)
    valid_rows = col_dates.dropna()
    mask = (valid_rows >= hd_day_date) & (valid_rows <= hd_end_date)
    rows_to_scan = mask[mask].index.tolist()

    if not rows_to_scan:
        st.error("⚠️ No rows in the 7‑day date range were found in column A.")
        st.stop()

    # 6️⃣ Within those rows, collect every provider to the right of "Hope Drive AM Continuity"
    providers = []
    for r in rows_to_scan:
        # scan all columns except the last
        for c in range(df.shape[1] - 1):
            if str(df.iat[r, c]).strip() == "Hope Drive AM Continuity":
                nxt = df.iat[r, c + 1]
                if pd.notna(nxt):
                    providers.append(str(nxt).strip())

    if not providers:
        st.error("⚠️ No 'Hope Drive AM Continuity' entries found in the 7‑day window.")
        st.stop()

    # 7️⃣ Build a single‑row REDCap import
    data = {"record_id": record_id, "hd_day_date": hd_day_date}
    for i, name in enumerate(providers, start=1):
        data[f"hd_am_d1_{i}"] = name

    out_df = pd.DataFrame([data])

    # 8️⃣ Display & download
    st.subheader("📋 REDCap Import Preview")
    st.dataframe(out_df)

    csv = out_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Download import CSV",
        data=csv,
        file_name="hope_drive_import.csv",
        mime="text/csv"
    )

    st.markdown(
        """
        **Next steps:**  
        1. In REDCap, define a repeating form/instrument named `hope_drive`.  
        2. Add fields:  
           - `hd_day_date` (Date Y‑M‑D)  
           - `hd_am_d1_1`, `hd_am_d1_2`, … (Text)  
        3. Use this CSV in the Data Import Tool or via the API.
        """
    )

elif uploaded_file and not record_id:
    st.info("Please enter a record_id so I can build the import row for you.")
else:
    st.info("Upload an Excel file and enter a record_id to get started.")

