import streamlit as st
import pandas as pd

st.set_page_config(page_title="Hope Drive → REDCap Import", layout="wide")
st.title("Hope Drive Preceptors → REDCap Import Template")

uploaded_file = st.file_uploader(
    "Upload your AGP calendar (Excel)", 
    type=["xlsx", "xls"]
)
record_id = st.text_input("Enter REDCap record_id for this session", "")

if uploaded_file and record_id:
    df = pd.read_excel(uploaded_file, header=None)

    # 1️⃣ Parse the session date from A5
    try:
        hd_day_date = pd.to_datetime(df.iat[4, 0]).date()
    except Exception:
        st.error("⚠️ Could not parse a valid date from cell A5.")
        st.stop()

    # 2️⃣ Vectorize column A → dates (NaT on non‑dates)
    col_dates = pd.to_datetime(df.iloc[:, 0], errors="coerce").dt.date

    # 3️⃣ Find the row index for that date
    date_rows = col_dates[col_dates == hd_day_date].index
    if date_rows.empty:
        st.error("⚠️ Could not locate the session date in column A.")
        st.stop()
    start_row = date_rows[0]

    # 4️⃣ Find where the **next** real date appears (so we know where to stop)
    differing = col_dates != hd_day_date
    next_rows = differing[differing].index
    # pick the first next‐date row that’s after start_row
    stop_row = next((r for r in next_rows if r > start_row), len(df))

    # 5️⃣ Scan **only** between start_row+1 and stop_row
    providers = []
    for r in range(start_row + 1, stop_row):
        for c in range(df.shape[1] - 1):
            if str(df.iat[r, c]).strip() == "Hope Drive AM Continuity":
                val = df.iat[r, c + 1]
                if pd.notna(val):
                    providers.append(str(val).strip())

    if not providers:
        st.error("⚠️ No 'Hope Drive AM Continuity' entries found in that date block.")
        st.stop()

    # 6️⃣ Build your single‐row REDCap import
    row = {
        "record_id": record_id,
        "hd_day_date": hd_day_date
    }
    for idx, name in enumerate(providers, start=1):
        row[f"hd_am_d1_{idx}"] = name

    out_df = pd.DataFrame([row])

    # 7️⃣ Show & let them download
    st.subheader("📋 REDCap Import Preview")
    st.dataframe(out_df)

    csv = out_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Download import CSV",
        data=csv,
        file_name="hope_drive_import.csv",
        mime="text/csv"
    )

elif uploaded_file and not record_id:
    st.info("Please enter a record_id so I can build the import row for you.")
else:
    st.info("Upload an Excel file and enter a record_id to get started.")

