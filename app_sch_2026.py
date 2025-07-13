import streamlit as st
import pandas as pd

# ─── Page setup ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="Hope Drive → REDCap Import", layout="wide")
st.title("Hope Drive Preceptors → REDCap Import Template")

# ─── Inputs ───────────────────────────────────────────────────────────────────
uploaded_file = st.file_uploader(
    "Upload your AGP calendar (Excel)", 
    type=["xlsx", "xls"]
)
record_id = st.text_input("Enter REDCap record_id for this session", "")

# ─── Main logic ────────────────────────────────────────────────────────────────
if uploaded_file and record_id:
    # 1️⃣ Read raw sheet (no headers)
    df = pd.read_excel(uploaded_file, header=None)

    # 2️⃣ Parse the session date from A5 (row 4)
    try:
        hd_day_date = pd.to_datetime(df.iat[4, 0]).date()
    except Exception:
        st.error("⚠️ Could not parse a valid date from cell A5.")
        st.stop()

    # 3️⃣ Vectorize column A → dates (NaT on non‑dates)
    col_dates = (
        pd.to_datetime(df.iloc[:, 0], errors="coerce")
          .dt.date
    )

    # 4️⃣ Find the first row matching that date
    matches = col_dates[col_dates == hd_day_date].index
    if matches.empty:
        st.error("⚠️ Could not locate the session date in column A.")
        st.stop()
    first_date_row = matches[0]

    # 5️⃣ Find the very next row where the date changes
    diff_mask = col_dates != hd_day_date
    after = diff_mask[first_date_row + 1 :].index
    next_date_row = after[0] if len(after) else len(df)

    # 6️⃣ Scan only between those two rows for “Hope Drive AM Continuity”
    providers = []
    for r in range(first_date_row + 1, next_date_row):
        for c in range(df.shape[1] - 1):
            if str(df.iat[r, c]).strip() == "Hope Drive AM Continuity":
                val = df.iat[r, c + 1]
                if pd.notna(val):
                    providers.append(str(val).strip())

    if not providers:
        st.error("⚠️ No ‘Hope Drive AM Continuity’ entries found in that date block.")
        st.stop()

    # 7️⃣ Build a single‑row REDCap import DataFrame
    data = {
        "record_id": record_id,
        "hd_day_date": hd_day_date
    }
    for i, name in enumerate(providers, start=1):
        data[f"hd_am_d1_{i}"] = name

    out_df = pd.DataFrame([data])

    # 8️⃣ Display & CSV download
    st.subheader("📋 REDCap Import Preview")
    st.dataframe(out_df)

    csv = out_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Download import CSV",
        data=csv,
        file_name="hope_drive_import.csv",
        mime="text/csv"
    )

    # 9️⃣ Instructions
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

# ─── Help messages ─────────────────────────────────────────────────────────────
elif uploaded_file and not record_id:
    st.info("Please enter a record_id so I can build the import row for you.")
else:
    st.info("Upload an Excel file and enter a record_id to get started.")

