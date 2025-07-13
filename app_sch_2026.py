import streamlit as st
import pandas as pd

st.set_page_config(page_title="Hope Drive → REDCap Import", layout="wide")
st.title("Hope Drive Preceptors → REDCap Import Template")

# 1️⃣ Upload + Record ID input
uploaded_file = st.file_uploader(
    "Upload your AGP calendar (Excel)", 
    type=["xlsx", "xls"]
)
record_id = st.text_input("Enter REDCap record_id for this session", "")

if uploaded_file:
    df = pd.read_excel(uploaded_file, header=None)

    # 2️⃣ Extract the date from A5
    try:
        hd_day_date = pd.to_datetime(df.iat[4, 0]).date()
    except Exception:
        st.error("⚠️ Could not parse a valid date from cell A5.")
        st.stop()

    # 3️⃣ Find every “Hope Drive AM Continuity” header in rows 0–19
    header_row = None
    cols = []
    for r in range(min(20, df.shape[0])):
        row = df.iloc[r].astype(str)
        if row.str.contains("Hope Drive AM Continuity", na=False).any():
            header_row = r
            cols = [c for c,v in row.items() if "Hope Drive AM Continuity" in v]
            break

    if header_row is None:
        st.error("⚠️ No ‘Hope Drive AM Continuity’ found in the first 20 rows.")
        st.stop()

    # 4️⃣ Grab the provider from the cell to the right of each header
    all_providers = []
    for c in cols:
        raw = df.iat[header_row + 1, c + 1]  # <-- note the “+1” here
        names = [n.strip() for n in str(raw).split(",") if n.strip()]
        all_providers.extend(names)

    if not all_providers:
        st.warning("Found headers, but no provider names to the right of them.")
        st.stop()

    # 5️⃣ Build a single-row DataFrame
    row = {
        "record_id": record_id,
        "hd_day_date": hd_day_date
    }
    for idx, name in enumerate(all_providers, start=1):
        row[f"hd_am_d1_{idx}"] = name

    out_df = pd.DataFrame([row])

    # 6️⃣ Show & download
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
        1. In REDCap, define a *single* repeating instrument (or classic form) called `hope_drive`.  
        2. Add fields:  
           - `hd_day_date` (Date Y‑M‑D)  
           - `hd_am_d1_1`, `hd_am_d1_2`, … (Text)  
        3. Use this CSV in the Data Import Tool or via the API.
        """
    )

elif not record_id:
    st.info("Enter a record_id so I can build the import row for you.")
else:
    st.info("Upload an Excel file to get started.")
