import streamlit as st
import pandas as pd

st.set_page_config(page_title="Hope Drive → REDCap Import", layout="wide")
st.title("Hope Drive Preceptors → REDCap Import Template")

uploaded_file = st.file_uploader(
    "Upload your AGP calendar (Excel)", 
    type=["xlsx", "xls"]
)
record_id = st.text_input("Enter REDCap record_id for this session", "")

if uploaded_file:
    df = pd.read_excel(uploaded_file, header=None)

    # 1. Extract date from A5
    try:
        hd_day_date = pd.to_datetime(df.iat[4, 0]).date()
    except Exception:
        st.error("⚠️ Could not parse a valid date from cell A5.")
        st.stop()

    # 2. Scan for every "Hope Drive AM Continuity" in rows 0–19
    providers = []
    max_rows = min(20, df.shape[0])
    for r in range(max_rows):
        for c in range(df.shape[1] - 1):
            if str(df.iat[r, c]).strip() == "Hope Drive AM Continuity":
                prov = df.iat[r, c + 1]
                if pd.notna(prov):
                    providers.append(str(prov).strip())

    if not providers:
        st.error("⚠️ No 'Hope Drive AM Continuity' rows found in the first 20 rows.")
        st.stop()

    # 3. Build single‐row REDCap import
    row = {
        "record_id": record_id,
        "hd_day_date": hd_day_date
    }
    for idx, name in enumerate(providers, start=1):
        row[f"hd_am_d1_{idx}"] = name

    out_df = pd.DataFrame([row])

    # 4. Display & download
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
        1. In REDCap, define a repeating form/instrument called `hope_drive`.  
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

