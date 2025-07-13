import streamlit as st
import pandas as pd

st.set_page_config(page_title="Hope Drive → REDCap Import", layout="wide")

st.title("Hope Drive Preceptors → REDCap Import Template")

# 1. File uploader
uploaded_file = st.file_uploader(
    "Upload your Academic General Pediatrics calendar (Excel)", 
    type=["xlsx", "xls"]
)

if uploaded_file:
    # 2. Read sheet without headers
    df_raw = pd.read_excel(uploaded_file, header=None)

    # 3. Extract the first date from cell A5 (row 4, col 0)
    try:
        hd_day_date = pd.to_datetime(df_raw.iat[4, 0]).date()
    except Exception as e:
        st.error(f"Could not parse date in A5: {e}")
        st.stop()

    # 4. Identify "Hope Drive AM Continuity" columns (assume header on row 4)
    header_row = 3
    hm_cols = [
        col
        for col, val in df_raw.iloc[header_row].items()
        if isinstance(val, str) and "Hope Drive AM Continuity" in val
    ]

    if not hm_cols:
        st.warning("No columns found with 'Hope Drive AM Continuity' in row 4.")
        st.stop()

    # 5. Parse providers from each column on day 1 (row 5)
    assignments = []
    for col in hm_cols:
        cell = df_raw.iat[4, col]
        providers = [p.strip() for p in str(cell).split(",") if p.strip()]
        assignments.append(providers)

    # 6. Build REDCap import DataFrame
    rows = []
    for instance, providers in enumerate(assignments, start=1):
        row = {
            "record_id": "YOUR_RECORD_ID",
            "redcap_repeat_instrument": "hope_drive",
            "redcap_repeat_instance": instance,
            "hd_day_date": hd_day_date,
        }
        for i, name in enumerate(providers, start=1):
            row[f"hd_am_d1_{i}"] = name
        rows.append(row)

    import_df = pd.DataFrame(rows)

    # 7. Show results
    st.subheader("📋 Import Template Preview")
    st.dataframe(import_df)

    st.markdown(
        """
        **Next steps:**  
        1. Define a repeating instrument in REDCap named **hope_drive**  
        2. Add fields:  
           - `hd_day_date` (Date/YMD)  
           - `hd_am_d1_1`, `hd_am_d1_2`, … (Text)  
        3. Export this table to CSV via **File → Download** or use the REDCap API to push it.  
        """
    )
else:
    st.info("Upload an Excel file to get started.")
