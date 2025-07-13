import streamlit as st
import pandas as pd

st.set_page_config(page_title="Hope Drive → REDCap Import", layout="wide")

st.title("Hope Drive Preceptors → REDCap Import Template")

# 1. File uploader
uploaded_file = st.file_uploader(
    "Upload your Academic General Pediatrics calendar (Excel)", 
    type=["xlsx", "xls"]
)

    # 1. Read raw, no headers
    df_raw = pd.read_excel(uploaded_file, header=None)

    # 2. Date cell (A5)
    hd_day_date = pd.to_datetime(df_raw.iat[4, 0]).date()

    # 3. Locate header row and columns
    header_row = None
    hm_cols = []
    for i in range(min(20, df_raw.shape[0])):  # look in first 20 rows
        row_vals = df_raw.iloc[i].astype(str)
        if row_vals.str.contains("Hope Drive AM Continuity", na=False).any():
            header_row = i
            hm_cols = [
                col for col, val in row_vals.items()
                if "Hope Drive AM Continuity" in val
            ]
            break

    if header_row is None:
        st.error("Could not find any 'Hope Drive AM Continuity' header.")
        st.stop()

    # 4. Parse Day 1 providers from row header_row + 1
    assignments = []
    for col in hm_cols:
        cell = df_raw.iat[header_row + 1, col]
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
