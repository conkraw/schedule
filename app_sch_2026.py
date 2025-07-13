import streamlit as st
import pandas as pd

st.set_page_config(page_title="Hope Drive → REDCap Import", layout="wide")
st.title("Hope Drive Preceptor Extractor → REDCap Format")

uploaded_file = st.file_uploader("Upload your AGP calendar (Excel)", type=["xlsx", "xls"])
record_id = st.text_input("Enter REDCap record_id for this session", "")

if uploaded_file and record_id:
    df = pd.read_excel(uploaded_file, header=None, dtype=str)

    # 1. Extract the session date from cell A5
    try:
        session_date = pd.to_datetime(df.iat[4, 0]).date()
    except Exception:
        st.error("⚠️ Could not parse a valid date from cell A5.")
        st.stop()

    # 2. Define designation → REDCap field prefix
    designation_map = {
        "hope drive am continuity": "hd_am_d1_",
        "hope drive pm continuity": "hd_pm_d1_",
        "hope drive am acute precept": "hd_am_acute_d1_",
        "hope drive pm acute precept": "hd_pm_acute_d1_",
    }

    # 3. Initialize result dictionary
    redcap_row = {"record_id": record_id, "hd_day_date1": session_date}
    field_counters = {prefix: 1 for prefix in designation_map.values()}

    # 4. Loop through rows starting at row 5 until we hit the next "Monday"
    for i in range(5, len(df)):
        designation = str(df.iat[i, 0]).replace("\xa0", " ").strip().lower()
        provider = str(df.iat[i, 1]).replace("\xa0", " ").strip()

        if "monday" in designation and i > 5:
            break

        if designation in designation_map and provider:
            field_prefix = designation_map[designation]
            field_name = f"{field_prefix}{field_counters[field_prefix]}"
            redcap_row[field_name] = provider
            field_counters[field_prefix] += 1

    # 5. Create and display output DataFrame
    out_df = pd.DataFrame([redcap_row])
    st.subheader("✅ REDCap Import Preview")
    st.dataframe(out_df)

    # 6. Download option
    csv = out_df.to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Download CSV", csv, "hope_drive_import.csv", "text/csv")

elif uploaded_file and not record_id:
    st.info("Please enter a record_id so I can build the import row for you.")
else:
    st.info("Upload an Excel file and enter a record_id to get started.")


