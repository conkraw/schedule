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
        "hope drive am continuity": ["hd_am_d1_"],
        "hope drive pm continuity": ["hd_pm_d1_"],
        "hope drive am acute precept": ["hd_am_acute_d1_"],
        "hope drive pm acute precept": ["hd_pm_acute_d1_"],
        "etown am continuity": ["etown_am_d1_"],
        "etown pm continuity": ["etown_pm_d1_"],
        "nyes rd am continuity": ["nyes_am_d1_"],
        "nyes rd pm continuity": ["nyes_pm_d1_"],
        "nursery weekday 8a-6p": ["nursery_am_d1_", "nursery_pm_d1_"]
    }

    # 3. Initialize result dictionary
    redcap_row = {"record_id": record_id, "hd_day_date1": session_date}

    # 4. Collect all providers by designation until next "Monday"
    grouped_data = {k: [] for k in designation_map.keys()}

    for i in range(5, len(df)):
        designation = str(df.iat[i, 0]).replace("\xa0", " ").strip().lower()
        provider = str(df.iat[i, 1]).replace("\xa0", " ").strip()

        if "monday" in designation and i > 5:
            break

        if designation in grouped_data and provider:
            grouped_data[designation].append(provider)

    # 5. Post-process: Ensure minimums by duplicating if needed
    min_required = {
        "hope drive am acute precept": 2,
        "hope drive pm acute precept": 2,
        "nursery weekday 8a-6p": 2,
    }

    for designation, providers in grouped_data.items():
        required = min_required.get(designation, len(providers))

        while len(providers) < required and providers:
            providers.append(providers[0])  # Duplicate the first provider

        for i, provider in enumerate(providers, start=1):
            for prefix in designation_map[designation]:
                field_name = f"{prefix}{i}"
                redcap_row[field_name] = provider

    # 6. Create and display output DataFrame
    out_df = pd.DataFrame([redcap_row])
    st.subheader("✅ REDCap Import Preview")
    st.dataframe(out_df)

    # 7. Download option
    csv = out_df.to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Download CSV", csv, "hope_drive_import.csv", "text/csv")

elif uploaded_file and not record_id:
    st.info("Please enter a record_id so I can build the import row for you.")
else:
    st.info("Upload an Excel file and enter a record_id to get started.")


