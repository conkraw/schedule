import streamlit as st
import pandas as pd

st.set_page_config(page_title="Hope Drive → REDCap Import", layout="wide")
st.title("Hope Drive Preceptor Extractor")

uploaded_file = st.file_uploader("Upload your AGP calendar (Excel)", type=["xlsx", "xls"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, header=None, dtype=str)

    # 1. Extract date from cell A5
    try:
        session_date = pd.to_datetime(df.iat[4, 0]).date()
    except Exception:
        st.error("⚠️ Could not parse a valid date from cell A5.")
        st.stop()

    # 2. Iterate through column A and B from row 5 onward
    designations = []
    for i in range(5, len(df)):
        designation = str(df.iat[i, 0]).replace("\xa0", " ").strip().lower()
        provider = str(df.iat[i, 1]).replace("\xa0", " ").strip()

        if "monday" in designation and i > 5:
            break  # stop when we hit the next week's block

        if designation:  # skip blanks
            designations.append({
                "date": session_date,
                "designation": designation.title(),
                "provider": provider
            })

    if not designations:
        st.warning("No designation-provider pairs found.")
    else:
        output_df = pd.DataFrame(designations)
        st.subheader("✅ Extracted Assignments")
        st.dataframe(output_df)

        # Download
        csv = output_df.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Download CSV", csv, "hope_drive_extracted.csv", "text/csv")

else:
    st.info("Please upload an Excel file to begin.")


