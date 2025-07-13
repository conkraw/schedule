import streamlit as st
import pandas as pd

st.set_page_config(page_title="Hope Drive → REDCap Import", layout="wide")
st.title("Hope Drive Preceptors → REDCap Import Template")

uploaded_file = st.file_uploader("Upload your AGP calendar (Excel)", type=["xlsx","xls"])
record_id    = st.text_input("Enter REDCap record_id for this session", "")

if uploaded_file and record_id:
    df = pd.read_excel(uploaded_file, header=None)

    # 1️⃣ Get the session date from A5
    try:
        hd_day_date = pd.to_datetime(df.iat[4, 0]).date()
    except Exception:
        st.error("⚠️ Could not parse a valid date from cell A5.")
        st.stop()

    # 2️⃣ Find all "Hope Drive AM Continuity" cells and pull the provider below
    providers = []
    rows, cols = df.shape
    for r in range(rows - 1):       # stop at rows-1 so r+1 is in-bounds
        for c in range(cols - 1):   # stop at cols-1 so c+1 exists
            cell = str(df.iat[r, c]).strip()
            if "Hope Drive AM Continuity" in cell:
                prov = df.iat[r + 1, c]
                if pd.notna(prov):
                    providers.append(str(prov).strip())

    if not providers:
        st.error("⚠️ No 'Hope Drive AM Continuity' entries found in the sheet.")
        st.stop()

    # 3️⃣ Build your REDCap import row
    data = {
        "record_id": record_id,
        "hd_day_date": hd_day_date
    }
    for i, name in enumerate(providers, start=1):
        data[f"hd_am_d1_{i}"] = name

    out_df = pd.DataFrame([data])

    # 4️⃣ Show & download
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
