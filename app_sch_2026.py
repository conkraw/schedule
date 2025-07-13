import streamlit as st
import pandas as pd

st.set_page_config(page_title="Hope Drive → REDCap Import", layout="wide")
st.title("Hope Drive Preceptors → REDCap Import Template")

# upload + record_id
uploaded_file = st.file_uploader("Upload your AGP calendar (Excel)", type=["xlsx","xls"])
record_id    = st.text_input("Enter REDCap record_id for this session", "")

if uploaded_file:
    df = pd.read_excel(uploaded_file, header=None)

    # 1️⃣ Get the session date from A5 (row 4)
    try:
        hd_day_date = pd.to_datetime(df.iat[4, 0]).date()
    except Exception:
        st.error("⚠️ Could not parse a valid date in cell A5.")
        st.stop()

    # 2️⃣ Find the row index of that date (first_date_row)
    first_date_row = None
    for r in range(df.shape[0]):
        try:
            if pd.to_datetime(df.iat[r, 0]).date() == hd_day_date:
                first_date_row = r
                break
        except:
            continue

    if first_date_row is None:
        st.error("⚠️ Could not locate the date row in column A.")
        st.stop()

    # 3️⃣ Find where the *next* date appears, so we know where to stop
    next_date_row = df.shape[0]
    for r in range(first_date_row + 1, df.shape[0]):
        try:
            d = pd.to_datetime(df.iat[r, 0]).date()
            if d != hd_day_date:
                next_date_row = r
                break
        except:
            continue

    # 4️⃣ Scan *only* between those two rows for your header + provider to the right
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

    # 5️⃣ Build the single-row REDCap import
    data = {
        "record_id": record_id,
        "hd_day_date": hd_day_date
    }
    for i, name in enumerate(providers, start=1):
        data[f"hd_am_d1_{i}"] = name

    out_df = pd.DataFrame([data])

    # 6️⃣ Display & download
    st.subheader("📋 REDCap Import Preview")
    st.dataframe(out_df)

    csv = out_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Download import CSV",
        data=csv,
        file_name="hope_drive_import.csv",
        mime="text/csv"
    )

elif not record_id:
    st.info("Enter a record_id so I can build the import row for you.")
else:
    st.info("Upload an Excel file to get started.")

