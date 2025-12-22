import streamlit as st
import pandas as pd

st.title("Datetime Fix + Duration Calculator")

uploaded_file = st.file_uploader(
    "Upload Excel file",
    type=["xlsx"]
)

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.subheader("Original data")
    st.dataframe(df)

    # ===============================
    # 1. Force datetime conversion
    # ===============================
    date_cols = [
        c for c in df.columns
        if "date" in c.lower() or "time" in c.lower()
    ]

    for col in date_cols:
        df[col] = pd.to_datetime(df[col], errors="coerce")

    st.success(f"Converted to datetime: {date_cols}")

    # ===============================
    # 2. Example difference (edit these)
    # ===============================
    START_COL = "Created"
    END_COL = "Sample_Completed_Date"

    if START_COL in df.columns and END_COL in df.columns:
        df["Duration_hours"] = (
            (df[END_COL] - df[START_COL])
            .dt.total_seconds()
            / 3600
        )

        df["Duration_hh_mm"] = (
            pd.to_timedelta(df["Duration_hours"], unit="h")
            .astype(str)
            .str.split(" ").str[-1]
        )

    # ===============================
    # 3. Show result
    # ===============================
    st.subheader("Fixed data")
    st.dataframe(df)

    # ===============================
    # 4. Download
    # ===============================
    output_file = "fixed_datetime_output.xlsx"
    df.to_excel(output_file, index=False)

    with open(output_file, "rb") as f:
        st.download_button(
            "Download fixed Excel",
            f,
            file_name=output_file
        )
