import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Excel Datetime Fix Tool")

uploaded_file = st.file_uploader(
    "Upload Excel file",
    type=["xlsx"]
)

if uploaded_file:
    df = pd.read_excel(uploaded_file, dtype=str)

    st.subheader("Original data")
    st.dataframe(df)

    # ===============================
    # 1. Detect date/time columns
    # ===============================
    date_cols = [
        c for c in df.columns
        if any(k in c.lower() for k in ["date", "time", "created", "completed"])
    ]

    # ===============================
    # 2. Force ISO + text → datetime
    # ===============================
    for col in date_cols:
        df[col] = (
            df[col]
            .str.replace("T", " ", regex=False)
            .str.replace("Z", "", regex=False)
            .pipe(pd.to_datetime, errors="coerce")
        )

    st.success(f"Fixed datetime columns: {date_cols}")

    # ===============================
    # 3. Example duration
    # ===============================
    START_COL = "Created"
    END_COL = "Pattern_Complete_Date"

    if START_COL in df.columns and END_COL in df.columns:
        df["Duration_hours"] = (
            (df[END_COL] - df[START_COL])
            .dt.total_seconds()
            / 3600
        )

    # ===============================
    # 4. Export (Excel-safe)
    # ===============================
    buffer = BytesIO()
    df.to_excel(
        buffer,
        index=False,
        engine="openpyxl",
        datetime_format="yyyy-mm-dd hh:mm"
    )
    buffer.seek(0)

    st.subheader("Fixed data")
    st.dataframe(df)

    st.download_button(
        "Download fixed Excel",
        buffer,
        file_name="fixed_datetime_output.xlsx"
    )
