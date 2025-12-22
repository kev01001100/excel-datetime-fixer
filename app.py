import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook

st.set_page_config(page_title="Excel Datetime Fix Tool")
st.title("Excel Datetime Fix Tool")

uploaded_file = st.file_uploader(
    "Upload Excel file",
    type=["xlsx"]
)

if uploaded_file:
    # -------------------------------------------------
    # 1. Read everything as text first (important)
    # -------------------------------------------------
    df = pd.read_excel(uploaded_file, dtype=str)

    st.subheader("Original data")
    st.dataframe(df)

    # -------------------------------------------------
    # 2. Detect date / datetime columns
    # -------------------------------------------------
    date_cols = [
        c for c in df.columns
        if any(k in c.lower() for k in ["date", "time", "created", "completed"])
    ]

    # -------------------------------------------------
    # 3. Fix ISO strings → real datetime
    # -------------------------------------------------
    for col in date_cols:
        df[col] = (
            df[col]
            .str.replace("T", " ", regex=False)
            .str.replace("Z", "", regex=False)
            .pipe(pd.to_datetime, errors="coerce")
        )

    st.success(f"Converted to datetime: {date_cols}")

    # -------------------------------------------------
    # 4. Example duration (hours)
    # -------------------------------------------------
    START_COL = "Created"
    END_COL = "Pattern_Complete_Date"

    if START_COL in df.columns and END_COL in df.columns:
        df["Duration_hours"] = (
            (df[END_COL] - df[START_COL])
            .dt.total_seconds()
            / 3600
        )

    st.subheader("Fixed data (preview)")
    st.dataframe(df)

    # -------------------------------------------------
    # 5. Export to Excel (FORCE datetime format)
    # -------------------------------------------------
    buffer = BytesIO()
    df.to_excel(buffer, index=False, engine="openpyxl")
    buffer.seek(0)

    wb = load_workbook(buffer)
    ws = wb.active

    for col in ws.columns:
        header = col[0].value
        if header and any(k in header.lower() for k in ["date", "time", "created", "completed"]):
            for cell in col[1:]:
                if cell.value is not None:
                    cell.number_format = "yyyy-mm-dd hh:mm"

    final_buffer = BytesIO()
    wb.save(final_buffer)
    final_buffer.seek(0)

    st.download_button(
        "Download fixed Excel",
        final_buffer,
        file_name="fixed_datetime_output.xlsx"
    )
