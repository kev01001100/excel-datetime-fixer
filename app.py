import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
import datetime as dt

st.set_page_config(page_title="Excel Datetime Fix Tool")
st.title("Excel Datetime Fix Tool")

uploaded_file = st.file_uploader(
    "Upload Excel file",
    type=["xlsx"]
)

if uploaded_file:
    # 1) Read normally (let pandas guess types)
    df = pd.read_excel(uploaded_file)

    st.subheader("Original data")
    st.dataframe(df)

    # 2) Find date/datetime-like columns
    date_cols = [
        c for c in df.columns
        if any(k in c.lower() for k in ["date", "time", "created", "completed"])
    ]

    # 3) Force everything in those columns to real datetime
    for col in date_cols:
        s = df[col].astype(str)
        s = s.replace({
            "NaT": None,
            "nan": None,
            "None": None,
            "(blank)": None,
            ""
        })
        s = s.str.replace("T", " ", regex=False)\
             .str.replace("Z", "", regex=False)

        df[col] = pd.to_datetime(s, errors="coerce")

    st.success(f"Converted to datetime: {date_cols}")

    # 4) Example duration (hours) between Created and Pattern_Complete_Date
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

    # 5) Export to Excel with proper datetime format
    buffer = BytesIO()
    df.to_excel(buffer, index=False, engine="openpyxl")
    buffer.seek(0)

    wb = load_workbook(buffer)
    ws = wb.active

    for col_cells in ws.columns:
        header = col_cells[0].value
        if header and any(k in str(header).lower() for k in ["date", "time", "created", "completed"]):
            for cell in col_cells[1:]:
                if isinstance(cell.value, dt.datetime):
                    cell.number_format = "yyyy-mm-dd hh:mm"

    final_buffer = BytesIO()
    wb.save(final_buffer)
    final_buffer.seek(0)

    st.download_button(
        "Download fixed Excel",
        final_buffer,
        file_name="fixed_datetime_output.xlsx"
    )
