from io import BytesIO
from openpyxl import load_workbook

buffer = BytesIO()

df.to_excel(
    buffer,
    index=False,
    engine="openpyxl"
)

buffer.seek(0)

# Force Excel datetime format (date + time)
wb = load_workbook(buffer)
ws = wb.active

for col in ws.columns:
    header = col[0].value
    if header and any(k in header.lower() for k in ["date", "time", "created", "completed"]):
        for cell in col[1:]:
            if cell.value:
                cell.number_format = "yyyy-mm-dd hh:mm"

buffer = BytesIO()
wb.save(buffer)
buffer.seek(0)

st.download_button(
    "Download fixed Excel",
    buffer,
    file_name="fixed_datetime_output.xlsx"
)
