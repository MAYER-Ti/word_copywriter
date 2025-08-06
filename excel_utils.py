from io import BytesIO
from openpyxl import load_workbook


def replace_placeholders(wb, data):
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    for key, value in data.items():
                        placeholder = f"{{{{{key}}}}}"
                        if placeholder in cell.value:
                            cell.value = cell.value.replace(placeholder, value)


def create_document(template_bytes, ext, data, output_path):
    if ext != ".xlsx":
        raise ValueError("Only .xlsx templates are supported")
    wb = load_workbook(BytesIO(template_bytes))
    replace_placeholders(wb, data)
    for ws in wb.worksheets:
        ws.sheet_state = "visible"
    wb.save(output_path)
