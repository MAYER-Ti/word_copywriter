from io import BytesIO
from openpyxl import load_workbook, Workbook
from openpyxl.utils.exceptions import InvalidFileException
import xlrd
import xlwt


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
    if ext == ".xls":
        book = xlrd.open_workbook(file_contents=template_bytes)
        wb = Workbook()
        wb.remove(wb.active)
        for sheet in book.sheets():
            ws = wb.create_sheet(title=sheet.name)
            for r in range(sheet.nrows):
                for c in range(sheet.ncols):
                    ws.cell(row=r + 1, column=c + 1, value=sheet.cell_value(r, c))
    else:
        try:
            wb = load_workbook(BytesIO(template_bytes))
        except InvalidFileException:
            book = xlrd.open_workbook(file_contents=template_bytes)
            wb = Workbook()
            wb.remove(wb.active)
            for sheet in book.sheets():
                ws = wb.create_sheet(title=sheet.name)
                for r in range(sheet.nrows):
                    for c in range(sheet.ncols):
                        ws.cell(row=r + 1, column=c + 1, value=sheet.cell_value(r, c))
    replace_placeholders(wb, data)
    if output_path.lower().endswith(".xls"):
        out_wb = xlwt.Workbook()
        for ws in wb.worksheets:
            out_ws = out_wb.add_sheet(ws.title[:31])
            for r_idx, row in enumerate(ws.iter_rows(values_only=True)):
                for c_idx, value in enumerate(row):
                    out_ws.write(r_idx, c_idx, value)
        out_wb.save(output_path)
    else:
        wb.save(output_path)
