from io import BytesIO
from openpyxl import load_workbook, Workbook
import xlrd
import xlwt
from xlutils.copy import copy as xl_copy


def replace_placeholders(wb, data):
    if isinstance(wb, Workbook):
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str):
                        for key, value in data.items():
                            placeholder = f"{{{{{key}}}}}"
                            if placeholder in cell.value:
                                cell.value = cell.value.replace(placeholder, value)
    elif isinstance(wb, xlwt.Workbook):
        book = getattr(wb, "_xlrd_book")
        for idx, sheet in enumerate(book.sheets()):
            ws = wb.get_sheet(idx)
            for r in range(sheet.nrows):
                for c in range(sheet.ncols):
                    cell_value = sheet.cell_value(r, c)
                    if isinstance(cell_value, str):
                        new_value = cell_value
                        for key, value in data.items():
                            placeholder = f"{{{{{key}}}}}"
                            if placeholder in new_value:
                                new_value = new_value.replace(placeholder, value)
                        if new_value != cell_value:
                            ws.write(r, c, new_value)
    else:
        raise TypeError("Unsupported workbook type")


def create_document(template_bytes, ext, data, output_path):
    if ext == ".xls":
        book = xlrd.open_workbook(file_contents=template_bytes, formatting_info=True)
        wb = xl_copy(book)
        wb._xlrd_book = book
    else:
        wb = load_workbook(BytesIO(template_bytes))
    replace_placeholders(wb, data)
    wb.save(output_path)
