import sys
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QFileDialog, QMessageBox
from docx import Document
import re


def read_data_from_docx(path):
    """Parse contract-like document and return values for placeholders."""
    doc = Document(path)

    data = {
        "Данные заказчика": "",
        "ИНН получателя": "",
        "ОГРН получателя": "",
        "Номер документа": "",
        "Адрес загрузки": "",
        "Адрес разгрузки": "",
        "Марка автомобиля": "",
        "Номер полуприцепа": "",
        "ФИО водителя": "",
        "Дата погрузки": "",
        "Дата разгрузки": "",
        "Стоимость перевозки": "",
    }

    # First paragraph with number
    for para in doc.paragraphs:
        text = para.text.strip()
        if text.startswith("Договор-заявка на перевозку груза") and text.endswith("года."):
            num_start = text.find("№")
            if num_start != -1:
                data["Номер документа"] = text[num_start:].strip()
            break

    # First table with route and cost info
    if doc.tables:
        t = doc.tables[0]
        try:
            data["Адрес загрузки"] = t.cell(8, 0).text.strip()
            data["Адрес разгрузки"] = t.cell(8, 4).text.strip()
            data["Дата погрузки"] = t.cell(10, 0).text.strip()
            data["Дата разгрузки"] = t.cell(10, 4).text.strip()
            data["Стоимость перевозки"] = t.cell(11, 4).text.strip()
        except IndexError:
            pass

    # Second table with vehicle info
    if len(doc.tables) > 1:
        t = doc.tables[1]
        try:
            data["Марка автомобиля"] = t.cell(0, 1).text.strip()
            data["Номер полуприцепа"] = t.cell(0, 2).text.strip()
            data["ФИО водителя"] = t.cell(1, 1).text.strip()
        except IndexError:
            pass

    # Third table with customer info
    if len(doc.tables) > 2:
        t = doc.tables[2]
        try:
            cell_text = t.cell(0, 1).text
        except IndexError:
            cell_text = ""
        if cell_text:
            start = cell_text.find("Заказчик:")
            end = cell_text.find("Почтовый адрес")
            if start != -1 and end != -1 and end > start:
                data["Данные заказчика"] = cell_text[start + len("Заказчик:"):end].strip()
            inn_match = re.search(r"ИНН получателя \d+", cell_text)
            if inn_match:
                data["ИНН получателя"] = inn_match.group(0).strip()
            ogrn_match = re.search(r"ОГРН \d+", cell_text)
            if ogrn_match:
                data["ОГРН получателя"] = ogrn_match.group(0).strip()

    return data


def replace_placeholders(doc, data):
    """Replace placeholders in doc with values from data."""
    for para in doc.paragraphs:
        for run in para.runs:
            for key, value in data.items():
                placeholder = f"{{{{{key}}}}}"
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, value)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_placeholders(cell, data)


def format_preview(data):
    """Return formatted string for preview widget."""
    lines = [
        data.get("Данные заказчика", ""),
        data.get("ИНН получателя", ""),
        data.get("ОГРН получателя", ""),
        f"Транспортные услуги по договору-заявке № {data.get('Номер документа', '')}",
        f"По маршруту {data.get('Адрес загрузки', '')} - {data.get('Адрес разгрузки', '')}",
        f"Автомобиль: {data.get('Марка автомобиля', '')} {data.get('Номер полуприцепа', '')}",
        f"Водитель: {data.get('ФИО водителя', '')}",
        f"Дата погрузки: {data.get('Дата погрузки', '')}",
        f"Дата разгрузки: {data.get('Дата разгрузки', '')}",
        f"Стоимость перевозки: {data.get('Стоимость перевозки', '')}",
    ]
    return "\n".join(lines)


class MainWindow(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.data = {}
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Word Copywriter")
        layout = QtWidgets.QGridLayout()

        # Source document
        self.source_edit = QtWidgets.QLineEdit()
        self.source_edit.setReadOnly(True)
        source_btn = QtWidgets.QPushButton("Browse")
        source_btn.clicked.connect(self.browse_source)
        layout.addWidget(QtWidgets.QLabel("Source"), 0, 0)
        layout.addWidget(self.source_edit, 1, 0)
        layout.addWidget(source_btn, 2, 0)
        self.preview_edit = QtWidgets.QTextEdit()
        self.preview_edit.setReadOnly(True)
        layout.addWidget(self.preview_edit, 3, 0)

        # Template document
        self.template_edit = QtWidgets.QLineEdit()
        self.template_edit.setReadOnly(True)
        template_btn = QtWidgets.QPushButton("Browse")
        template_btn.clicked.connect(self.browse_template)
        layout.addWidget(QtWidgets.QLabel("Template"), 0, 1)
        layout.addWidget(self.template_edit, 1, 1)
        layout.addWidget(template_btn, 2, 1)

        # Save button
        self.save_btn = QtWidgets.QPushButton("Save")
        self.save_btn.clicked.connect(self.save_document)
        self.save_btn.setEnabled(False)
        layout.addWidget(self.save_btn, 3, 1)

        self.setLayout(layout)

    def browse_source(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select source", filter="Word Documents (*.docx)")
        if path:
            self.source_edit.setText(path)
            self.data = read_data_from_docx(path)
            self.preview_edit.setPlainText(format_preview(self.data))
        self.update_save_button_state()

    def browse_template(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select template", filter="Word Documents (*.docx)")
        if path:
            self.template_edit.setText(path)
        self.update_save_button_state()

    def save_document(self):
        source_path = self.source_edit.text()
        template_path = self.template_edit.text()
        if not (source_path and template_path):
            QMessageBox.warning(self, "Warning", "Please select source and template files")
            return
        output_path, _ = QFileDialog.getSaveFileName(self, "Save document", filter="Word Documents (*.docx)")
        if not output_path:
            return
        if not output_path.lower().endswith('.docx'):
            output_path += '.docx'
        data = getattr(self, 'data', None)
        if not data:
            data = read_data_from_docx(source_path)
        doc = Document(template_path)
        replace_placeholders(doc, data)
        try:
            doc.save(output_path)
        except PermissionError:
            QMessageBox.critical(
                self,
                "Error",
                "Cannot save file. It may be open in another program."
            )
            return
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save file: {e}")
            return

    def update_save_button_state(self):
        if self.source_edit.text() and self.template_edit.text():
            self.save_btn.setEnabled(True)
        else:
            self.save_btn.setEnabled(False)


def main():
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
