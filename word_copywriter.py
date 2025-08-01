import sys
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QFileDialog, QMessageBox
from docx import Document


def read_data_from_docx(path):
    """Read key:value pairs from a docx document."""
    doc = Document(path)
    data = {}
    for para in doc.paragraphs:
        text = para.text.strip()
        if ':' in text:
            key, value = text.split(':', 1)
            key = key.strip().strip('{}')
            data[key] = value.strip()
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


class MainWindow(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Word Copywriter")
        layout = QtWidgets.QGridLayout()

        # Source document
        self.source_edit = QtWidgets.QLineEdit()
        source_btn = QtWidgets.QPushButton("Browse")
        source_btn.clicked.connect(self.browse_source)
        layout.addWidget(QtWidgets.QLabel("Source"), 0, 0)
        layout.addWidget(self.source_edit, 1, 0)
        layout.addWidget(source_btn, 2, 0)

        # Template document
        self.template_edit = QtWidgets.QLineEdit()
        template_btn = QtWidgets.QPushButton("Browse")
        template_btn.clicked.connect(self.browse_template)
        layout.addWidget(QtWidgets.QLabel("Template"), 0, 1)
        layout.addWidget(self.template_edit, 1, 1)
        layout.addWidget(template_btn, 2, 1)

        # Output path and copy button
        self.output_edit = QtWidgets.QLineEdit()
        output_btn = QtWidgets.QPushButton("Browse")
        output_btn.clicked.connect(self.browse_output)
        copy_btn = QtWidgets.QPushButton("Copy")
        copy_btn.clicked.connect(self.copy_data)
        layout.addWidget(QtWidgets.QLabel("Output"), 0, 2)
        layout.addWidget(self.output_edit, 1, 2)
        layout.addWidget(output_btn, 2, 2)
        layout.addWidget(copy_btn, 3, 1)

        self.setLayout(layout)

    def browse_source(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select source", filter="Word Documents (*.docx)")
        if path:
            self.source_edit.setText(path)

    def browse_template(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select template", filter="Word Documents (*.docx)")
        if path:
            self.template_edit.setText(path)

    def browse_output(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save document", filter="Word Documents (*.docx)")
        if path:
            if not path.lower().endswith('.docx'):
                path += '.docx'
            self.output_edit.setText(path)

    def copy_data(self):
        source_path = self.source_edit.text()
        template_path = self.template_edit.text()
        output_path = self.output_edit.text()
        if not (source_path and template_path and output_path):
            QMessageBox.warning(self, "Warning", "Please select all files")
            return
        data = read_data_from_docx(source_path)
        doc = Document(template_path)
        replace_placeholders(doc, data)
        doc.save(output_path)
        QMessageBox.information(self, "Success", f"Document saved to {output_path}")


def main():
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
