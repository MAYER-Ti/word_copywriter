from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QFileDialog, QMessageBox
from docx import Document

from parsers import read_data_from_file
from doc_utils import format_preview, replace_placeholders


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
        path, _ = QFileDialog.getOpenFileName(
            self, "Select source", filter="Documents (*.docx *.pdf)"
        )
        if path:
            self.source_edit.setText(path)
            self.data = read_data_from_file(path)
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
            data = read_data_from_file(source_path)
        doc = Document(template_path)
        replace_placeholders(doc, data)
        try:
            doc.save(output_path)
            QMessageBox.information(self, "Success", f"Document saved to {output_path}")
        except PermissionError:
            QMessageBox.critical(
                self,
                "Error",
                "Cannot save file. It may be open in another program.",
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
