from PyQt5 import QtWidgets, QtGui
from PyQt5.QtWidgets import QFileDialog, QMessageBox
from PyQt5.QtCore import QSettings, Qt
import os

from parsers import read_data_from_file
from doc_utils import format_preview
from excel_utils import create_document as generate_document


class AboutWidget(QtWidgets.QWidget):
    """Simple widget showing information about the program."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("О программе")

        layout = QtWidgets.QHBoxLayout()

        text = (
            "<h2>Copywriter</h2>"
            "<p>Утилита для заполнения шаблонов на основе данных из документов.</p>"
            "<p>Мы небольшой стартап, который специализируется на автоматизации производства</p>"
            "<p>Больше о нас на <a href='https://project14096453.tilda.ws/'>сайте</a></p>"
            "<p>Github разработчика <a href='https://github.com/MAYER-Ti'>здесь</a></p>"
        )
        text_label = QtWidgets.QLabel(text)
        text_label.setWordWrap(True)
        text_label.setOpenExternalLinks(True)

        icon_path = os.path.join(os.path.dirname(__file__), "resources", "icon.png")
        max_size = 128
        pixmap = QtGui.QPixmap(icon_path).scaled(
            max_size, max_size, Qt.KeepAspectRatio, Qt.SmoothTransformation
        )
        image_label = QtWidgets.QLabel()
        image_label.setPixmap(pixmap)
        image_label.setAlignment(Qt.AlignCenter)
        image_label.setFixedSize(pixmap.size())

        layout.addWidget(text_label, 1)
        layout.addWidget(image_label, 1)
        self.setLayout(layout)


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings("word_copywriter", "templates")
        self.templates = {}
        self.data = {}
        self.status_label = QtWidgets.QLabel()
        icon_path = os.path.join(os.path.dirname(__file__), "resources", "icon.png")
        self.setWindowIcon(QtGui.QIcon(icon_path))
        self.init_ui()
        self.load_templates()

    def init_ui(self):
        self.setWindowTitle("Word Copywriter")

        # Toolbar with settings and status
        toolbar = self.addToolBar("Main")
        settings_button = QtWidgets.QToolButton()
        settings_button.setText("Настройки")
        settings_menu = QtWidgets.QMenu()
        act_action = settings_menu.addAction("Загрузить шаблон акта")
        act_action.triggered.connect(self.browse_act_template)
        invoice_action = settings_menu.addAction("Загрузить шаблон счёта")
        invoice_action.triggered.connect(self.browse_invoice_template)
        settings_button.setMenu(settings_menu)
        settings_button.setPopupMode(QtWidgets.QToolButton.InstantPopup)
        toolbar.addWidget(settings_button)
        toolbar.addSeparator()
        toolbar.addWidget(self.status_label)

        spacer = QtWidgets.QWidget()
        spacer.setSizePolicy(
            QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding
        )
        toolbar.addWidget(spacer)

        about_button = QtWidgets.QToolButton()
        about_button.setText("О нас")
        about_button.clicked.connect(self.show_about)
        toolbar.addWidget(about_button)

        # Central widget layout
        central = QtWidgets.QWidget()
        self.setCentralWidget(central)
        layout = QtWidgets.QVBoxLayout()

        source_layout = QtWidgets.QHBoxLayout()
        source_layout.addWidget(QtWidgets.QLabel("Документ"))
        self.source_edit = QtWidgets.QLineEdit()
        self.source_edit.setReadOnly(True)
        source_layout.addWidget(self.source_edit)
        source_btn = QtWidgets.QPushButton("Найти")
        source_btn.clicked.connect(self.browse_source)
        source_layout.addWidget(source_btn)
        layout.addLayout(source_layout)

        self.preview_edit = QtWidgets.QTextEdit()
        self.preview_edit.setReadOnly(True)
        layout.addWidget(self.preview_edit)

        buttons_layout = QtWidgets.QHBoxLayout()
        self.create_act_btn = QtWidgets.QPushButton("Создать акт")
        self.create_act_btn.clicked.connect(self.create_act)
        self.create_act_btn.setEnabled(False)
        buttons_layout.addWidget(self.create_act_btn)

        self.create_invoice_btn = QtWidgets.QPushButton("Создать счёт")
        self.create_invoice_btn.clicked.connect(self.create_invoice)
        self.create_invoice_btn.setEnabled(False)
        buttons_layout.addWidget(self.create_invoice_btn)
        layout.addLayout(buttons_layout)
        
        central.setLayout(layout)

    def set_status(self, message):
        self.status_label.setText(message)

    def load_template(self, key):
        path = self.settings.value(f"{key}_template", "", type=str)
        name = "акта" if key == "act" else "счёта"
        if path and os.path.exists(path):
            ext = os.path.splitext(path)[1].lower()
            if ext != ".xlsx":
                self.templates[key] = None
                self.set_status(f"Неверный формат шаблона {name}")
                return
            try:
                with open(path, "rb") as f:
                    self.templates[key] = {"bytes": f.read(), "ext": ext}
                self.set_status(f"Шаблон {name} загружен")
            except Exception as e:
                self.templates[key] = None
                self.set_status(f"Ошибка загрузки {name}: {e}")
        else:
            self.templates[key] = None

    def load_templates(self):
        self.templates = {}
        for key in ("act", "invoice"):
            self.load_template(key)
        self.update_create_buttons_state()

    def browse_source(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select source",
            filter="Documents (*.docx *.pdf)",
        )
        if path:
            self.source_edit.setText(path)
            self.data = read_data_from_file(path)
            self.preview_edit.setPlainText(format_preview(self.data))
        self.update_create_buttons_state()

    def browse_act_template(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select act template",
            filter="Excel Files (*.xlsx)",
        )
        if path:
            self.settings.setValue("act_template", path)
            self.load_template("act")
        self.update_create_buttons_state()

    def browse_invoice_template(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select invoice template",
            filter="Excel Files (*.xlsx)",
        )
        if path:
            self.settings.setValue("invoice_template", path)
            self.load_template("invoice")
        self.update_create_buttons_state()

    def create_act(self):
        self.create_document("act")

    def create_invoice(self):
        self.create_document("invoice")

    def create_document(self, key):
        template_info = self.templates.get(key)
        source_path = self.source_edit.text()
        if not (template_info and source_path):
            QMessageBox.warning(self, "Warning", "Please select source and templates")
            return
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save document",
            filter="Excel Workbook (*.xlsx)",
        )
        if not output_path:
            return
        if not output_path.lower().endswith(".xlsx"):
            output_path += ".xlsx"
        data = self.data or read_data_from_file(source_path)
        try:
            generate_document(
                template_info["bytes"], template_info["ext"], data, output_path
            )
            QMessageBox.information(self, "Success", f"Document saved to {output_path}")
        except PermissionError:
            QMessageBox.critical(
                self,
                "Error",
                "Cannot save file. It may be open in another program.",
            )
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save file: {e}")

    def update_create_buttons_state(self):
        source_selected = bool(self.source_edit.text())
        self.create_act_btn.setEnabled(source_selected and bool(self.templates.get("act")))
        self.create_invoice_btn.setEnabled(
            source_selected and bool(self.templates.get("invoice"))
        )

    def show_about(self):
        self.about_widget = AboutWidget()
        self.about_widget.show()
