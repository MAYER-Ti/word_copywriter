import sys
import os
from PyQt5 import QtWidgets, QtGui

from gui import MainWindow


def main():
    app = QtWidgets.QApplication(sys.argv)
    icon_path = os.path.join(os.path.dirname(__file__), "resources", "icon.png")
    app.setWindowIcon(QtGui.QIcon(icon_path))
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
