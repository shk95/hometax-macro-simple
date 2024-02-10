import sys

from PySide6 import QtWidgets

from hometax_macro_simple.gui import MyWidget
from hometax_macro_simple.webdriver import WebDriverManager


def main():
    app = QtWidgets.QApplication([])
    webdriver = WebDriverManager()

    widget = MyWidget(webdriver)
    widget.resize(800, 600)
    widget.show()

    app.aboutToQuit.connect(webdriver.close)
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
