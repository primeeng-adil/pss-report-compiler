import sys
from PyQt5.QtWidgets import QApplication
from interface import Interface
from pdf_inserter import *

if __name__ == "__main__":
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        app_path = Path(sys._MEIPASS)
    else:
        app_path = Path(sys.argv[0]).parent
    app = QApplication(sys.argv)
    window = Interface(app_path)
    sys.exit(app.exec())

