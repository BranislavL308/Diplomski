import sys, os
from PyQt5.QtWidgets import QMainWindow, QMessageBox, QGridLayout, QTableWidget, QTableWidgetItem, QVBoxLayout, QApplication, QCheckBox, QComboBox
from mainWindow import *
from PyQt5.QtCore import QUrl
from mainWindowUI import *
from settingsDialog import *
from communicationThread import *


if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = MyForm()
    w.show()
    sys.exit(app.exec_())
    