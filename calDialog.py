from PyQt5.QtWidgets import QDialog, QMessageBox, QComboBox, QLabel, QInputDialog, QLineEdit
from calDialogUI import *
import serial
from serial import Serial
from serial.tools import list_ports
from communicationThread import *
#from settingsDialog import *
from PyQt5 import QtWidgets
from PyQt5.QtCore import QPoint
import PyQt5.QtCore as QtCore
from PyQt5 import QtGui
from PyQt5.QtGui import QIcon, QPixmap
from settingsDialog import *
from mainWindow import *
import time
import os,sys



class CalDialog(QDialog):
    def __init__(self,dataType,currentRng):
        super().__init__()
        self.cal = Ui_Dialog2()
        self.cal.setupUi(self)
        self.data = dataType
        self.currentVal = currentRng
        
        

        label = QLabel(self)
        self.left = 70
        self.top = 50
        self.width = 321
        self.height = 371
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.cal.pushButtonNext.clicked.connect(self.accept)
        label = QLabel(self)
        label.move(self.left, self.top)
        
        def img_resource_path(relative_path):
            """ Get absolute path to resource, works for dev and for PyInstaller """
            try:
                # PyInstaller creates a temp folder and stores path in _MEIPASS
                base_path = sys._MEIPASS
            except Exception:
                base_path = os.path.abspath(".")
            return os.path.join(base_path, relative_path)
        
        img_pathRES = img_resource_path("img/kalRES.jpg")
        img_pathACI = img_resource_path("img/kalACI.jpg")
        img_pathACI20 = img_resource_path("img/kalACI20.jpg")
        img_pathACV = img_resource_path("img/kalACI.jpg")
        img_pathDCV = img_resource_path("img/kalACI.jpg")
        img_pathDCI = img_resource_path("img/kalACI.jpg")
        img_pathDCI20 = img_resource_path("img/kalACI20.jpg")
        img_pathTEMP = img_resource_path("img/kalACI.jpg")
        img_pathFREQ = img_resource_path("img/kalFREQ.jpg")
        
        if(self.data == 'RES'):
            pixmap = QPixmap(img_pathRES)
            label.setPixmap(pixmap)

        elif(self.data == 'ACI'):
            if(self.currentVal == 20):
                pixmap = QPixmap(img_pathACI20)
                label.setPixmap(pixmap)
            else:
                pixmap = QPixmap(img_pathACI)
                label.setPixmap(pixmap)

        elif(self.data == 'ACV'):
            pixmap = QPixmap(img_pathACV)
            label.setPixmap(pixmap)

        elif(self.data == 'DCV'):
            pixmap = QPixmap(img_pathDCV)
            label.setPixmap(pixmap)

        elif(self.data == 'FREQ'):
            pixmap = QPixmap(img_pathFREQ)
            label.setPixmap(pixmap)
            
        elif(self.data == 'TEMP'):
            pixmap = QPixmap(img_pathTEMP)
            label.setPixmap(pixmap)

        elif(self.data == 'DCI'):
            if(self.currentVal == 20):
                pixmap = QPixmap(img_pathDCI20)
                label.setPixmap(pixmap)
            else:
                pixmap = QPixmap(img_pathDCI)
                label.setPixmap(pixmap)
         


    




       