from PyQt5.QtWidgets import QDialog, QMessageBox, QComboBox
from settingsDialogUI import *
import serial
from serial import Serial
from serial.tools import list_ports
from communicationThread import *
from PyQt5 import QtWidgets
from PyQt5.QtCore import QPoint
import PyQt5.QtCore as QtCore
from PyQt5 import QtGui

class SettingsDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.dialog = Ui_Dialog1()
        self.dialog.setupUi(self)

        self.returnValue = {None}

        self.red = QtGui.QBrush(QtGui.QColor(255, 0, 0))
        self.red.setStyle(QtCore.Qt.SolidPattern)
        self.green = QtGui.QBrush(QtGui.QColor(0, 85, 0))
        self.green.setStyle(QtCore.Qt.SolidPattern)
        self.blue = QtGui.QBrush(QtGui.QColor(0, 0, 255))
        self.blue.setStyle(QtCore.Qt.SolidPattern)

        self.dialog.labelNeKal.setStyleSheet("background-color:#ff0000;")
        self.dialog.labelKal.setStyleSheet("background-color:#005500;")
        self.dialog.label_31.setStyleSheet("background-color:#0000ff;")
        

        #Primena filtera na tableWidget
        self.item = QtWidgets.QTableWidgetItem()
        self.dialog.tableWidget.viewport().installEventFilter(self)

        #self.dialog.pushButtonOK_next.clicked.connect(self.accept)

        self.dialog.tabWidget.setTabEnabled(0, True)
        self.dialog.tabWidget.setTabEnabled(1, False)
        self.dialog.tabWidget.setTabEnabled(2, False)


        
        self.dialog.pushButtonCheck.clicked.connect(self.checkData)
        

        self.dialog.pushButtonOK_next.setEnabled(False)#ZAMENITI u FALSE !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        self.dialog.pushButtonOKnext1.setEnabled(False) 
        
       
    

        self.dialog.pushButtonOKnext1.clicked.connect(self.nextTab1)
        self.dialog.pushButtonOK_next.clicked.connect(self.nextTab2)
        self.dialog.pushButtonTestComm.clicked.connect(self.testCom)
        self.dialog.pushButtonOK_next.clicked.connect(self.accept)


        self.podaci = {
            'DCV':[],
            'ACV': [], #[vrednosts 10, 90%, range, freq] #[[vr1, vr2,],freq,opseg],[[vr1, vr2,],freq,opseg],[[vr1, vr2,],freq,opseg]
            'DCI':[],
            'ACI':[],
            'RES':[],
            'FREQ':[],
            'TEMP':[]
        }
        #Kada kliknemo OK vrati nam 0
        #self.dialog.pushButton.clicked.connect(self.reject)
        #Kada kliknemo OK vrati nam 1
        #self.dialog.pushButtonOK.clicked.connect(self.accept)
        #Kada kliknemo Cancel  vrati nam 0
        #self.dialog.pushButtonCancel.clicked(self.reject)


        #self.dialog.pushButtonStartComm.setEnabled(False)
        #self.dialog.pushButtonStopComm.setEnabled(False)

        #self.dialog.pushButtonOK_next.clicked.connect(self.config)
        #self.dialog.pushButtonStartComm.clicked.connect(self.open_com)
        #self.dialog.pushButtonStopComm.clicked.connect(self.close_com)

        availablePorts = list_ports.comports()
        #Prodji kroz svaki objekat i upisi njegov device kod u listu
        a = list_ports.comports()
        for i in a:
            self.dialog.comboBoxPort.addItem(str(i.device))
        self.returnVal = {}
    
    
    def get_data(self):
        self.returnVal['com_port']=self.dialog.comboBoxPort.currentText()
        self.returnVal['baud_rate']=int(self.dialog.comboBoxBaudRate.currentText())
        
        tmp = self.dialog.comboBoxDataBits.currentText()
        if(tmp == '5'):
            self.returnVal['data_bits']=serial.FIVEBITS
        if(tmp == '6'):
            self.returnVal['data_bits']=serial.SIXBITS
        if(tmp == '7'):
            self.returnVal['data_bits']=serial.SEVENBITS
        if(tmp == '8'):
            self.returnVal['data_bits']=serial.EIGHTBITS

        tmp1 = self.dialog.comboBoxParityBits.currentText()
        if(tmp1 == 'None'):
            self.returnVal['parity']=serial.PARITY_NONE
        if(tmp1 == 'Odd'):
            self.returnVal['parity']=serial.PARITY_ODD
        if(tmp1 == 'Even'):
            self.returnVal['parity']=serial.PARITY_EVEN
        if(tmp1 == 'Mark'):
            self.returnVal['parity']=serial.PARITY_MARK
        if(tmp1 == 'Space'):
            self.returnVal['parity']=serial.PARITY_SPACE

        tmp2 = self.dialog.comboBoxStopBit.currentText()
        if(tmp2 =='1'):
            self.returnVal['stop_bit']=serial.STOPBITS_ONE
        if(tmp2 =='1.5'):
            self.returnVal['stop_bit']=serial.STOPBITS_ONE_POINT_FIVE
        if(tmp2 =='2'):
            self.returnVal['stop_bit']=serial.STOPBITS_TWO

        return self.returnVal

    def eventFilter(self, obj, event):
        cnt = 0
        cnt1 = 0
        cnt2 = 0
        cnt3 = 0
        if(event.type() == QtCore.QEvent.MouseButtonPress and event.button() == QtCore.Qt.RightButton):
            item = self.dialog.tableWidget.itemAt(event.pos())
            if(item != None):
 
                if(item.background().color() == QtGui.QColor(255, 0 , 0)): #Crvena boja:  Smallest range
                    item.setBackground(self.green)
                    item.setCheckState(2)
                    if(item.column() == 6):
                        self.dialog.tableWidget.item(item.row(), 7).setBackground(self.green)
                        self.dialog.tableWidget.item(item.row(), 7).setCheckState(2)
                    elif(item.column() == 7):
                        self.dialog.tableWidget.item(item.row(), 6).setBackground(self.green)
                        self.dialog.tableWidget.item(item.row(), 6).setCheckState(2)
                    if(item.column() == 5):
                        for i in range(4):
                            if(self.dialog.tableWidget.item(i,5).checkState() == 2):
                                cnt3 +=1
                            else:
                                pass
                        if(cnt3 > 1):
                            self.errorSonda()
                            item.setBackground(self.red)
                            item.setCheckState(0)


                elif(item.background().color() == QtGui.QColor(0, 85, 0)): 
                    if(item.column() < 3):
                        item.setBackground(self.blue)
                        item.setCheckState(1)
                        if(item.column() == 0):
                            for i in range(9):
                                if(self.dialog.tableWidget.item(i,0).checkState() == 1):
                                    cnt += 1
                                else:
                                    pass
                            if(cnt > 1):
                                self.selectOneIntermediateRangeDCV()
                                item.setBackground(self.red)
                                item.setCheckState(0)
                        if(item.column() == 1):
                            for i in range(9):
                                if(self.dialog.tableWidget.item(i,1).checkState() == 1):
                                    cnt1 += 1
                                else:
                                    pass
                            if(cnt1 > 1):
                                self.selectOneIntermediateRangeACV()
                                item.setBackground(self.red)
                                item.setCheckState(0)
                        if(item.column() == 2):
                            for i in range(9):
                                if(self.dialog.tableWidget.item(i,2).checkState() == 1):
                                    cnt2 += 1
                                else:
                                    pass
                            if(cnt2 > 1):
                                self.selectOneIntermediateRangeDCI()
                                item.setBackground(self.red)
                                item.setCheckState(0)     
                    else:
                        item.setBackground(self.red)
                        item.setCheckState(0)
                        if(item.column() == 6):
                            self.dialog.tableWidget.item(item.row(), 7).setBackground(self.red)
                            self.dialog.tableWidget.item(item.row(), 7).setCheckState(0)
                        elif(item.column() == 7):
                            self.dialog.tableWidget.item(item.row(), 6).setBackground(self.red)
                            self.dialog.tableWidget.item(item.row(), 6).setCheckState(0)

                elif(item.background().color() == QtGui.QColor(0, 0, 255)): 
                    item.setBackground(self.red)
                    item.setCheckState(0)
                    if(item.column() == 6):
                        self.dialog.tableWidget.item(item.row(), 7).setBackground(self.red)
                        self.dialog.tableWidget.item(item.row(), 7).setCheckState(0)
                    elif(item.column() == 7):
                        self.dialog.tableWidget.item(item.row(), 6).setBackground(self.red)
                        self.dialog.tableWidget.item(item.row(), 6).setCheckState(0)
                        
        return QtCore.QObject.event(obj, event)
    #Tip sonde
    def sensorData(self):
        sensorDataList = []
        val1 = 0
        val2 = 0
        val3 = 0
        for i in range(4):
            if(self.dialog.tableWidget.item(i,5).checkState() == 0):
                pass
            elif(self.dialog.tableWidget.item(i,5).checkState() == 2):
                try:
                    sensorDataList.append(self.dialog.tableWidget.item(i,5).text())
                    MIN = round(float(self.dialog.tableWidget.item(6,5).text()),2) 
                    MAX = round(float(self.dialog.tableWidget.item(8,5).text()),2)
                    if(MIN < MAX):
                        val1 = MIN+0.1*(MAX-MIN)
                        val2 = MIN+0.5*(MAX-MIN)
                        val3 = MIN+0.9*(MAX-MIN)
                        sensorDataList.append(val1)
                        sensorDataList.append(val2)
                        sensorDataList.append(val3)
                    else:
                        self.ressMinMax()
                except:
                    self.errorSondaUnos()
            else:
                pass
        self.podaci['TEMP'] = sensorDataList

    #DCV data
    def dataDCV(self):
        tmp1Data = []
        for i in range(9):
            dcvDataList = []
            tmpData = []
            if(self.dialog.tableWidget.item(i,0).checkState() == 0):
                pass
            elif(self.dialog.tableWidget.item(i,0).checkState() == 1): #Intermediate
                try:
                    val = float(self.dialog.tableWidget.item(i,0).text())
                    dcvDataList.append(round(val*0.1,6))
                    dcvDataList.append(round(val*(-0.1),6))
                    dcvDataList.append(round(val*0.5,6))
                    dcvDataList.append(round(val*0.9,6))
                    dcvDataList.append(round(val*(-0.9),6))
                    dcvDataList.sort()
                    tmpData.append(dcvDataList)
                    tmpData.append(val)
                except:
                    self.erorInputDataFormatDCV()
            elif(self.dialog.tableWidget.item(i,0).checkState() == 2): #All ranges
                try:
                    val = float(self.dialog.tableWidget.item(i,0).text())
                    dcvDataList.append(round(val*0,6))
                    dcvDataList.append(round(val*0.1,6))
                    dcvDataList.append(round(val*0.9,6))
                    dcvDataList.append(round(val*(-0.9),6))
                    dcvDataList.sort()
                    tmpData.append(dcvDataList)
                    tmpData.append(val)
                except:
                    self.erorInputDataFormatDCV()
            if(tmpData != []):
                tmp1Data.append(tmpData)

        self.podaci['DCV'] = tmp1Data

    #DCI data
    def dataDCI(self):
        tmp1Data = []
        for i in range(9):
            tmpData=[]
            dciDataList=[]

            if(self.dialog.tableWidget.item(i,2).checkState() == 0):
                pass
            elif(self.dialog.tableWidget.item(i,2).checkState() == 1): #Intermediate
                try:
                    val = round(float(self.dialog.tableWidget.item(i,2).text()),6)
                    dciDataList.append(round(val*0.1,6))
                    dciDataList.append(round(val*(0.5),6))
                    dciDataList.append(round(val*(0.9),6))
                    dciDataList.append(round(val*(-0.9),6))
                    dciDataList.sort()
                    tmpData.append(dciDataList)
                    tmpData.append(val)
                except:
                    self.erorInputDataFormatDCI()
                    self.dialog.pushButtonOKnext1.setEnabled(False)
            elif(self.dialog.tableWidget.item(i,2).checkState() == 2): #All ranges
                try:
                    val = round(float(self.dialog.tableWidget.item(i,2).text()),4)
                    dciDataList.append(round(val*0,6))
                    dciDataList.append(round(val*0.9,6))
                    #dciDataList.sort() 
                    tmpData.append(dciDataList)
                    tmpData.append(val)
                except:
                    self.erorInputDataFormatDCI()
                    self.dialog.pushButtonOKnext1.setEnabled(False)
            if(tmpData != []):
                tmp1Data.append(tmpData)
        self.podaci['DCI'] = tmp1Data
    #ACI data
    def dataACI(self):
        opseg = []
        for i in range(9):
            if(self.dialog.tableWidget.item(i,3).checkState() == 0):
                pass  
            elif(self.dialog.tableWidget.item(i,3).checkState() == 2): #All ranges
                try:
                    if(self.dialog.comboBoxSmallest10ACI.currentText() != 'None'):
                        value = round(float(self.dialog.tableWidget.item(i,3).text()),6)
                        opseg.append(value)
                    else:
                        self.smallestFreqACI()
                        self.dialog.pushButtonOKnext1.setEnabled(False)
                except:
                    self.erorInputDataFormatACI()
                    self.dialog.pushButtonOKnext1.setEnabled(False)               
        opseg.sort()

        freq = []
        
        if (self.dialog.comboBoxSmallest10ACI.currentText() != 'None'):
            freq.append(float(self.dialog.comboBoxSmallest10ACI.currentText()))
        else:
            freq.append(0)
            
        if(self.dialog.comboBoxAll1ACI.currentText() != 'None'):
            val = float(self.dialog.comboBoxAll1ACI.currentText())
            if(val not in freq):
                freq.append(val)
        
        if(self.dialog.comboBoxAll2ACI.currentText() != 'None'):
            val = float(self.dialog.comboBoxAll2ACI.currentText())
            if(val not in freq):
                freq.append(val)
                  
        if(self.dialog.comboBoxAll3ACI.currentText() != 'None'):
            val = float(self.dialog.comboBoxAll3ACI.currentText())
            if(val not in freq):
                freq.append(val)
        
        if(self.dialog.comboBoxALL4ACI.currentText() != 'None'):
            val = float(self.dialog.comboBoxALL4ACI.currentText())
            if(val not in freq):
                freq.append(val)
        
        flag = True
        tmpData1 = []
        for rng in opseg:
            for f in freq:
                tmpData = []
                if(flag):
                    flag = False
                    tmpData.append(round(rng*0.1,6))
                tmpData.append(round(rng*0.9,6))
                tmpData1.append([tmpData, f, rng])
        self.podaci['ACI'] = tmpData1

    def dataACV(self):
        opseg = []
        intermediate_rng = -1
        intermediate_freq = -1
        smallest_rng = -1
        freq = []
        for i in range(9):
            if(self.dialog.tableWidget.item(i,1).checkState() == 0):
                pass
            elif(self.dialog.tableWidget.item(i,1).checkState() == 2):
                try:
                    if(self.dialog.comboBoxSmallest10.currentText() != 'None'):
                        value = round(float(self.dialog.tableWidget.item(i,1).text()),6)
                        opseg.append(value)
                    else:
                        self.smallestFreqACV()
                        self.dialog.pushButtonOKnext1.setEnabled(False)
                except:
                    self.erorInputDataFormatACV()
                    self.dialog.pushButtonOKnext1.setEnabled(False)
            elif(self.dialog.tableWidget.item(i,1).checkState() == 1): #Intermediate range
                try:
                    float(self.dialog.tableWidget.item(i,1).text())
                    if(self.dialog.comboBoxInt10.currentText() != 'None'):
                        intermediate_rng = round(float(self.dialog.tableWidget.item(i,1).text()),6)
                        intermediate_freq = float(self.dialog.comboBoxInt10.currentText())+1
                        freq.append(intermediate_freq)
                        opseg.append(intermediate_rng)
                    else:
                        self.intermediateFreqACV()
                        self.dialog.pushButtonOKnext1.setEnabled(False)
                except:
                    self.erorInputDataFormatACV()
                    self.dialog.pushButtonOKnext1.setEnabled(False)
           
        opseg.sort()
        if(opseg == []):
            return

        if(self.dialog.comboBoxSmallest10.currentText() != 'None'):

            smallest_rng = float(self.dialog.comboBoxSmallest10.currentText())-1
            if(intermediate_freq != smallest_rng and intermediate_rng != opseg[0]):
                freq.append(smallest_rng)
        
        if(self.dialog.comboBoxAll1.currentText() != 'None'):
            val = float(self.dialog.comboBoxAll1.currentText())
            if(val not in freq):
                freq.append(val)
        
        if(self.dialog.comboBoxAll2.currentText() != 'None'):
            val = float(self.dialog.comboBoxAll2.currentText())
            if(val not in freq):
                freq.append(val)
        
        if(self.dialog.comboBoxAll3.currentText() != 'None'):
            val = float(self.dialog.comboBoxAll3.currentText())
            if(val not in freq):
                freq.append(val)
        
        if(self.dialog.comboBoxAll4.currentText() != 'None'):
            val = float(self.dialog.comboBoxAll4.currentText())
            if(val not in freq):
                freq.append(val)
  
        flag = True
        tmpData1 = []
        for rng in opseg:
            for f in freq:
                tmpData = []
                if(rng == intermediate_rng and f == intermediate_freq and f != smallest_rng):
                    tmpData.append(round(rng*0.1,6))
                    tmpData.append(round(rng*0.5,6))
                    tmpData.append(round(rng*0.9,6))
                    f -= 1
                elif(rng != intermediate_rng and f != intermediate_freq and f != smallest_rng):
                    tmpData.append(round(rng*0.9,6))
                elif(flag and f != intermediate_freq and f == smallest_rng):
                    flag = False
                    tmpData.append(round(rng*0.1,6))
                    tmpData.append(round(rng*0.9,6))
                    f +=1
                elif(rng == intermediate_rng and f != intermediate_freq and f != smallest_rng):
                    if(f == intermediate_freq-1):
                        continue
                    tmpData.append(round(rng*0.9,6))
                else:
                    continue
                tmpData1.append([tmpData, f, rng])

        self.podaci['ACV'] = tmpData1
    #Ress data
    def dataRess(self):
        tmp1Data = []
        opseg = []
        #tmp2Data = []
        for i in range(9):
            resistanceDataList = []
            tmpData = []
            if(self.dialog.tableWidget.item(i,4).checkState() == 0):
                pass
            elif(self.dialog.tableWidget.item(i,4).checkState() == 2):
                try:
                    val = round(float(self.dialog.tableWidget.item(i,4).text()),2)
                    opseg.append(val)
                except:
                    self.erorInputDataFormatRESS()
                    self.dialog.pushButtonOKnext1.setEnabled(False)
                    
 
        opseg.sort()
        for i in range(9):
            tmpData = []
            resistanceDataList = []
            if(self.dialog.tableWidget.item(i,4).checkState() == 0):
                pass
            elif(self.dialog.tableWidget.item(i,4).checkState() == 2):
                try:
                    val = round(float(self.dialog.tableWidget.item(i,4).text()),2)
                    if(opseg[0] == val):
                        resistanceDataList.append(val*0.0)
                        resistanceDataList.append(val*0.1)
                        resistanceDataList.append(val*0.9)
                        tmpData.append(resistanceDataList)
                        tmpData.append(val)
                    else:
                        resistanceDataList.append(val*0.1)
                        resistanceDataList.append(val*0.9)
                        tmpData.append(resistanceDataList)
                        tmpData.append(val)
                except:
                    self.erorInputDataFormatRESS()
                    self.dialog.pushButtonOKnext1.setEnabled(False)
                    

            if(tmpData != []):
                tmp1Data.append(tmpData)
            
        self.podaci['RES'] = tmp1Data

    #Freq data
    def dataFreq(self):
        tmpData = []
        val = 0
        val1 = 0
        for i in range(9):
            freqDataList = []
            if(self.dialog.tableWidget.item(i,6).checkState() == 0):
                pass
            elif(self.dialog.tableWidget.item(i,6).checkState() == 2):
                try:
                    val = round(float(self.dialog.tableWidget.item(i,6).text()),2)
                    val1 = round(float(self.dialog.tableWidget.item(i,7).text()),2)
                    if(val <= val1):
                        freqDataList.append(val)
                        freqDataList.append(val1)
                        tmpData.append(freqDataList) 
                    else:
                        self.freqRangeError()
                        self.dialog.pushButtonOKnext1.setEnabled(False)                  
                except:
                    self.erorInputDataFormatFREQ()
                    self.dialog.pushButtonOKnext1.setEnabled(False)
        if(len(tmpData) >= 1):
            self.podaci['FREQ'] = tmpData
    
    def nextTab1(self):
        self.dialog.tabWidget.setTabEnabled(0, False)
        self.dialog.tabWidget.setTabEnabled(1, True)
        self.dialog.tabWidget.setTabEnabled(2, False)
        self.dialog.tabWidget.setCurrentIndex(1)
        
    def nextTab2(self):
        self.dialog.tabWidget.setTabEnabled(0, False)
        self.dialog.tabWidget.setTabEnabled(1, False)
        self.dialog.tabWidget.setTabEnabled(2, True)
        self.dialog.tabWidget.setCurrentIndex(2)

    def testCom(self):
        self.returnValue = {None}
        self.returnValue = self.get_data()
        self.s = serial.Serial(port = self.returnValue['com_port'], baudrate = self.returnValue['baud_rate'], parity = self.returnValue['parity'], stopbits = self.returnValue['stop_bit'], bytesize = self.returnValue['data_bits'], timeout = 5)
        self.s.write(b'SYST:VERS?\r\n')
        returnMsg = ''
        returnMsg = self.s.readline()
        if returnMsg != b'': #1 je test iz konzole
            self.dialog.pushButtonTestComm.setStyleSheet("background-color: green")
            self.dialog.pushButtonOK_next.setEnabled(True)
            self.s.close()
        else:
            self.dialog.pushButtonTestComm.setStyleSheet("background-color: red")
            self.dialog.pushButtonOK_next.setEnabled(False)
            self.errorMsg()
            self.s.close()
            return

    def errorMsg(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText('Greška')
        msg.setInformativeText('Uređaj ne odgovara!\nZamenite port ili povežite uređaj ponovo.')
        msg.setWindowTitle('Greška')
        msg.exec_()

    def errorSonda(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText('[T]')
        msg.setInformativeText('Izaberite samo jednu sondu!')
        msg.setWindowTitle('Greška')
        msg.exec_()
        self.dialog.pushButtonOKnext1.setEnabled(False)

    def errorSondaUnos(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText('[T]')
        msg.setInformativeText('Unesi MIN i MAX range!')
        msg.setWindowTitle('Greška')
        msg.exec_()
        self.dialog.pushButtonOKnext1.setEnabled(False)

    def erorInputDataFormatDCI(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText('Nepravilan unos podataka!')
        msg.setInformativeText('Unesite podatak u odgovarajućem formatu!')
        msg.setWindowTitle('DCI')
        msg.exec_()
        #self.dialog.pushButtonOKnext1.setEnabled(False)

    def erorInputDataFormatACI(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText('Nepravilan unos podataka!')
        msg.setInformativeText('Unesite podatak u odgovarajućem formatu!')
        msg.setWindowTitle('ACI')
        msg.exec_()
        #self.dialog.pushButtonOKnext1.setEnabled(False)
    
    def erorInputDataFormatACV(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText('Nepravilan unos podataka!')
        msg.setInformativeText('Unesite podatak u odgovarajućem formatu!')
        msg.setWindowTitle('ACV')
        msg.exec_()
        #self.dialog.pushButtonOKnext1.setEnabled(False)

    def erorInputDataFormatDCV(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText('Nepravilan unos podataka!')
        msg.setInformativeText('Unesite podatak u odgovarajućem formatu!')
        msg.setWindowTitle('DCV')
        msg.exec_()
        #self.dialog.pushButtonOKnext1.setEnabled(False)

    def erorInputDataFormatFREQ(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText('Nepravilan unos podataka!')
        msg.setInformativeText('Unesite podatak u odgovarajućem formatu!')
        msg.setWindowTitle('FREQ')
        msg.exec_()
        #self.dialog.pushButtonOKnext1.setEnabled(False)

    def erorInputDataFormatRESS(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText('Nepravilan unos podataka!')
        msg.setInformativeText('Unesite podatak u odgovarajućem formatu!')
        msg.setWindowTitle('RESS')
        msg.exec_()
        #self.dialog.pushButtonOKnext1.setEnabled(False)

    def freqRangeError(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText('Greska u odabiru opsega [FREQ RANGE]!')
        msg.setInformativeText('Opseg frekvencije veci ili jednak vrednosti!')
        msg.setWindowTitle('Greška')
        msg.exec_()
        #self.dialog.pushButtonOKnext1.setEnabled(False)
        
    def selectFreqComboACI(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText('Niste izabrali frekvenciju!')
        msg.setInformativeText('Odaberite frekvenciju!')
        msg.setWindowTitle('ACI')
        msg.exec_()
        #self.dialog.pushButtonOKnext1.setEnabled(False)


    def selectFreqComboACV(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText('Niste izabrali frekvenciju!')
        msg.setInformativeText('Odaberite frekvenciju!')
        msg.setWindowTitle('ACV')
        msg.exec_()
        #


    #Funkcija koja provjerava chekirana polja u tableWidget-u
    #Ne mozemo preci na sledeci tab dok check ne bude True
    def checkData(self):
        
        self.dataDCV()
        self.dataACV()
        self.dataDCI()
        self.dataACI()
        self.dataRess()
        self.sensorData()
        self.dataFreq()
        self.dialog.pushButtonOKnext1.setEnabled(True)

        return self.podaci

    def selectOneIntermediateRangeDCV(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText('I!')
        msg.setInformativeText('Mozete izabrati samo 1 Intermediate range!')
        msg.setWindowTitle('Intermediate range DCV')
        msg.exec_()

    def selectOneIntermediateRangeACV(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText('I!')
        msg.setInformativeText('Mozete izabrati samo 1 Intermediate range!')
        msg.setWindowTitle('Intermediate range ACV')
        msg.exec_()

    def selectOneIntermediateRangeDCI(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText('I!')
        msg.setInformativeText('Mozete izabrati samo 1 Intermediate range!')
        msg.setWindowTitle('Intermediate range DCI')
        msg.exec_()

    def ressMinMax(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText('I!')
        msg.setInformativeText('Max mora biti veci od MIN!')
        msg.setWindowTitle('MIN max')
        msg.exec_()

    def smallestFreqACV(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText('Izaberite frekvenciju!')
        msg.setInformativeText('Niste izabrali frekvenciju za smallest range!')
        msg.setWindowTitle('ACV:Freq error')
        msg.exec_()

    def intermediateFreqACV(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText('Izaberite frekvenciju!')
        msg.setInformativeText('Niste izabrali frekvenciju za intermediate range!')
        msg.setWindowTitle('ACV:Freq error')
        msg.exec_()

    def smallestFreqACI(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText('Izaberite frekvenciju!')
        msg.setInformativeText('Niste izabrali frekvenciju za smallest range!')
        msg.setWindowTitle('ACI:Freq error')
        msg.exec_()





    
