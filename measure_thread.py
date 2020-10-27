from PyQt5.QtCore import QThread, pyqtSignal, QWaitCondition, QWaitCondition, QMutex
from PyQt5.QtWidgets import QInputDialog
import time
import copy

class MeasureThread(QThread):
    set_data = pyqtSignal(str, str, str, str)
    send_data = pyqtSignal(str)
    send_measure = pyqtSignal(dict)
    def __init__(self):
        super().__init__()
        self.waitCondition = QWaitCondition()
        self.mutex = QMutex()
        self.podaciMain = {}
        self.flag = False
        self.dataType = ''
        self.flag_end = False
        self.izmereno = 0
        
        
    def collect_data(self, data):
        self.podaciMain = data
        self.podaciMereno = copy.deepcopy(self.podaciMain)

    def collect_measure(self, dataM):
        self.izmereno = dataM

    def update_flag(self):
        self.waitCondition.wakeAll()
        

    #Funkcije za setovanje izlaza
    def resData(self):
        
        self.dataRES = self.podaciMain.get("RES")
        self.dataRES_M = self.podaciMereno.get("RES")
        r =len(self.dataRES)
        if(r !=0):
            str1 = 'RES'
            self.dataType = str1
            for i in range(r):
                tmpData = self.dataRES[i][0] #vrednosti 10% 90% 50% ...
                tmpData1 = self.dataRES[i][1] #Opseg
                tmpData_M = self.dataRES_M[i][0]
                
               
                a = len(tmpData)
                for j in range(a):
                    str2 = str(tmpData1)   
                    str3 = str(tmpData[j])
                    self.set_data.emit(str1,str2,str3,'')
                    self.send_data.emit('sres' +' '+ str(tmpData[j])+ '\r\n')
                    
                    self.mutex.lock()
                    self.waitCondition.wait(self.mutex)
                    self.mutex.unlock()
                   
                    self.getMessure()
                    
                    self.mutex.lock()
                    self.waitCondition.wait(self.mutex)
                    self.mutex.unlock()
                    tmpData_M[j] = self.izmereno
        else:
            pass

    
 

    def acvData(self):
        self.dataACV = self.podaciMain.get("ACV")
        self.dataACV_M = self.podaciMereno.get("ACV")
        r = len(self.dataACV)
        if(r != 0):
            self.send_data.emit('func sin \r\n')
            time.sleep(0.01)
            str1 = 'ACV'
            self.dataType = str1
            for i in range(r):
                tmpData = self.dataACV[i][0]
                tmpData1 = self.dataACV[i][1] #freq
                tmpData2 = self.dataACV[i][2] #opseg
                tmpData_M = self.dataACV_M[i][0]
                self.send_data.emit('freq' +' '+ str(tmpData1)+ '\r\n')
                time.sleep(0.01)
                self.send_data.emit('volt:rang' +' '+ str(tmpData2)+ '\r\n')
                
                a = len(tmpData)
                for j in range(a):
                    str2 = str(tmpData2)
                    str3 = str(tmpData[j])
                    str4 = str(tmpData1)
                    self.set_data.emit(str1,str2,str3,str4)
                    self.send_data.emit('VOLT' +' '+ str(tmpData[j])+ '\r\n')
    
                    self.mutex.lock()
                    self.waitCondition.wait(self.mutex)
                    self.mutex.unlock()
                   
                    self.getMessure()

                    self.mutex.lock()
                    self.waitCondition.wait(self.mutex)
                    self.mutex.unlock()
                    tmpData_M[j] = self.izmereno
        else:
            pass

    def aciData(self):
        self.dataACI = self.podaciMain.get("ACI")
        self.dataACI_M = self.podaciMereno.get("ACI")
        r = len(self.dataACI)
        if(r != 0):
            self.send_data.emit('func sin \r\n')
            time.sleep(0.01)
            str1 = 'ACI'
            self.dataType = str1
            for i in range(r):
                tmpData = self.dataACI[i][0]
                tmpData1 = self.dataACI[i][1] #freq
                tmpData2 = self.dataACI[i][2] #opseg
                tmpData_M = self.dataACI_M[i][0]
                self.send_data.emit('freq ' +' '+ str(tmpData1)+ '\r\n')
                time.sleep(0.01)
                self.send_data.emit('curr:rang' +' '+ str(tmpData2)+ '\r\n')
                time.sleep(0.01)
               


                a = len(tmpData)
                for j in range(a):
                    str2 = str(tmpData2)
                    str3 = str(str(tmpData[j]))
                    str4 = str(tmpData1)
                    self.set_data.emit(str1,str2,str3,str4)
                    self.send_data.emit('curr' +' '+ str(tmpData[j])+ '\r\n')
                    time.sleep(0.01)

                    self.mutex.lock()
                    self.waitCondition.wait(self.mutex)
                    self.mutex.unlock()
                   
                    self.getMessure()

                    self.mutex.lock()
                    self.waitCondition.wait(self.mutex)
                    self.mutex.unlock()
                    tmpData_M[j] = self.izmereno      
        else:
            pass

    def dcvData(self):
        self.dataDCV = self.podaciMain.get("DCV")
        self.dataDCV_M = self.podaciMereno.get("DCV")
        r = len(self.dataDCV)
        if(r != 0):
            self.send_data.emit('func dc \r\n')
            time.sleep(0.01)

            str1 = 'DCV'
            self.dataType = str1
            for i in range(r):
                tmpData = self.dataDCV[i][0]
                tmpData1 = self.dataDCV[i][1] #opseg
                tmpData_M = self.dataDCV_M[i][0]
                self.send_data.emit('volt:rang' +' '+ str(tmpData1)+ '\r\n')
                time.sleep(0.01)
                
                a = len(tmpData)
                for j in range(a):
                    str2 = str(tmpData1)
                    str3 = str(tmpData[j])
                    self.set_data.emit(str1,str2,str3,'')
                    self.send_data.emit('volt' +' '+ str(tmpData[j])+ '\r\n')
                    time.sleep(0.01)

                    self.mutex.lock()
                    self.waitCondition.wait(self.mutex)
                    self.mutex.unlock()
                   
                    self.getMessure()

                    self.mutex.lock()
                    self.waitCondition.wait(self.mutex)
                    self.mutex.unlock()
                    tmpData_M[j] = self.izmereno 
        else:
            pass
    
    def dciData(self):
        self.dataDCI = self.podaciMain.get("DCI")
        self.dataDCI_M = self.podaciMereno.get("DCI")
        r = len(self.dataDCI)
        if(r != 0):
            self.send_data.emit('func dc \r\n')
            time.sleep(0.01)
            str1 = 'DCI'
            self.dataType = str1
            
            for i in range(r):
                tmpData = self.dataDCI[i][0]
                tmpData1 = self.dataDCI[i][1] #opseg
                tmpData_M = self.dataDCI_M[i][0]
                self.send_data.emit('curr:rang' +' '+ str(tmpData1)+ '\r\n')
                time.sleep(0.01)
                time.sleep(0.01)
                
                a = len(tmpData)
                for j in range(a):
                    str2 = str(tmpData1)
                    str3 = str(tmpData[j])
                    self.set_data.emit(str1,str2,str3,'')
                    self.send_data.emit('curr' +' '+ str(tmpData[j])+ '\r\n')
                    time.sleep(0.01) 
                
                    self.mutex.lock()
                    self.waitCondition.wait(self.mutex)
                    self.mutex.unlock()
                   
                    self.getMessure()

                    self.mutex.lock()
                    self.waitCondition.wait(self.mutex)
                    self.mutex.unlock()
                    tmpData_M[j] = self.izmereno                 
        else:
            pass
    
    def freqData(self):
        self.dataFREQ = self.podaciMain.get("FREQ")
        self.dataFREQ_M = self.podaciMereno.get("FREQ")
        r = len(self.dataFREQ)
        if(r != 0):
            str1 = 'FREQ'
            self.dataType = str1
            for i in range(r):
                tmpData = self.dataFREQ[i][0] #vrednost
                tmpData1 = self.dataFREQ[i][1] #opseg
                 
               
            
                str2 = str(tmpData1)
                str3 = str(tmpData)
                self.set_data.emit(str1,str2,str3,'')
                
                self.send_data.emit('puls:sfr' +' '+ str(tmpData)+ '\r\n')
                time.sleep(0.01)

                self.mutex.lock()
                self.waitCondition.wait(self.mutex)
                self.mutex.unlock()
                
                self.getMessure()
                
                self.mutex.lock()
                self.waitCondition.wait(self.mutex)
                self.mutex.unlock()
                self.dataFREQ_M[i][0] = self.izmereno
                
        else:
            pass
   
    def tempData(self):
        self.dataTEMP = self.podaciMain.get("TEMP")
        self.dataTEMP_M = self.podaciMereno.get("TEMP")

        r = len(self.dataTEMP)
        if(r != 0):
            str1 = 'TEMP'
            self.dataType = str1
            tmpData = self.dataTEMP[0] #tip sonde
            
            if(tmpData == 'PT100'):
                self.send_data.emit('rtd:type pt100 \r\n')
                
                time.sleep(0.01)
                for i in range(3):
                    tmpData = self.dataTEMP[i+1]
                    str3 = str(tmpData)
                    self.set_data.emit(str1,'',str3,'')
                    self.send_data.emit('rtd ' + str(tmpData)+ '\r\n')
                    time.sleep(0.01)

                    self.mutex.lock()
                    self.waitCondition.wait(self.mutex)
                    self.mutex.unlock()
                    
                    self.getMessure()

                    self.mutex.lock()
                    self.waitCondition.wait(self.mutex)
                    self.mutex.unlock()

                    self.dataTEMP_M[i+1] = self.izmereno

            elif(tmpData == 'TC K type'):
                self.send_data.emit('ther:type K \r\n')
                
                time.sleep(0.01)
                for i in range(3):
                    tmpData = self.dataTEMP[i+1]
                    str3 = str(tmpData)
                    self.set_data.emit(str1,'',str3,'')
                    self.send_data.emit('ther ' + str(tmpData)+ '\r\n')
                    time.sleep(0.01)

                    self.mutex.lock()
                    self.waitCondition.wait(self.mutex)
                    self.mutex.unlock()
                    
                    self.getMessure()

                    self.mutex.lock()
                    self.waitCondition.wait(self.mutex)
                    self.mutex.unlock()
                    self.dataTEMP_M[i+1] = self.izmereno

            elif(tmpData == 'TC J type'):
                self.send_data.emit('ther:type J \r\n')
                time.sleep(0.01)
                for i in range(3):
                    tmpData = self.dataTEMP[i+1]
                    str3 = str(tmpData)
                    self.set_data.emit(str1,'',str3,'')
                    self.send_data.emit('ther ' + str(tmpData)+ '\r\n')
                    time.sleep(0.01)

                    self.mutex.lock()
                    self.waitCondition.wait(self.mutex)
                    self.mutex.unlock()
                    
                    self.getMessure()

                    self.mutex.lock()
                    self.waitCondition.wait(self.mutex)
                    self.mutex.unlock()
                    self.dataTEMP_M[i+1] = self.izmereno
        else:
            pass


    def printDataMer(self):
        self.send_measure.emit(self.podaciMereno)

    def run(self):
        self.dcvData()
        self.acvData()
        self.dciData()
        self.aciData()
        self.resData()
        self.tempData()
        self.freqData()
        self.printDataMer()
        
        self.flag_end = True

    
    def getMessure(self):
        pass

    def clearTextLine(self):
        self.flag_end = True

    def exit(self):
        self.terminate()
