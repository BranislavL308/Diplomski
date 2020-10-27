from PyQt5.QtCore import QThread, pyqtSignal  
from serial import Serial


class MyThread(QThread):
    change_value = pyqtSignal(list)
    error = pyqtSignal()

    def __init__(self, com_data):
        super().__init__()


    def init_com(self, com_data):
        
        self.data = com_data
        self.com_port = com_data['com_port']
        self.baud = com_data['baud_rate']
        self.data_bit = com_data['data_bits']
        self.parity = com_data['parity']
        self.stop_bit = com_data['stop_bit']
        try:
            self.s = Serial(self.com_port, self.baud, self.data_bit, self.parity, self.stop_bit)
        except:
            self.error.emit()
            return

    def run(self): 
        while(1):
            d = self.s.readline().decode()
            data = list(d.split())
            self.change_value.emit(data)


    def sendData(self,tmp):
        self.s.write((tmp).encode())

    def exit(self):
        self.terminate()
        if self.s.is_open:
            self.s.close()