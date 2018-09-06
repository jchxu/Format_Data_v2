# coding=utf-8
from ReadFunc import *
from WriteFunc import *
import sys
from PyQt5 import QtGui,QtCore,QtWidgets

from openfile import Ui_MainWindow

#################### Templete ################################
class MyWindows(QtWidgets.QMainWindow,Ui_MainWindow):
    def __init__(self):
        super(MyWindows, self).__init__()
        self.setupUi(self)

    def button_click(self):
        self.textEdit.setText("fjsahfksdf")


######################################

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = MyWindows()
    window.show()

    #PortFiles =
    #StdPort = ['曹妃甸实业', '曹妃甸弘毅', '曹妃甸矿三', '京唐', '岚桥', '岚山', '连云港', '青岛', '日照']
    #### 读取各港口信息 ###
    #PortsData = CollectPortData(PortFiles, StdPort)
    #### 汇总各港口信息，统一格式输出 ###
    #PortList = PortOrder(PortsData)
    #WritePortsData(PortList, PortsData)

    sys.exit(app.exec_())