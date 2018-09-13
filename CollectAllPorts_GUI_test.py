# coding=utf-8
from ReadFunc import *
from WriteFunc import *
import sys
from PyQt5 import QtGui,QtCore,QtWidgets
from openfile import Ui_MainWindow

StdPort = ['曹妃甸实业', '曹妃甸弘毅', '曹妃甸矿三', '京唐', '岚桥', '岚山', '连云港', '青岛', '日照']

#################### Templete ################################
PortFiles = []
class MyWindows(QtWidgets.QMainWindow,Ui_MainWindow):
    def __init__(self):
        super(MyWindows, self).__init__()
        self.setupUi(self)
    def button_click(self):
        fileName, filetype = QtWidgets.QFileDialog.getOpenFileNames(self, "选取港口库存数据文件", "./")
        for item in fileName:
            PortFiles.append(item.split('/')[-1])
        test = '计划读取以下港口数据文件：'
        self.textEdit.setText('计划读取以下港口数据文件：')
        for item in PortFiles:
            self.textEdit.append(item)
        #光标移至最后一行
        cursor = self.textEdit.textCursor()
        cursor.setPosition(len(self.textEdit.toPlainText())-1)
        #self.textEdit.ensureCursorVisible()
        self.textEdit.setTextCursor(cursor)
    def button2_click(self):
        #print(PortFiles)
        ##### 读取各港口信息 ###
        PortsData = CollectPortData(PortFiles, StdPort)
        self.textEdit.append('已读取各港口数据')
        ##### 汇总各港口信息，统一格式输出 ###
        PortList = PortOrder(PortsData)
        WritePortsData(PortList, PortsData)
        self.textEdit.append('已汇总输出各港口数据，请关闭窗口')
        cursor = self.textEdit.textCursor()
        cursor.setPosition(len(self.textEdit.toPlainText()) - 1)
        self.textEdit.setTextCursor(cursor)
######################################
if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = MyWindows()
    window.show()
    sys.exit(app.exec_())