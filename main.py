# Author:尉玉林(Mr.Wei)

# Create Date:2019/10/14

# Edition:V1.0.0

# Python自带库
import sys
# 第三方库
from PySide2.QtWidgets import QMainWindow, QApplication, QMessageBox
# 自己的包
from config import Config
from HslCommunication import SiemensS7Net
from HslCommunication import SiemensPLCS
from UI2PY.MainWindow import Ui_MainWindow


class MyWindow(QMainWindow):
    def __init__(self):
        super(MyWindow, self).__init__()
        self.Ui_MainWindow = Ui_MainWindow()
        self.Ui_MainWindow.setupUi(self)
        self.Ui_MainWindow.lineEdit_scanning.clear()
        self.Ui_MainWindow.lineEdit_previous_barcode.clear()

        self.conf = Config()

        self.IP = self.conf.read_config('PLC', 'IP')
        self.Ui_MainWindow.lineEdit_IP.setText(self.IP)

        # 创建PLC实例
        self.siemens = SiemensS7Net(SiemensPLCS.S200Smart, self.IP)
        if self.siemens.ConnectServer().IsSuccess:  # 连接成功
            self.Ui_MainWindow.label_status.setText('等待扫码...')
        else:
            QMessageBox.critical(self, '错误！', 'PLC连接失败')
            self.Ui_MainWindow.label_status.setText('PLC连接失败！')
            self.Ui_MainWindow.label_status.setStyleSheet("background-color: rgb(255, 0, 0);")

    # 槽函数
    def barcode_scanning(self):
        self.Ui_MainWindow.label_status.setStyleSheet("background-color: rgb(255, 255, 127);")
        self.Ui_MainWindow.label_status.setText('已扫码，正在处理...')
        QApplication.processEvents()

        # 获取扫码信息
        barcode = self.Ui_MainWindow.lineEdit_scanning.text()[-5:]
        # 清空扫描区
        self.Ui_MainWindow.lineEdit_scanning.clear()
        # 显示当前已扫描的条码
        self.Ui_MainWindow.lineEdit_previous_barcode.setText(barcode)
        # 检测
        # 发送信息给激光刻字机

        # 通知PLC可以刻字了

    def change_ip(self):
        self.IP = self.Ui_MainWindow.lineEdit_IP.text()
        self.conf.update_config(section='PLC', name='IP', value=self.IP)

    def connect_test(self):
        self.Ui_MainWindow.label_status.setStyleSheet("background-color: rgb(255, 255, 127);")
        self.Ui_MainWindow.label_status.setText('正在连接PLC...')
        QApplication.processEvents()
        # 创建PLC实例
        siemens = SiemensS7Net(SiemensPLCS.S200Smart, self.IP)
        if siemens.ConnectServer().IsSuccess:  # 连接成功
            self.Ui_MainWindow.label_status.setText('PLC连接成功！')
        else:  # 若连接失败
            self.Ui_MainWindow.label_status.setStyleSheet("background-color: rgb(255, 0, 0);")
            self.Ui_MainWindow.label_status.setText('PLC连接失败!')
        QApplication.processEvents()

    # 功能函数


if __name__ == '__main__':
    # 创建一个应用程序对象
    app = QApplication(sys.argv)

    # 创建控件(容器)
    window = MyWindow()

    # 设置标题
    # window.setWindowTitle('title')

    # 显示窗口
    window.show()

    # 进入消息循环
    sys.exit(app.exec_())
