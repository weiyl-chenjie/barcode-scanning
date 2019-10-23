# Author:尉玉林(Mr.Wei)

# Create Date:2019/10/14

# Edition:V1.0.0

# Python自带库
import os
import sys
from time import sleep
import re
# 第三方库
from PySide2.QtWidgets import QMainWindow, QApplication, QMessageBox
from PySide2.QtCore import QThread, Signal
import win32api, win32con, win32gui
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

        self._thread = MyThread()
        self._thread.start()

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

        # 将刻字内容发送到.txt文件
        txtPath = r"D:\刻字.txt"
        if not os.path.exists(txtPath):  # 如果文件不存在，则创建相应文件并输入刻字内容
            with open(txtPath, "w", encoding="utf-8") as f:
                f.write(barcode)
        else:  # 如果文件已存在，则清空文件内容，并输入新的刻字内容
            with open(txtPath, "w+", encoding="utf-8") as f:
                f.write(barcode)

        self._thread.working = True

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


#  定义窗口类，用于窗口前置等操作
class cWindow:
    def __init__(self):
        self._hwnd = None

    def SetAsForegroundWindow(self):
        # First, make sure all (other) always-on-top windows are hidden.
        self.hide_always_on_top_windows()
        win32gui.SetForegroundWindow(self._hwnd)

    def Maximize(self):
        win32gui.ShowWindow(self._hwnd, win32con.SW_MAXIMIZE)

    def _window_enum_callback(self, hwnd, regex):
        '''Pass to win32gui.EnumWindows() to check all open windows'''
        if self._hwnd is None and re.match(regex, str(win32gui.GetWindowText(hwnd))) is not None:
            self._hwnd = hwnd

    def find_window_regex(self, regex):
        self._hwnd = None
        win32gui.EnumWindows(self._window_enum_callback, regex)

    def hide_always_on_top_windows(self):
        win32gui.EnumWindows(self._window_enum_callback_hide, None)

    def _window_enum_callback_hide(self, hwnd, unused):
        if hwnd != self._hwnd: # ignore self
            # Is the window visible and marked as an always-on-top (topmost) window?
            if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE) & win32con.WS_EX_TOPMOST:
                # Ignore windows of class 'Button' (the Start button overlay) and
                # 'Shell_TrayWnd' (the Task Bar).
                className = win32gui.GetClassName(hwnd)
                if not (className == 'Button' or className == 'Shell_TrayWnd'):
                    # Force-minimize the window.
                    # Fortunately, this seems to work even with windows that
                    # have no Minimize button.
                    # Note that if we tried to hide the window with SW_HIDE,
                    # it would disappear from the Task Bar as well.
                    win32gui.ShowWindow(hwnd, win32con.SW_FORCEMINIMIZE)


class MyThread(QThread):
    signal = Signal()

    def __init__(self):
        super(MyThread, self).__init__()
        self.working = False  # 工作状态标志量

        self.conf = Config()
        self.IP = self.conf.read_config('PLC', 'IP')
        self.delay = int(self.conf.read_config('laser', 'delay'))
        # 创建PLC实例
        self.siemens = SiemensS7Net(SiemensPLCS.S200Smart, self.IP)
        if self.siemens.ConnectServer().IsSuccess:  # 连接成功
            pass
        else:
            QMessageBox.critical(self, '错误！', 'PLC连接失败')

    def __del__(self):
        self.working = False

    # 激光刻字流程函数--通过调用刻字程序控制刻字机刻字
    def laser_marking(self):
        # 将刻字程序前置，并按下F2进行刻字
        print("前置激光刻字机程序:" + self.ezdName)
        self.cW.find_window_regex(self.ezdName)  # 找到指定窗口
        self.cW.Maximize()  # 窗口最大化
        self.cW.SetAsForegroundWindow()  # 窗口前置
        print("激光机开始刻字")
        win32api.keybd_event(113, 0, 0, 0)  # 按下F2 刻字机刻字
        win32api.keybd_event(113, 0, win32con.KEYEVENTF_KEYUP, 0)  # 松开F2
        sleep(1)
        # 刻字完成后，将本程序前置
        print("刻字完成，将读钥匙程序前置")
        self.cW.find_window_regex("扫码")  # 找到指定窗口
        self.cW.Maximize()  # 窗口最大化
        self.cW.SetAsForegroundWindow()  # 窗口前置

        # 刻字完成，将标志位置为False，等待下一次触发
        self.working = False

    def run(self):
        # 进行线程任务
        while self.working:
            # 判断是哪个项目,I0.3是转换开关（True代表280B，False代表480B）,待刻印件从右往左依次为I0.0、I0.1、I0.2
            if self.siemens.ReadBool("I0.3").Content:  # 如果I0.3为True(即，280B项目--该项目刻字工位需放2个件,I0.0、I0.1)
                # 判断零件是否到位
                if self.siemens.ReadBool('I0.0').Content and self.siemens.ReadBool('I0.1').Content:  # 如果零件放到位
                    sleep(self.delay)
                    self.laser_marking()
                else:
                    print('有零件未放到位')
            else:  # 如果时是480B项目--该项目刻字工位需放3个件,I0.0、I0.1、I0.2
                if self.siemens.ReadBool("I0.0").Content and self.siemens.ReadBool("I0.1").Content and self.siemens.ReadBool("I0.2").Content:  # 零件放到位
                    self.laser_marking()  # 通知刻字机刻字
                else:
                    print('有零件未放到位')


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
