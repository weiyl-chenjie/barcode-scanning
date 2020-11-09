# Author:尉玉林(Mr.Wei)

# Create Date:2019/10/14

# Edition:V1.0.0

# 自带库
import os
import sys
from time import sleep
import re
# 第三方库
from PySide2.QtWidgets import QMainWindow, QApplication, QMessageBox
from PySide2.QtCore import QThread, Signal
import win32api
import win32con
import win32gui
# 自己的包
from config import Config
from HslCommunication import SiemensS7Net
from HslCommunication import SiemensPLCS
from UI2PY.MainWindow import Ui_MainWindow
from access import ODBC_MS


class MyWindow(QMainWindow):
    def __init__(self):
        super(MyWindow, self).__init__()
        self.cur_path = os.getcwd()
        self.Ui_MainWindow = Ui_MainWindow()
        self.Ui_MainWindow.setupUi(self)
        self.Ui_MainWindow.lineEdit_scanning.clear()
        self.Ui_MainWindow.lineEdit_previous_barcode.clear()
        self.Ui_MainWindow.lineEdit_next_barcode.clear()

        self.conf = Config()
        self.IP = self.conf.read_config('PLC', 'IP')
        self.ezdName = self.conf.read_config('laser', 'ezdname')
        self.Ui_MainWindow.lineEdit_IP.setText(self.IP)

        self.cW = cWindow()

        self._thread = MyThread()
        
        # 数据库变量
        self.driver = '{Microsoft Access Driver (*.mdb)}'
        self.dbq = os.path.join(self.cur_path, "save.mdb")  # 数据库文件路径
        
        # 创建数据库实例
        self.db = ODBC_MS(self.driver, self.dbq)  # 创建数据库连接实例
        
        if self._thread.connect_to_plc:
            pass
        else:
            self.Ui_MainWindow.label_status.setStyleSheet("background-color: rgb(255, 0, 0);")
            self.Ui_MainWindow.label_status.setText('PLC连接失败!')
        self._thread.start()

        self.working_barcode = ''  # 工作中的钥匙条码
        self.waiting_barcode = ''  # 在挡停位停留的钥匙条码
        self.temp_barcode = ''  # 暂存的barcode

        # 绑定刻字函数
        self._thread.signal.connect(self.laser_marking)

    # 槽函数
    def barcode_scanning(self):
        if self._thread.connect_to_plc:  # 如果PLC连接成功
            self.Ui_MainWindow.label_status.setStyleSheet("background-color: rgb(255, 255, 127);")
            self.Ui_MainWindow.label_status.setText('已扫码，等待刻字信号...')
            QApplication.processEvents()

            # 获取扫码信息
            barcode = self.Ui_MainWindow.lineEdit_scanning.text()[-5:]
            # 清空扫描区
            self.Ui_MainWindow.lineEdit_scanning.clear()
            # # 显示当前已扫描的条码
            # self.Ui_MainWindow.lineEdit_previous_barcode.setText(barcode)
            # print(barcode)
            if self.working_barcode == '':  # 如果working_barcode为空，则将条码赋值给working_barcode
                self.working_barcode = barcode
                self.Ui_MainWindow.lineEdit_next_barcode.setText(self.working_barcode)
            elif self.waiting_barcode == '':  # 如果working_barcode不为空，且waiting_barcode为空，则将条码赋值给waiting_barcode
                if barcode == self.working_barcode:  # 如果重复扫码
                    pass
                else:
                    self.waiting_barcode = barcode
                    self.Ui_MainWindow.lineEdit_next_barcode.setText(self.working_barcode + '   ' + self.waiting_barcode)
            else:
                self.temp_barcode = barcode  # 否则将条码赋值给temp_barcode

    def change_ip(self):
        self.IP = self.Ui_MainWindow.lineEdit_IP.text()
        self.conf.update_config(section='PLC', name='IP', value=self.IP)

    def connect_test(self):
        self.Ui_MainWindow.label_status.setStyleSheet("background-color: rgb(255, 255, 127);")
        self.Ui_MainWindow.label_status.setText('正在连接PLC...')
        QApplication.processEvents()
        # 创建PLC实例
        self._thread.siemens = SiemensS7Net(SiemensPLCS.S200Smart, self.IP)
        if self._thread.siemens.ConnectServer().IsSuccess:  # 连接成功
            self.Ui_MainWindow.label_status.setText('PLC连接成功！')
            self._thread.connect_to_plc = True
        else:  # 若连接失败
            self._thread.connect_to_plc = False
            self.Ui_MainWindow.label_status.setStyleSheet("background-color: rgb(255, 0, 0);")
            self.Ui_MainWindow.label_status.setText('PLC连接失败!')
        QApplication.processEvents()

    # 功能函数
    # 将刻字内容发送到.txt文件
    def send_barcode(self):
        txtPath = r"D:\刻字.txt"
        barcode_to_laser = self.working_barcode
        self.working_barcode = self.waiting_barcode
        self.waiting_barcode = ''

        if not os.path.exists(txtPath):  # 如果文件不存在，则创建相应文件并输入刻字内容
            with open(txtPath, "w", encoding="utf-8") as f:
                f.write(barcode_to_laser)
        else:  # 如果文件已存在，则清空文件内容，并输入新的刻字内容
            with open(txtPath, "w+", encoding="utf-8") as f:
                f.write(barcode_to_laser)

        # barcode存放到数据库中
        save_barcode_sql = "INSERT INTO barcode(barcode) VALUES('" + barcode_to_laser + "')"
        self.db.insert_query(save_barcode_sql)
        
        # 显示当前刻印的条码
        self.Ui_MainWindow.lineEdit_previous_barcode.setText(barcode_to_laser)
        # 显示待刻印的条码
        self.Ui_MainWindow.lineEdit_next_barcode.setText(self.working_barcode)

    # 激光刻字流程函数--通过调用刻字程序控制刻字机刻字
    def laser_marking(self):
        self.send_barcode()
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
        print("刻字完成，将扫码程序前置")
        self.cW.find_window_regex("条码扫描")  # 找到指定窗口
        # self.cW.Maximize()  # 窗口最大化
        self.cW.SetAsForegroundWindow()  # 窗口前置

        # 刻字完成，将标志位置为False，等待下一次触发
        self._thread.working = False
        # 刻字完成，将暂停标志置为True，等待下一次触发
        self._thread.pause = True

        self.Ui_MainWindow.label_status.setStyleSheet("background-color: rgb(255, 255, 127);")
        self.Ui_MainWindow.label_status.setText('刻字完成,请取走零件...')
        QApplication.processEvents()


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
        if hwnd != self._hwnd:  # ignore self
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
    signal = Signal(bool)

    def __init__(self):
        super(MyThread, self).__init__()
        self.working = True  # 工作状态标志量
        self.pause = False  # 刻字之后暂停

        self.conf = Config()
        self.IP = self.conf.read_config('PLC', 'IP')
        # 零件放到位，刻字延迟时间
        self.delay = int(self.conf.read_config('laser', 'delay'))

        # 创建PLC实例
        self.siemens = SiemensS7Net(SiemensPLCS.S200Smart, self.IP)
        if self.siemens.ConnectServer().IsSuccess:  # 连接成功
            self.connect_to_plc = True
        else:
            self.connect_to_plc = False

    # def __del__(self):
    #     self.working = False

    def run(self):
        # 进行线程任务
        while True:
            if self.working:  # 如果处于刻字工作状态
                # 判断是哪个项目,I0.3是转换开关（True代表280B，False代表480B）,待刻印件从右往左依次为I0.0、I0.1、I0.2
                if self.siemens.ReadBool("I0.3").Content:  # 如果I0.3为True(即，280B项目--该项目刻字工位需放2个件,I0.0、I0.1)
                    # 判断零件是否到位
                    if self.siemens.ReadBool('I0.0').Content and self.siemens.ReadBool('I0.1').Content:  # 如果零件放到位
                        sleep(self.delay)
                        self.signal.emit(True)
                        sleep(5)
                    else:
                        pass
                        # print('有零件未放到位')
                else:  # 如果时是480B项目--该项目刻字工位需放3个件,I0.0、I0.1、I0.2
                    if self.siemens.ReadBool("I0.0").Content and self.siemens.ReadBool("I0.1").Content and self.siemens.ReadBool("I0.2").Content:  # 零件放到位
                        self.signal.emit(True)
                        sleep(5)
                    else:
                        pass
                        # print('有零件未放到位')
            elif self.pause:  # 如果处于刻字之后的暂停状态
                # 判断是哪个项目,I0.3是转换开关（True代表280B，False代表480B）,待刻印件从右往左依次为I0.0、I0.1、I0.2
                if self.siemens.ReadBool("I0.3").Content:  # 如果I0.3为True(即，280B项目--该项目刻字工位需放2个件,I0.0、I0.1)
                    # 判断零件是否被拿走
                    if not self.siemens.ReadBool('I0.0').Content and not self.siemens.ReadBool('I0.1').Content:  # 如果零件都被取走
                        self.pause = False
                        self.working = True
                    else:
                        pass
                        # print('有零件未取走')
                else:  # 如果时是480B项目--该项目刻字工位需放3个件,I0.0、I0.1、I0.2
                    # 判断零件是否被拿走
                    if not self.siemens.ReadBool("I0.0").Content and not self.siemens.ReadBool(
                            "I0.1").Content and not self.siemens.ReadBool("I0.2").Content:  # 如果零件都被取走
                        self.pause = False
                        self.working = True
                    else:
                        pass
                        # print('有零件未取走')


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
