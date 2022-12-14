# -*- coding: utf-8 -*-
"""
Created on Thu Aug  9 21:05:37 2022

@author: REC3WX

Tool to mass input the VAT on GTE Tax Recon website
CMD:
"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --remote-debugging-port=9222 --user-data-dir="C:\selenium\EdgeProfile"
"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\selenium\ChromeProfile"

ChromeDriverManager-chrome.py
            url: str = "http://npm.taobao.org/mirrors/chromedriver",
            latest_release_url: str = "http://npm.taobao.org/mirrors/chromedriver/LATEST_RELEASE",

"""


from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome  import ChromeDriverManager
import pandas as pd
import sys
import time
from datetime import datetime
from PySide6.QtWidgets import (QWidget, QPushButton, QFileDialog, QApplication, QLineEdit, QGridLayout, QLabel, QMessageBox, QPlainTextEdit, QFrame, QStyle, QComboBox)
from PySide6.QtGui import QFont
from PySide6.QtCore import Slot, Qt, QThread, Signal
import socket
import tempfile

opt = Options()
sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
port = ("localhost", 9222)
open = sock.connect_ex(port)
temp = tempfile.gettempdir()
if open == 0:
    opt.add_experimental_option("debuggerAddress", "localhost:9222")
else:
    opt.add_argument("--remote-debugging-port=9222")
    opt.add_argument("--start-maximized")
    opt.add_argument(f'user-data-dir={temp}//ChromeProfile')  

##driver_path = ChromeService(r'./chromedriver.exe')
##driver = webdriver.Chrome(service=driver_path, options=opt)
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=opt)
driver.set_page_load_timeout(5)

class Worker(QThread):
  sinOut = Signal(str)

  def __init__(self, parent=None):
    super(Worker, self).__init__(parent)

  def getdata(self, path, year, month):
    self.filepath = path
    self.year = year
    self.month = month
  
  def run(self):
    #?????????
    year = self.year
    month = self.month
    df = pd.read_csv(self.filepath, dtype=str, header=0)
    col_list = df.values.tolist()
    row = 0
    num1 = 1
    def autofill(num1, fphm, net, vat):
        num2 = num1 + 6
        num3 = num1 + 7
        path1 = f'//input[@id="_easyui_textbox_input{num1}"]'
        path2 = f'//input[@id="_easyui_textbox_input{num2}"]'
        path3 = f'//input[@id="_easyui_textbox_input{num3}"]'
        driver.find_element(By.XPATH, path1).click()
        driver.find_element(By.XPATH, path1).send_keys(fphm)
        driver.find_element(By.XPATH, path2).click()
        driver.find_element(By.XPATH, path2).send_keys(net)
        driver.find_element(By.XPATH, path3).click()
        driver.find_element(By.XPATH, path3).send_keys(vat)
    
    try:
        message = f'{datetime.now()} - ???????????????????????????...'
        self.sinOut.emit(message)
        driver.get(url='http://192.168.10.47:8080/glaf/loginApp.do')
        time.sleep(2)
        driver.find_element(By.NAME, "x").click()
        driver.find_element(By.NAME, "x").clear()
        driver.find_element(By.NAME, "x").send_keys('3334')
        driver.find_element(By.NAME, "y1").click()
        driver.find_element(By.NAME, "y1").clear()
        driver.find_element(By.NAME, "y1").send_keys('KLnA67LW')
        driver.find_element(By.XPATH, '//button[@onclick="doLogin()"]').click()
        message = f'{datetime.now()} - ??????????????????! '
        self.sinOut.emit(message)            
        message = f'{datetime.now()} - ???????????????????????????...'
        self.sinOut.emit(message)
        time.sleep(2)
        driver.get(url='http://192.168.10.47:8080/glaf/apps/bill.do?flag=billConfirm')
        time.sleep(2)
        message = f'{datetime.now()} - ??????????????????: {year} ??? {month} ???'
        self.sinOut.emit(message)
        driver.find_element(By.XPATH, '//span[@id="select2-iyear-container"]').click()
        driver.find_element(By.XPATH, '//input[@class="select2-search__field"]').send_keys(year)
        driver.find_element(By.XPATH, '//li[@class="select2-results__option select2-results__option--highlighted"]').click()
        driver.find_element(By.XPATH, '//span[@id="select2-imonth-container"]').click()
        driver.find_element(By.XPATH, '//input[@class="select2-search__field"]').send_keys(month)
        driver.find_element(By.XPATH, '//li[@class="select2-results__option select2-results__option--highlighted"]').click()
        driver.find_element(By.XPATH, '//button[@id="Search_Btn"]').click()
        time.sleep(1)
        driver.find_element(By.XPATH, '//input[@name="ck"]').click()
        driver.find_element(By.XPATH, '//button[@onclick="javascript:invoice();"]').click()
        time.sleep(2)

        message = f'{datetime.now()} - ??????????????????...'
        self.sinOut.emit(message)
        iframe = driver.find_element(By.XPATH, '//iframe[@id="layui-layer-iframe1"]')
        driver.switch_to.frame("layui-layer-iframe1")

        for item in col_list:
            try: 
                driver.find_element(By.XPATH, '//button[@id="Add_Btn"]').click()
                autofill(num1, col_list[row][0], col_list[row][1],col_list[row][2])
                message = f'?????????{col_list[row][0]} ???????????????{col_list[row][1]} ??????{col_list[row][2]} ????????????!'
                self.sinOut.emit(message)
                num1 = num1 + 8
                row = row + 1
            except:
                message = f'?????????{col_list[row][0]} ???????????????{col_list[row][1]} ??????{col_list[row][2]} ????????????!'
                self.sinOut.emit(message)
        time.sleep(1)
        message = f'{datetime.now()} - ????????????, ???{df.shape[0]}???, ??????{row}???. '
        self.sinOut.emit(message)
        message = 'DONE'
        self.sinOut.emit(message)

    except Exception:
        driver.execute_script('window.stop()')
        message = f'{datetime.now()} - ??????????????????, ?????????VPN??????????????????! '
        self.sinOut.emit(message)
        #driver.quit()

class MyWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.thread = Worker()
        title = 'GTE???????????????????????? v0.3   - Made by REC3WX'
        self.setWindowTitle(title)
        pixmapi = QStyle.SP_FileDialogDetailedView
        icon = self.style().standardIcon(pixmapi)
        self.setWindowIcon(icon)
        self.setFixedSize(700, 300)
        self.setWindowFlags(
            Qt.Window |
            Qt.CustomizeWindowHint |
            Qt.WindowTitleHint |
            Qt.WindowCloseButtonHint |
            Qt.WindowStaysOnTopHint
            )

        self.fld_csv= QLabel('??????CSV??????:')
        self.btn_csv= QPushButton('??????')
        self.btn_csv.clicked.connect(self.opencsvDialog)
        self.line_csv = QLineEdit()
        #self.line_csv.setClearButtonEnabled(True) #????????????

        self.line_csv.setAcceptDrops(True)
        
        self.fld_year= QLabel('????????????:')
        self.cb_year= QComboBox()
        year = datetime.today().year
        self.cb_year.addItems(['', str(year-1), str(year), str(year+1)])
        self.cb_year.currentTextChanged[str].connect(self.get_year_month)

        self.fld_month= QLabel('????????????:')
        self.cb_month = QComboBox()
        self.cb_month.addItems(['', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12'])
        self.cb_month.currentTextChanged[str].connect(self.get_year_month)

        self.btn_start = QPushButton('??????')
        self.btn_start.setEnabled(False)
        self.btn_start.clicked.connect(self.execute)
        
        self.fld_result = QLabel('????????????:')
        self.text_result = QPlainTextEdit()
        self.text_result.setReadOnly(True)

        self.line = QFrame()
        self.line.setFrameShape(QFrame.HLine)
        self.line.setFrameShadow(QFrame.Sunken)

        self.btn_reset = QPushButton('??????')
        self.btn_reset.clicked.connect(self.reset)

        self.layout = QGridLayout()
        self.layout.addWidget((self.fld_csv), 0, 0)
        self.layout.addWidget((self.line_csv), 0, 1, 1, 6)
        self.layout.addWidget((self.btn_csv), 0, 7)
        self.layout.addWidget((self.fld_year), 1, 3)
        self.layout.addWidget((self.cb_year), 1, 4)
        self.layout.addWidget((self.fld_month), 1, 5)
        self.layout.addWidget((self.cb_month), 1, 6)

        self.layout.addWidget((self.line), 2, 0, 1, 7)
        self.layout.addWidget((self.fld_result), 3, 0)
        self.layout.addWidget((self.text_result), 4, 0, 5, 7)
        self.layout.addWidget((self.btn_start), 5, 7, 1, 1)
        self.layout.addWidget((self.btn_reset), 6, 7, 1, 1)

        self.setLayout(self.layout)
        
        self.thread.sinOut.connect(self.Addmsg)  #????????????emit

    @Slot()
    def Addmsg(self, message):
        if not message == 'DONE':
            self.text_result.appendPlainText(message)
        else:
            self.msgbox('DONE', '????????????, ??????????????????!! ')

    def get_year_month(self):
        year = str(self.cb_year.currentText())
        month = str(self.cb_month.currentText())
        title = 'GTE???????????????????????? v0.3   - Made by REC3WX'
        if not year == '' and not month == '':
            #self.text_result.appendPlainText(f'??????????????????????????????: {year}???{month}???')
            self.setWindowTitle(f'{title} - ????????????: {year}???{month}???')

    def reset(self):
        self.line_csv.setText('')
        self.cb_year.setCurrentText('')
        self.cb_month.setCurrentText('')
        self.text_result.clear()
        self.btn_start.setEnabled(False)

    def opencsvDialog(self):
        dialog = QFileDialog()
        dialog.setFileMode(QFileDialog.ExistingFile)
        dialog.setViewMode(QFileDialog.Detail)
        dialog.setNameFilter("CSV?????? (*.csv)")
        if dialog.exec():
            fileNames = dialog.selectedFiles()
            self.line_csv.setText(fileNames[0])
            self.btn_start.setEnabled(True)

    def msgbox(self, title, text):
        tip = QMessageBox(self)
        if title == 'error':
            tip.setIcon(QMessageBox.Critical)
        elif title == 'DONE' :
            tip.setIcon(QMessageBox.Warning)
        tip.setWindowFlag(Qt.FramelessWindowHint)
        font = QFont()
        font.setFamily("Microsoft YaHei")
        font.setPointSize(9)
        tip.setFont(font)
        tip.setText(text)
        tip.exec()  

    def execute(self):
        year = str(self.cb_year.currentText())
        month = str(self.cb_month.currentText())
        if year == '' or month == '':
            self.msgbox('error', '??????????????????????????????!! ')
        else:
            self.thread.getdata(self.line_csv.text(), year, month)
            self.thread.start()

def main():
    if not QApplication.instance():
        app = QApplication(sys.argv)
    else:
        app = QApplication.instance()
    app.setStyle('fusion')
    font = QFont()
    font.setFamily("Microsoft YaHei")
    font.setPointSize(10)
    app.setFont(font)
    widget = MyWidget()
    widget.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()