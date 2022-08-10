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


from argparse import FileType
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome  import ChromeDriverManager
import pandas as pd
import sys
import time
from tkinter import simpledialog, messagebox
from PySide6.QtWidgets import QWidget, QPushButton, QFileDialog, QApplication, QLineEdit, QGridLayout, QLabel, QMessageBox, QPlainTextEdit, QFrame, QStyle, QComboBox
from PySide6.QtGui import QFont
from PySide6.QtCore import Slot, Qt, QThread, QObject, Signal
import qdarktheme
import threading

class MyWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setWindowTitle('GTE发票批量录入工具 v0.1   - Made by REC3WX')
        pixmapi = QStyle.SP_FileDialogDetailedView
        icon = self.style().standardIcon(pixmapi)
        self.setWindowIcon(icon)
        self.setFixedSize(700, 300)

        self.fld_csv= QLabel('发票CSV文件:')
        self.btn_csv= QPushButton('打开')
        self.btn_csv.clicked.connect(self.opencsvDialog)
        self.line_csv = QLineEdit()
        self.line_csv.setClearButtonEnabled(True) #清空按钮
        
        self.fld_year= QLabel('发票年份:')
        self.cb_year= QComboBox()
        self.cb_year.addItems(['', '2022', '2023', '2024', '2025', '2026', '2027', '2028', '2029', '2030'])
        self.cb_year.currentTextChanged[str].connect(self.get_year)

        self.fld_month= QLabel('发票月份:')
        self.cb_month = QComboBox()
        self.cb_month.addItems(['', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12'])
        self.cb_month.currentTextChanged[str].connect(self.get_month)

        self.btn_start = QPushButton('开始')
        self.btn_start.setEnabled(False)
        self.btn_start.clicked.connect(self.execute)


        self.fld_result = QLabel('运行日志:')
        self.text_result = QPlainTextEdit()
        self.text_result.setReadOnly(True)

        self.line = QFrame()
        self.line.setFrameShape(QFrame.HLine)
        self.line.setFrameShadow(QFrame.Sunken)

        self.btn_reset = QPushButton('清空')
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

    @Slot()
    def get_year(self):
        year = str(self.cb_year.currentText())
        self.text_result.appendPlainText(r"当前选择的发票年份为: {}年".format(year))

    def get_month(self):
        month = str(self.cb_month.currentText())
        self.text_result.appendPlainText(r"当前选择的发票月份为: {}月".format(month))

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
        dialog.setNameFilter("CSV文件 (*.csv)")
        if dialog.exec():
            fileNames = dialog.selectedFiles()
            self.line_csv.setText(fileNames[0])
            self.btn_start.setEnabled(True)

    def msgbox(self, title, text):
        tip = QMessageBox(self)
        if title == 'error':
            tip.setIcon(QMessageBox.Critical)
        elif title == 'done' :
            tip.setIcon(QMessageBox.Warning)
        tip.setWindowFlag(Qt.FramelessWindowHint)
        font = QFont()
        font.setFamily("Microsoft YaHei")
        font.setPointSize(9)
        tip.setFont(font)
        tip.setText(text)
        tip.exec()

    def execute(self):
        def autofill(num1, fphm, net, vat):
            num2 = num1 + 6
            num3 = num1 + 7
            path1 = (r'//input[@id="_easyui_textbox_input{}"]'.format(num1))
            path2 = (r'//input[@id="_easyui_textbox_input{}"]'.format(num2))
            path3 = (r'//input[@id="_easyui_textbox_input{}"]'.format(num3)) 
            driver.find_element(By.XPATH, path1).click()
            driver.find_element(By.XPATH, path1).send_keys(fphm)
            driver.find_element(By.XPATH, path2).click()
            driver.find_element(By.XPATH, path2).send_keys(net)
            driver.find_element(By.XPATH, path3).click()
            driver.find_element(By.XPATH, path3).send_keys(vat)

        year = str(self.cb_year.currentText())
        month = str(self.cb_month.currentText())

        if year == '' or month == '':
                self.msgbox('error', '请选择并确认发票年月!! ')
        else:
            df = pd.read_csv(self.line_csv.text(), dtype=str, header=0)
            
            col_list = df.values.tolist()
            row = 0
            num1 = 1
            opt = Options()
            #opt.add_experimental_option("debuggerAddress", "localhost:9222")

            opt.add_argument("--remote-debugging-port=9222")
            opt.add_argument("--start-maximized")
            opt.add_argument('user-data-dir=C:\\selenium\\ChromeProfile')
            ##driver_path = ChromeService(r'./chromedriver.exe')
            ##driver = webdriver.Chrome(service=driver_path, options=opt)
            self.text_result.appendPlainText('打开Chrome并自动登录...')
            driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=opt)
            driver.set_page_load_timeout(25)
            try:
                driver.get(url='http://192.168.10.47:8080/glaf/loginApp.do')
                time.sleep(2)
                driver.find_element(By.NAME, "x").click()
                driver.find_element(By.NAME, "x").clear()
                driver.find_element(By.NAME, "x").send_keys('3334')
                driver.find_element(By.NAME, "y1").click()
                driver.find_element(By.NAME, "y1").clear()
                driver.find_element(By.NAME, "y1").send_keys('KLnA67LW')
                driver.find_element(By.XPATH, '//button[@onclick="doLogin()"]').click()
                self.text_result.appendPlainText('打开发票录入单窗口...')
                time.sleep(2)
                driver.get(url='http://192.168.10.47:8080/glaf/apps/bill.do?flag=billConfirm')
                time.sleep(2)
                self.text_result.appendPlainText('筛选发票年月...')
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
                iframe = driver.find_element(By.XPATH, '//iframe[@id="layui-layer-iframe1"]')
                driver.switch_to.frame("layui-layer-iframe1")
                self.text_result.appendPlainText('开始录入发票...')
                for item in col_list:
                    try: 
                        driver.find_element(By.XPATH, '//button[@id="Add_Btn"]').click()
                        autofill(num1, col_list[row][0], col_list[row][1],col_list[row][2])
                        num1 = num1 + 8
                        row = row + 1
                    except:
                        self.text_result.appendPlainText(col_list[row],'录入失败!')
                time.sleep(1)
                self.text_result.appendPlainText('录入完成，请确认后保存!! ')

            except Exception:
                driver.execute_script('window.stop()')
                self.text_result.appendPlainText('网页无法加载, 请确认VPN连接是否正常! ')
                driver.quit()

def main():
    if not QApplication.instance():
        app = QApplication(sys.argv)
    else:
        app = QApplication.instance()
    app.setStyleSheet(qdarktheme.load_stylesheet())
    font = QFont()
    font.setFamily("Microsoft YaHei")
    font.setPointSize(10)
    app.setFont(font)
    widget = MyWidget()
    widget.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()