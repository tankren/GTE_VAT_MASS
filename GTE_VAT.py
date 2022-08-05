# -*- coding: utf-8 -*-
"""
Created on Thu Aug  4 12:42:37 2022

@author: REC3WX

Tool to mass input the VAT on GTE Tax Recon website
CMD:
chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\selenum\Profile"
"""


from selenium import webdriver
from selenium.webdriver.common.by import By
#from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service as EdgeService
#from webdriver_manager.microsoft import EdgeChromiumDriverManager
#from requests import get
import pandas as pd
import os
import time
from tkinter import Tk  # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askopenfilename

os.popen("chcp 936")

opt = Options()
opt.add_experimental_option("debuggerAddress", "localhost:9222")
driver_path = EdgeService(r'./msedgedriver.exe')
# opt.add_argument("--remote-debugging-port=9222")
# opt.add_argument('user-data-dir=C:\\selenium\\Profile')
#driver = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()), options=opt)
driver = webdriver.Edge(service=driver_path, options=opt)
print(driver)


driver.get(url='https://www.xd.com/users/register/?')
driver.implicitly_wait(10)


def autofill(fphm, net, vat):
    driver.find_element(By.NAME, 'username').send_keys(fphm)
    driver.find_element(By.NAME,  'password').send_keys('12345AAAAA!')
    driver.find_element(By.NAME, 'confirm').send_keys('12345AAAAA!')
    driver.find_element(By.NAME, 'realname').send_keys(net)
    driver.find_element(By.NAME, 'realid').send_keys('110101199001013195')
    driver.find_element(By.NAME, 'email').send_keys(vat, '@123.com')
    driver.find_element(By.NAME, 'mobile').send_keys('13800138000')
    driver.find_element(By.NAME, 'agreement').click()
    time.sleep(2)
    #driver.find_element(By.CLASS_NAME, 'geetest_radar_tip').click()


Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing

csvfile = askopenfilename(
    filetypes=[('CSV文件', '.csv')], title='选择发票信息CSV文件'
)

if not csvfile == '':
    df = pd.read_csv(csvfile, dtype=str, header=0)

    col_list = df.values.tolist()
    row = 0
    fail = []
    for item in col_list:
        try:
            autofill(col_list[row][0], col_list[row][1], col_list[row][2])
            row = row + 1
        except:
            fail.append(col_list[row])
