# -*- coding: utf-8 -*-
"""
Created on Thu Aug  4 12:42:37 2022

@author: REC3WX

Tool to mass input the VAT on GTE Tax Recon website
CMD:
"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --remote-debugging-port=9222 --user-data-dir="C:\selenium\EdgeProfile"
"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\selenium\ChromeProfile"
<<<<<<< HEAD
ChromeDriverManager-chrome.py
            url: str = "http://npm.taobao.org/mirrors/chromedriver",
            latest_release_url: str = "http://npm.taobao.org/mirrors/chromedriver/LATEST_RELEASE",
=======
>>>>>>> b258b2fe63b7d83af2067dce0956d130ba7c5da0
"""


from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
<<<<<<< HEAD
from webdriver_manager.chrome  import ChromeDriverManager
=======
##from webdriver_manager.chrome  import ChromeDriverManager
>>>>>>> b258b2fe63b7d83af2067dce0956d130ba7c5da0
import pandas as pd
import os
from tkinter import Tk 
from tkinter.filedialog import askopenfilename

os.popen("chcp 936")
#year = input("请输入发票年份(YYYY):")
#month = input("请输入发票月份(M):")
year = ('2022')
month = ('7')

opt = Options()
opt.add_experimental_option("debuggerAddress", "localhost:9222")
"""
opt.add_argument("--remote-debugging-port=9222")
opt.add_argument('user-data-dir=C:\\selenium\\ChromeProfile')
<<<<<<< HEAD
driver_path = ChromeService(r'./chromedriver.exe')
driver = webdriver.Chrome(service=driver_path, options=opt)
"""
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), chrome_options=opt)
=======
"""
driver_path = ChromeService(r'./chromedriver.exe')
driver = webdriver.Chrome(service=driver_path, options=opt)
#driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), chrome_options=opt)
>>>>>>> b258b2fe63b7d83af2067dce0956d130ba7c5da0
driver.set_window_size('1920', '1080')
print(driver)
"""
driver.get(url='https://www.xd.com/users/register/?')
driver.get(url='http://192.168.10.47:8080/glaf/loginApp.do')
driver.implicitly_wait(5)
driver.find_element(By.NAME, "x").click()
driver.find_element(By.NAME, "x").clear()
driver.find_element(By.NAME, "x").send_keys('3334')
driver.find_element(By.NAME, "y1").click()
driver.find_element(By.NAME, "y1").clear()
driver.find_element(By.NAME, "y1").send_keys('KLnA67LW')
driver.find_element(By.XPATH, '//button[@onclick="doLogin()"]').click()
driver.implicitly_wait(5)
driver.find_element(By.XPATH, '/html/body/aside/nav/ul/li[4]/a').click()
"""
driver.implicitly_wait(1)
##driver.find_element(By.XPATH, '/html/body/aside/nav/ul/li[4]/ul/li[1]/a').click()
driver.implicitly_wait(1)
##driver.get(url='http://192.168.10.47:8080/glaf/index.do')
driver.implicitly_wait(1)
driver.find_element(By.XPATH, '//span[@id="select2-iyear-container"]').click()
driver.implicitly_wait(1)
driver.find_element(By.XPATH, '//span[@id="select2-imonth-container"]').click()
driver.implicitly_wait(1)
driver.find_element(By.XPATH, '//button[@id="Search_Btn"]').click()


"""
def autofill(fphm, net, vat):
    driver.find_element(By.NAME, 'username').send_keys(fphm)
    driver.find_element(By.NAME,  'password').send_keys('12345AAAAA!')
    driver.find_element(By.NAME, 'confirm').send_keys('12345AAAAA!')
    driver.find_element(By.NAME, 'realname').send_keys(net)
    driver.find_element(By.NAME, 'realid').send_keys('110101199001013195')
    driver.find_element(By.NAME, 'email').send_keys(vat,'@123.com')
    driver.find_element(By.NAME, 'mobile').send_keys('13800138000')
    driver.find_element(By.NAME, 'agreement').click()
    driver.implicitly_wait(2)
    #driver.find_element(By.CLASS_NAME, 'geetest_radar_tip').click()

Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing

csvfile = askopenfilename(
    filetypes=[('CSV文件', '.csv')], title='选择发票信息CSV文件'
)

if not csvfile == '':
    df = pd.read_csv(csvfile, dtype=str, header=0)
    
    col_list = df.values.tolist()
    row = 0
    fail=[]
    for item in col_list:
        try: 
            autofill(col_list[row][0], col_list[row][1],col_list[row][2])
            row = row + 1
        except:
            fail.append(col_list[row])
<<<<<<< HEAD
"""
=======
"""
>>>>>>> b258b2fe63b7d83af2067dce0956d130ba7c5da0
