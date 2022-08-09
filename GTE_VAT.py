# -*- coding: utf-8 -*-
"""
Created on Thu Aug  4 12:42:37 2022

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
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome  import ChromeDriverManager
import pandas as pd
import os
import time
from tkinter import Tk 
from tkinter.filedialog import askopenfilename
from tkinter import simpledialog, messagebox

os.popen("chcp 936")

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


Tk().withdraw()

csvfile = askopenfilename(
    filetypes=[('CSV文件', '.csv')], title='选择发票信息CSV文件'
)
print(csvfile)

if not csvfile == '':
    year = simpledialog.askstring("请输入", "发票年份(YYYY):")
    month = simpledialog.askstring("请输入", "发票月份(M):")
                    
    #year = input("请输入发票年份(YYYY):")
    #month = input("请输入发票月份(M):")
    #year = ('2022')
    #month = ('7')


    df = pd.read_csv(csvfile, dtype=str, header=0)
    
    col_list = df.values.tolist()
    row = 0
    num1 = 1
    fail=[]
    opt = Options()
    #opt.add_experimental_option("debuggerAddress", "localhost:9222")

    opt.add_argument("--remote-debugging-port=9222")
    opt.add_argument('user-data-dir=C:\\selenium\\ChromeProfile')
    """
    driver_path = ChromeService(r'./chromedriver.exe')
    driver = webdriver.Chrome(service=driver_path, options=opt)
    """
    print('打开Chrome并自动登录...')
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=opt)
    driver.fullscreen_window()
    driver.get(url='http://192.168.10.47:8080/glaf/loginApp.do')
    time.sleep(2)
    driver.find_element(By.NAME, "x").click()
    driver.find_element(By.NAME, "x").clear()
    driver.find_element(By.NAME, "x").send_keys('3334')
    driver.find_element(By.NAME, "y1").click()
    driver.find_element(By.NAME, "y1").clear()
    driver.find_element(By.NAME, "y1").send_keys('KLnA67LW')
    driver.find_element(By.XPATH, '//button[@onclick="doLogin()"]').click()
    print('打开发票录入单窗口...')
    time.sleep(2)
    driver.get(url='http://192.168.10.47:8080/glaf/apps/bill.do?flag=billConfirm')
    ##driver.find_element(By.XPATH, '/html/body/aside/nav/ul/li[4]/a').click()

    ##driver.implicitly_wait(1)
    ##driver.find_element(By.XPATH, '/html/body/aside/nav/ul/li[4]/ul/li[1]/a').click()
    print('筛选发票年月...')
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
    print('开始录入发票...')
    for item in col_list:
        try: 
            driver.find_element(By.XPATH, '//button[@id="Add_Btn"]').click()
            autofill(num1, col_list[row][0], col_list[row][1],col_list[row][2])
            num1 = num1 + 8
            row = row + 1
        except:
            fail.append(col_list[row],'录入失败!')

    if not fail:
        print(fail)
    #driver.find_element(By.XPATH, '//button[@id="Save_Btn"]).click()  #自动保存，暂不开启
    time.sleep(1)
    messagebox.showinfo(title='成功', message='录入完成，请确认后保存!')
    
else:
    messagebox.showerror(title='错误', message='请选择csv文件!')