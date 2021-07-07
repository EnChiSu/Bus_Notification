#!/usr/bin/env python
# coding: utf-8

# In[ ]:


#傳送line訊息通知5分鐘內車將到站
import requests
from bs4 import BeautifulSoup
import smtplib
import time
from email.mime.text import MIMEText
from requests_html import HTMLSession
from selenium import webdriver
import time
import xlwings as xw
from xlwings.constants import Direction
import numpy as np
import pandas as pd
from pandas import read_csv
# matplotlib
import matplotlib.pyplot as plt
get_ipython().run_line_magic('matplotlib', 'inline')
#解決不能輸出中文字
from pylab import mpl
mpl.rcParams['font.sans-serif'] = ['Microsoft YaHei']  
mpl.rcParams['axes.unicode_minus'] = False
# seaborn
import seaborn as sns

def send_line():
    # Line Notify api 網址
    url = "https://notify-api.line.me/api/notify"
    # 圖片路徑
    pic_url = r"C:\Users\steve\Documents\Python\test.png"
    # 訊息
    msg = "606即將進站"
    token = "YA0UN0zW31KGoyYC2iSr2BKamZp02Etl1MMfQI3ssLA"
    
    headers = {
        "Authorization": "Bearer " + token
    }
    
    # 設定訊息内容
    payload = { "message": msg }
    # 設定圖檔路徑  #跟之前不同的地方
    files = { "imageFile": open(pic_url, "rb") }
    # 發送請求
    r = requests.post(url, headers = headers, params = payload, files = files)

    

def check_bus():
    wd = webdriver.Chrome()
    wd.get("http://www.e-bus.taipei.gov.tw/newmap/Tw/Map?rid=11815&sec=1")
    html = wd.page_source
    soup = BeautifulSoup(html, "html.parser")
    resultcol = soup.findAll("span", {"class": "eta_onroad"})
    time = soup.find("span", {"id": "plLastUpdateTime"})
    time = time.text
    time = time.replace("更新時間：","")
    
    cmt = []
    for resultcol in soup.findAll("div", {"class": "eta"}):
        cmt.append(resultcol.text)
    
    for i in range(len(cmt)):
        if cmt[i]=="將到站":
            cmt[i]=1
        elif len(cmt[i])==6 or len(cmt[i])==8 or len(cmt[i])==12 or len(cmt[i])==14 or len(cmt[i])==16:
            cmt[i]=0
        else:
            cmt[i]=cmt[i].replace("約", "")
            cmt[i]=cmt[i].replace("分", "")
            cmt[i]=int(cmt[i])
    
    print(cmt)
            
    wb = xw.Book(r"bus_data.xlsx") #加寫一個r代表後面雙引號的文字是原始的文字，叫python在讀的時候如果看到\不要視為跳脫字元
    sheet = wb.sheets["星期一"] #記得星期二要改為"星期二"、星期三改為"星期三"...
    last_row = sheet.range("A1").end(Direction.xlDown).row
    sheet.range(f"A{last_row+1}").value =time
    sheet.range(f"B{last_row+1}").value =cmt
    wb.save()    
    
    dta = pd.read_excel('bus_data.xlsx', sheet_name="星期一", header=0, index_col="Time", encoding="big5",na_values=" NaN")
    dta = dta.dropna().astype(int)
    dta = dta[(last_row-36):last_row]
    plt.figure(figsize=(12, 9))
    sns.heatmap(dta.astype(float), cmap="BuPu")
    plt.savefig("C:\\Users\\steve\\Documents\\Python\\test.png")
    
    
    if (cmt[9]<=15):
        send_line()

while(True):
    check_bus()
    time.sleep(293)

