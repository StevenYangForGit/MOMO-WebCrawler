# -*- coding: utf-8 -*-
"""
Created on Thu Aug 12 21:57:58 2021

@author: user
"""

import tkinter as tk
import tkinter.ttk
import tkinter.filedialog
import os
import pandas as pd
from openpyxl import load_workbook
from selenium import webdriver
from bs4 import BeautifulSoup
import random
import time
from tkinter.messagebox import showinfo

def RandomTimeSleep():
   RandomNum = random.uniform(1, 10)
   time.sleep(RandomNum)

def GetVAData():
    return r"VA"

def GetETMallData():
    return r"ETMall"

def GetMOMOData(itemno):
    url = 'https://www.momoshop.com.tw/goods/GoodsDetail.jsp?i_code={}&Area=search&mdiv=403&oid=1_4&cid=index&kw=%E4%BA%9E%E5%B8%9D%E8%8A%AC%E5%A5%87'.format(itemno)
  
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('User-Agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.70 Safari/537.36"')
    chrome_options.add_argument('--disable-gpu') #谷歌文档提到需要加上这个属性来规避bug
    chrome_options.add_argument('blink-settings=imagesEnabled=false') #不加载图片, 提升速度
    chrome_options.add_argument('--headless') #浏览器不提供可视化页面. linux下如果系统不支持可视化不加这条会启动失败
    driver = webdriver.Chrome(chrome_options=chrome_options)
    driver.get(url)

    soup = BeautifulSoup(driver.page_source, "html.parser")

    a = [link.find_all('li')[-1].text for link in soup.find_all('table',id ="attributesTable")]
    
    if a:
        result = a[0].split("\u25C6")[-1]
    else:
        result = '很抱歉，查無'+itemno+'的相關商品'
    
    RandomTimeSleep()
    
    return result

def OverWriteFile():
    result = []
    
    df_dict = pd.read_excel(filePath.get(),header=4,usecols=[0,1],names=['class','itemno'],dtype=str, engine='openpyxl').to_dict('records')
    
    for count, data in enumerate(df_dict):
        if data['class'] == 'MO':
            GetData = GetMOMOData(data['itemno'])  
            
        if data['class'] == 'EHS':
            GetData = GetETMallData()
            
        if data['class'] == 'VA':
            GetData = GetVAData()
            
        result.append(GetData)
        
        progressbarOne['value'] = int((count/(len(df_dict)-1))*100)
        value_label['text'] = progressbarOne['value'],'%'
        root.update_idletasks()
        time.sleep(0.5)
    
    df_data = pd.DataFrame(result,columns=['result'])
    
    with pd.ExcelWriter(filePath.get(), mode='a', engine='openpyxl') as writer:
        book = load_workbook(filePath.get())
    
        writer.book = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        df_data.to_excel(writer, sheet_name='Sheet1', header=None, index=False,startcol=6,startrow=5)
            
    showinfo(message='The progress completed!')
        
def GetFile():
    default_dir = r"選擇檔案"
    filePaths = tk.filedialog.askopenfilename(title=u'選擇檔案', initialdir=(os.path.expanduser(default_dir)),filetypes=[('Excel', '.xlsx')])
    filePath.delete(0, "end")
    filePath.insert(0, filePaths)

root = tk.Tk() # 產生 root 主視窗
root.geometry("600x300")
root.title("") # 標題
#root.iconbitmap('diamond.png')
#root.iconphoto(False, tk.PhotoImage(file='diamond.png'))


filePath = tk.Entry(root)
filePath.pack(anchor=tk.SW)
filePath.place(relx = 0.22, rely = 0.2,height=40, width=260)


getFile = tk.Button(root, text='選擇檔案',command=GetFile)
getFile.pack(anchor=tk.SE)
getFile.place(relx = 0.65, rely = 0.2, height=40, width=80)

submit = tk.Button(root, text='確定',command=OverWriteFile)
submit.place(relx = 0.4, rely = 0.6,height=50, width=100)

value_label = tk.Label(root)
value_label.place(x = 240, y = 235,height=50, width=100)

progressbarOne = tk.ttk.Progressbar(root,orient = 'horizontal',length = 100,mode = 'determinate')
progressbarOne.pack(fill='x',side="bottom")
# 进度值最大值
progressbarOne['maximum'] = 100
# 进度值初始值
progressbarOne['value'] = 0

root.mainloop()  