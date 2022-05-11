#!/usr/bin/env python

from selenium import webdriver
from selenium.webdriver.common.by import By
from urllib import parse
from time import *
from config import Config

#引入自訂類別
import sheet

sh_name = input("輸入試算表名稱\n")
wks_name = input("輸入工作表名稱\n")
#parse.quote 轉換成url編碼
search_key = parse.quote(input("搜尋關鍵字\n"),encoding= "UTF-8")    

while True:
    confirm = input("確認關鍵字:T/F\n").upper().strip()
    if confirm == "T":
        break
    else:
        search_key = input("搜尋關鍵字\n")

print("開始搜尋")

# 各網站元素
sites = [
    {
        'web':"PCHome24h",
        'url':"https://ecshweb.pchome.com.tw/search/v3.3/?q=%s"%(search_key),
        'title_path':"//dd[@class='c2f']/h5[@class='prod_name']/a",
        'price_path':"//span[@class='price']/span[@class='value']",
        'link_path':"//h5[@class='prod_name']/a"
    },
    {
        'web':"蝦皮",
        'url':"https://shopee.tw/search?keyword=%s"%(search_key),
        'title_path':"//div[@class='ie3A+n bM+7UW Cve6sh']",
        'price_path':"//span[@class='ZEgDH9']",
        'link_path':"//div[@class='col-xs-2-4 shopee-search-item-result__item']/a"
    },
]
# selenium 環境建置
driver_path = Config["driver"]
option = webdriver.ChromeOptions()
# 開啟無痕模式
option.add_argument('headless')
option.add_argument('--log-level=3')
driver = webdriver.Chrome(driver_path, chrome_options=option)
driver.implicitly_wait(10)

# 開啟google表單
gs = sheet.GoogleSheet(sh_name, wks_name)
# 清空表單
gs._wks.clear()

# 開啟網站讀取元素
def open_page(title_path, price_path, link_path):
    driver.get(url)
    titles = driver.find_elements(By.XPATH,title_path)
    prices = driver.find_elements(By.XPATH,price_path)
    links  = driver.find_elements(By.XPATH,link_path)
    return(titles, prices, links)

# 編寫標題 "網站", "時間", "標題", "價錢", "連結"
write_header = True


def save_google_sheet(titles, prices, links):
    global write_header
    if write_header:
        gs.append_row(["網站", "時間", "標題", "價錢", "連結"])
        write_header = False
    
    for t, p, l in zip(titles, prices, links):
        _w = web
        dt = strftime("%Y/%m/%d %H:%M")
        _t = t.text
        _p = "$ "+p.text
        _l = l.get_attribute('href')
        gs.append_row([_w, dt, _t, _p, _l])


for s in sites:
    web = s['web']
    url = s['url']
    title_path = s['title_path']
    price_path = s['price_path']
    link_path = s['link_path']
    
    # 將元素儲存進變數
    titles, prices, links = open_page(title_path, price_path, link_path)
    # 將變數儲存進google表單
    save_google_sheet(titles, prices, links)

print("搜尋完成")

driver.close()