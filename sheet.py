#!/usr/bin/env python

import gspread
from oauth2client.service_account import ServiceAccountCredentials
import sys
from config import Config

class GoogleSheet():    #試算表名稱 #工作表名稱     #憑證檔名
    def __init__(self, wks_name, wks_title=None, oauth=Config["oauth"]):
        #設定存取範疇的區域變數
        scopes = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive"
        ]

        try:
            # 嘗試讀取憑證檔
            cr = ServiceAccountCredentials.from_json_keyfile_name(oauth, scopes)
        except:
            print("無法開啟憑證檔")
            # 關閉程式
            sys.exit(1)

        try:
            # 嘗試開啟試算表
            gc = gspread.authorize(cr)
            sh = gc.open(wks_name)
        except:
            print("無法開啟試算表")
            sys.exit(1)

        # 若無指定工作表
        if wks_title is None:
            # 工作表為表單1
            self._wks = sh.sheet1
        else:
            #嘗試開啟指定工作表
            try:
                self._wks = sh.worksheet(wks_title)
            except:
                print("無法開啟工作表")
                sys.exit(1)
    
    @property
    # 回傳第一列資料
    def header(self):
        return self._wks.row_values(1)

    def update_header(self, data, delete=True):
        # 刪除第一列
        if delete:
            self._wks.delete_row(1)
        #插入第一列
        self._wks.append_row(data,1)

    # 插入新的參數一列
    def append_row(self, data):
        self._wks.append_row(data)
    
    #調整試算表大小 預設縮小至一列
    def resize(self ,n=1):
        self._wks.resize(n)