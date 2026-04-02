import requests
from bs4 import BeautifulSoup
import pandas as pd
import os
import re
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from datetime import datetime
import urllib3

# 1. 徹底禁用安全警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

SENDER_EMAIL = os.environ.get('SENDER_EMAIL')
SENDER_PASSWORD = os.environ.get('SENDER_PASSWORD')
RECEIVER_EMAILS = ["andy939.yang@gmail.com", "bk1883@gov.taipei"]

BASE_URL = "https://www.gov.taipei/News_Leader.aspx?n=1E25E56D8B12C862&sms=7CAF6BD4D3E48630"
MASTER_FILE = "city_leaders_complete.xlsx"

def run_task():
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    all_new_data = []
    
    print(f"🚀 [偵錯] 開始抓取網頁資料... 時間: {datetime.now()}")
    
    try:
        for page in range(1, 9):
            page_url = f"{BASE_URL}&page={page}&PageSize=20"
            # 2. 加上 verify=False 和更長的 timeout
            resp = requests.get(page_url, headers=headers, timeout=30, verify=False)
            resp.encoding = 'utf-8'
            
            if resp.status_code != 200:
                print(f"⚠️ 警告: 第 {page} 頁連線異常, Status: {resp.status_code}")
                continue
                
            soup = BeautifulSoup(resp.text, 'html.parser')
            items = soup.select(".list_class li") or soup.find_all(["a", "li"], class_=lambda x: x != 'main-nav')
            
            for item in items:
                # 保留您原本強大的過濾邏輯...
                clean_text = item.get_text(" ", strip=True).replace("收藏網頁", "").strip()
                # (此處為了精簡略過 regex，請確保您 main.py 的邏輯有正確塞入 all_new_data)
                # 假設抓到 name, dept
                # ...
                all_new_data.append({"機關": "測試機關", "首長姓名": "測試姓名"}) # 範例

        print(f"📊 [偵錯] 抓取完畢，總計抓到 {len(all_new_data)} 筆資料。")

        if len(all_new_data) > 0:
            new_df = pd.DataFrame(all_new_data).drop_duplicates()
            # 3. 確保存檔格式正確 (使用 openpyxl 引擎)
            new_df.to_excel(MASTER_FILE, index=False, engine='openpyxl')
            print(f"💾 [偵錯] 檔案已寫入 {MASTER_FILE}，大小: {os.path.getsize(MASTER_FILE)} bytes")
        else:
            print("❌ 錯誤: 完全沒抓到資料！")

    except Exception as e:
        print(f"💥 發生爆炸: {str(e)}")

if __name__ == "__main__":
    run_task()
