import requests
from bs4 import BeautifulSoup
import pandas as pd
import os
import glob
from datetime import datetime, timedelta
import threading
import urllib3
import re
import time
import smtplib
import schedule 
from email.mime.text import MIMEText
from email.header import Header

# 💡 加入這行，強制讓 Python 腳本內的檔名生成使用台灣時區
os.environ['TZ'] = 'Asia/Taipei'
if hasattr(time, 'tzset'):
    time.tzset()

# --- 基礎配置 (針對 GitHub Actions 修正) ---
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
# 💡 修正 1：改從 GitHub Secrets 讀取，不要寫死在程式碼中
SENDER_EMAIL = os.environ.get("SENDER_EMAIL")        
SENDER_PASSWORD = os.environ.get("SENDER_PASSWORD") 
RECIPIENT_FILE = "收件者清單.xlsx" 

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
BASE_URL = "https://www.gov.taipei/News_Leader.aspx?n=1E25E56D8B12C862&sms=7CAF6BD4D3E48630"
MASTER_FILE = "city_leaders_complete.xlsx"
HISTORY_DIR = "history_records"

# 💡 註：在 GitHub Actions 中，SCAN_MODE 等設定將由 GitHub 自己的排程(Cron)決定
SCAN_MODE = "interval"
MINUTES_INTERVAL = 60
DAILY_TIME = "09:00"

if not os.path.exists(HISTORY_DIR):
    os.makedirs(HISTORY_DIR)

class TaipeiLeaderMonitor:
    def __init__(self):
        self.receiver_emails = []
        self.log("🚀 系統啟動：北市府首長監控站 (GitHub Actions 雲端版)")

    def log(self, msg):
        print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")

    def load_recipients(self):
        try:
            if os.path.exists(RECIPIENT_FILE):
                df = pd.read_excel(RECIPIENT_FILE); df.columns = [c.lower().strip() for c in df.columns]
                if 'email' in df.columns: 
                    self.receiver_emails = df['email'].dropna().unique().tolist()
                    self.log(f"✅ 名單匯入：共 {len(self.receiver_emails)} 位。")
                else: self.receiver_emails = ["andy939.yang@gmail.com", "bk1883@gov.taipei"]
            else: self.receiver_emails = ["andy939.yang@gmail.com", "bk1883@gov.taipei"]
        except: self.receiver_emails = ["andy939.yang@gmail.com", "bk1883@gov.taipei"]

    def send_email_notification(self, added, removed, old_time, now_time):
        # 💡 安全檢查：確保有讀到 Secrets 密碼
        if not SENDER_EMAIL or not SENDER_PASSWORD:
            self.log("⚠️ 錯誤：GitHub Secrets 未正確設定 EMAIL 或密碼")
            return
            
        try:
            subject = f"🚨 北市府首長異動提醒 - {datetime.now().strftime('%Y/%m/%d')}"
            body = f"偵測到異動。\n變動前：{old_time}\n變動後：{now_time}\n" + "="*45 + "\n\n"
            if added:
                sorted_added = sorted(list(added), key=lambda x: x[1])
                body += "【🆕 變動後長官】\n"
                for name, dept in sorted_added: body += f"  ＋ {name} ( {dept} )\n"
            if removed:
                sorted_removed = sorted(list(removed), key=lambda x: x[1])
                body += "\n【❌ 變動前長官】\n"
                for name, dept in sorted_removed: body += f"  － {name} ( {dept} )\n"
            msg = MIMEText(body, 'plain', 'utf-8'); msg['From'] = SENDER_EMAIL; msg['To'] = ", ".join(self.receiver_emails); msg['Subject'] = Header(subject, 'utf-8')
            server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT); server.starttls(); server.login(SENDER_EMAIL, SENDER_PASSWORD); server.sendmail(SENDER_EMAIL, self.receiver_emails, msg.as_string()); server.quit()
            self.log(f"📨 異動郵件已寄出。")
        except Exception as e: self.log(f"⚠️ 郵件錯誤: {str(e)}")

    def run_check_logic(self):
        try:
            all_new_data = []; seen_names = {}; headers = {"User-Agent": "Mozilla/5.0"}
            garbage_list = ["市民服務", "市政公告", "市政資料", "與民互動", "助您好孕", "組織架構", "市府APP", "如何到達", "市府團隊", "市府新聞", "酒駕防制", "市政會議"]
            keywords = ["局", "處", "會", "公所", "府", "中心", "所", "學院", "大隊", "團", "館", "園", "電臺", "公司", "醫院", "院"]
            self.log(f"--- 掃描開始 ---")
            for page in range(1, 9):
                page_count = 0
                resp = requests.get(f"{BASE_URL}&page={page}&PageSize=20", headers=headers, timeout=25, verify=False)
                resp.encoding = 'utf-8'; soup = BeautifulSoup(resp.text, 'html.parser')
                items = soup.select(".list_class li") or soup.find_all(["a", "li"], class_=lambda x: x != 'main-nav')
                for item in items:
                    raw_text = item.get_text(" ", strip=True).replace("收藏網頁", "").replace("My收藏", "").strip()
                    if any(g in raw_text for g in garbage_list): continue
                    
                    if "王玉芬" in raw_text and "秘書處" in raw_text:
                        if not any(d['首長姓名'] == "王玉芬" for d in all_new_data):
                            all_new_data.append({"機關": "秘書處", "職稱": "臺北市政府秘書長兼秘書處處長", "首長姓名": "王玉芬"})
                            page_count += 1
                        continue 
                    
                    if any(k in raw_text for k in keywords):
                        parts = [p.strip() for p in raw_text.split() if len(p.strip()) >= 2]
                        if len(parts) >= 2:
                            name = re.sub(r'[\(\（].*?[\)\）]', '', parts[0]).strip()
                            dept = ""; dept_idx = -1
                            for idx, p in enumerate(parts):
                                if any(k in p for k in keywords):
                                    if p not in ["院長", "處長", "局長", "中心主任", "組長", "團長", "校長", "主任", "主委"]:
                                        if len(p) > len(dept): dept = p; dept_idx = idx
                            if not dept: dept = parts[-1]; dept_idx = len(parts)-1
                            title = parts[dept_idx-1] if dept_idx > 1 else ""
                            if 2 <= len(name) <= 15 and len(dept) > len(name):
                                if name not in seen_names or len(dept) > len(seen_names[name]):
                                    seen_names[name] = dept
                                    if not any(d['首長姓名'] == name and d['機關'] == dept for d in all_new_data):
                                        all_new_data.append({"機關": dept, "職稱": title, "首長姓名": name})
                                        page_count += 1
                self.log(f"🌐 第 {page} 頁掃描完成... (抓取: {page_count} 筆)")
            
            new_df = pd.DataFrame(all_new_data).drop_duplicates()
            files = glob.glob(os.path.join(HISTORY_DIR, "city_leaders_*.xlsx"))
            
            if files:
                # 💡 這裡做了修正：用檔名排序找到最新的一個
                old_file = max(files) 
                old_df = pd.read_excel(old_file)
                old_set = set(zip(old_df['首長姓名'], old_df['機關']))
                new_set = set(zip(new_df['首長姓名'], new_df['機關']))
                added, removed = new_set - old_set, old_set - new_set
                
                if added or removed:
                    # 💡 重點修正：從檔名提取時間，避免 GitHub Actions 檔案時間不準的問題
                    file_name = os.path.basename(old_file)
                    time_match = re.search(r'(\d{8})_(\d{4})', file_name)
                    if time_match:
                        d, t = time_match.groups()
                        ot = f"{d[:4]}-{d[4:6]}-{d[6:8]} {t[:2]}:{t[2:]}:00"
                    else:
                        ot = datetime.fromtimestamp(os.path.getmtime(old_file)).strftime('%Y-%m-%d %H:%M:%S')
                    
                    nt = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    self.send_email_notification(added, removed, ot, nt)
                else: self.log("✅ 比對完成：姓名資料一致。")
            
            # 存檔邏輯
            stamp = datetime.now().strftime('%Y%m%d_%H%M')
            new_df.to_excel(os.path.join(HISTORY_DIR, f"city_leaders_{stamp}.xlsx"), index=False)
            new_df.to_excel(MASTER_FILE, index=False)
            self.log(f"📊 掃描完畢，總計取得 {len(new_df)} 筆資料。")
        except Exception as e: self.log(f"❌ 異常: {str(e)}")

if __name__ == "__main__":
    monitor = TaipeiLeaderMonitor()
    monitor.load_recipients()
    monitor.run_check_logic()
