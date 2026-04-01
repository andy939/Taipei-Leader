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

# 禁用安全警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 從 GitHub Secrets 讀取敏感資訊
SENDER_EMAIL = os.environ.get('SENDER_EMAIL')
SENDER_PASSWORD = os.environ.get('SENDER_PASSWORD')
# 收件者清單
RECEIVER_EMAILS = ["andy939.yang@gmail.com", "bk1883@gov.taipei"]

BASE_URL = "https://www.gov.taipei/News_Leader.aspx?n=1E25E56D8B12C862&sms=7CAF6BD4D3E48630"
MASTER_FILE = "city_leaders_complete.xlsx"

def send_email(added, removed, old_time, now_time):
    """ 發送多人郵件通知 """
    try:
        subject = f"🚨 北市府首長異動提醒 - {datetime.now().strftime('%Y/%m/%d')}"
        body = f"偵測到北市府首長名單有異動。\n"
        body += f"變動前參考基準點：{old_time}\n"
        body += f"變動後最新掃描點：{now_time}\n"
        body += "=" * 45 + "\n\n"
        
        if added:
            sorted_added = sorted(list(added), key=lambda x: x[1])
            body += f"【🆕 變動後人員 (資料時間: {now_time})】\n"
            for name, dept in sorted_added: body += f"  ＋ {name} ( {dept} )\n"
        
        if removed:
            sorted_removed = sorted(list(removed), key=lambda x: x[1])
            body += f"\n【❌ 變動前人員 (資料時間: {old_time})】\n"
            for name, dept in sorted_removed: body += f"  － {name} ( {dept} )\n"

        msg = MIMEText(body, 'plain', 'utf-8')
        msg['From'] = SENDER_EMAIL
        msg['To'] = ", ".join(RECEIVER_EMAILS)
        msg['Subject'] = Header(subject, 'utf-8')

        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(SENDER_EMAIL, RECEIVER_EMAILS, msg.as_string())
        print("📧 異動通知信已成功寄出！")
    except Exception as e:
        print(f"⚠️ 郵件發送失敗: {e}")

def run_task():
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    all_new_data = []
    seen_names = {}
    garbage_list = ["市民服務", "市政公告", "市政資料", "與民互動", "助您好孕", "組織架構", "市府APP", "如何到達", "市府團隊", "市府新聞"]

    print(f"🚀 開始掃描北市府首長名單...")
    now_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    for page in range(1, 9):
        page_url = f"{BASE_URL}&page={page}&PageSize=20"
        resp = requests.get(page_url, headers=headers, timeout=25, verify=False)
        resp.encoding = 'utf-8'
        soup = BeautifulSoup(resp.text, 'html.parser')
        
        items = soup.select(".list_class li") or soup.find_all(["a", "li"], class_=lambda x: x != 'main-nav')

        for item in items:
            raw_text = item.get_text(" ", strip=True)
            clean_text = raw_text.replace("收藏網頁", "").replace("My收藏", "").strip()
            if any(g in clean_text for g in garbage_list): continue

            keywords = ["局", "處", "會", "公所", "府", "中心", "所", "學院", "大隊", "團", "館", "園", "電臺", "公司", "醫院", "院"]
            if any(k in clean_text for k in keywords):
                parts = [p.strip() for p in clean_text.split() if len(p.strip()) >= 2]
                if len(parts) >= 2:
                    name = re.sub(r'[\(\（].*?[\)\）]', '', parts[0]).strip()
                    dept = ""
                    for p in parts:
                        if any(k in p for k in keywords):
                            if p not in ["院長", "處長", "局長", "中心主任", "組長", "團長", "校長", "主任", "主委"]:
                                if len(p) > len(dept): dept = p
                    if not dept: dept = parts[-1]

                    if 2 <= len(name) <= 15 and len(dept) > len(name):
                        if name == "王玉芬" and "秘書處" not in dept: continue
                        if name not in seen_names or len(dept) > len(seen_names[name]):
                            seen_names[name] = dept
                            all_new_data.append({"機關": dept, "首長姓名": name})
    
    new_df = pd.DataFrame(all_new_data).drop_duplicates()
    
    if os.path.exists(MASTER_FILE):
        old_df = pd.read_excel(MASTER_FILE)
        old_time = datetime.fromtimestamp(os.path.getmtime(MASTER_FILE)).strftime('%Y-%m-%d %H:%M:%S')
        
        old_set = set(zip(old_df['首長姓名'], old_df['機關']))
        new_set = set(zip(new_df['首長姓名'], new_df['機關']))
        
        added = new_set - old_set
        removed = old_set - new_set
        
        if added or removed:
            send_email(added, removed, old_time, now_time)
        else:
            print(f"✅ 資料一致 (基準點: {old_time})。")
    else:
        print("ℹ️ 首次執行，建立初始 Excel。")
    
    new_df.to_excel(MASTER_FILE, index=False)
    print(f"📊 任務完成。")

if __name__ == "__main__":
    run_task()
