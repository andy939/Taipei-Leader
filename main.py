# -*- coding: utf-8 -*-
"""
臺北市政府機關首長異動監控 (GitHub Actions 版)
- 抓取首長資料（div.figure 精準解析，PageSize=200 一次取完）
- 與 repo 內最新的 city_leaders_YYYYMMDD_HHMMSS.xlsx 比對
- 有異動才產生新的帶日期時間 Excel 並寄 HTML 通知信
- 無異動：靜默結束，不寫檔、不 commit
"""

import os
import glob
import smtplib
import warnings
import requests
import pandas as pd
from bs4 import BeautifulSoup
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from urllib3.exceptions import InsecureRequestWarning

warnings.filterwarnings("ignore", category=InsecureRequestWarning)

# ── 設定區 ────────────────────────────────────────────────────────────────────
URL_BASE   = "https://www.gov.taipei/News_Leader.aspx"
URL_PARAMS = {
    "n":        "1E25E56D8B12C862",
    "sms":      "7CAF6BD4D3E48630",
    "PageSize": "200",
}

SMTP_SERVER     = os.environ.get("SMTP_SERVER",     "smtp.gmail.com")
SMTP_PORT       = int(os.environ.get("SMTP_PORT",   "587"))
SENDER_EMAIL    = os.environ.get("SENDER_EMAIL",    "")
SENDER_PASSWORD = os.environ.get("SENDER_PASSWORD", "")

FILE_PREFIX    = "city_leaders_"   # 檔名前綴，後接 YYYYMMDD_HHMMSS.xlsx
RECIPIENT_FILE = "收件者清單.xlsx"
# ─────────────────────────────────────────────────────────────────────────────


def log(msg: str):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


def latest_file() -> "str | None":
    """找 repo 內最新的 city_leaders_*.xlsx，沒有則回傳 None"""
    files = sorted(glob.glob(f"{FILE_PREFIX}????????_??????.xlsx"))
    return files[-1] if files else None


def new_filename() -> str:
    """產生帶當前日期時間的檔名，例如 city_leaders_20260518_093012.xlsx"""
    return f"{FILE_PREFIX}{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"


# ── 1. 抓取首長資料 ───────────────────────────────────────────────────────────
def fetch_leaders() -> pd.DataFrame:
    log("開始抓取首長資料...")
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        )
    }
    resp = requests.get(URL_BASE, params=URL_PARAMS, headers=headers,
                        timeout=30, verify=False)
    resp.encoding = "utf-8"
    soup = BeautifulSoup(resp.text, "html.parser")

    records = []
    for figure in soup.select("div.figure"):
        essay = figure.find("div", class_="essay")
        if not essay:
            continue
        link  = essay.find("a")
        spans = essay.find_all("span")
        name  = link.get_text(strip=True) if link else ""
        title = spans[0].get_text(strip=True) if len(spans) > 0 else ""
        org   = spans[1].get_text(strip=True) if len(spans) > 1 else ""
        if name:
            records.append({"機關": org, "職稱": title, "姓名": name})

    if not records:
        raise RuntimeError("無法解析首長資料，請確認網頁結構是否異動")

    df = pd.DataFrame(records, columns=["機關", "職稱", "姓名"])
    log(f"抓取完成，共 {len(df)} 筆")
    return df


# ── 2. 讀取收件者清單 ─────────────────────────────────────────────────────────
def load_recipients() -> list:
    fallback = ["bk1883@gov.taipei"]
    try:
        if not os.path.exists(RECIPIENT_FILE):
            log(f"收件者清單不存在，使用預設：{fallback}")
            return fallback
        df = pd.read_excel(RECIPIENT_FILE)
        df.columns = [c.lower().strip() for c in df.columns]
        if "email" not in df.columns:
            log("收件者清單無 email 欄位，使用預設")
            return fallback
        emails = df["email"].dropna().str.strip().unique().tolist()
        log(f"收件者清單載入：共 {len(emails)} 位")
        return emails
    except Exception as e:
        log(f"讀取收件者清單失敗：{e}，使用預設")
        return fallback


# ── 3. 比對新舊資料 ───────────────────────────────────────────────────────────
def compare(old_df: pd.DataFrame, new_df: pd.DataFrame):
    old_t = set(old_df.apply(tuple, axis=1))
    new_t = set(new_df.apply(tuple, axis=1))
    added   = new_df[~new_df.apply(tuple, axis=1).isin(old_t)].copy()
    removed = old_df[~old_df.apply(tuple, axis=1).isin(new_t)].copy()

    old_d = {r["機關"]: r for _, r in old_df.iterrows()}
    new_d = {r["機關"]: r for _, r in new_df.iterrows()}
    chg = []
    for org in set(old_d) & set(new_d):
        o, n = old_d[org], new_d[org]
        if o["職稱"] != n["職稱"] or o["姓名"] != n["姓名"]:
            chg.append({
                "機關":   org,
                "舊職稱": o["職稱"], "新職稱": n["職稱"],
                "舊姓名": o["姓名"], "新姓名": n["姓名"],
            })
    return added, removed, pd.DataFrame(chg)


# ── 4. 存 Excel ───────────────────────────────────────────────────────────────
def save_excel(df: pd.DataFrame, path: str):
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="首長名單")
        ws = writer.sheets["首長名單"]
        for col in ws.columns:
            w = max(len(str(c.value)) if c.value else 0 for c in col)
            ws.column_dimensions[col[0].column_letter].width = w + 4
    log(f"Excel 已儲存：{path}")


# ── 5. 組裝 HTML 郵件內容 ─────────────────────────────────────────────────────
def build_html(now_str, added, removed, changed, excel_name):
    S_TABLE = "border-collapse:collapse;width:100%;font-size:13px;margin-top:8px;"
    S_TH    = "background:#2F5496;color:#fff;padding:7px 12px;text-align:center;border:1px solid #aaa;"
    S_TD    = "padding:6px 12px;border:1px solid #ccc;text-align:center;"
    S_OLD   = "padding:6px 12px;border:1px solid #ccc;text-align:center;color:#900;font-weight:bold;"
    S_NEW   = "padding:6px 12px;border:1px solid #ccc;text-align:center;color:#060;font-weight:bold;"
    S_RADD  = "background:#e6f4ea;"
    S_RDEL  = "background:#fce8e6;"

    def make_table(df, row_style, old_cols, new_cols):
        cols = list(df.columns)
        hdr  = "".join('<th style="' + S_TH + '">' + c + "</th>" for c in cols)
        body = ""
        for _, row in df.iterrows():
            cells = ""
            for c in cols:
                v = str(row[c]) if (row[c] and str(row[c]) != "nan") else "—"
                s = S_OLD if c in old_cols else S_NEW if c in new_cols else S_TD
                cells += '<td style="' + s + '">' + v + "</td>"
            body += '<tr style="' + row_style + '">' + cells + "</tr>"
        return ('<table style="' + S_TABLE + '"><thead><tr>' + hdr +
                "</tr></thead><tbody>" + body + "</tbody></table>")

    def make_section(title, color, bg, df, row_style, old_cols=None, new_cols=None):
        old_cols = old_cols or []
        new_cols = new_cols or []
        bar = ('<div style="background:' + bg + ';border-left:5px solid ' + color +
               ';padding:8px 14px;font-weight:bold;font-size:15px;color:' + color + ';">' +
               title + ("　共 " + str(len(df)) + " 筆" if not df.empty else "") + "</div>")
        if df.empty:
            return ('<div style="margin-top:20px;">' + bar +
                    '<p style="color:#999;padding:4px 14px;">（無）</p></div>')
        return ('<div style="margin-top:20px;">' + bar +
                make_table(df, row_style, old_cols, new_cols) + "</div>")

    def make_bullets(df, kind, color):
        if df.empty:
            return ""
        label = {"add": "新增", "del": "移除", "chg": "異動"}[kind]
        items = ""
        for _, r in df.iterrows():
            if kind in ("add", "del"):
                txt = r["機關"] + "　" + r["職稱"] + "　" + r["姓名"]
            else:
                txt = (r["機關"] + "　職稱：" + r["舊職稱"] + " → " + r["新職稱"] +
                       "　姓名：" + r["舊姓名"] + " → " + r["新姓名"])
            items += ('<li style="margin:5px 0;">'
                      '<span style="color:' + color + ';font-weight:bold;">【' + label + '】</span>'
                      + txt + "</li>")
        return items

    total   = len(added) + len(removed) + len(changed)
    bullets = (make_bullets(added,   "add", "#1a7f37") +
               make_bullets(removed, "del", "#c0392b") +
               make_bullets(changed, "chg", "#E67E00"))
    summary = '<ul style="line-height:1.9;margin:8px 0;">' + bullets + "</ul>"

    sec_add = make_section("▲ 新增首長",      "#1a7f37", "#e6f4ea", added,   S_RADD)
    sec_del = make_section("▼ 移除首長",      "#c0392b", "#fce8e6", removed, S_RDEL)
    sec_chg = make_section("◆ 職稱／姓名異動", "#E67E00", "#fff8e6", changed, "",
                           old_cols=["舊職稱", "舊姓名"], new_cols=["新職稱", "新姓名"])

    return (
        "<html><body style=\"font-family:Arial,'Microsoft JhengHei',sans-serif;"
        "font-size:14px;color:#333;max-width:960px;margin:auto;\">"
        "<div style=\"background:#2F5496;color:#fff;padding:14px 22px;\">"
        "<h2 style=\"margin:0;font-size:18px;\">臺北市政府機關首長異動通知</h2></div>"
        "<div style=\"border:1px solid #ccc;padding:22px;\">"
        "<p>您好，</p>"
        "<p>系統於 <b>" + now_str + "</b> 偵測到首長資料異動，共 <b>" + str(total) + "</b> 項：</p>"
        "<div style=\"background:#f8f8f8;border:1px solid #ddd;border-radius:4px;"
        "padding:14px 18px;margin-bottom:16px;\">"
        "<b style=\"font-size:15px;\">異動摘要</b>" + summary + "</div>"
        + sec_add + sec_del + sec_chg +
        "<p style=\"margin-top:24px;color:#555;\">附件：<b>" + excel_name +
        "</b>（最新完整首長名單）</p>"
        "<p><a href=\"https://www.gov.taipei/News_Leader.aspx"
        "?n=1E25E56D8B12C862&amp;sms=7CAF6BD4D3E48630\""
        " style=\"color:#2F5496;\">臺北市政府機關首長頁面</a></p>"
        "<p style=\"color:#aaa;font-size:12px;margin-top:20px;\">"
        "── 臺北市政府首長異動自動通知系統</p>"
        "</div></body></html>"
    )


# ── 6. 寄信 ──────────────────────────────────────────────────────────────────
def send_mail(recipients, now_str, added, removed, changed, excel_path):
    if not SENDER_EMAIL or not SENDER_PASSWORD:
        log("⚠ 未設定 SENDER_EMAIL / SENDER_PASSWORD，跳過寄信")
        return

    subject   = (f"【首長異動通知】{now_str}"
                 f"（新增{len(added)}/移除{len(removed)}/異動{len(changed)}筆）")
    html_body = build_html(now_str, added, removed, changed,
                           os.path.basename(excel_path))

    msg_outer = MIMEMultipart("mixed")
    msg_outer["Subject"] = subject
    msg_outer["From"]    = SENDER_EMAIL
    msg_outer["To"]      = ", ".join(recipients)

    msg_alt = MIMEMultipart("alternative")
    msg_alt.attach(MIMEText(html_body, "html", "utf-8"))
    msg_outer.attach(msg_alt)

    fname = os.path.basename(excel_path)
    with open(excel_path, "rb") as f:
        part = MIMEBase("application",
                        "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", "attachment",
                    filename=("utf-8", "", fname))
    msg_outer.attach(part)

    log(f"寄信中... SMTP={SMTP_SERVER}:{SMTP_PORT}  收件者={recipients}")
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=15) as server:
        server.ehlo()
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.send_message(msg_outer)
    log(f"郵件已寄出至 {len(recipients)} 位收件者")


# ── 主程式 ────────────────────────────────────────────────────────────────────
def main():
    log("=" * 55)
    log("臺北市政府首長資料監控啟動")
    log("=" * 55)

    new_df     = fetch_leaders()
    recipients = load_recipients()
    now_str    = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    prev_file  = latest_file()

    # ── 第一次執行（repo 內尚無 Excel）─────────────────────────────────────
    if prev_file is None:
        out = new_filename()
        log(f"首次執行，建立初始首長名單：{out}")
        save_excel(new_df, out)
        log("初始名單已建立，本次不寄信")
        return

    # ── 後續執行：比對 ──────────────────────────────────────────────────────
    log(f"比對基準：{prev_file}")
    old_df     = pd.read_excel(prev_file, sheet_name="首長名單", dtype=str).fillna("")
    new_df_str = new_df.astype(str).fillna("")

    added, removed, changed = compare(old_df, new_df_str)

    if added.empty and removed.empty and changed.empty:
        log("比對完成：資料無異動")
        return

    # ── 有異動：產生新 Excel（帶日期時間）+ 寄信 ───────────────────────────
    log(f"偵測到異動 → 新增 {len(added)} / 移除 {len(removed)} / 異動 {len(changed)} 筆")
    out = new_filename()
    save_excel(new_df, out)

    try:
        send_mail(recipients, now_str, added, removed, changed, out)
    except Exception as e:
        log(f"⚠ 寄信失敗：{e}")

    log(f"完成！新檔案：{out}")


if __name__ == "__main__":
    main()
