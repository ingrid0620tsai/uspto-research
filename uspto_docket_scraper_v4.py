"""
檔案名稱：uspto_docket_scraper_v4.py
撰寫目的：抓取 regulations.gov 中指定單一 PTO docket 的所有資料，
         包含完整的 docket 元資料、documents（含全文）、comments（含附件連結）。
資料來源：regulations.gov v4 REST API
輸出格式：Excel（4 個工作表：Summary / Docket / Documents / Comments）
         + 3 個 CSV 備份檔
撰寫日期：2025-03-20
修改紀錄：
  - v4：修正 PDF/HTML 下載需帶瀏覽器 headers（之前 bytes=0 問題）
        修正 comments 附件欄位全空（list API 不含 included，需 detail API）
依賴套件：pip install requests pandas openpyxl tqdm beautifulsoup4 pdfplumber
"""

import io
import os
import json
import time
import requests
import pandas as pd
from tqdm import tqdm
from pathlib import Path
from datetime import datetime
from bs4 import BeautifulSoup

# ════════════════════════════════════════════════════════
# ★ 設定區 — 只需修改這裡，不需動下方程式碼
# ════════════════════════════════════════════════════════

API_KEY = "og1wltTsDKFQB8tTUdMMFS4akawWzkULBrt18hDg"
# 申請網址：https://api.data.gov/signup/

# ──────────────────────────────────────────────────────
# 【抓取目標】
# 模式 A：單一 docket（預設）
#   → 只填 DOCKET_ID，其餘篩選條件留空
#   → 範例：DOCKET_ID = "PTO-P-2023-0048"
#
# 模式 B：批次抓多個 docket
#   → 把 DOCKET_ID 改成 list，例如：
#   → DOCKET_ID = ["PTO-P-2023-0048", "PTO-P-2024-0003"]
#   → 程式會依序抓取每個 docket
#
# 模式 C：依條件篩選（見下方篩選設定）
#   → 把 DOCKET_ID 設為 None，再填寫篩選條件
#   → DOCKET_ID = None
# ──────────────────────────────────────────────────────
DOCKET_ID = None

# ──────────────────────────────────────────────────────
# 【篩選條件】僅在 DOCKET_ID = None 時生效
#
# FILTER_YEAR：依 Docket ID 中的年份篩選
#   → 填字串年份，例如 "2024"
#   → 留空 "" 代表不篩選年份
#   → 原理：PTO-P-2024-0003 的第三段就是年份，
#           比用 modifyDate 更準確（避免舊案修改後被誤納）
#
# FILTER_DOCKET_TYPE：依 docket 類型篩選
#   → "Rulemaking"    = 只抓正式規則制定案件（有 comment letters）
#   → "Nonrulemaking" = 只抓非規則制定案件
#   → ""              = 不篩選，抓全部
#
# 範例一：抓 2024 年所有 Rulemaking dockets
#   DOCKET_ID        = None
#   FILTER_YEAR      = "2024"
#   FILTER_DOCKET_TYPE = "Rulemaking"
#
# 範例二：抓所有 PTO dockets（不限年份與類型）
#   DOCKET_ID        = None
#   FILTER_YEAR      = ""
#   FILTER_DOCKET_TYPE = ""
# ──────────────────────────────────────────────────────
FILTER_YEAR        = "2024"
FILTER_DOCKET_TYPE = "Rulemaking"

# ──────────────────────────────────────────────────────
# 【抓取模式】
# ──────────────────────────────────────────────────────
FETCH_COMMENT_DETAIL = True
# True  = 每筆 comment 打 detail API，取得完整內文 + 附件連結（推薦）
# False = 快速測試模式，comment 附件欄位會是空的

FETCH_DOCUMENT_FULLTEXT = True
# True  = 下載 document 的 PDF/HTML 全文並解析成純文字
# False = 只保留 metadata，不下載全文

FETCH_COMMENT_ATTACHMENT_TEXT = True
# True  = 下載 comment 附件（PDF/HTML）並解析成純文字，存入 attachment_fulltext 欄位
#         注意：每筆 comment 可能有多個附件，會全部下載，速度較慢
# False = 只保留附件連結欄位，不下載全文（預設）

SAVE_RAW_JSON = True
# True  = 備份每筆原始 JSON 到 {docket_id}_raw_json/ 資料夾
#         建議保留，確保原始資料可回溯，未來欄位有變動也不怕

OUTPUT_DIR = Path(os.path.dirname(os.path.abspath(__file__)))
# 輸出到腳本所在資料夾
# 如需指定其他路徑，改為：OUTPUT_DIR = Path("/Users/你的名字/Desktop/output")

# ════════════════════════════════════════════════════════
# 以下為程式本體，一般情況不需修改
# ════════════════════════════════════════════════════════

BASE_URL = "https://api.regulations.gov/v4"
DELAY    = 7.5
# 說明：免費 Key 每小時 500 次，每次間隔 7.5 秒保留緩衝

# ★ 修正一：下載 PDF/HTML 必須帶瀏覽器 headers
# 原因：downloads.regulations.gov 不帶 headers 時回傳 HTTP 200 但 content 為空
# 診斷確認：帶 headers 後正常下載（HTTP 200，349,958 bytes，開頭 %PDF-1.7）
DOWNLOAD_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
                  "Chrome/120.0.0.0 Safari/537.36",
    "Referer":    "https://www.regulations.gov/",
    "Accept":     "text/html,application/pdf,*/*",
}


# ══════════════════════════════════════════════════════
# SECTION 1：API 層
# 說明：所有 HTTP 請求集中在此。
#       api_get()         = regulations.gov API 請求
#       download_file()   = 附件下載（帶瀏覽器 headers）
#       fetch_all_pages() = documents/comments 自動翻頁
# ══════════════════════════════════════════════════════

def api_get(endpoint, params={}, retries=5):
    """
    發送單次 regulations.gov API GET 請求。

    Args:
        endpoint (str) : API 路徑，不含 BASE_URL
        params   (dict): 查詢參數，api_key 自動附加
        retries  (int) : 失敗時最多重試幾次

    Returns:
        dict: API 回傳的完整 JSON，失敗時回傳空字典 {}
    """
    p   = {**params, "api_key": API_KEY}
    url = f"{BASE_URL}/{endpoint}"

    for attempt in range(retries):
        try:
            r         = requests.get(url, params=p, timeout=30)
            remaining = int(r.headers.get("X-RateLimit-Remaining", 999))

            if remaining < 10:
                print(f"\n  [Rate Limit 警告] 剩餘 {remaining} 次，暫停 65 秒...")
                time.sleep(65)

            if r.status_code == 200:
                time.sleep(DELAY)
                return r.json()
            elif r.status_code == 429:
                wait = 70 * (attempt + 1)
                print(f"\n  [429] 等待 {wait}s...")
                time.sleep(wait)
            elif r.status_code == 404:
                return {}
            else:
                print(f"\n  [HTTP {r.status_code}] {url}")
                print(f"  [訊息] {r.text[:200]}")
                time.sleep(10)

        except requests.RequestException as e:
            print(f"\n  [連線錯誤] {e}，等待 30s...")
            time.sleep(30)

    return {}


def download_file(url, retries=3):
    """
    下載 regulations.gov 附件（PDF 或 HTML）。

    ★ 修正說明：必須帶瀏覽器 headers，否則回傳 HTTP 200 但 content 為空。

    Args:
        url     (str): 附件下載網址
        retries (int): 失敗時最多重試幾次

    Returns:
        bytes: 檔案原始 bytes，失敗時回傳 None
    """
    if not url:
        return None

    for attempt in range(retries):
        try:
            r = requests.get(url, headers=DOWNLOAD_HEADERS, timeout=60)
            if r.status_code == 200 and len(r.content) > 0:
                return r.content
            elif r.status_code == 404:
                return None
            else:
                print(f"\n  [下載失敗] HTTP {r.status_code} {url}")
                time.sleep(5)
        except requests.RequestException as e:
            print(f"\n  [下載連線錯誤] {e}")
            time.sleep(10)

    return None


def fetch_all_pages(endpoint, params, label=""):
    """
    自動翻頁抓取 documents 或 comments。
    支援超過 5,000 筆的 lastModifiedDate cursor 分段策略。

    Args:
        endpoint (str) : API 路徑
        params   (dict): 查詢參數
        label    (str) : 進度顯示標籤

    Returns:
        list: 所有頁面合併後的完整資料清單
    """
    all_items = []
    cursor    = None
    batch     = 0

    while True:
        batch += 1
        p = {
            **params,
            "page[size]": 250,
            "sort":       "lastModifiedDate,documentId",
        }
        if cursor:
            p["filter[lastModifiedDate][ge]"] = cursor

        segment = []
        for page in range(1, 21):
            p["page[number]"] = page
            resp  = api_get(endpoint, p)
            data  = resp.get("data", [])
            total = resp.get("meta", {}).get("totalElements", "?")

            if not data:
                break

            segment.extend(data)
            print(
                f"\r  {label} batch {batch} p{page} | "
                f"{len(all_items)+len(segment)}/{total}",
                end="", flush=True
            )

            if len(data) < 250:
                break

        all_items.extend(segment)

        if len(segment) == 5000:
            last_ts = segment[-1].get("attributes", {}).get("lastModifiedDate", "")
            cursor  = last_ts.replace("T", " ").replace("Z", "")[:19] if last_ts else None
            if not cursor:
                break
        else:
            break

    print()
    return all_items


# ══════════════════════════════════════════════════════
# SECTION 2：全文解析層
# 說明：將下載的 bytes 解析成純文字。
#       優先用 HTML（乾淨），備援用 PDF（需 pdfplumber）。
# ══════════════════════════════════════════════════════

def parse_html_to_text(content_bytes):
    """
    將 HTML bytes 解析成純文字。

    Args:
        content_bytes (bytes): HTML 檔案的原始 bytes

    Returns:
        str: 清理後的純文字，失敗時回傳空字串
    """
    try:
        soup = BeautifulSoup(content_bytes, "html.parser")
        for tag in soup(["script", "style", "nav", "header", "footer"]):
            tag.decompose()
        lines = [l.strip() for l in soup.get_text(separator="\n").splitlines() if l.strip()]
        return "\n".join(lines)
    except Exception as e:
        print(f"\n  [HTML 解析失敗] {e}")
        return ""


def parse_pdf_to_text(content_bytes):
    """
    將 PDF bytes 解析成純文字。

    Args:
        content_bytes (bytes): PDF 檔案的原始 bytes

    Returns:
        str: 所有頁面合併的純文字，失敗時回傳空字串
    """
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(content_bytes)) as pdf:
            pages = [p.extract_text() for p in pdf.pages if p.extract_text()]
        return "\n".join(pages)
    except ImportError:
        print("\n  [提示] 未安裝 pdfplumber，執行：pip install pdfplumber")
        return ""
    except Exception as e:
        print(f"\n  [PDF 解析失敗] {e}")
        return ""


def fetch_fulltext(html_url, pdf_url):
    """
    抓取並解析 document 全文，優先 HTML，備援 PDF。

    Args:
        html_url (str): HTML 附件網址
        pdf_url  (str): PDF 附件網址

    Returns:
        tuple: (full_text, source)，source 為 "html"、"pdf" 或 ""
    """
    if html_url:
        content = download_file(html_url)
        if content:
            text = parse_html_to_text(content)
            if text:
                return text, "html"

    if pdf_url:
        content = download_file(pdf_url)
        if content:
            text = parse_pdf_to_text(content)
            if text:
                return text, "pdf"

    return "", ""


# ══════════════════════════════════════════════════════
# SECTION 3：資料解析層
# 說明：將 API 原始 JSON 轉成結構化字典。
#       所有欄位完整保留，禁止刪除未使用的欄位。
# ══════════════════════════════════════════════════════

def extract_attachment_urls(raw):
    """
    從 raw JSON 提取所有附件 URL，依副檔名分類。

    ★ 注意：list API 不含 included，附件只有 detail API 才有。

    Args:
        raw (dict): document 或 comment 的完整 JSON

    Returns:
        dict: { "pdf": [...], "html": [...], "other": [...] }
    """
    all_urls = []

    for item in raw.get("included", []):
        if item.get("type") == "attachments":
            for fmt in item.get("attributes", {}).get("fileFormats") or []:
                u = fmt.get("fileUrl", "") if isinstance(fmt, dict) else fmt
                if u:
                    all_urls.append(u)

    for fmt in raw.get("attributes", {}).get("fileFormats") or []:
        u = fmt.get("fileUrl", "") if isinstance(fmt, dict) else fmt
        if u and u not in all_urls:
            all_urls.append(u)

    pdf_urls   = [u for u in all_urls if u.lower().endswith(".pdf")]
    html_urls  = [u for u in all_urls if u.lower().endswith((".html", ".htm"))]
    other_urls = [u for u in all_urls if u not in pdf_urls and u not in html_urls]

    return {"pdf": pdf_urls, "html": html_urls, "other": other_urls}


def parse_docket(raw):
    """
    解析單一 docket 的完整元資料。

    Args:
        raw (dict): docket detail API 回傳的原始 JSON

    Returns:
        dict: 結構化的 docket 資料
    """
    a = raw.get("attributes", {})
    return {
        "docket_id":       raw.get("id", ""),
        "agency_id":       a.get("agencyId", ""),
        "rin":             a.get("rin", ""),
        "docket_type":     a.get("docketType", ""),
        "category":        a.get("category", ""),
        "program":         a.get("program", ""),
        "sub_type":        a.get("subType", ""),
        "title":           a.get("title", ""),
        "description":     a.get("description", ""),
        "keywords":        ", ".join(a.get("keywords") or []),
        "effective_date":  a.get("effectiveDate", ""),
        "modify_date":     a.get("modifyDate", ""),
        "regulations_url": f"https://www.regulations.gov/docket/{raw.get('id','')}",
        "scraped_at":      datetime.utcnow().isoformat() + "Z",
    }


def parse_document(raw, fetch_fulltext_flag=False):
    """
    解析單一 document 的完整資料，含附件連結與全文。

    Args:
        raw                (dict): document detail API（?include=attachments）的原始 JSON
        fetch_fulltext_flag(bool): 是否下載並解析全文

    Returns:
        dict: 結構化的 document 資料
    """
    a     = raw.get("attributes", {})
    links = raw.get("links", {})
    urls  = extract_attachment_urls(raw)

    pdf_url  = urls["pdf"][0]  if urls["pdf"]  else ""
    html_url = urls["html"][0] if urls["html"] else ""

    full_text, fulltext_source = "", ""
    if fetch_fulltext_flag:
        full_text, fulltext_source = fetch_fulltext(html_url, pdf_url)

    return {
        "document_id":           raw.get("id", ""),
        "object_id":             a.get("objectId", ""),
        "docket_id":             a.get("docketId", ""),
        "agency_id":             a.get("agencyId", ""),
        "document_type":         a.get("documentType", ""),
        "subtype":               a.get("subtype", ""),
        "exhibit_type":          a.get("exhibitType", ""),
        "title":                 a.get("title", ""),
        "summary":               a.get("summary", ""),
        "fr_doc_num":            a.get("frDocNum", ""),
        "cfr_parts":             ", ".join(a.get("cfrPart") or []),
        "full_text_xml_url":     a.get("fullTextXmlUrl", ""),
        "attachment_url_pdf":    "; ".join(urls["pdf"]),
        "attachment_url_html":   "; ".join(urls["html"]),
        "attachment_url_other":  "; ".join(urls["other"]),
        "attachment_count":      len(urls["pdf"]) + len(urls["html"]) + len(urls["other"]),
        "document_full_text":    full_text,
        "fulltext_source":       fulltext_source,
        "posted_date":           a.get("postedDate", ""),
        "last_modified_date":    a.get("lastModifiedDate", ""),
        "comment_start_date":    a.get("commentStartDate", ""),
        "comment_end_date":      a.get("commentEndDate", ""),
        "open_for_comment":      a.get("openForComment", False),
        "num_comments_received": a.get("numberOfCommentsReceived", 0),
        "withdrawn":             a.get("withdrawn", False),
        "regulations_url":       f"https://www.regulations.gov/document/{raw.get('id','')}",
        "api_url":               links.get("self", ""),
        "scraped_at":            datetime.utcnow().isoformat() + "Z",
    }


def fetch_all_attachment_texts(urls):
    """
    下載並解析所有 PDF/HTML 附件為純文字。
    多個附件間以分隔線區隔。

    Args:
        urls (dict): extract_attachment_urls() 回傳的 { "pdf": [...], "html": [...], ... }

    Returns:
        str: 所有附件合併後的純文字，無附件時回傳空字串
    """
    texts = []
    # HTML 優先，PDF 備援；other 類型不解析
    for url in urls["html"]:
        content = download_file(url)
        if content:
            text = parse_html_to_text(content)
            if text:
                texts.append(text)
    for url in urls["pdf"]:
        content = download_file(url)
        if content:
            text = parse_pdf_to_text(content)
            if text:
                texts.append(text)
    return "\n\n---\n\n".join(texts)


def parse_comment(raw, docket_id="", document_id="", fetch_att_text=False):
    """
    解析單一 comment 的完整資料，含附件連結與（可選）附件純文字。

    ★ 修正說明：附件欄位需 FETCH_COMMENT_DETAIL=True 才有值。
      list API 不含 included，附件需打 detail API 才能取得。

    Args:
        raw           (dict): comment detail API（?include=attachments）的原始 JSON
        docket_id     (str) : 補填用
        document_id   (str) : 補填用
        fetch_att_text(bool): 是否下載附件並解析成純文字

    Returns:
        dict: 結構化的 comment 資料
    """
    a    = raw.get("attributes", {})
    urls = extract_attachment_urls(raw)

    att_titles = [
        item.get("attributes", {}).get("title", "")
        for item in raw.get("included", [])
        if item.get("type") == "attachments"
    ]

    first = (a.get("firstName") or "").strip()
    last  = (a.get("lastName")  or "").strip()

    attachment_fulltext = ""
    if fetch_att_text and (urls["pdf"] or urls["html"]):
        attachment_fulltext = fetch_all_attachment_texts(urls)

    return {
        "comment_id":           raw.get("id", ""),
        "tracking_nbr":         a.get("trackingNbr", ""),
        "legacy_id":            a.get("legacyId", ""),
        "docket_id":            a.get("docketId", "")   or docket_id,
        "document_id":          a.get("documentId", "") or document_id,
        "comment_on_id":        a.get("commentOnId", ""),
        "document_type":        a.get("documentType", ""),
        "subtype":              a.get("subtype", ""),
        "commenter_name":       f"{first} {last}".strip(),
        "organization":         a.get("organization", ""),
        "gov_agency":           a.get("govAgency", ""),
        "gov_agency_type":      a.get("govAgencyType", ""),
        "city":                 a.get("city", ""),
        "state":                a.get("stateProvinceRegion", ""),
        "country":              a.get("country", ""),
        "zip":                  a.get("zip", ""),
        "title":                a.get("title", ""),
        "comment_text":         a.get("comment", ""),
        "doc_abstract":         a.get("docAbstract", ""),
        "page_count":           a.get("pageCount", ""),
        "attachment_count":     len(urls["pdf"]) + len(urls["html"]) + len(urls["other"]),
        "attachment_titles":    "; ".join(filter(None, att_titles)),
        "attachment_url_pdf":   "; ".join(urls["pdf"]),
        "attachment_url_html":  "; ".join(urls["html"]),
        "attachment_url_other": "; ".join(urls["other"]),
        "attachment_fulltext":  attachment_fulltext,
        "posted_date":          a.get("postedDate", ""),
        "receive_date":         a.get("receiveDate", ""),
        "postmark_date":        a.get("postmarkDate", ""),
        "last_modified_date":   a.get("lastModifiedDate", ""),
        "withdrawn":            a.get("withdrawn", False),
        "restrict_reason_type": a.get("restrictReasonType", ""),
        "reason_withdrawn":     a.get("reasonWithdrawn", ""),
        "regulations_url":      f"https://www.regulations.gov/comment/{raw.get('id','')}",
        "scraped_at":           datetime.utcnow().isoformat() + "Z",
    }


# ══════════════════════════════════════════════════════
# SECTION 4：斷點續跑
# ══════════════════════════════════════════════════════

def load_progress(f):
    """載入已完成的 ID 清單，首次執行回傳空集合。"""
    return set(Path(f).read_text(encoding="utf-8").splitlines()) if Path(f).exists() else set()

def save_progress(f, item_id):
    """追加一筆已完成的 ID 到進度檔。"""
    with open(f, "a", encoding="utf-8") as fp:
        fp.write(item_id + "\n")


# ══════════════════════════════════════════════════════
# SECTION 5：主程式
# ══════════════════════════════════════════════════════

def resolve_docket_ids():
    """
    依設定區的條件決定要抓取的 docket ID 清單。

    回傳邏輯：
      - DOCKET_ID 是字串 → 單一 docket，直接回傳 [DOCKET_ID]
      - DOCKET_ID 是 list → 批次模式，直接回傳該 list
      - DOCKET_ID 是 None → 篩選模式，抓全部 PTO dockets 後依條件過濾

    Returns:
        list: 要抓取的 docket ID 清單
    """
    # 模式 A / B：直接指定
    if DOCKET_ID is not None:
        if isinstance(DOCKET_ID, str):
            return [DOCKET_ID]
        if isinstance(DOCKET_ID, list):
            return DOCKET_ID

    # 模式 C：依條件篩選
    print("  篩選模式：抓取所有 PTO dockets 後過濾...")
    all_dockets = []
    for page in range(1, 100):
        resp  = api_get("dockets", {
            "filter[agencyId]": "PTO",
            "page[size]":       250,
            "page[number]":     page,
        })
        data  = resp.get("data", [])
        total = resp.get("meta", {}).get("totalElements", "?")
        if not data:
            break
        all_dockets.extend(data)
        print(f"\r  已抓 {len(all_dockets)}/{total}", end="", flush=True)
        if len(data) < 250:
            break
    print()

    result = []
    for d in all_dockets:
        docket_id   = d.get("id", "")
        docket_type = d.get("attributes", {}).get("docketType", "")

        # 依年份篩選：從 Docket ID 第三段取年份
        # 例如 PTO-P-2024-0003 → parts[2] = "2024"
        if FILTER_YEAR:
            parts = docket_id.split("-")
            year  = parts[2] if len(parts) >= 4 else ""
            if year != FILTER_YEAR:
                continue

        # 依 docket 類型篩選
        if FILTER_DOCKET_TYPE and docket_type != FILTER_DOCKET_TYPE:
            continue

        result.append(docket_id)

    print(f"  篩選結果：{len(result)} 個 dockets")
    return result


def main():
    print("=" * 65)
    print(f"  USPTO Docket Scraper v4")
    print(f"  開始時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 65)

    if API_KEY == "你的API_KEY":
        print("\n請先填入 API Key")
        return

    # ── Step 0：決定要抓哪些 dockets ─────────────────
    print(f"\n[準備] 決定抓取目標...")
    docket_ids = resolve_docket_ids()

    if not docket_ids:
        print("  沒有符合條件的 dockets，請確認設定區的篩選條件")
        return

    print(f"  ✓ 共 {len(docket_ids)} 個 dockets 待抓取：")
    for did in docket_ids:
        print(f"    {did}")

    # 依序抓取每個 docket
    for docket_id in docket_ids:
        print(f"\n{'='*65}")
        print(f"  抓取：{docket_id}")
        print(f"{'='*65}")
        fetch_single_docket(docket_id)


def fetch_single_docket(docket_id):
    """
    抓取單一 docket 的完整資料（documents + comments）。

    Args:
        docket_id (str): 要抓取的 docket ID
    """
    safe_id = docket_id.replace("-", "_").lower()

    raw_dir       = OUTPUT_DIR / f"{safe_id}_raw_json"
    progress_file = OUTPUT_DIR / f"{safe_id}_progress.txt"
    comments_csv  = OUTPUT_DIR / f"{safe_id}_comments.csv"

    if SAVE_RAW_JSON:
        raw_dir.mkdir(exist_ok=True)

    # ── Step 1：Docket 元資料 ────────────────────────
    print(f"\n[1/3] 抓取 Docket 元資料...")
    resp = api_get(f"dockets/{docket_id}")
    if not resp or "data" not in resp:
        print(f"  找不到此 Docket：{docket_id}，跳過")
        return

    docket_row = parse_docket(resp["data"])
    if SAVE_RAW_JSON:
        (raw_dir / "docket.json").write_text(
            json.dumps(resp, indent=2, ensure_ascii=False), encoding="utf-8"
        )

    print(f"  ✓ {docket_row['title']}")
    print(f"  ✓ 類型：{docket_row['docket_type']}  RIN：{docket_row['rin']}")

    # ── Step 2：Documents（含全文）───────────────────
    print(f"\n[2/3] 抓取 Documents（FETCH_FULLTEXT={FETCH_DOCUMENT_FULLTEXT}）...")
    docs_list = fetch_all_pages(
        "documents",
        {"filter[docketId]": docket_id},
        label="Documents"
    )

    all_documents = []
    for d in tqdm(docs_list, desc="  Document detail + fulltext"):
        detail   = api_get(f"documents/{d['id']}", {"include": "attachments"})
        doc_data = detail.get("data", d)

        if SAVE_RAW_JSON:
            (raw_dir / f"document_{d['id']}.json").write_text(
                json.dumps(detail, indent=2, ensure_ascii=False), encoding="utf-8"
            )

        all_documents.append(
            parse_document(doc_data, fetch_fulltext_flag=FETCH_DOCUMENT_FULLTEXT)
        )

    df_documents = pd.DataFrame(all_documents)
    print(f"\n  ✓ 共 {len(all_documents)} 份 Documents")
    for _, doc in df_documents.iterrows():
        has_text = "✓全文" if doc.get("document_full_text") else "✗無全文"
        print(f"    [{doc['document_type']:28s}] {doc['document_id']}  {has_text}")

    # ── Step 3：Comments（含附件，斷點續跑）──────────
    print(f"\n[3/3] 抓取所有 Comments（FETCH_DETAIL={FETCH_COMMENT_DETAIL}）...")
    done_ids     = load_progress(progress_file)
    all_comments = []

    if done_ids and comments_csv.exists():
        all_comments = pd.read_csv(comments_csv, encoding="utf-8-sig").to_dict("records")
        print(f"  ↩ 斷點續跑：已有 {len(done_ids)} 筆")

    valid_docs = df_documents[
        df_documents["object_id"].notna() & (df_documents["object_id"] != "")
    ]

    for _, doc_row in valid_docs.iterrows():
        object_id   = doc_row["object_id"]
        document_id = doc_row["document_id"]

        print(f"\n  → {document_id}")
        comments_list = fetch_all_pages(
            "comments",
            {"filter[commentOnId]": object_id},
            label="    Comments"
        )

        if not comments_list:
            print("    （無 comments）")
            continue

        pending = [c for c in comments_list if c["id"] not in done_ids]
        print(f"    待處理 {len(pending)} 筆 / 已完成 {len(comments_list)-len(pending)} 筆")

        if FETCH_COMMENT_DETAIL:
            for c in tqdm(pending, desc="    Detail", leave=False):
                detail = api_get(f"comments/{c['id']}", {"include": "attachments"})
                # ★ 修正：included（附件資訊）在 API 回傳的最外層，
                #   不在 data 裡面，必須手動合併進去再傳給 parse_comment，
                #   否則 extract_attachment_urls 永遠找不到附件
                comment_data = detail.get("data", c)
                comment_data["included"] = detail.get("included", [])

                if SAVE_RAW_JSON:
                    (raw_dir / f"comment_{c['id']}.json").write_text(
                        json.dumps(detail, indent=2, ensure_ascii=False), encoding="utf-8"
                    )

                all_comments.append(parse_comment(comment_data, docket_id, document_id,
                                                  fetch_att_text=FETCH_COMMENT_ATTACHMENT_TEXT))
                save_progress(progress_file, c["id"])

                if len(all_comments) % 50 == 0:
                    pd.DataFrame(all_comments).to_csv(
                        comments_csv, index=False, encoding="utf-8-sig"
                    )
        else:
            for c in pending:
                all_comments.append(parse_comment(c, docket_id, document_id,
                                                  fetch_att_text=False))
                save_progress(progress_file, c["id"])

    df_comments = pd.DataFrame(all_comments) if all_comments else pd.DataFrame()

    # ── 輸出 CSV ─────────────────────────────────────
    df_docket = pd.DataFrame([docket_row])
    df_docket.to_csv(
        OUTPUT_DIR / f"{safe_id}_docket.csv", index=False, encoding="utf-8-sig"
    )
    df_documents.to_csv(
        OUTPUT_DIR / f"{safe_id}_documents.csv", index=False, encoding="utf-8-sig"
    )
    if not df_comments.empty:
        df_comments.to_csv(comments_csv, index=False, encoding="utf-8-sig")

    # ── 輸出 Excel ────────────────────────────────────
    excel_path = OUTPUT_DIR / f"{safe_id}_full_data.xlsx"
    print(f"\n正在產生 Excel：{excel_path.name}")

    with pd.ExcelWriter(str(excel_path), engine="openpyxl") as writer:
        summary = {
            "欄位": ["Docket ID", "Title", "Type", "RIN",
                    "Documents", "Comments", "產生時間"],
            "值":   [docket_row["docket_id"], docket_row["title"],
                    docket_row["docket_type"], docket_row["rin"],
                    len(all_documents), len(all_comments),
                    datetime.now().strftime("%Y-%m-%d %H:%M")],
        }
        pd.DataFrame(summary).to_excel(writer, sheet_name="Summary",   index=False)
        df_docket.to_excel(            writer, sheet_name="Docket",    index=False)
        df_documents.to_excel(         writer, sheet_name="Documents", index=False)
        if not df_comments.empty:
            df_comments.to_excel(      writer, sheet_name="Comments",  index=False)

        for ws in writer.sheets.values():
            for col in ws.columns:
                w = max((len(str(c.value)) if c.value else 0) for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(w + 4, 80)

    # ── 完成摘要 ──────────────────────────────────────
    print("\n" + "=" * 65)
    print(f"  ✓ 完成！")
    print(f"  Documents： {len(all_documents)}")
    print(f"  Comments：  {len(all_comments)}")
    print(f"  輸出：{excel_path.name}")
    print(f"  結束時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 65)

    if progress_file.exists():
        progress_file.unlink()


if __name__ == "__main__":
    main()