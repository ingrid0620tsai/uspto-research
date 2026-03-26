"""
USPTO Docket Full Scraper v3 — regulations.gov v4 API
======================================================
v3 新增：
  - attachment_urls 拆成 attachment_url_pdf / attachment_url_html 兩個欄位
  - 自動抓取 document 全文（優先用 HTML，備援用 PDF）
  - document 全文存入 document_full_text 欄位
  - comment 附件同樣拆欄位

使用方式：
  1. pip install requests pandas openpyxl tqdm beautifulsoup4 pdfplumber
  2. 填入 API_KEY
  3. 設定 DOCKET_ID
  4. python uspto_docket_scraper_v3.py

申請免費 API Key：https://api.data.gov/signup/
速率限制：每小時 500 次 GET 請求（regulations.gov API）
全文抓取不佔 API 配額（直接下載 downloads.regulations.gov）
"""

import os
import io
import time
import json
import requests
import pandas as pd
from tqdm import tqdm
from datetime import datetime
from pathlib import Path
from bs4 import BeautifulSoup

# ─────────────────────────────────────────────────────
# ★ 設定區（只需修改這裡）
# ─────────────────────────────────────────────────────
API_KEY   = "og1wltTsDKFQB8tTUdMMFS4akawWzkULBrt18hDg"
DOCKET_ID = "PTO-P-2023-0048"

# 是否抓每筆 comment 的詳細內文 + 附件（建議 True）
FETCH_COMMENT_DETAIL = True

# 是否抓 document 全文（NPRM、Final Rule 等的正文）
FETCH_DOCUMENT_FULLTEXT = True

# 是否備份原始 JSON
SAVE_RAW_JSON = True
# ─────────────────────────────────────────────────────

BASE_URL   = "https://api.regulations.gov/v4"
DELAY      = 7.5
OUTPUT_DIR = Path(os.path.dirname(os.path.abspath(__file__)))


# ══════════════════════════════════════════════════════
# 全文抓取函式（不佔 API 配額）
# ══════════════════════════════════════════════════════

def fetch_fulltext_from_html(url: str) -> str:
    """
    從 .html 連結抓取純文字全文。
    regulations.gov 的 HTML 版本是最乾淨的全文來源。
    """
    if not url:
        return ""
    try:
        r = requests.get(url, timeout=30)
        if r.status_code != 200:
            return ""
        soup = BeautifulSoup(r.content, "html.parser")
        # 移除 script / style 標籤
        for tag in soup(["script", "style", "nav", "header", "footer"]):
            tag.decompose()
        text = soup.get_text(separator="\n")
        # 清理多餘空白行
        lines = [l.strip() for l in text.splitlines() if l.strip()]
        return "\n".join(lines)
    except Exception as e:
        print(f"\n  [HTML 全文抓取失敗] {url} — {e}")
        return ""


def fetch_fulltext_from_pdf(url: str) -> str:
    """
    從 .pdf 連結抓取純文字全文（備援方案）。
    需要 pdfplumber 套件。
    """
    if not url:
        return ""
    try:
        import pdfplumber
        r = requests.get(url, timeout=60)
        if r.status_code != 200:
            return ""
        with pdfplumber.open(io.BytesIO(r.content)) as pdf:
            pages_text = []
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    pages_text.append(t)
        return "\n".join(pages_text)
    except ImportError:
        print("\n  [提示] 未安裝 pdfplumber，PDF 全文抓取略過。執行：pip install pdfplumber")
        return ""
    except Exception as e:
        print(f"\n  [PDF 全文抓取失敗] {url} — {e}")
        return ""


def fetch_fulltext(html_url: str, pdf_url: str) -> str:
    """
    優先用 HTML，HTML 失敗才用 PDF。
    """
    if html_url:
        text = fetch_fulltext_from_html(html_url)
        if text:
            return text
    if pdf_url:
        return fetch_fulltext_from_pdf(pdf_url)
    return ""


# ══════════════════════════════════════════════════════
# 附件 URL 拆分函式
# ══════════════════════════════════════════════════════

def split_attachment_urls(urls: list) -> dict:
    """
    將附件 URL 清單依副檔名拆成：
      - pdf_urls：list of PDF 連結
      - html_urls：list of HTML 連結
      - other_urls：其他格式

    範例輸入：
      ["https://downloads.regulations.gov/PTO-P-2023-0048-0112/content.pdf",
       "https://downloads.regulations.gov/PTO-P-2023-0048-0112/content.html"]

    範例輸出：
      {
        "pdf_urls":   ["https://.../content.pdf"],
        "html_urls":  ["https://.../content.html"],
        "other_urls": []
      }
    """
    pdf_urls   = []
    html_urls  = []
    other_urls = []

    for url in urls:
        if not url:
            continue
        lower = url.lower()
        if lower.endswith(".pdf"):
            pdf_urls.append(url)
        elif lower.endswith(".htm") or lower.endswith(".html"):
            html_urls.append(url)
        else:
            other_urls.append(url)

    return {
        "pdf_urls":   pdf_urls,
        "html_urls":  html_urls,
        "other_urls": other_urls,
    }


# ══════════════════════════════════════════════════════
# HTTP 核心函式（同 v2）
# ══════════════════════════════════════════════════════

def api_get(endpoint: str, params: dict = {}, retries: int = 6) -> dict:
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
            elif r.status_code == 503:
                time.sleep(30 * (attempt + 1))
            else:
                print(f"\n  [HTTP {r.status_code}] {url}")
                time.sleep(10)
        except requests.RequestException as e:
            print(f"\n  [連線錯誤] {e}")
            time.sleep(30)

    return {}


def fetch_all_pages(endpoint: str, params: dict, label: str = "") -> list:
    all_items = []
    cursor    = None
    batch     = 0

    while True:
        batch += 1
        p = {**params, "page[size]": 250, "sort": "lastModifiedDate,documentId"}
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
            print(f"\r  {label} batch {batch} p{page} | {len(all_items)+len(segment)}/{total}",
                  end="", flush=True)
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
# 資料解析函式
# ══════════════════════════════════════════════════════

def parse_docket(raw: dict) -> dict:
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


def extract_att_urls(raw: dict) -> list:
    """從 document 或 comment 的 raw JSON 中提取所有附件 URL"""
    urls = []
    for item in raw.get("included", []):
        if item.get("type") == "attachments":
            for fmt in item.get("attributes", {}).get("fileFormats") or []:
                u = fmt.get("fileUrl", "") if isinstance(fmt, dict) else fmt
                if u:
                    urls.append(u)
    for fmt in raw.get("attributes", {}).get("fileFormats") or []:
        u = fmt.get("fileUrl", "") if isinstance(fmt, dict) else fmt
        if u and u not in urls:
            urls.append(u)
    return urls


def parse_document(raw: dict, fetch_fulltext_flag: bool = False) -> dict:
    a     = raw.get("attributes", {})
    links = raw.get("links", {})

    # 附件 URL 拆分
    all_urls = extract_att_urls(raw)
    split    = split_attachment_urls(all_urls)

    pdf_url  = split["pdf_urls"][0]  if split["pdf_urls"]  else ""
    html_url = split["html_urls"][0] if split["html_urls"] else ""

    # 全文抓取
    full_text = ""
    if fetch_fulltext_flag:
        full_text = fetch_fulltext(html_url, pdf_url)

    return {
        # 識別
        "document_id":              raw.get("id", ""),
        "object_id":                a.get("objectId", ""),
        "docket_id":                a.get("docketId", ""),
        "agency_id":                a.get("agencyId", ""),

        # 類型
        "document_type":            a.get("documentType", ""),
        "subtype":                  a.get("subtype", ""),
        "exhibit_type":             a.get("exhibitType", ""),

        # 內容
        "title":                    a.get("title", ""),
        "summary":                  a.get("summary", ""),
        "fr_doc_num":               a.get("frDocNum", ""),
        "cfr_parts":                ", ".join(a.get("cfrPart") or []),
        "full_text_xml_url":        a.get("fullTextXmlUrl", ""),

        # ★ 附件欄位（拆開）
        "attachment_url_pdf":       "; ".join(split["pdf_urls"]),
        "attachment_url_html":      "; ".join(split["html_urls"]),
        "attachment_url_other":     "; ".join(split["other_urls"]),
        "attachment_count":         len(all_urls),

        # ★ 全文
        "document_full_text":       full_text,
        "fulltext_source":          "html" if (full_text and html_url) else
                                    "pdf"  if (full_text and pdf_url)  else "",

        # 時間
        "posted_date":              a.get("postedDate", ""),
        "last_modified_date":       a.get("lastModifiedDate", ""),
        "comment_start_date":       a.get("commentStartDate", ""),
        "comment_end_date":         a.get("commentEndDate", ""),

        # 狀態
        "open_for_comment":         a.get("openForComment", False),
        "num_comments_received":    a.get("numberOfCommentsReceived", 0),
        "withdrawn":                a.get("withdrawn", False),

        # 連結
        "regulations_url":          f"https://www.regulations.gov/document/{raw.get('id','')}",
        "api_url":                  links.get("self", ""),
        "scraped_at":               datetime.utcnow().isoformat() + "Z",
    }


def parse_comment(raw: dict, docket_id: str = "", document_id: str = "") -> dict:
    a     = raw.get("attributes", {})

    # 附件 URL 拆分
    all_urls     = extract_att_urls(raw)
    split        = split_attachment_urls(all_urls)

    att_titles = []
    for item in raw.get("included", []):
        if item.get("type") == "attachments":
            att_titles.append(item.get("attributes", {}).get("title", ""))

    first = (a.get("firstName") or "").strip()
    last  = (a.get("lastName")  or "").strip()

    return {
        # 識別
        "comment_id":           raw.get("id", ""),
        "tracking_nbr":         a.get("trackingNbr", ""),
        "legacy_id":            a.get("legacyId", ""),

        # 三層關聯
        "docket_id":            a.get("docketId", "")   or docket_id,
        "document_id":          a.get("documentId", "") or document_id,
        "comment_on_id":        a.get("commentOnId", ""),

        # 類型
        "document_type":        a.get("documentType", ""),
        "subtype":              a.get("subtype", ""),

        # Commenter 資訊
        "commenter_name":       f"{first} {last}".strip(),
        "organization":         a.get("organization", ""),
        "gov_agency":           a.get("govAgency", ""),
        "gov_agency_type":      a.get("govAgencyType", ""),
        "city":                 a.get("city", ""),
        "state":                a.get("stateProvinceRegion", ""),
        "country":              a.get("country", ""),
        "zip":                  a.get("zip", ""),

        # 內容
        "title":                a.get("title", ""),
        "comment_text":         a.get("comment", ""),
        "doc_abstract":         a.get("docAbstract", ""),
        "page_count":           a.get("pageCount", ""),

        # ★ 附件欄位（拆開）
        "attachment_count":     len(all_urls),
        "attachment_titles":    "; ".join(filter(None, att_titles)),
        "attachment_url_pdf":   "; ".join(split["pdf_urls"]),
        "attachment_url_html":  "; ".join(split["html_urls"]),
        "attachment_url_other": "; ".join(split["other_urls"]),

        # 時間
        "posted_date":          a.get("postedDate", ""),
        "receive_date":         a.get("receiveDate", ""),
        "postmark_date":        a.get("postmarkDate", ""),
        "last_modified_date":   a.get("lastModifiedDate", ""),

        # 狀態
        "withdrawn":            a.get("withdrawn", False),
        "restrict_reason_type": a.get("restrictReasonType", ""),
        "reason_withdrawn":     a.get("reasonWithdrawn", ""),

        # 連結
        "regulations_url":      f"https://www.regulations.gov/comment/{raw.get('id','')}",
        "scraped_at":           datetime.utcnow().isoformat() + "Z",
    }


# ══════════════════════════════════════════════════════
# 斷點續跑
# ══════════════════════════════════════════════════════

def load_progress(f: Path) -> set:
    return set(f.read_text().splitlines()) if f.exists() else set()

def save_progress(f: Path, cid: str):
    with open(f, "a") as fp:
        fp.write(cid + "\n")


# ══════════════════════════════════════════════════════
# 主程式
# ══════════════════════════════════════════════════════

def main():
    safe_id = DOCKET_ID.replace("-", "_").lower()

    print("=" * 65)
    print(f"  USPTO Docket Full Scraper v3")
    print(f"  Docket：{DOCKET_ID}")
    print(f"  開始時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 65)

    if API_KEY == "YOUR_API_KEY_HERE":
        print("\n⚠️  請先填入 API Key（申請：https://api.data.gov/signup/）")
        return

    raw_dir       = OUTPUT_DIR / f"{safe_id}_raw_json"
    progress_file = OUTPUT_DIR / f"{safe_id}_progress.txt"
    comments_csv  = OUTPUT_DIR / f"{safe_id}_comments.csv"

    if SAVE_RAW_JSON:
        raw_dir.mkdir(exist_ok=True)

    # ── STEP 1：Docket ──────────────────────────────
    print(f"\n[1/3] 抓取 Docket 元資料...")
    resp = api_get(f"dockets/{DOCKET_ID}")
    if not resp or "data" not in resp:
        print("  ✗ 找不到此 Docket")
        return

    docket_row = parse_docket(resp["data"])
    if SAVE_RAW_JSON:
        (raw_dir / "docket.json").write_text(
            json.dumps(resp, indent=2, ensure_ascii=False), encoding="utf-8"
        )
    print(f"  ✓ {docket_row['title']}")
    print(f"  ✓ 類型：{docket_row['docket_type']}  RIN：{docket_row['rin']}")

    # ── STEP 2：Documents（含全文）──────────────────
    print(f"\n[2/3] 抓取 Documents（FETCH_FULLTEXT={FETCH_DOCUMENT_FULLTEXT}）...")
    docs_list = fetch_all_pages("documents", {"filter[docketId]": DOCKET_ID}, "Documents")

    documents = []
    for d in tqdm(docs_list, desc="  Document detail + fulltext"):
        detail   = api_get(f"documents/{d['id']}", {"include": "attachments"})
        doc_data = detail.get("data", d)

        if SAVE_RAW_JSON:
            (raw_dir / f"document_{d['id']}.json").write_text(
                json.dumps(detail, indent=2, ensure_ascii=False), encoding="utf-8"
            )

        documents.append(parse_document(doc_data, fetch_fulltext_flag=FETCH_DOCUMENT_FULLTEXT))

    df_documents = pd.DataFrame(documents)
    print(f"\n  ✓ 共 {len(documents)} 份 Documents")
    for _, doc in df_documents.iterrows():
        has_text = "✓全文" if doc.get("document_full_text") else "✗無全文"
        print(f"    [{doc['document_type']:28s}] {doc['document_id']}  {has_text}")

    # ── STEP 3：Comments ────────────────────────────
    print(f"\n[3/3] 抓取所有 Comments（FETCH_DETAIL={FETCH_COMMENT_DETAIL}）...")
    done_ids     = load_progress(progress_file)
    all_comments = []

    if done_ids and comments_csv.exists():
        all_comments = pd.read_csv(comments_csv).to_dict("records")
        print(f"  ↩ 斷點續跑：已有 {len(done_ids)} 筆")

    valid_docs = df_documents[
        df_documents["object_id"].notna() & (df_documents["object_id"] != "")
    ]

    for _, doc_row in valid_docs.iterrows():
        object_id   = doc_row["object_id"]
        document_id = doc_row["document_id"]

        print(f"\n  → {document_id}")
        comments_list = fetch_all_pages(
            "comments", {"filter[commentOnId]": object_id}, "    Comments"
        )
        if not comments_list:
            print("    （無 comments）")
            continue

        pending = [c for c in comments_list if c["id"] not in done_ids]
        print(f"    待處理 {len(pending)} 筆 / 已完成 {len(comments_list)-len(pending)} 筆")

        if FETCH_COMMENT_DETAIL:
            for c in tqdm(pending, desc="    Detail", leave=False):
                detail       = api_get(f"comments/{c['id']}", {"include": "attachments"})
                comment_data = detail.get("data", c)

                if SAVE_RAW_JSON:
                    (raw_dir / f"comment_{c['id']}.json").write_text(
                        json.dumps(detail, indent=2, ensure_ascii=False), encoding="utf-8"
                    )

                all_comments.append(parse_comment(comment_data, DOCKET_ID, document_id))
                save_progress(progress_file, c["id"])

                if len(all_comments) % 50 == 0:
                    pd.DataFrame(all_comments).to_csv(comments_csv, index=False, encoding="utf-8-sig")
        else:
            for c in pending:
                all_comments.append(parse_comment(c, DOCKET_ID, document_id))
                save_progress(progress_file, c["id"])

    df_comments = pd.DataFrame(all_comments) if all_comments else pd.DataFrame()

    # ── 輸出 CSV ────────────────────────────────────
    df_dockets = pd.DataFrame([docket_row])
    df_dockets.to_csv(  OUTPUT_DIR / f"{safe_id}_docket.csv",    index=False, encoding="utf-8-sig")
    df_documents.to_csv(OUTPUT_DIR / f"{safe_id}_documents.csv", index=False, encoding="utf-8-sig")
    if not df_comments.empty:
        df_comments.to_csv(comments_csv, index=False, encoding="utf-8-sig")

    # ── 輸出 Excel ──────────────────────────────────
    excel_path = OUTPUT_DIR / f"{safe_id}_full_data.xlsx"
    print(f"\n正在產生 Excel：{excel_path.name}")

    with pd.ExcelWriter(str(excel_path), engine="openpyxl") as writer:
        summary = {
            "欄位": ["Docket ID","Title","Docket Type","RIN","Category","Program",
                    "Keywords","URL","Total Documents","Total Comments","Scrape Date"],
            "值":   [docket_row["docket_id"], docket_row["title"], docket_row["docket_type"],
                    docket_row["rin"], docket_row["category"], docket_row["program"],
                    docket_row["keywords"], docket_row["regulations_url"],
                    len(documents), len(all_comments), datetime.now().strftime("%Y-%m-%d %H:%M")],
        }
        pd.DataFrame(summary).to_excel(writer, sheet_name="Summary",   index=False)
        df_dockets.to_excel(           writer, sheet_name="Docket",    index=False)
        df_documents.to_excel(         writer, sheet_name="Documents", index=False)
        if not df_comments.empty:
            df_comments.to_excel(      writer, sheet_name="Comments",  index=False)

        for ws in writer.sheets.values():
            for col in ws.columns:
                w = max((len(str(c.value)) if c.value else 0) for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(w + 4, 80)

    # ── 完成 ────────────────────────────────────────
    print("\n" + "=" * 65)
    print(f"  ✓ 完成！  Documents：{len(documents)}  Comments：{len(all_comments)}")
    print(f"  主要輸出：{excel_path.name}")
    print(f"  結束時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 65)

    if progress_file.exists():
        progress_file.unlink()


if __name__ == "__main__":
    main()