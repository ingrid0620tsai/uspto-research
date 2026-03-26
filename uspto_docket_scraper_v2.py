"""
USPTO Docket Full Scraper v2 — regulations.gov v4 API
======================================================
研究目的：完整抓取一個 USPTO docket 的所有資料，
         包含 docket 元資料、所有 documents、所有 comment letters 全文。

修正紀錄（v2）：
  - Rate limit 更正為每小時 500 次（delay 改為 7.5 秒）
  - Document detail API 加入 ?include=attachments
  - 補齊所有 agency-configurable 欄位
  - 新增 raw JSON 備份（確保不遺漏任何未知欄位）
  - 新增斷點續跑功能（中途中斷可從上次繼續）

使用方式：
  1. pip install requests pandas openpyxl tqdm
  2. 填入 API_KEY
  3. 設定 DOCKET_ID
  4. python uspto_docket_scraper_v2.py

申請免費 API Key：https://api.data.gov/signup/
速率限制：每小時 500 次 GET 請求
"""

import os
import time
import json
import requests
import pandas as pd
from tqdm import tqdm
from datetime import datetime
from pathlib import Path

# ─────────────────────────────────────────────────────
# ★ 設定區（只需修改這裡）
# ─────────────────────────────────────────────────────
API_KEY   = "og1wltTsDKFQB8tTUdMMFS4akawWzkULBrt18hDg"       # 填入你的 API Key
DOCKET_ID = "PTO-P-2023-0048"         # 要抓的 Docket ID

# 是否抓每筆 comment 的詳細內文 + 附件連結
# True  = 完整資料（推薦，但較慢）
# False = 快速模式（comment 全文可能被截斷）
FETCH_COMMENT_DETAIL = True

# 是否同時備份原始 JSON（建議 True，確保資料完整）
SAVE_RAW_JSON = True
# ─────────────────────────────────────────────────────

BASE_URL = "https://api.regulations.gov/v4"
DELAY    = 7.5   # 秒（500 requests/hr = 7.2 秒/request，加緩衝）

# 輸出目錄：與腳本同資料夾
OUTPUT_DIR = Path(os.path.dirname(os.path.abspath(__file__)))


# ══════════════════════════════════════════════════════
# HTTP 核心函式
# ══════════════════════════════════════════════════════

def api_get(endpoint: str, params: dict = {}, retries: int = 6) -> dict:
    """
    發送 GET 請求，含自動重試與 rate limit 處理。
    回應 header 中的 X-RateLimit-Remaining 若快到底會主動暫停。
    """
    p = {**params, "api_key": API_KEY}
    url = f"{BASE_URL}/{endpoint}"

    for attempt in range(retries):
        try:
            r = requests.get(url, params=p, timeout=30)

            # 檢查 rate limit 剩餘次數
            remaining = int(r.headers.get("X-RateLimit-Remaining", 999))
            if remaining < 10:
                print(f"\n  [Rate Limit 警告] 剩餘 {remaining} 次，暫停 60 秒...")
                time.sleep(60)

            if r.status_code == 200:
                time.sleep(DELAY)
                return r.json()

            elif r.status_code == 429:
                wait = 70 * (attempt + 1)
                print(f"\n  [429 Rate Limit] 等待 {wait}s 後重試...")
                time.sleep(wait)

            elif r.status_code == 404:
                print(f"\n  [404 Not Found] {url}")
                return {}

            elif r.status_code == 503:
                wait = 30 * (attempt + 1)
                print(f"\n  [503 Service Unavailable] 等待 {wait}s...")
                time.sleep(wait)

            else:
                print(f"\n  [HTTP {r.status_code}] {url} | {r.text[:200]}")
                time.sleep(10)

        except requests.RequestException as e:
            print(f"\n  [連線錯誤] {e} — 等待 30s...")
            time.sleep(30)

    print(f"  [失敗] 超過最大重試次數：{url}")
    return {}


def fetch_all_pages(endpoint: str, params: dict, label: str = "") -> list:
    """
    自動翻頁（每頁 250 筆，最多 20 頁 = 5,000 筆）。
    超過 5,000 筆時，用 lastModifiedDate cursor 分段繼續抓。
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
                f"\r  {label} batch {batch} page {page} | "
                f"本批 {len(segment)} / 總計 {len(all_items)+len(segment)}/{total}",
                end="", flush=True
            )

            if len(data) < 250:
                break

        all_items.extend(segment)

        # 若本 batch 滿 5,000，代表還有更多 → 移動 cursor
        if len(segment) == 5000:
            last_ts = segment[-1].get("attributes", {}).get("lastModifiedDate", "")
            if last_ts:
                cursor = last_ts.replace("T", " ").replace("Z", "")[:19]
            else:
                break
        else:
            break

    print()
    return all_items


# ══════════════════════════════════════════════════════
# 資料解析函式（完整欄位版）
# ══════════════════════════════════════════════════════

def parse_docket(raw: dict) -> dict:
    """TABLE 1 — dockets：完整解析所有公開欄位"""
    a = raw.get("attributes", {})
    return {
        # 識別
        "docket_id":          raw.get("id", ""),
        "agency_id":          a.get("agencyId", ""),
        "rin":                a.get("rin", ""),

        # 類型（★ 研究核心：Rulemaking vs Nonrulemaking）
        "docket_type":        a.get("docketType", ""),
        "category":           a.get("category", ""),
        "program":            a.get("program", ""),
        "sub_type":           a.get("subType", ""),

        # 內容
        "title":              a.get("title", ""),
        "description":        a.get("description", ""),
        "keywords":           ", ".join(a.get("keywords") or []),

        # 時間
        "effective_date":     a.get("effectiveDate", ""),
        "modify_date":        a.get("modifyDate", ""),

        # 連結
        "regulations_url":    f"https://www.regulations.gov/docket/{raw.get('id','')}",

        # 爬取時間
        "scraped_at":         datetime.utcnow().isoformat() + "Z",
    }


def parse_document(raw: dict) -> dict:
    """
    TABLE 2 — documents：解析文件，含附件連結。
    傳入的 raw 應是 detail API（?include=attachments）回傳的資料。
    """
    a     = raw.get("attributes", {})
    links = raw.get("links", {})

    # 附件（需 ?include=attachments）
    att_urls = []
    for item in raw.get("included", []):
        if item.get("type") == "attachments":
            for fmt in item.get("attributes", {}).get("fileFormats") or []:
                url = fmt.get("fileUrl", "") if isinstance(fmt, dict) else fmt
                if url:
                    att_urls.append(url)

    # fileFormats 有時也在 attributes 裡
    for fmt in a.get("fileFormats") or []:
        url = fmt.get("fileUrl", "") if isinstance(fmt, dict) else fmt
        if url and url not in att_urls:
            att_urls.append(url)

    return {
        # 識別
        "document_id":               raw.get("id", ""),
        "object_id":                 a.get("objectId", ""),   # ★ 查 comments 用
        "docket_id":                 a.get("docketId", ""),
        "agency_id":                 a.get("agencyId", ""),

        # 類型（★ 追蹤 ANPRM→NPRM→Final Rule 流程）
        "document_type":             a.get("documentType", ""),
        "subtype":                   a.get("subtype", ""),
        "exhibit_type":              a.get("exhibitType", ""),

        # 內容
        "title":                     a.get("title", ""),
        "summary":                   a.get("summary", ""),
        "fr_doc_num":                a.get("frDocNum", ""),     # Federal Register 文件號
        "cfr_parts":                 ", ".join(a.get("cfrPart") or []),
        "full_text_xml_url":         a.get("fullTextXmlUrl", ""),
        "attachment_urls":           "; ".join(att_urls),
        "attachment_count":          len(att_urls),

        # 時間
        "posted_date":               a.get("postedDate", ""),
        "last_modified_date":        a.get("lastModifiedDate", ""),
        "comment_start_date":        a.get("commentStartDate", ""),
        "comment_end_date":          a.get("commentEndDate", ""),

        # 狀態
        "open_for_comment":          a.get("openForComment", False),
        "num_comments_received":     a.get("numberOfCommentsReceived", 0),
        "withdrawn":                 a.get("withdrawn", False),

        # 連結
        "regulations_url":           f"https://www.regulations.gov/document/{raw.get('id','')}",
        "api_url":                   links.get("self", ""),

        "scraped_at":                datetime.utcnow().isoformat() + "Z",
    }


def parse_comment(raw: dict, docket_id: str = "", document_id: str = "") -> dict:
    """
    TABLE 3 — comments：完整解析所有公開欄位。
    傳入 detail API（?include=attachments）的回傳以取得全文與附件。
    """
    a     = raw.get("attributes", {})
    links = raw.get("links", {})

    # 附件（只在 detail API 回傳）
    att_urls  = []
    att_names = []
    for item in raw.get("included", []):
        if item.get("type") == "attachments":
            att_a = item.get("attributes", {})
            att_names.append(att_a.get("title", ""))
            for fmt in att_a.get("fileFormats") or []:
                url = fmt.get("fileUrl", "") if isinstance(fmt, dict) else fmt
                if url:
                    att_urls.append(url)

    # 姓名
    first = (a.get("firstName") or "").strip()
    last  = (a.get("lastName")  or "").strip()
    name  = f"{first} {last}".strip() or ""

    return {
        # ── 識別欄位 ──────────────────────────
        "comment_id":           raw.get("id", ""),
        "tracking_nbr":         a.get("trackingNbr", ""),
        "legacy_id":            a.get("legacyId", ""),

        # ── 三層關聯 key ───────────────────────
        "docket_id":            a.get("docketId", "")    or docket_id,
        "document_id":          a.get("documentId", "")  or document_id,
        "comment_on_id":        a.get("commentOnId", ""),   # parent document 的 objectId

        # ── Comment 類型 ───────────────────────
        "document_type":        a.get("documentType", ""),  # 通常是 "Public Submission"
        "subtype":              a.get("subtype", ""),

        # ── Commenter 資訊（★ 分析誰在回應）──
        "commenter_name":       name,
        "organization":         a.get("organization", ""),
        "gov_agency":           a.get("govAgency", ""),
        "gov_agency_type":      a.get("govAgencyType", ""),
        # govAgencyType 常見值：Federal / State / Local / Tribal / Foreign

        # ── 地理資訊 ───────────────────────────
        "city":                 a.get("city", ""),
        "state":                a.get("stateProvinceRegion", ""),
        "country":              a.get("country", ""),
        "zip":                  a.get("zip", ""),

        # ── Comment 內容（★ 研究核心）──────────
        "title":                a.get("title", ""),
        "comment_text":         a.get("comment", ""),        # 全文（detail API）
        "doc_abstract":         a.get("docAbstract", ""),
        "page_count":           a.get("pageCount", ""),

        # ── 附件 ───────────────────────────────
        "attachment_count":     len(att_urls),
        "attachment_titles":    "; ".join(filter(None, att_names)),
        "attachment_urls":      "; ".join(att_urls),

        # ── 時間戳 ─────────────────────────────
        "posted_date":          a.get("postedDate", ""),
        "receive_date":         a.get("receiveDate", ""),
        "postmark_date":        a.get("postmarkDate", ""),
        "last_modified_date":   a.get("lastModifiedDate", ""),

        # ── 狀態 ───────────────────────────────
        "withdrawn":            a.get("withdrawn", False),
        "restrict_reason_type": a.get("restrictReasonType", ""),
        "reason_withdrawn":     a.get("reasonWithdrawn", ""),

        # ── 連結 ───────────────────────────────
        "regulations_url":      f"https://www.regulations.gov/comment/{raw.get('id','')}",

        "scraped_at":           datetime.utcnow().isoformat() + "Z",
    }


# ══════════════════════════════════════════════════════
# 斷點續跑輔助
# ══════════════════════════════════════════════════════

def load_progress(progress_file: Path) -> set:
    """載入已完成的 comment ID 清單（斷點續跑用）"""
    if progress_file.exists():
        with open(progress_file) as f:
            return set(line.strip() for line in f if line.strip())
    return set()


def save_progress(progress_file: Path, comment_id: str):
    """追加已完成的 comment ID"""
    with open(progress_file, "a") as f:
        f.write(comment_id + "\n")


# ══════════════════════════════════════════════════════
# 主程式
# ══════════════════════════════════════════════════════

def main():
    safe_id = DOCKET_ID.replace("-", "_").lower()

    print("=" * 65)
    print(f"  USPTO Docket Full Scraper v2")
    print(f"  Docket：{DOCKET_ID}")
    print(f"  開始時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 65)

    if API_KEY == "YOUR_API_KEY_HERE":
        print("\n⚠️  請先填入 API Key（申請：https://api.data.gov/signup/）")
        return

    raw_dir = OUTPUT_DIR / f"{safe_id}_raw_json"
    if SAVE_RAW_JSON:
        raw_dir.mkdir(exist_ok=True)

    progress_file = OUTPUT_DIR / f"{safe_id}_progress.txt"

    # ────────────────────────────────────────────
    # STEP 1：Docket 元資料
    # ────────────────────────────────────────────
    print(f"\n[1/3] 抓取 Docket 元資料...")
    resp = api_get(f"dockets/{DOCKET_ID}")
    if not resp or "data" not in resp:
        print("  ✗ 找不到此 Docket，請確認 DOCKET_ID")
        return

    docket_raw = resp["data"]
    docket_row = parse_docket(docket_raw)

    if SAVE_RAW_JSON:
        with open(raw_dir / "docket.json", "w", encoding="utf-8") as f:
            json.dump(resp, f, indent=2, ensure_ascii=False)

    print(f"  ✓ 標題：    {docket_row['title']}")
    print(f"  ✓ 類型：    {docket_row['docket_type']}")
    print(f"  ✓ RIN：     {docket_row['rin']}")
    print(f"  ✓ Category：{docket_row['category']}")

    # ────────────────────────────────────────────
    # STEP 2：Documents（含 detail + attachments）
    # ────────────────────────────────────────────
    print(f"\n[2/3] 抓取 Documents...")
    docs_list_raw = fetch_all_pages(
        "documents",
        {"filter[docketId]": DOCKET_ID},
        label="Documents"
    )

    documents = []
    print(f"  取得 document 詳細資料（含附件連結）...")
    for d in tqdm(docs_list_raw, desc="  Document detail"):
        detail = api_get(f"documents/{d['id']}", {"include": "attachments"})
        doc_data = detail.get("data", d)

        if SAVE_RAW_JSON:
            with open(raw_dir / f"document_{d['id']}.json", "w", encoding="utf-8") as f:
                json.dump(detail, f, indent=2, ensure_ascii=False)

        documents.append(parse_document(doc_data))

    df_documents = pd.DataFrame(documents)
    print(f"\n  ✓ 共 {len(documents)} 份 Documents")
    for _, doc in df_documents.iterrows():
        nc = doc.get("num_comments_received", 0)
        print(f"    [{doc['document_type']:30s}] {doc['document_id']}  ({nc} comments)")

    # ────────────────────────────────────────────────────────
    # STEP 3：Comments（含 detail + attachments + 斷點續跑）
    # ────────────────────────────────────────────────────────
    print(f"\n[3/3] 抓取所有 Comments（FETCH_DETAIL={FETCH_COMMENT_DETAIL}）...")
    done_ids  = load_progress(progress_file)
    all_comments = []

    # 如果有斷點進度，先讀已儲存的 CSV
    comments_csv = OUTPUT_DIR / f"{safe_id}_comments.csv"
    if done_ids and comments_csv.exists():
        all_comments = pd.read_csv(comments_csv).to_dict("records")
        print(f"  ↩ 斷點續跑：已有 {len(done_ids)} 筆，從上次繼續...")

    valid_docs = df_documents[
        df_documents["object_id"].notna() & (df_documents["object_id"] != "")
    ]

    for _, doc_row in valid_docs.iterrows():
        object_id   = doc_row["object_id"]
        document_id = doc_row["document_id"]
        doc_type    = doc_row["document_type"]

        print(f"\n  → [{doc_type}] {document_id}  (objectId={object_id})")

        comments_list = fetch_all_pages(
            "comments",
            {"filter[commentOnId]": object_id},
            label="    Comments"
        )

        if not comments_list:
            print("    （此 document 無 comments）")
            continue

        # 過濾已完成的（斷點續跑）
        pending = [c for c in comments_list if c["id"] not in done_ids]
        print(f"    待處理：{len(pending)} 筆（已完成：{len(comments_list)-len(pending)} 筆）")

        if FETCH_COMMENT_DETAIL:
            for c in tqdm(pending, desc="    Detail", leave=False):
                detail = api_get(
                    f"comments/{c['id']}",
                    {"include": "attachments"}
                )
                comment_data = detail.get("data", c)

                if SAVE_RAW_JSON:
                    with open(raw_dir / f"comment_{c['id']}.json", "w", encoding="utf-8") as f:
                        json.dump(detail, f, indent=2, ensure_ascii=False)

                all_comments.append(parse_comment(comment_data, DOCKET_ID, document_id))
                save_progress(progress_file, c["id"])

                # 每 50 筆存一次 CSV（防止意外中斷遺失資料）
                if len(all_comments) % 50 == 0:
                    pd.DataFrame(all_comments).to_csv(
                        comments_csv, index=False, encoding="utf-8-sig"
                    )
        else:
            for c in pending:
                all_comments.append(parse_comment(c, DOCKET_ID, document_id))
                save_progress(progress_file, c["id"])

    df_comments = pd.DataFrame(all_comments) if all_comments else pd.DataFrame()

    # ────────────────────────────────────────────
    # 輸出 CSV
    # ────────────────────────────────────────────
    df_dockets  = pd.DataFrame([docket_row])
    dockets_csv  = OUTPUT_DIR / f"{safe_id}_docket.csv"
    docs_csv     = OUTPUT_DIR / f"{safe_id}_documents.csv"

    df_dockets.to_csv(dockets_csv, index=False, encoding="utf-8-sig")
    df_documents.to_csv(docs_csv, index=False, encoding="utf-8-sig")
    if not df_comments.empty:
        df_comments.to_csv(comments_csv, index=False, encoding="utf-8-sig")

    # ────────────────────────────────────────────
    # 輸出 Excel（4 個工作表）
    # ────────────────────────────────────────────
    excel_path = OUTPUT_DIR / f"{safe_id}_full_data.xlsx"
    print(f"\n正在產生 Excel：{excel_path.name}")

    with pd.ExcelWriter(str(excel_path), engine="openpyxl") as writer:

        # Sheet 0：Summary
        summary = {
            "欄位": ["Docket ID", "Title", "Docket Type", "RIN", "Category",
                    "Program", "Keywords", "URL",
                    "Total Documents", "Total Comments", "Scrape Date"],
            "值":   [docket_row["docket_id"], docket_row["title"],
                    docket_row["docket_type"], docket_row["rin"],
                    docket_row["category"], docket_row["program"],
                    docket_row["keywords"], docket_row["regulations_url"],
                    len(documents), len(all_comments),
                    datetime.now().strftime("%Y-%m-%d %H:%M")],
        }
        pd.DataFrame(summary).to_excel(writer, sheet_name="Summary",   index=False)
        df_dockets.to_excel(           writer, sheet_name="Docket",    index=False)
        df_documents.to_excel(         writer, sheet_name="Documents", index=False)
        if not df_comments.empty:
            df_comments.to_excel(      writer, sheet_name="Comments",  index=False)

        # 自動調整欄寬
        for ws in writer.sheets.values():
            for col in ws.columns:
                max_w = max(
                    (len(str(c.value)) if c.value is not None else 0) for c in col
                )
                ws.column_dimensions[col[0].column_letter].width = min(max_w + 4, 80)

    # ────────────────────────────────────────────
    # 完成摘要
    # ────────────────────────────────────────────
    print("\n" + "=" * 65)
    print("  ✓ 完成！")
    print(f"  Docket：    {DOCKET_ID}")
    print(f"  Documents： {len(documents)} 份")
    print(f"  Comments：  {len(all_comments)} 筆")
    print(f"\n  主要輸出：  {excel_path.name}")
    if SAVE_RAW_JSON:
        print(f"  原始 JSON： {raw_dir.name}/ 資料夾")
    print(f"\n  結束時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 65)

    # 完成後清除斷點進度檔
    if progress_file.exists():
        progress_file.unlink()
        print("  （斷點進度檔已清除）")


if __name__ == "__main__":
    main()
