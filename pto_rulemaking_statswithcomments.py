"""
檔案名稱：pto_rulemaking_stats.py
撰寫目的：統計 regulations.gov 中所有 PTO Rulemaking dockets 底下，
         各種 document type 的文件數量，以及每份 document 對應的
         真實 comment 筆數（透過 comments API 的 totalElements 取得）。
資料來源：regulations.gov v4 REST API
輸出格式：終端機統計表
         pto_rulemaking_stats.csv        — 各 document type 的彙總統計
         pto_rulemaking_all_documents.csv — 每份 document 的詳細資料（含 comment 數）
撰寫日期：2025-03-20
修改紀錄：
  - 2025-03-20：修正 comment 數來源，改用 comments API totalElements 直接計數
                （list API 的 numberOfCommentsReceived 欄位永遠回傳 0）
依賴套件：pip install requests pandas tqdm
注意事項：
  - 需要 regulations.gov API Key（申請：https://api.data.gov/signup/）
  - 每份 document 需打 1 次 comments API 查詢筆數
  - 388 份 documents × 7.5 秒 ≈ 49 分鐘
"""

import time
import requests
import pandas as pd
from collections import defaultdict
from tqdm import tqdm

# ════════════════════════════════════════════════════════
# ★ 設定區 — 只需修改這裡
# ════════════════════════════════════════════════════════

API_KEY = "og1wltTsDKFQB8tTUdMMFS4akawWzkULBrt18hDg"
# 申請網址：https://api.data.gov/signup/

DELAY = 7.5
# 說明：免費 Key 每小時 500 次，每次間隔 7.5 秒保留緩衝

# ════════════════════════════════════════════════════════
# 以下為程式本體，一般情況不需修改
# ════════════════════════════════════════════════════════

BASE_URL = "https://api.regulations.gov/v4"


# ══════════════════════════════════════════════════════
# SECTION 1：API 層
# 說明：所有 HTTP 請求集中在此，含錯誤處理與重試機制。
#       分為三個函式，分別對應三種查詢需求，
#       避免混用 sort 參數導致 400 錯誤。
# ══════════════════════════════════════════════════════

def api_get(endpoint, params={}, retries=5):
    """
    發送單次 GET 請求，含自動重試與 rate limit 保護。

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

            # 剩餘配額快耗盡時主動暫停，避免觸發 429
            if remaining < 10:
                print(f"\n  [Rate Limit 警告] 剩餘 {remaining} 次，暫停 65 秒...")
                time.sleep(65)

            if r.status_code == 200:
                time.sleep(DELAY)
                return r.json()
            elif r.status_code == 429:
                # Too Many Requests，等待時間隨重試次數遞增
                wait = 70 * (attempt + 1)
                print(f"\n  [429] 等待 {wait}s...")
                time.sleep(wait)
            elif r.status_code == 404:
                return {}
            else:
                print(f"\n  [HTTP {r.status_code}] {url}")
                print(f"  [錯誤訊息] {r.text[:300]}")
                time.sleep(10)

        except requests.RequestException as e:
            print(f"\n  [連線錯誤] {e}，等待 30s...")
            time.sleep(30)

    return {}


def fetch_dockets(agency_id):
    """
    抓取指定 agency 的所有 dockets（自動翻頁）。
    注意：dockets endpoint 不支援 sort=lastModifiedDate 參數，
          只用 page 翻頁，不加 sort。

    Args:
        agency_id (str): 機構代碼，例如 "PTO"

    Returns:
        list: 所有 docket 的原始 JSON 清單
    """
    all_dockets = []

    for page in range(1, 100):
        resp  = api_get("dockets", {
            "filter[agencyId]": agency_id,
            "page[size]":       250,
            "page[number]":     page,
        })
        data  = resp.get("data", [])
        total = resp.get("meta", {}).get("totalElements", "?")

        if not data:
            break

        all_dockets.extend(data)
        print(f"\r  Dockets：{len(all_dockets)}/{total}", end="", flush=True)

        if len(data) < 250:
            break

    print()
    return all_dockets


def fetch_documents_for_docket(docket_id):
    """
    抓取單一 docket 底下的所有 documents（自動翻頁）。
    documents endpoint 支援 lastModifiedDate 排序，
    用 cursor 策略處理超過 5,000 筆的情況。

    Args:
        docket_id (str): docket ID，例如 "PTO-P-2023-0048"

    Returns:
        list: 該 docket 所有 document 的原始 JSON 清單
    """
    all_docs = []
    cursor   = None

    while True:
        p = {
            "filter[docketId]": docket_id,
            "page[size]":       250,
            "sort":             "lastModifiedDate,documentId",
        }
        if cursor:
            p["filter[lastModifiedDate][ge]"] = cursor

        segment = []
        for page in range(1, 21):
            p["page[number]"] = page
            resp = api_get("documents", p)
            data = resp.get("data", [])
            if not data:
                break
            segment.extend(data)
            if len(data) < 250:
                break

        all_docs.extend(segment)

        # 若剛好 5,000 筆，移動 cursor 繼續抓下一批
        if len(segment) == 5000:
            last_ts = segment[-1].get("attributes", {}).get("lastModifiedDate", "")
            cursor  = last_ts.replace("T", " ").replace("Z", "")[:19] if last_ts else None
            if not cursor:
                break
        else:
            break

    return all_docs


def get_comment_count(object_id):
    """
    ★ 方案 B 核心函式：查詢單一 document 的真實 comment 筆數。

    做法：只取第 1 頁 1 筆，讀 meta.totalElements 就能知道總數，
          不需要把所有 comments 都抓下來，只用 1 次 API 請求。

    Args:
        object_id (str): document 的 objectId（非 documentId），
                         這是 comments API 的查詢 key

    Returns:
        int: 該 document 收到的 comment 總筆數，查詢失敗時回傳 0
    """
    if not object_id:
        # objectId 為空代表這份 document 無法被 comment，直接回傳 0
        return 0

    resp  = api_get("comments", {
        "filter[commentOnId]": object_id,
        "page[size]":          1,
        "page[number]":        1,
    })
    return resp.get("meta", {}).get("totalElements", 0) or 0


# ══════════════════════════════════════════════════════
# SECTION 2：主程式
# ══════════════════════════════════════════════════════

def main():
    print("=" * 65)
    print("  PTO Rulemaking 統計腳本（方案 B：直接數 comments 筆數）")
    print("=" * 65)

    if API_KEY == "你的API_KEY":
        print("\n請先填入 API Key")
        return

    # ── Step 1：抓所有 PTO dockets，Python 裡篩 Rulemaking ──
    # 說明：API 不支援直接用 docketType 篩選（會回傳 400），
    #       先抓全部 PTO dockets，再在 Python 裡過濾
    print("\n[1/3] 抓取所有 PTO dockets...")
    all_dockets = fetch_dockets("PTO")

    rulemaking_dockets = [
        d for d in all_dockets
        if d.get("attributes", {}).get("docketType", "") == "Rulemaking"
    ]
    docket_ids = [d.get("id") for d in rulemaking_dockets]

    print(f"  PTO 總 dockets：{len(all_dockets)} 個")
    print(f"  其中 Rulemaking：{len(docket_ids)} 個")

    # ── Step 2：抓每個 docket 底下的 documents ──────────
    print(f"\n[2/3] 抓取 documents...")
    all_documents_raw = []

    for docket_id in tqdm(docket_ids, desc="  Dockets"):
        docs = fetch_documents_for_docket(docket_id)
        for doc in docs:
            all_documents_raw.append((docket_id, doc))

    print(f"  共 {len(all_documents_raw)} 份 documents")

    # ── Step 3：查每份 document 的真實 comment 數 ────────
    # 說明：list API 的 numberOfCommentsReceived 永遠回傳 0，
    #       必須用 objectId 打 comments API 才能取得真實筆數。
    #       每份 document 打 1 次，共需約 49 分鐘。
    print(f"\n[3/3] 查詢每份 document 的真實 comment 筆數...")
    print(f"  預計時間：{len(all_documents_raw) * DELAY / 60:.0f} 分鐘")

    all_documents = []  # 保留完整資料，含所有原始欄位與 comment 數
    stats = defaultdict(lambda: {"doc_count": 0, "comment_count": 0})

    for docket_id, doc in tqdm(all_documents_raw, desc="  Documents"):
        a         = doc.get("attributes", {})
        doc_type  = a.get("documentType", "Unknown")
        object_id = a.get("objectId", "")

        # ★ 用 objectId 直接查 comment 真實筆數（方案 B 核心）
        n_comments = get_comment_count(object_id)

        # 累計各 document type 的統計
        stats[doc_type]["doc_count"]     += 1
        stats[doc_type]["comment_count"] += n_comments

        # 保留完整原始資料（所有欄位），供後續擴充使用
        all_documents.append({
            "docket_id":             docket_id,
            "document_id":           doc.get("id", ""),
            "object_id":             object_id,
            "document_type":         doc_type,
            "subtype":               a.get("subtype", ""),
            "title":                 a.get("title", ""),
            "posted_date":           a.get("postedDate", ""),
            "comment_start_date":    a.get("commentStartDate", ""),
            "comment_end_date":      a.get("commentEndDate", ""),
            "open_for_comment":      a.get("openForComment", False),
            "withdrawn":             a.get("withdrawn", False),
            # ★ 真實 comment 數（來自 comments API totalElements）
            "comment_count_actual":  n_comments,
            "regulations_url":       f"https://www.regulations.gov/document/{doc.get('id','')}",
        })

    # ── Step 4：整理統計表並輸出 ────────────────────────
    if not stats:
        print("  沒有抓到任何資料，請檢查 API Key")
        return

    rows = []
    for doc_type, v in stats.items():
        rows.append({
            "document_type":        doc_type,
            "doc_count":            v["doc_count"],
            "comment_count":        v["comment_count"],
            "avg_comments_per_doc": round(v["comment_count"] / v["doc_count"], 1)
                                    if v["doc_count"] > 0 else 0,
        })

    df_stats = pd.DataFrame(rows).sort_values(
        "comment_count", ascending=False
    ).reset_index(drop=True)

    # 加入合計列
    total_row = pd.DataFrame([{
        "document_type":        "合計",
        "doc_count":            df_stats["doc_count"].sum(),
        "comment_count":        df_stats["comment_count"].sum(),
        "avg_comments_per_doc": "",
    }])
    df_stats = pd.concat([df_stats, total_row], ignore_index=True)

    # 印出統計表
    print("\n" + "=" * 65)
    print(f"  PTO Rulemaking Dockets：{len(docket_ids)} 個")
    print(f"  總 Documents：{len(all_documents)} 份")
    print("=" * 65)
    print(f"\n  {'Document Type':<32} {'文件數':>6} {'Comments':>10} {'平均':>8}")
    print("  " + "-" * 60)
    for _, row in df_stats.iterrows():
        avg = f"{row['avg_comments_per_doc']}" if row["avg_comments_per_doc"] != "" else ""
        print(
            f"  {str(row['document_type']):<32} "
            f"{int(row['doc_count']):>6,} "
            f"{int(row['comment_count']):>10,} "
            f"{avg:>8}"
        )

    # 輸出 CSV
    df_stats.to_csv("pto_rulemaking_stats.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(all_documents).to_csv(
        "pto_rulemaking_all_documents.csv", index=False, encoding="utf-8-sig"
    )

    print(f"\n  pto_rulemaking_stats.csv        — 各類型彙總統計")
    print(f"  pto_rulemaking_all_documents.csv — 每份 document 詳細資料")
    print("=" * 65)


if __name__ == "__main__":
    main()