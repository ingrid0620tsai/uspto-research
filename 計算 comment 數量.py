"""
檔案名稱：pto_rulemaking_comment_count.py
撰寫目的：統計所有 PTO Rulemaking dockets 底下，
         每個 docket 的真實 comment 筆數，
         以及各 document type 的 comment 筆數分佈。
資料來源：regulations.gov v4 REST API
輸出格式：終端機印出統計表 + pto_rulemaking_comment_count.csv
撰寫日期：2025-03-20
原理說明：
  對每份 document 的 object_id 打一次 comments API，
  只取 meta.totalElements（第1頁1筆即可），不下載 comment 內容。
  這是取得真實 comment 筆數最快、最準確的方法。
依賴套件：pip install requests pandas tqdm
"""

import time
import requests
import pandas as pd
from tqdm import tqdm

# ════════════════════════════════════════════════════════
# ★ 設定區
# ════════════════════════════════════════════════════════

API_KEY = "og1wltTsDKFQB8tTUdMMFS4akawWzkULBrt18hDg"
DELAY   = 7.5

# ════════════════════════════════════════════════════════

BASE_URL = "https://api.regulations.gov/v4"


# ══════════════════════════════════════════════════════
# SECTION 1：API 層
# ══════════════════════════════════════════════════════

def api_get(endpoint, params={}, retries=5):
    """
    發送單次 GET 請求。

    Args:
        endpoint (str) : API 路徑
        params   (dict): 查詢參數，api_key 自動附加
        retries  (int) : 最多重試幾次

    Returns:
        dict: API 回傳的完整 JSON，失敗時回傳 {}
    """
    p   = {**params, "api_key": API_KEY}
    url = f"{BASE_URL}/{endpoint}"

    for attempt in range(retries):
        try:
            r         = requests.get(url, params=p, timeout=30)
            remaining = int(r.headers.get("X-RateLimit-Remaining", 999))
            if remaining < 10:
                print(f"\n  [Rate Limit] 剩餘 {remaining} 次，暫停 65s...")
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
                time.sleep(10)
        except requests.RequestException as e:
            print(f"\n  [連線錯誤] {e}，等待 30s...")
            time.sleep(30)
    return {}


def fetch_dockets_all():
    """
    抓取所有 PTO dockets（自動翻頁），在 Python 裡篩選 Rulemaking。
    注意：dockets endpoint 不支援 sort 參數，只用 page 翻頁。

    Returns:
        list: 所有 PTO Rulemaking docket 的原始 JSON 清單
    """
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
    return [
        d for d in all_dockets
        if d.get("attributes", {}).get("docketType", "") == "Rulemaking"
    ]


def fetch_documents_for_docket(docket_id):
    """
    抓取單一 docket 底下的所有 documents（自動翻頁）。

    Args:
        docket_id (str): docket ID

    Returns:
        list: 所有 document 的原始 JSON 清單
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
    ★ 核心函式：用 object_id 查詢單一 document 的真實 comment 筆數。
    只取第 1 頁 1 筆，讀 meta.totalElements，1 次 API 就得到總數。

    Args:
        object_id (str): document 的 objectId

    Returns:
        int: comment 真實筆數，查詢失敗回傳 0
    """
    if not object_id:
        return 0
    resp = api_get("comments", {
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
    print("  PTO Rulemaking — Comment 筆數統計")
    print("=" * 65)

    if API_KEY == "你的API_KEY":
        print("\n請先填入 API Key")
        return

    # ── Step 1：取得所有 Rulemaking dockets ──────────
    print("\n[1/3] 抓取所有 PTO Rulemaking dockets...")
    rulemaking_dockets = fetch_dockets_all()
    docket_ids = [d.get("id") for d in rulemaking_dockets]
    print(f"  ✓ 共 {len(docket_ids)} 個 Rulemaking dockets")

    # ── Step 2：抓每個 docket 的 documents ───────────
    print(f"\n[2/3] 抓取每個 docket 的 documents...")
    all_rows = []   # 保留完整資料（每份 document 一列）

    for docket_id in tqdm(docket_ids, desc="  Dockets"):
        docket_data = next(
            (d.get("attributes", {}) for d in rulemaking_dockets if d.get("id") == docket_id), {}
        )
        docs = fetch_documents_for_docket(docket_id)
        for doc in docs:
            a = doc.get("attributes", {})
            all_rows.append({
                "docket_id":     docket_id,
                "docket_title":  docket_data.get("title", ""),
                "document_id":   doc.get("id", ""),
                "object_id":     a.get("objectId", ""),
                "document_type": a.get("documentType", ""),
                "posted_date":   a.get("postedDate", ""),
                "comment_count": 0,   # Step 3 填入
            })

    print(f"  ✓ 共 {len(all_rows)} 份 documents")

    # ── Step 3：查每份 document 的真實 comment 筆數 ──
    # 每份打 1 次 API，只取 totalElements，不下載 comment 內容
    print(f"\n[3/3] 查詢每份 document 的 comment 筆數...")
    print(f"  預計時間：{len(all_rows) * DELAY / 60:.0f} 分鐘")

    for row in tqdm(all_rows, desc="  Documents"):
        row["comment_count"] = get_comment_count(row["object_id"])

    df = pd.DataFrame(all_rows)

    # ── 彙總統計 ─────────────────────────────────────
    # 1. 依 document_type 彙總
    by_type = df.groupby("document_type").agg(
        doc_count    = ("document_id",   "count"),
        comment_count= ("comment_count", "sum"),
    ).reset_index()
    by_type["avg_per_doc"] = (by_type["comment_count"] / by_type["doc_count"]).round(1)
    by_type = by_type.sort_values("comment_count", ascending=False)

    # 2. 依 docket 彙總
    by_docket = df.groupby(["docket_id", "docket_title"]).agg(
        doc_count    = ("document_id",   "count"),
        comment_count= ("comment_count", "sum"),
    ).reset_index().sort_values("comment_count", ascending=False)

    # ── 印出統計表 ────────────────────────────────────
    print("\n" + "=" * 65)
    print(f"  PTO Rulemaking dockets：{len(docket_ids)} 個")
    print(f"  總 documents：{len(all_rows)} 份")
    print(f"  總 comments： {df['comment_count'].sum():,} 筆")
    print("=" * 65)

    print(f"\n【依 Document Type 彙總】")
    print(f"  {'Document Type':<32} {'文件數':>6} {'Comments':>10} {'平均':>8}")
    print("  " + "-" * 58)
    for _, row in by_type.iterrows():
        print(f"  {row['document_type']:<32} {row['doc_count']:>6,} {row['comment_count']:>10,} {row['avg_per_doc']:>8}")
    print(f"  {'合計':<32} {by_type['doc_count'].sum():>6,} {by_type['comment_count'].sum():>10,}")

    print(f"\n【依 Docket 彙總（前20名）】")
    print(f"  {'Docket ID':<28} {'Comments':>10}  標題")
    print("  " + "-" * 75)
    for _, row in by_docket.head(20).iterrows():
        title = str(row["docket_title"])[:40]
        print(f"  {row['docket_id']:<28} {row['comment_count']:>10,}  {title}")

    # ── 輸出 CSV ──────────────────────────────────────
    # 保留完整資料（每份 document 一列，含 comment_count）
    df.to_csv("pto_rulemaking_comment_count.csv", index=False, encoding="utf-8-sig")

    # 彙總統計
    by_type.to_csv("pto_rulemaking_by_doctype.csv", index=False, encoding="utf-8-sig")
    by_docket.to_csv("pto_rulemaking_by_docket.csv", index=False, encoding="utf-8-sig")

    print(f"\n  ✓ pto_rulemaking_comment_count.csv  （每份 document 一列）")
    print(f"  ✓ pto_rulemaking_by_doctype.csv     （依 document type 彙總）")
    print(f"  ✓ pto_rulemaking_by_docket.csv      （依 docket 彙總）")
    print("=" * 65)


if __name__ == "__main__":
    main()