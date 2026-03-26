"""
檔案名稱：pto_rulemaking_stats.py
撰寫目的：統計 regulations.gov 中所有 PTO Rulemaking dockets 底下，
         各種 document type 的文件數量，以及對應收到的 comment 總數。
資料來源：regulations.gov v4 REST API
輸出格式：終端機統計表 + pto_rulemaking_stats.csv + pto_rulemaking_all_documents.csv
撰寫日期：2025-03-20
依賴套件：pip install requests pandas tqdm
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
# 說明：免費 Key 每小時 500 次，每次間隔 7.5 秒

# ════════════════════════════════════════════════════════

BASE_URL = "https://api.regulations.gov/v4"


# ══════════════════════════════════════════════════════
# SECTION 1：API 層
# ══════════════════════════════════════════════════════

def api_get(endpoint, params={}, retries=5):
    """
    發送單次 GET 請求。

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
                print(f"  [錯誤訊息] {r.text[:300]}")
                time.sleep(10)

        except requests.RequestException as e:
            print(f"\n  [連線錯誤] {e}，等待 30s...")
            time.sleep(30)

    return {}


def fetch_dockets(agency_id):
    """
    抓取指定 agency 的所有 dockets（自動翻頁）。
    注意：dockets endpoint 不支援 sort=lastModifiedDate,documentId，
          只用 page[size] 和 page[number] 翻頁。

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
    documents endpoint 支援 lastModifiedDate 排序，用於處理超過 5,000 筆。

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

        if len(segment) == 5000:
            last_ts = segment[-1].get("attributes", {}).get("lastModifiedDate", "")
            cursor  = last_ts.replace("T", " ").replace("Z", "")[:19] if last_ts else None
            if not cursor:
                break
        else:
            break

    return all_docs


# ══════════════════════════════════════════════════════
# SECTION 2：主程式
# ══════════════════════════════════════════════════════

def main():
    print("=" * 60)
    print("  PTO Rulemaking 統計腳本")
    print("=" * 60)

    if API_KEY == "你的API_KEY":
        print("\n請先填入 API Key")
        return

    # ── Step 1：抓所有 PTO dockets，Python 裡篩 Rulemaking ──
    # 說明：API 不支援直接用 docketType 篩選，
    #       先抓全部再在 Python 裡過濾
    print("\n[1/3] 抓取所有 PTO dockets...")
    all_dockets = fetch_dockets("PTO")

    rulemaking_dockets = [
        d for d in all_dockets
        if d.get("attributes", {}).get("docketType", "") == "Rulemaking"
    ]
    docket_ids = [d.get("id") for d in rulemaking_dockets]

    print(f"  PTO 總 dockets：{len(all_dockets)} 個")
    print(f"  其中 Rulemaking：{len(docket_ids)} 個")

    # ── Step 2：逐一抓每個 docket 底下的 documents ──────
    # 說明：用 numberOfCommentsReceived 欄位加總估算 comment 數
    print(f"\n[2/3] 抓取每個 Rulemaking docket 底下的 documents...")

    stats         = defaultdict(lambda: {"doc_count": 0, "comment_count": 0})
    all_documents = []

    for docket_id in tqdm(docket_ids, desc="  Dockets 進度"):
        docs = fetch_documents_for_docket(docket_id)

        for doc in docs:
            a          = doc.get("attributes", {})
            doc_type   = a.get("documentType", "Unknown")
            n_comments = a.get("numberOfCommentsReceived", 0) or 0

            stats[doc_type]["doc_count"]     += 1
            stats[doc_type]["comment_count"] += n_comments

            all_documents.append({
                "docket_id":             docket_id,
                "document_id":           doc.get("id", ""),
                "document_type":         doc_type,
                "subtype":               a.get("subtype", ""),
                "title":                 a.get("title", ""),
                "posted_date":           a.get("postedDate", ""),
                "comment_start_date":    a.get("commentStartDate", ""),
                "comment_end_date":      a.get("commentEndDate", ""),
                "open_for_comment":      a.get("openForComment", False),
                "num_comments_received": n_comments,
                "withdrawn":             a.get("withdrawn", False),
                "object_id":             a.get("objectId", ""),
                "regulations_url":       f"https://www.regulations.gov/document/{doc.get('id','')}",
            })

    # ── Step 3：整理統計並輸出 ──────────────────────────
    print(f"\n[3/3] 整理統計結果...")

    if not stats:
        print("  沒有抓到任何 documents，請檢查 API Key 是否正確")
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

    total_row = pd.DataFrame([{
        "document_type":        "合計",
        "doc_count":            df_stats["doc_count"].sum(),
        "comment_count":        df_stats["comment_count"].sum(),
        "avg_comments_per_doc": "",
    }])
    df_stats = pd.concat([df_stats, total_row], ignore_index=True)

    print("\n" + "=" * 65)
    print(f"  PTO Rulemaking Dockets：{len(docket_ids)} 個")
    print(f"  總 Documents：{len(all_documents)} 份")
    print("=" * 65)
    print(f"\n  {'Document Type':<32} {'文件數':>6} {'Comments':>10} {'平均':>8}")
    print("  " + "-" * 60)
    for _, row in df_stats.iterrows():
        avg = f"{row['avg_comments_per_doc']}" if row["avg_comments_per_doc"] != "" else ""
        print(f"  {str(row['document_type']):<32} {int(row['doc_count']):>6,} {int(row['comment_count']):>10,} {avg:>8}")

    df_stats.to_csv("pto_rulemaking_stats.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(all_documents).to_csv(
        "pto_rulemaking_all_documents.csv", index=False, encoding="utf-8-sig"
    )

    print(f"\n  pto_rulemaking_stats.csv")
    print(f"  pto_rulemaking_all_documents.csv")
    print("=" * 65)


if __name__ == "__main__":
    main()