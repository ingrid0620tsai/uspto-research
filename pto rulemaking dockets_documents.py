import requests
import pandas as pd
import time

API_KEY = "og1wltTsDKFQB8tTUdMMFS4akawWzkULBrt18hDg"
BASE    = "https://api.regulations.gov/v4"

# Step 1：抓所有 PTO Rulemaking dockets
print("抓取所有 PTO Rulemaking dockets...")
all_dockets = []
for page in range(1, 50):
    r = requests.get(f"{BASE}/dockets", params={
        "filter[agencyId]":   "PTO",
        "filter[docketType]": "Rulemaking",
        "page[size]":         250,
        "page[number]":       page,
        "api_key":            API_KEY
    })
    data = r.json()
    rows = data.get("data", [])
    if not rows:
        break
    all_dockets.extend(rows)
    total = data.get("meta", {}).get("totalElements", "?")
    print(f"  第 {page} 頁，目前 {len(all_dockets)}/{total} 筆")
    time.sleep(7.5)

print(f"✓ 共 {len(all_dockets)} 個 Rulemaking dockets")

# Step 2：整理成表格，每個 docket 一列
rows = []
for d in all_dockets:
    a = d.get("attributes", {})
    rows.append({
        "docket_id":   d.get("id", ""),
        "title":       a.get("title", ""),
        "docket_type": a.get("docketType", ""),
        "modify_date": a.get("modifyDate", ""),
        "regulations_url": f"https://www.regulations.gov/docket/{d.get('id','')}",
    })

df = pd.DataFrame(rows)
df.to_csv("pto_rulemaking_dockets.csv", index=False, encoding="utf-8-sig")
print(f"✓ 已存成 pto_rulemaking_dockets.csv")
print(df[["docket_id", "title"]].to_string())