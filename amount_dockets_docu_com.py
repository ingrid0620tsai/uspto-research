import requests

API_KEY = "og1wltTsDKFQB8tTUdMMFS4akawWzkULBrt18hDg"
BASE = "https://api.regulations.gov/v4"

for endpoint in ["dockets", "documents", "comments"]:
    r = requests.get(
        f"{BASE}/{endpoint}",
        params={
            "filter[agencyId]": "PTO",
            "page[size]": 5,
            "page[number]": 1,
            "api_key": API_KEY
        }
    )
    data = r.json()
    total = data.get("meta", {}).get("totalElements", "?")
    print(f"PTO {endpoint}：{total} 筆")
    
    # 如果還是 ?，印出完整回應幫助診斷
    if total == "?":
        print(f"  完整 meta：{data.get('meta')}")
        print(f"  HTTP 狀態：{r.status_code}")