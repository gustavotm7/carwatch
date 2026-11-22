import json, os, re, time, requests
D = os.path.dirname(__file__)
s = requests.Session()
s.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36"})
API = "https://bringatrailer.com/wp-json/bringatrailer/1.0/data/listings-filter"
ids = [20556842, 63227464, 1833027, 63227594, 12875122, 1833018, 25108927, 1833019, 25077613, 25078037, 25078204, 25033702, 25077848, 12870552,
       1833234, 25109108, 2904138, 1833028, 25109133, 51981770, 25109065, 80784901, 25553604, 81159716, 49155935]
out = {"ids": [], "pages": [], "links": []}
for k in ids:
    r = s.get(f"{API}?page=1&per_page=12&get_items=1&get_stats=0&sort=td&state=sold&base_filter[keyword_pages][]={k}", timeout=30)
    d = r.json()
    out["ids"].append([k, d.get("items_total"), [i["title"] for i in d.get("items", [])],
                       [(i.get("excerpt") or "")[:240] for i in d.get("items", [])[:2]]])
    time.sleep(1.2)
for u in ["https://bringatrailer.com/bmw/e90-e92-m3/"]:
    h = s.get(u, timeout=30).text
    m = re.search(r'auctionsCompletedInitialData\s*=\s*(\{.*?\});\s*\n', h, re.S)
    out["pages"].append({"url": u, "base_filter": json.loads(m.group(1)).get("base_filter") if m else None, "total": json.loads(m.group(1)).get("items_total") if m else None})
out["note"]="per12"
json.dump(out, open(os.path.join(D, "inspect7.json"), "w"), indent=1)
