import json, os, re, time, requests
D = os.path.dirname(__file__)
s = requests.Session()
s.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36"})
API = "https://bringatrailer.com/wp-json/bringatrailer/1.0/data/listings-filter"
out = {"pages": [], "items": [], "listing_links": [], "checks": []}
for u in ["https://bringatrailer.com/bmw/m3/", "https://bringatrailer.com/bmw/3-series/", "https://bringatrailer.com/bmw/e46-m3/", "https://bringatrailer.com/bmw/"]:
    h = s.get(u, timeout=30).text
    m = re.search(r'auctionsCompletedInitialData\s*=\s*(\{.*?\});\s*\n', h, re.S)
    o = {"url": u, "bmw_links": sorted(set(re.findall(r'https://bringatrailer\.com/bmw/[a-z0-9\-]+/', h)))[:200],
         "filter_snips": re.findall(r'.{0,80}keyword_pages.{0,80}', h)[:6]}
    if m:
        d = json.loads(m.group(1))
        o["base_filter"] = d.get("base_filter"); o["keys"] = [k for k in d.keys() if k != "items"]
        for it in d.get("items", []):
            out["items"].append({"from": u, "title": it["title"], "kw": it.get("keyword_pages"), "url": it["url"]})
    out["pages"].append(o); time.sleep(2)
cands = [i for i in out["items"] if re.search(r"20(0[7-9]|1[0-3]) BMW M3|330C?i|Wagon|E9[0-3]", i["title"])][:4]
for c in cands:
    h = s.get(c["url"], timeout=30).text
    out["listing_links"].append({"title": c["title"], "links": sorted(set(re.findall(r'https://bringatrailer\.com/(?:bmw|nissan)/[a-z0-9\-]+/', h))),
                                 "kw_snips": re.findall(r'.{0,60}keyword_pages.{0,60}', h)[:3]})
    time.sleep(2)
json.dump(out, open(os.path.join(D, "inspect6.json"), "w"), indent=1)
