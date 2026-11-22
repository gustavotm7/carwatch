import json, os, re, time, requests
D = os.path.dirname(__file__)
s = requests.Session()
s.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36"})
API = "https://bringatrailer.com/wp-json/bringatrailer/1.0/data/listings-filter"
out = []
def summ(r):
    try:
        d = r.json()
        return {"total": d.get("items_total"), "n": len(d.get("items", [])), "titles": [i.get("title") for i in d.get("items", [])[:6]],
                "keys": list(d["items"][0].keys()) if d.get("items") else []}
    except Exception as e:
        return {"err": str(e)[:80], "start": r.text[:150]}
tests = [
  ("page", "https://bringatrailer.com/bmw/e9x-m3/"), ("page", "https://bringatrailer.com/bmw/e90-e92-e93-m3/"),
  ("page", "https://bringatrailer.com/bmw/m3/"), ("page", "https://bringatrailer.com/bmw/3-series/"),
  ("page", "https://bringatrailer.com/bmw/e46-330ci/"), ("page", "https://bringatrailer.com/bmw/e91/"),
  ("page", "https://bringatrailer.com/auctions/results/?search=e92+m3"),
  ("api", API + "?page=1&per_page=12&get_items=1&get_stats=0&sort=td&state=sold&base_filter[keyword_pages][]=1833234"),
  ("api", API + "?page=1&per_page=12&get_items=1&get_stats=0&sort=td&state=sold&base_filter[keyword_s]=e92%20m3"),
  ("api", API + "?page=1&per_page=12&get_items=1&get_stats=0&sort=td&state=sold&search=e92%20m3"),
  ("api", API + "?page=1&per_page=12&get_items=1&get_stats=0&sort=td&state=sold&s=e92%20m3"),
  ("api", API + "?page=2&per_page=60&get_items=1&get_stats=0&sort=td&state=sold&base_filter[keyword_pages][]=1833234"),
]
for kind, u in tests:
    o = {"kind": kind, "url": u}
    try:
        r = s.get(u, timeout=30); o["http"] = r.status_code; o["final"] = r.url
        if kind == "api":
            o.update(summ(r))
        else:
            h = r.text
            m = re.search(r'auctionsCompletedInitialData\s*=\s*(\{.*?\});\s*\n', h, re.S)
            o["title"] = (re.search(r"<title>(.*?)</title>", h, re.S) or [0, ""])[1][:80]
            o["kw"] = re.findall(r'keyword_pages[^0-9]{0,20}(\d+)', h)[:3]
            o["links"] = sorted(set(re.findall(r'https://bringatrailer\.com/bmw/[a-z0-9\-]+/', h)))[:80]
            if m:
                d = json.loads(m.group(1)); o["initial_total"] = d.get("items_total"); o["initial_titles"] = [i["title"] for i in d.get("items", [])[:6]]
    except Exception as e:
        o["error"] = str(e)[:120]
    out.append(o); time.sleep(2)
json.dump(out, open(os.path.join(D, "inspect3.json"), "w"), indent=1)
