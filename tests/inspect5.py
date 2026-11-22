import json, os, re, time, requests
D = os.path.dirname(__file__)
s = requests.Session()
s.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36"})
API = "https://bringatrailer.com/wp-json/bringatrailer/1.0/data/listings-filter"
out = {"dbg": [], "kw": {}, "samples": []}
def get(pid, page, per=60):
    u = f"{API}?page={page}&per_page={per}&get_items=1&get_stats=0&sort=td&state=sold&base_filter[keyword_pages][]={pid}"
    r = s.get(u, timeout=30)
    try: d = r.json()
    except Exception: d = {}
    if not d.get("items"): out["dbg"].append({"u": u, "http": r.status_code, "start": r.text[:200]})
    return d
for per in (36, 60, 100):
    d = get(1833018, 1, per); out["dbg"].append({"per": per, "n": len(d.get("items", [])), "total": d.get("items_total")}); time.sleep(1.5)
want = {"E9X": (1833018, r"\bM3\b", 2008, 2013), "330": (20556842, r"330C?i", 2001, 2006), "E91": (20556842, r"Wagon", 2006, 2012)}
for tag, (pid, rx, y0, y1) in want.items():
    ids = {}
    for p in range(1, 6):
        for it in get(pid, p).get("items", []):
            m = re.search(r"\b(19|20)\d{2}\b", it["title"]); y = int(m.group(0)) if m else None
            if y and y0 <= y <= y1 and re.search(rx, it["title"]):
                for k in it.get("keyword_pages") or []: ids[str(k)] = ids.get(str(k), 0) + 1
                out["samples"].append({"tag": tag, "title": it["title"], "kw": it.get("keyword_pages"), "sold": it["sold_text"][:60],
                                       "excerpt": (it.get("excerpt") or "")[:300], "cc": it.get("country_code"), "cur": it.get("currency")})
        time.sleep(1.5)
    out["kw"][tag] = sorted(ids.items(), key=lambda x: -x[1])[:12]
json.dump(out, open(os.path.join(D, "inspect5.json"), "w"), indent=1)
