"""Open every listed ad with a real request; also look for the 2012 328i xDrive wagon on Kijiji."""
import csv, json, os, re, time, importlib.util
D = os.path.dirname(__file__)
spec = importlib.util.spec_from_file_location("cc", os.path.join(D, "car_collector.py"))
cc = importlib.util.module_from_spec(spec); spec.loader.exec_module(cc)
import requests
s = requests.Session()
s.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36",
                  "Accept-Language": "en-CA,en;q=0.9"})
rows = [r for r in csv.DictReader(open(os.path.join(D, "car_listings_20261118.csv"), encoding="utf-8")) if r["status"] in ("active", "unverified")]
out = {"links": [], "kijiji_search": []}
for r in rows:
    o = {"id": r["listing_id"], "name": r["name"], "year": r["year"], "price": r["price"], "km": r["km"], "status": r["status"], "url": r["url"]}
    try:
        g = s.get(r["url"], timeout=30, allow_redirects=True)
        info = cc.parse_listing(g.text) if g.status_code == 200 else {}
        o.update(http=g.status_code, final=g.url, id_in_final=cc.id_key(r) in g.url,
                 page_price=info.get("price"), page_km=info.get("km"), page_title=(info.get("title") or "")[:80])
    except Exception as e:
        o["error"] = str(e)[:200]
    out["links"].append(o)
    time.sleep(1.5)
for kw in ("bmw-328i-xdrive-touring", "bmw-328xi-wagon", "bmw-wagon-manual", "bmw-3-series-wagon", "bmw-328i-touring"):
    url = f"https://www.kijiji.ca/b-cars-trucks/canada/{kw}/k0c174l0"
    g = s.get(url, timeout=30)
    links = list(dict.fromkeys(cc.HOSTS["kijiji"] + m.group(1) for m in cc.LINK_RE["kijiji"].finditer(g.text)))
    out["kijiji_search"].append({"search": url, "http": g.status_code, "n": len(links)})
    for u in links:
        if not re.search(r"328|wagon|touring|wgn", u):
            continue
        a = s.get(u, timeout=30); info = cc.parse_listing(a.text)
        out["kijiji_search"].append({"url": u, "http": a.status_code, "title": info.get("title"), "price": info.get("price"),
                                     "km": info.get("km"), "year": info.get("year"), "trans": info.get("trans_field"),
                                     "place": info.get("place")})
        time.sleep(1.5)
    time.sleep(1.5)
json.dump(out, open(os.path.join(D, "links.json"), "w"), indent=1)
