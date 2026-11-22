"""One-off: listing dates, VIN/CARFAX fields, auction-result sources, FX rate."""
import json, os, re, time, requests
D = os.path.dirname(__file__)
cars = json.load(open(os.path.join(D, "..", "site", "data", "cars.json")))["cars"]
s = requests.Session()
s.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36",
                  "Accept-Language": "en-CA,en;q=0.9"})
KEYS = re.compile(r"date|time|created|publish|since|online|posted|activat|vin|carfax|history|accident|report|badge|seller|dealer", re.I)
def walk(o, p=""):
    if isinstance(o, dict):
        for k, v in o.items(): yield from walk(v, f"{p}.{k}")
    elif isinstance(o, list):
        for i, v in enumerate(o[:3]): yield from walk(v, f"{p}[{i}]")
    else: yield p, o
out = {"autotrader": [], "kijiji": [], "auctions": [], "fx": None, "vin": []}
at = [c for c in cars if c["source"] == "autotrader"][:4]
for c in at:
    r = s.get(c["url"], timeout=30); h = r.text
    m = re.search(r'(?is)<script[^>]+id=["\']__NEXT_DATA__["\'][^>]*>(.*?)</script>', h)
    ld = json.loads(m.group(1))["props"]["pageProps"].get("listingDetails") if m else {}
    hits = [(p, str(v)[:120]) for p, v in walk(ld) if KEYS.search(p.split(".")[-1]) and v not in (None, "", [], {})]
    out["autotrader"].append({"url": c["url"][-50:], "http": r.status_code, "hits": hits[:60],
                              "carfax_links": list(set(re.findall(r'https?://[^"\'\s<>]*carfax[^"\'\s<>]*', h)))[:5]})
    time.sleep(1.5)
kj = [c for c in cars if c["source"] == "kijiji"][:2]
for c in kj:
    r = s.get(c["url"], timeout=30); h = r.text
    lds = re.findall(r'(?is)<script[^>]+application/ld\+json[^>]*>(.*?)</script>', h)
    dates = re.findall(r'"(datePosted|dateCreated|activationDate|sortingDate|postedDate|creationDate|createdAt|startDate)"\s*:\s*"?([^",}]{6,40})', h)[:10]
    vins = re.findall(r'"(?:vin|vehicleIdentificationNumber)"\s*:\s*"([A-HJ-NPR-Z0-9]{17})"', h)[:3]
    out["kijiji"].append({"url": c["url"][-50:], "http": r.status_code, "dates": dates, "vins": vins,
                          "carfax_links": list(set(re.findall(r'https?://[^"\'\s<>]*carfax[^"\'\s<>]*', h)))[:5],
                          "ld_keys": [list(json.loads(x).keys())[:30] for x in lds if x.strip().startswith("{")][:2]})
    time.sleep(1.5)
BAT = ["https://bringatrailer.com/bmw/e46-m3/", "https://bringatrailer.com/bmw/e92-m3/", "https://bringatrailer.com/bmw/e90-m3/",
       "https://bringatrailer.com/bmw/e39-m5/", "https://bringatrailer.com/bmw/e60-m5/", "https://bringatrailer.com/bmw/e46-3-series/",
       "https://bringatrailer.com/bmw/e91-3-series/", "https://bringatrailer.com/bmw/e90-3-series/",
       "https://bringatrailer.com/nissan/350z/", "https://bringatrailer.com/nissan/370z/"]
for u in BAT:
    o = {"url": u}
    try:
        r = s.get(u, timeout=30); h = r.text
        m = re.search(r'auctionsCompletedInitialData\s*=\s*(\{.*?\});\s*\n', h, re.S)
        o.update(http=r.status_code, final=r.url, bytes=len(h), has_completed=bool(m),
                 kw_ids=re.findall(r'keyword_pages[^0-9]{0,20}(\d+)', h)[:3], title=(re.search(r"<title>(.*?)</title>", h, re.S) or [0, ""])[1][:80])
        if m:
            d = json.loads(m.group(1))
            items = d.get("items", [])
            o["n_items"], o["total"] = len(items), d.get("items_total") or d.get("total")
            o["sample"] = [{k: it.get(k) for k in ("title", "current_bid", "sold_text", "timestamp_end", "url", "country", "year")} for it in items[:3]]
            o["item_keys"] = list(items[0].keys())[:40] if items else []
    except Exception as e:
        o["error"] = str(e)[:150]
    out["auctions"].append(o); time.sleep(2)
for u in ["https://carsandbids.com/search?q=e46%20m3", "https://carsandbids.com/v2/autos/auctions?limit=20&status=closed&search=e46%20m3",
          "https://bringatrailer.com/wp-json/bringatrailer/1.0/data/listings-filter?page=1&per_page=36&get_items=1&get_stats=0&sort=td&state=sold&keyword=e46%20m3"]:
    try:
        r = s.get(u, timeout=30)
        out["auctions"].append({"url": u, "http": r.status_code, "bytes": len(r.content), "start": r.text[:300]})
    except Exception as e:
        out["auctions"].append({"url": u, "error": str(e)[:150]})
    time.sleep(2)
try:
    r = s.get("https://www.bankofcanada.ca/valet/observations/FXUSDCAD/json?recent=1", timeout=30)
    out["fx"] = {"http": r.status_code, "obs": r.json().get("observations")}
except Exception as e:
    out["fx"] = {"error": str(e)[:150]}
json.dump(out, open(os.path.join(D, "inspect2.json"), "w"), indent=1, default=str)
