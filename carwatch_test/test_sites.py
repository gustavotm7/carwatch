"""Can GitHub's servers read the car sites? Fetches a few pages, records what came back."""
import json, re, sys, time, os, importlib.util
sys.path.insert(0, os.path.dirname(__file__))
spec = importlib.util.spec_from_file_location("cc", os.path.join(os.path.dirname(__file__), "car_collector.py"))
cc = importlib.util.module_from_spec(spec); spec.loader.exec_module(cc)
import requests

TESTS = [
    ("autotrader", "search", "https://www.autotrader.ca/cars/nissan/370z/"),
    ("autotrader", "listing", "https://www.autotrader.ca/offers/bmw-m3-cabriolet-gasoline-white-cat_ma13gr202607va361-0c6eb378-89f4-4ff4-81c4-91e3c7bad91d"),
    ("autotrader", "photo", "https://prod.pictures.autoscout24.net/listing-images/0c6eb378-89f4-4ff4-81c4-91e3c7bad91d_48ae7343-a609-4ad2-8cd9-dac1c2fadada.jpg/250x188.jpg"),
    ("kijiji", "search", "https://www.kijiji.ca/b-cars-trucks/canada/nissan-370z/k0c174l0"),
    ("kijiji", "listing", "https://www.kijiji.ca/v-cars-trucks/markham-york-region/2012-bmw-3-series-4dr-touring-wgn-328i-xdrive-awd/m8422021"),
    ("carpages", "search", "https://www.carpages.ca/used-cars/nissan/370z/"),
    ("carpages", "listing", "https://www.carpages.ca/used-cars/british-columbia/surrey/2009-bmw-3-series-12823162"),
]
BLOCK = re.compile(r"captcha|access denied|cf-chl|challenge-platform|are you a human|unusual traffic|request blocked|incapsula|perimeterx|px-captcha|datadome|akamai", re.I)
s = requests.Session()
s.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36",
                  "Accept-Language": "en-CA,en;q=0.9", "Accept": "text/html,application/xhtml+xml,image/avif,image/webp,*/*;q=0.8"})
out = {"runner_ip": None, "results": []}
try:
    out["runner_ip"] = requests.get("https://api.ipify.org", timeout=10).text
except Exception as e:
    out["runner_ip"] = str(e)
for site, kind, url in TESTS:
    r0 = {"site": site, "kind": kind, "url": url}
    try:
        r = s.get(url, timeout=30, allow_redirects=True)
        r0.update(status=r.status_code, final_url=r.url, bytes=len(r.content),
                  content_type=r.headers.get("content-type", ""), server=r.headers.get("server", ""))
        if kind != "photo":
            t = r.text
            m = BLOCK.search(t[:20000])
            r0["block_marker"] = m.group(0) if m else ""
            r0["title"] = (re.search(r"(?is)<title[^>]*>(.*?)</title>", t) or [None, ""])[1].strip()[:120] if re.search(r"(?is)<title", t) else ""
            if kind == "search":
                r0["ad_links_found"] = len({m.group(2) for m in cc.LINK_RE[site].finditer(t)})
            else:
                info = cc.parse_listing(t)
                r0["parsed"] = {k: info.get(k) for k in ("title", "price", "km", "year", "trans_field", "body_raw", "place")}
                r0["parsed"]["desc_chars"] = len(info.get("desc") or "")
                r0["parsed"]["photo"] = cc.photo_of({"source": site, "listing_id": "0c6eb378-89f4-4ff4-81c4-91e3c7bad91d"}, t, info)[:120]
    except Exception as e:
        r0["error"] = f"{type(e).__name__}: {e}"
    out["results"].append(r0)
    print(json.dumps(r0)[:400], flush=True)
    time.sleep(2)
json.dump(out, open(os.path.join(os.path.dirname(__file__), "results.json"), "w"), indent=1)
