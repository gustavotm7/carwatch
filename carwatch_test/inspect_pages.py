"""Where do the description, year and photo live in Autotrader and Kijiji ad pages?"""
import json, os, re, requests
D = os.path.dirname(__file__)
s = requests.Session()
s.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36", "Accept-Language": "en-CA,en;q=0.9"})
PAGES = {
  "at": "https://www.autotrader.ca/offers/bmw-m3-cabriolet-gasoline-white-cat_ma13gr202607va361-0c6eb378-89f4-4ff4-81c4-91e3c7bad91d",
  "at_redirect": "https://www.autotrader.ca/offers/bmw-m3-6-speed-manual-s54-engine-clean-title-gasoline-black-cat_ma13gr202607va361-09a3f0f8-9b5a-48c0-b7a0-60f77a5bc0d6",
  "kj": "https://www.kijiji.ca/v-cars-trucks/laval-rive-nord/2017-nissan-370z-touring/1743340345",
}
out = {}
def walk(o, path=""):
    if isinstance(o, dict):
        for k, v in o.items(): yield from walk(v, f"{path}.{k}")
    elif isinstance(o, list):
        for i, v in enumerate(o[:50]): yield from walk(v, f"{path}[{i}]")
    else:
        yield path, o
for k, url in PAGES.items():
    r = s.get(url, timeout=30)
    h = r.text
    o = {"http": r.status_code, "final": r.url, "history": [x.status_code for x in r.history],
         "title_tag": (re.search(r"(?is)<title[^>]*>(.*?)</title>", h) or [0, ""])[1][:150],
         "og_image": re.findall(r'<meta[^>]+(?:property|name)=["\']og:image["\'][^>]*>', h)[:2],
         "ldjson": [m[:1500] for m in re.findall(r'(?is)<script[^>]+application/ld\+json[^>]*>(.*?)</script>', h)][:3]}
    m = re.search(r'(?is)<script[^>]+id=["\']__NEXT_DATA__["\'][^>]*>(.*?)</script>', h)
    if m:
        data = json.loads(m.group(1))
        longs = [(p, v[:160]) for p, v in walk(data) if isinstance(v, str) and len(v) > 150 and not v.startswith(("http", "data:", "{"))]
        imgs = [(p, v) for p, v in walk(data) if isinstance(v, str) and re.search(r"\.(jpe?g|png|webp)|listing-images|/images/", v, re.I)][:8]
        yrs = [(p, v) for p, v in walk(data) if re.search(r"year|modelDate", p, re.I)][:8]
        trans = [(p, v) for p, v in walk(data) if re.search(r"transmission|gearbox", p, re.I)][:6]
        body = [(p, v) for p, v in walk(data) if re.search(r"body", p, re.I) and isinstance(v, str) and len(v) < 40][:6]
        o.update(next_long_strings=longs[:12], next_images=imgs, next_years=yrs, next_trans=trans, next_body=body)
    else:
        o["no_next_data"] = True
        o["script_ids"] = re.findall(r'<script[^>]+id=["\']([^"\']+)', h)[:20]
    out[k] = o
json.dump(out, open(os.path.join(D, "inspect.json"), "w"), indent=1, default=str)
