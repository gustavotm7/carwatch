import json, os, re, time, requests
D = os.path.dirname(__file__)
s = requests.Session()
s.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36"})
API = "https://bringatrailer.com/wp-json/bringatrailer/1.0/data/listings-filter"
out = {"kw": {}, "samples": [], "per_page": {}, "vin": [], "carfax": [], "decoders": []}
def get(page_id, page, per=60):
    r = s.get(API, params={"page": page, "per_page": per, "get_items": 1, "get_stats": 0, "sort": "td", "state": "sold",
                           "base_filter[keyword_pages][]": page_id}, timeout=30)
    return r.json()
for per in (100, 200):
    try:
        d = get(1833018, 1, per); out["per_page"][per] = len(d.get("items", []))
    except Exception as e: out["per_page"][per] = str(e)[:80]
    time.sleep(1.5)
want = {"E9X": (1833018, r"\bM3\b", 2008, 2013), "330": (20556842, r"330C?i", 2001, 2006), "E91": (20556842, r"Wagon", 2006, 2012)}
for tag, (pid, rx, y0, y1) in want.items():
    ids = {}
    for p in range(1, 5):
        for it in get(pid, p).get("items", []):
            m = re.search(r"\b(19|20)\d{2}\b", it["title"])
            y = int(m.group(0)) if m else None
            if y and y0 <= y <= y1 and re.search(rx, it["title"]):
                for k in it.get("keyword_pages") or []: ids[k] = ids.get(k, 0) + 1
                if len(out["samples"]) < 40:
                    out["samples"].append({"tag": tag, "title": it["title"], "kw": it.get("keyword_pages"), "sold": it["sold_text"][:60],
                                           "excerpt": (it.get("excerpt") or "")[:260], "cc": it.get("country_code"), "cur": it.get("currency")})
        time.sleep(1.5)
    out["kw"][tag] = sorted(ids.items(), key=lambda x: -x[1])[:12]
# test candidate ids by fetching 1 item
for tag, lst in out["kw"].items():
    for k, n in lst[:6]:
        try:
            d = get(k, 1, 3)
            out["kw"].setdefault(tag + "_check", []).append([k, d.get("items_total"), [i["title"] for i in d.get("items", [])]])
        except Exception as e: pass
        time.sleep(1)
# VIN decoders
for vin in ["WBSBR93482PK00421", "WBSKG9C54BE367661", "WBSVA93508E041783", "JN1AZ34D54M100000"]:
    try:
        r = s.get(f"https://vpic.nhtsa.dot.gov/api/vehicles/DecodeVinValues/{vin}?format=json", timeout=30)
        row = r.json()["Results"][0]
        out["vin"].append({k: row.get(k) for k in ("VIN", "ModelYear", "Make", "Model", "Series", "Trim", "BodyClass", "TransmissionStyle",
                                                   "TransmissionSpeeds", "DisplacementL", "PlantCountry", "ErrorCode", "ErrorText")})
    except Exception as e: out["vin"].append({"vin": vin, "error": str(e)[:100]})
for u in ["https://www.mdecoder.com/decode/WBSBR93482PK00421", "https://bimmer.work/vin/WBSBR93482PK00421",
          "https://www.bmwvin.com/decode/WBSBR93482PK00421", "https://vin.bimmerprofs.com/WBSBR93482PK00421",
          "https://www.elcats.ru/bmw/default.aspx?vin=PK00421"]:
    try:
        r = s.get(u, timeout=30); h = r.text
        out["decoders"].append({"url": u, "http": r.status_code, "final": r.url, "bytes": len(h),
                                "has_smg": bool(re.search(r"SMG|2MK|\b206\b", h)), "has_6mt": bool(re.search(r"manual", h, re.I)),
                                "title": (re.search(r"<title>(.*?)</title>", h, re.S) or [0, ""])[1][:80], "snip": re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", h))[:300]})
    except Exception as e: out["decoders"].append({"url": u, "error": str(e)[:100]})
    time.sleep(1)
for u in ["https://vhr.carfax.ca/?id=x+V4SM289x14PxSiHmF05yBBr8cTVq3l", "https://vhr.carfax.ca/?id=G8nnyI+6hxzH/ywWsH8EscsuN0KB/bua"]:
    try:
        r = s.get(u, timeout=30); h = r.text
        txt = re.sub(r"\s+", " ", re.sub(r"(?s)<script.*?</script>|<style.*?</style>|<[^>]+>", " ", h))
        out["carfax"].append({"url": u, "http": r.status_code, "final": r.url, "bytes": len(h), "title": (re.search(r"<title>(.*?)</title>", h, re.S) or [0, ""])[1][:80],
                              "accident_words": re.findall(r"(?i)[^.]{0,60}(accident|damage|claim|stolen|registration|odometer|service record)[^.]{0,60}", txt)[:12],
                              "json_keys": re.findall(r'"(accident\w*|damage\w*|claim\w*|stolen\w*|odometer\w*|serviceRecord\w*|registration\w*|recall\w*)"', h)[:30],
                              "text": txt[:600]})
    except Exception as e: out["carfax"].append({"url": u, "error": str(e)[:100]})
json.dump(out, open(os.path.join(D, "inspect4.json"), "w"), indent=1)
