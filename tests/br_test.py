"""One-off: does Autotrader.ca take an automatic-gearbox filter in the path?"""
import json, os, sys, time, requests
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "collector"))
import car_collector as cc
A = "https://www.autotrader.ca/cars/"
TRY = [A + "honda/civic/tr_manual?fregfrom=2012&fregto=2015", A + "honda/civic/tr_automatic?fregfrom=2012&fregto=2015",
       A + "honda/civic/tr_auto?fregfrom=2012&fregto=2015", A + "honda/civic?fregfrom=2012&fregto=2015",
       A + "toyota/yaris/tr_automatic", A + "toyota/yaris"]
s = requests.Session(); s.headers.update({"User-Agent": cc.UA, "Accept-Language": "en-CA,en;q=0.8"})
out = []
for u in TRY:
    try:
        r = s.get(u, timeout=30)
        ids = {m.group(2) for m in cc.LINK_RE["autotrader"].finditer(r.text)}
        out.append({"url": u, "status": r.status_code, "final": r.url, "ads": len(ids),
                    "manual_words": r.text.lower().count("manual"), "auto_words": r.text.lower().count("automatic")})
    except Exception as e:
        out.append({"url": u, "error": str(e)[:200]})
    time.sleep(4)
json.dump(out, open(os.path.join(os.path.dirname(__file__), "br_results.json"), "w"), indent=1)
