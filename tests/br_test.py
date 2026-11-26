"""One-off: do the batch-4 search pages exist and what do they hold?"""
import json, os, sys, time, requests
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "collector"))
import car_collector as cc
A, NP, MB = "https://www.autotrader.ca/cars/", "https://napista.com.br/busca/", "https://www.mobiauto.com.br/comprar/carros-usados/brasil/"
TRY = [A + s for s in ["toyota/gr-corolla", "toyota/gr-corolla/tr_manual", "toyota/corolla/tr_manual?fregfrom=2023", "acura/integra/tr_manual?fregfrom=2023",
                       "acura/tsx/tr_manual", "dodge/challenger/tr_manual", "audi/s4/tr_manual", "mitsubishi/eclipse/tr_manual?fregfrom=1995&fregto=1999",
                       "lotus/elise", "lotus/exige"]] + \
      [NP + s for s in ["ford/maverick/1974", "dodge/charger", "dodge/dart", "volkswagen/sp2", "chevrolet/monza", "chevrolet/corsa/1995", "fiat/bravo"]] + \
      [MB + s for s in ["ford/maverick", "dodge/charger", "volkswagen/sp2", "chevrolet/monza", "fiat/bravo/s/cambio-manual"]]
s = requests.Session(); s.headers.update({"User-Agent": cc.UA, "Accept-Language": "pt-BR,pt;q=0.9,en;q=0.8"})
out = []
for u in TRY:
    try:
        r = s.get(u, timeout=30)
        if "autotrader" in u:
            ids = {m.group(2) for m in cc.LINK_RE["autotrader"].finditer(r.text)}
            titles = []
            import re
            titles = re.findall(r'"(?:name|title)":"([^"]{8,80})"', r.text)[:8]
            out.append({"url": u, "status": r.status_code, "final": r.url, "n": len(ids), "titles": titles})
        else:
            it = cc.br_search_items("napista" if "napista" in u else "mobiauto", r.text)
            out.append({"url": u, "status": r.status_code, "final": r.url, "n": len(it), "sample": [(x["title"], x["gear"]) for x in it[:8]]})
    except Exception as e:
        out.append({"url": u, "error": str(e)[:200]})
    time.sleep(3)
json.dump(out, open(os.path.join(os.path.dirname(__file__), "br_results.json"), "w"), indent=1, ensure_ascii=False)
