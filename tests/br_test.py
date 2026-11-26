"""One-off: which model slugs do NaPista and Mobiauto use, and do the empty searches really have no cars?"""
import json, os, re, sys, time, requests
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "collector"))
import car_collector as cc
NP, MB = "https://napista.com.br/busca/", "https://www.mobiauto.com.br/comprar/carros-usados/brasil/"
BRANDS = ["honda", "mazda", "toyota", "subaru", "nissan", "mitsubishi", "bmw", "audi"]
PAGES = [NP + b for b in BRANDS] + [MB + b for b in BRANDS]
TRY = [MB + "ford/escort/ano-1985", MB + "ford/escort/ano-1991", MB + "chevrolet/kadett/ano-1992", MB + "chevrolet/vectra/ano-1994",
       MB + "fiat/uno/ano-1995", MB + "volkswagen/gol/ano-1987", NP + "toyota/86", NP + "toyota/gt-86", NP + "mazda/mx5",
       NP + "mazda/miata", MB + "mazda/mx5", MB + "subaru/impreza", NP + "toyota/yaris/2018/manual", MB + "toyota/yaris/ano-2018"]
s = requests.Session(); s.headers.update({"User-Agent": cc.UA, "Accept-Language": "pt-BR,pt;q=0.9"})
out = {"slugs": {}, "tries": []}
def pp_of(html):
    m = re.search(r'(?is)<script[^>]+id=["\']__NEXT_DATA__["\'][^>]*>(.*?)</script>', html)
    return json.loads(m.group(1))["props"]["pageProps"] if m else {}
for u in PAGES:
    try:
        r = s.get(u, timeout=30); pp = pp_of(r.text); found = set()
        for d in cc._walk(pp):
            sl = d.get("slug") or d.get("urlName") or d.get("seoName") or d.get("value")
            nm = d.get("name") or d.get("label") or d.get("description")
            if isinstance(sl, str) and isinstance(nm, str) and len(sl) < 40:
                found.add(f'{nm}={sl}' + (f'({d.get("count") or d.get("total") or d.get("quantity")})' if (d.get("count") or d.get("total") or d.get("quantity")) else ""))
        out["slugs"][u] = {"status": r.status_code, "keys": list(pp)[:30], "found": sorted(found)[:400]}
    except Exception as e:
        out["slugs"][u] = {"error": str(e)[:200]}
    time.sleep(3)
for u in TRY:
    try:
        r = s.get(u, timeout=30); src = "napista" if "napista" in u else "mobiauto"
        it = cc.br_search_items(src, r.text)
        out["tries"].append({"url": u, "status": r.status_code, "final": r.url, "n": len(it), "sample": [(x["title"], x["gear"]) for x in it[:6]]})
    except Exception as e:
        out["tries"].append({"url": u, "error": str(e)[:200]})
    time.sleep(3)
json.dump(out, open(os.path.join(os.path.dirname(__file__), "br_results.json"), "w"), indent=1, ensure_ascii=False)
