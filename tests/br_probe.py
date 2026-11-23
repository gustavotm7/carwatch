"""One-off (3): Mobiauto filter names (gearbox), NaPista paging."""
import json, os, re, time, requests
D = os.path.dirname(__file__)
s = requests.Session()
s.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) personal-car-watch/2.0", "Accept-Language": "pt-BR,pt;q=0.9"})
def pp(h):
    m = re.search(r'(?is)<script[^>]+id="__NEXT_DATA__"[^>]*>(.*?)</script>', h)
    try: return json.loads(m.group(1))["props"]["pageProps"] if m else {}
    except Exception: return {}
out = {"aggs": None, "tries": []}
MB = "https://www.mobiauto.com.br/comprar/carros-usados/brasil/honda/civic"
p = pp(s.get(MB, timeout=30).text)
aggs = (p.get("deals") or {}).get("aggregations") or {}
out["aggs"] = {k: (json.dumps(v, ensure_ascii=False)[:700]) for k, v in aggs.items()}
out["hrefs"] = sorted(set(re.findall(r'href="(/comprar/carros-usados/brasil/honda/civic/[^"]+)"', s.get(MB, timeout=30).text)))[:60]
def t(u):
    r = s.get(u, timeout=30); q = pp(r.text)
    if "searchResult" in q:
        offs = q["searchResult"].get("offers") or []
        return {"url": u, "final": r.url[-70:], "n": len(offs), "ids": [o["id"][:8] for o in offs[:4]], "total": q["searchResult"].get("offersTotal")}
    d = q.get("deals") or {}
    res = d.get("results") or []
    return {"url": u, "final": r.url[-80:], "params": (q.get("query") or {}).get("params"), "n": len(res), "num": d.get("numResults"),
            "trans": [((x.get("trim") or {}).get("transmission") or {}).get("name") for x in res[:10]], "ids": [x.get("id") for x in res[:3]]}
for u in [MB + "/manual", MB + "/transmissao-manual", MB + "/cambio-manual", MB + "?trs=manual", MB + "?trs=2", MB + "/ano-2010/manual",
          "https://napista.com.br/busca/honda/civic/manual", "https://napista.com.br/busca/honda/civic/manual?pn=2", "https://napista.com.br/busca/honda/civic/manual?pagina=3"]:
    try: out["tries"].append(t(u))
    except Exception as e: out["tries"].append({"url": u, "error": str(e)[:100]})
    time.sleep(1.5)
json.dump(out, open(os.path.join(D, "br_probe.json"), "w"), indent=1, ensure_ascii=False)
