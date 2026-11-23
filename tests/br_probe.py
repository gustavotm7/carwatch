"""One-off (2): ad records, photo addresses, paging, gearbox/year filters, removed ads, model address names."""
import json, os, re, time, requests
D = os.path.dirname(__file__)
s = requests.Session()
s.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) personal-car-watch/2.0", "Accept-Language": "pt-BR,pt;q=0.9"})
out = {"slugs": [], "napista_ad": None, "mobi_ad": None, "napista_meta": None, "mobi_sop": None, "gone": [], "img": []}
def nd(h):
    m = re.search(r'(?is)<script[^>]+id="__NEXT_DATA__"[^>]*>(.*?)</script>', h)
    try: return json.loads(m.group(1))["props"]["pageProps"] if m else {}
    except Exception: return {}
def trim(o, depth=0):
    if isinstance(o, dict): return {k: trim(v, depth + 1) for k, v in list(o.items())[:60]}
    if isinstance(o, list): return [trim(v, depth + 1) for v in o[:3]] + (["...%d" % len(o)] if len(o) > 3 else [])
    if isinstance(o, str): return o[:300]
    return o
def get(u):
    r = s.get(u, timeout=30); time.sleep(1.5); return r
def summary(u):
    r = get(u); pp = nd(r.text)
    if "searchResult" in pp:
        offers = pp["searchResult"].get("offers") or []
        return {"url": u, "status": r.status_code, "final": r.url[:120], "n": len(offers),
                "first": [(o.get("versionName"), o.get("modelYear"), o.get("price")) for o in offers[:4]],
                "ids": [o.get("id")[:8] for o in offers[:3]]}
    deals = (pp.get("deals") or {}).get("results") or []
    return {"url": u, "status": r.status_code, "final": r.url[:120], "n": len(deals), "params": (pp.get("query") or {}).get("params"),
            "first": [((d.get("trim") or {}).get("name"), (d.get("trim") or {}).get("productionYear"), d.get("price"),
                       ((d.get("trim") or {}).get("transmission") or {}).get("name")) for d in deals[:4]],
            "ids": [d.get("id") for d in deals[:3]], "total": (pp.get("deals") or {}).get("total") or (pp.get("deals") or {}).get("count")}
NP = "https://napista.com.br/busca/"
MB = "https://www.mobiauto.com.br/comprar/carros-usados/brasil/"
for u in [NP + "bmw/series-3-m3", NP + "bmw/m3", NP + "bmw/series-3", NP + "bmw/series-3/manual", NP + "bmw/series-5", NP + "bmw/series-5-m5",
          NP + "honda/civic/manual", NP + "honda/civic/manual/2010", NP + "honda/civic/2010", NP + "toyota/yaris/manual", NP + "nissan/350z", NP + "nissan/370z",
          NP + "honda/civic/manual?page=2", NP + "honda/civic/manual?p=2",
          MB + "bmw/m3", MB + "bmw/m3/pagina-2", MB + "bmw/m5", MB + "bmw/serie-3", MB + "bmw/series-3", MB + "bmw/330i", MB + "honda/civic",
          MB + "honda/civic/ano-2010", MB + "honda/civic/pagina-2", MB + "toyota/yaris", MB + "nissan/350z", MB + "nissan/370z"]:
    try: out["slugs"].append(summary(u))
    except Exception as e: out["slugs"].append({"url": u, "error": str(e)[:120]})
# NaPista: search meta + full ad record
r = get(NP + "honda/civic/manual"); pp = nd(r.text)
sr = pp.get("searchResult", {})
out["napista_meta"] = {"sr_keys": {k: (v if not isinstance(v, (list, dict)) else type(v).__name__ + str(len(v))) for k, v in sr.items()},
                       "searchParams": trim(pp.get("searchParams"))}
ad_id = (sr.get("offers") or [{}])[0].get("id")
if ad_id:
    r = get("https://napista.com.br/anuncios/" + ad_id); off = nd(r.text).get("offer", {})
    out["napista_ad"] = trim({k: v for k, v in off.items() if k not in ("similarOffers", "sellerOffers")})
# Mobiauto: search params object + full ad record + photo address
r = get(MB + "honda/civic"); pp = nd(r.text); h = r.text
out["mobi_sop"] = {"requestSop": trim(pp.get("requestSop")), "sop": trim(pp.get("sop")), "canonical": pp.get("canonicalPath"),
                   "deals_keys": {k: (v if not isinstance(v, (list, dict)) else type(v).__name__) for k, v in (pp.get("deals") or {}).items()}}
deal = ((pp.get("deals") or {}).get("results") or [{}])[0]
img_id = str(((deal.get("images") or [{}])[0]).get("imageId", ""))
out["img"] = list(dict.fromkeys(re.findall(r'https?://[^"\s\\]*' + img_id + r'[^"\s\\]*', h)))[:5] if img_id else []
link = re.search(r'href="(https://www\.mobiauto\.com\.br/comprar/carros/[^"]+/detalhes/' + str(deal.get("id")) + r')', h)
if link:
    r = get(link.group(1)); pa = nd(r.text)
    out["mobi_ad"] = {"url": link.group(1), "deal": trim(pa.get("deal")), "images": trim(pa.get("images")),
                      "img_urls": list(dict.fromkeys(re.findall(r'https?://[^"\s\\]*' + img_id + r'[^"\s\\]*', r.text)))[:5],
                      "og_image": (re.search(r'og:image"\s+content="([^"]+)', r.text) or [0, ""])[1]}
# removed ads
for u in ["https://napista.com.br/anuncios/00000000-0000-4000-8000-000000000000",
          "https://www.mobiauto.com.br/comprar/carros/sp-sao-paulo/honda/civic/2011/new-lxl-1-8-i-vtec-couro-aut-flex/detalhes/1000001"]:
    r = s.get(u, timeout=30, allow_redirects=True)
    out["gone"].append({"url": u, "status": r.status_code, "final": r.url[:140], "hist": [x.status_code for x in r.history],
                        "title": (re.search(r"(?is)<title[^>]*>(.*?)</title>", r.text) or [0, ""])[1][:90], "pp_keys": list(nd(r.text).keys())[:12]})
json.dump(out, open(os.path.join(D, "br_probe.json"), "w"), indent=1, ensure_ascii=False)
