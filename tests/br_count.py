"""One-off: how many FB Civic and Yaris sedan ads (any gearbox) each site has, and Mobiauto gearbox/body filter names."""
import json, os, re, time, requests
D = os.path.dirname(__file__)
s = requests.Session()
s.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) personal-car-watch/2.0", "Accept-Language": "pt-BR,pt;q=0.9"})
def pp(h):
    m = re.search(r'(?is)<script[^>]+id="__NEXT_DATA__"[^>]*>(.*?)</script>', h)
    try: return json.loads(m.group(1))["props"]["pageProps"] if m else {}
    except Exception: return {}
out = []
NP, MB = "https://napista.com.br/busca/", "https://www.mobiauto.com.br/comprar/carros-usados/brasil/"
urls = [NP + f"honda/civic/{y}" for y in range(2012, 2017)] + [NP + f"toyota/yaris/{y}" for y in (2019, 2021, 2023)] + \
       [MB + f"honda/civic/ano-{y}" for y in range(2012, 2017)] + [MB + f"toyota/yaris/ano-{y}" for y in (2019, 2021, 2023)] + \
       [MB + "toyota/yaris", MB + "toyota/yaris/s/cambio-automatico", MB + "toyota/yaris/s/carroceria-sedan",
        MB + "toyota/yaris/s/carroceria-seda", MB + "honda/civic/ano-2014/s/cambio-automatico"]
for u in urls:
    try:
        r = s.get(u, timeout=30); q = pp(r.text)
        if "searchResult" in q:
            sr = q["searchResult"]; offs = sr.get("offers") or []
            out.append({"url": u, "total": sr.get("offersTotal") or sr.get("total"), "n": len(offs),
                        "aut": sum(1 for o in offs if re.search(r"aut|cvt", o.get("versionName", ""), re.I)),
                        "sedan": sum(1 for o in offs if re.search(r"sed", o.get("versionName", ""), re.I))})
        else:
            dd = q.get("deals") or {}; res = dd.get("results") or []
            out.append({"url": u, "final": r.url[-60:], "total": dd.get("numResults"), "n": len(res),
                        "aut": sum(1 for x in res if "autom" in str(((x.get("trim") or {}).get("transmission") or {}).get("name", "")).lower()),
                        "sedan": sum(1 for x in res if "sed" in str(((x.get("trim") or {}).get("bodystyle") or {}).get("name", "")).lower()),
                        "img_in_html": bool(res) and str(((res[0].get("images") or [{}])[0]).get("imageId", "x")) in r.text.split("__NEXT_DATA__")[0],
                        "img_urls": list(dict.fromkeys(re.findall(r'https://[^"\s]+/images/v1\.0/\d+/transform/[^"\s]+', r.text)))[:2]})
    except Exception as e:
        out.append({"url": u, "error": str(e)[:100]})
    time.sleep(1.5)
json.dump(out, open(os.path.join(D, "br_count.json"), "w"), indent=1)
