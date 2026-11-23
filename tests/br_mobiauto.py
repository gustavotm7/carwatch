"""One-off: find Mobiauto's search page address and check whether a plain request returns listings."""
import json, os, re, time, requests
out = {"links": [], "tries": []}
s = requests.Session()
s.headers.update({"User-Agent": "personal-car-watch/2.0 (one-off check)", "Accept-Language": "pt-BR,pt;q=0.9"})
h = s.get("https://www.mobiauto.com.br/", timeout=30).text
links = sorted(set(re.findall(r'href="(/comprar[^"#?]{0,120})"', h)))
out["links"] = links[:60]
cands = ["/comprar/carros/brasil/bmw/m3", "/comprar/carros-usados/brasil/bmw/m3", "/comprar/carros/bmw-m3"]
cands += [l for l in links if "bmw" in l.lower()][:3]
for path in cands:
    u = "https://www.mobiauto.com.br" + path
    r = s.get(u, timeout=30); t = r.text
    out["tries"].append({"url": u, "status": r.status_code, "final": r.url[:140], "bytes": len(t),
                         "title": (re.search(r"(?is)<title[^>]*>(.*?)</title>", t) or [0, ""])[1].strip()[:90],
                         "prices": len(re.findall(r"R\$\s?\d{1,3}(?:\.\d{3})+", t)), "m3": len(re.findall(r"(?i)\bm3\b", t)),
                         "next_data": "__NEXT_DATA__" in t, "ldjson": t.count("application/ld+json")})
    time.sleep(2)
json.dump(out, open(os.path.join(os.path.dirname(__file__), "br_mobiauto.json"), "w"), indent=1, ensure_ascii=False)
