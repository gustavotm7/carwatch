"""Can GitHub's servers read Brazilian car sites? (one-off test)"""
import json, os, re, time, requests
URLS = [
 ("webmotors", "page", "https://www.webmotors.com.br/carros/estoque/bmw/m3"),
 ("webmotors", "api", "https://www.webmotors.com.br/api/search/car?url=https://www.webmotors.com.br/carros/estoque/bmw/m3&actualPage=1&displayPerPage=24&order=1"),
 ("olx", "page", "https://www.olx.com.br/autos-e-pecas/carros-vans-e-utilitarios/bmw/m3"),
 ("mercadolivre", "page", "https://lista.mercadolivre.com.br/veiculos/carros-caminhonetes/bmw/m3/"),
 ("mercadolivre", "api", "https://api.mercadolibre.com/sites/MLB/search?category=MLB1744&q=bmw%20m3"),
 ("icarros", "page", "https://www.icarros.com.br/comprar/bmw/m3"),
 ("mobiauto", "page", "https://www.mobiauto.com.br/comprar/carros/bmw/m3"),
 ("napista", "page", "https://napista.com.br/busca/bmw-m3"),
]
BLOCK = re.compile(r"captcha|access denied|cf-chl|challenge-platform|just a moment|are you a human|request blocked|perimeterx|px-captcha|datadome|akamai|incapsula|attention required", re.I)
s = requests.Session()
s.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36",
                  "Accept-Language": "pt-BR,pt;q=0.9,en;q=0.5", "Accept": "text/html,application/json;q=0.9,*/*;q=0.8"})
out = []
for site, kind, url in URLS:
    o = {"site": site, "kind": kind, "url": url}
    try:
        r = s.get(url, timeout=30)
        t = r.text
        m = BLOCK.search(t[:30000])
        o.update(status=r.status_code, final=r.url[:120], bytes=len(r.content), ctype=r.headers.get("content-type", "")[:40],
                 server=r.headers.get("server", ""), block=m.group(0) if m else "",
                 title=(re.search(r"(?is)<title[^>]*>(.*?)</title>", t) or [0, ""])[1].strip()[:100] if "<title" in t.lower() else "",
                 prices=len(re.findall(r"R\$\s?\d{1,3}(?:\.\d{3})+", t)), m3_mentions=len(re.findall(r"(?i)\bm3\b", t)),
                 has_next_data="__NEXT_DATA__" in t, ldjson=t.count("application/ld+json"), start=t[:200] if "json" in o.get("ctype", "") or kind == "api" else "")
    except Exception as e:
        o["error"] = f"{type(e).__name__}: {e}"[:200]
    out.append(o); time.sleep(2)
json.dump(out, open(os.path.join(os.path.dirname(__file__), "br_results.json"), "w"), indent=1, ensure_ascii=False)
