"""One-off: page structure of NaPista and Mobiauto (search + ad pages), model address slugs, paging and gearbox filters."""
import json, os, re, time, requests
D = os.path.dirname(__file__)
s = requests.Session()
s.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) personal-car-watch/2.0", "Accept-Language": "pt-BR,pt;q=0.9"})
out = {"pages": [], "models": {}}

def nd(html):
    m = re.search(r'(?is)<script[^>]+id="__NEXT_DATA__"[^>]*>(.*?)</script>', html)
    try:
        return json.loads(m.group(1)) if m else None
    except ValueError:
        return None

def lists_of_ads(o, path="", found=None):
    found = [] if found is None else found
    if isinstance(o, list) and len(o) >= 2 and all(isinstance(x, dict) for x in o[:3]):
        ks = set().union(*[set(x.keys()) for x in o[:3]])
        if any(re.search(r"price|preco|valor", k, re.I) for k in ks):
            found.append((path, len(o), sorted(ks)[:60], json.dumps(o[0], ensure_ascii=False)[:2500]))
    if isinstance(o, dict):
        for k, v in o.items(): lists_of_ads(v, f"{path}.{k}", found)
    elif isinstance(o, list):
        for i, v in enumerate(o[:3]): lists_of_ads(v, f"{path}[{i}]", found)
    return found

def probe(url, detail_rx=None):
    o = {"url": url}
    try:
        r = s.get(url, timeout=30); h = r.text
        o.update(status=r.status_code, final=r.url[:150], bytes=len(h),
                 title=(re.search(r"(?is)<title[^>]*>(.*?)</title>", h) or [0, ""])[1].strip()[:100])
        d = nd(h)
        if d:
            pp = d.get("props", {}).get("pageProps", {})
            o["pp_keys"] = list(pp.keys())[:30]
            o["query"] = d.get("query"); o["page"] = d.get("page")
            o["lists"] = lists_of_ads(pp)[:3]
        o["ld_types"] = [re.findall(r'"@type"\s*:\s*"(\w+)"', x)[:6] for x in re.findall(r'(?is)<script[^>]+ld\+json[^>]*>(.*?)</script>', h)][:6]
        if detail_rx:
            o["detail_links"] = list(dict.fromkeys(re.findall(detail_rx, h)))[:12]
        o["_h"] = h
    except Exception as e:
        o["error"] = str(e)[:150]
    return o

# NaPista: make pages -> model slugs; one search; one ad
for mk in ("bmw", "honda", "toyota", "nissan"):
    o = probe(f"https://napista.com.br/busca/{mk}")
    out["models"][f"napista/{mk}"] = sorted(set(re.findall(rf'href="(/busca/{mk}/[a-z0-9\-]+)"', o.pop("_h", ""))))[:120]
    time.sleep(2)
for mk in ("bmw", "honda", "toyota", "nissan"):
    o = probe(f"https://www.mobiauto.com.br/comprar/carros-usados/brasil/{mk}")
    out["models"][f"mobiauto/{mk}"] = sorted(set(re.findall(rf'href="(?:https://www\.mobiauto\.com\.br)?(/comprar/carros-usados/brasil/{mk}/[a-z0-9\-]+)"', o.get("_h", ""))))[:120]
    out["models"][f"mobiauto/{mk}_any"] = sorted(set(re.findall(r'href="([^"]*' + mk + r'[^"]*)"', o.pop("_h", ""))))[:40]
    time.sleep(2)
for url, rx in [("https://napista.com.br/busca/bmw/series-3-m3", r'href="(/anuncios?/[^"]+|/carro/[^"]+|/oferta/[^"]+|/[a-z0-9\-]+/\d{5,}[^"]*)"'),
                ("https://napista.com.br/busca/honda/civic", r'href="(/anuncios?/[^"]+|/carro/[^"]+|/oferta/[^"]+|/[a-z0-9\-]+/\d{5,}[^"]*)"'),
                ("https://www.mobiauto.com.br/comprar/carros-usados/brasil/bmw/m3", r'href="([^"]*/detalhes/[^"]*|[^"]*/\d{6,}[^"]*)"'),
                ("https://www.mobiauto.com.br/comprar/carros-usados/brasil/honda/civic", r'href="([^"]*/detalhes/[^"]*|[^"]*/\d{6,}[^"]*)"')]:
    o = probe(url, rx); h = o.pop("_h", "")
    o["href_sample"] = list(dict.fromkeys(re.findall(r'href="([^"]{12,160})"', h)))[40:90]
    out["pages"].append(o); time.sleep(2)
# paging / gearbox filter guesses
for u in ["https://napista.com.br/busca/honda/civic?page=2", "https://napista.com.br/busca/honda/civic?pagina=2",
          "https://napista.com.br/busca/honda/civic?cambio=manual", "https://napista.com.br/busca/honda/civic/manual",
          "https://www.mobiauto.com.br/comprar/carros-usados/brasil/honda/civic?page=2",
          "https://www.mobiauto.com.br/comprar/carros-usados/brasil/honda/civic?pagina=2",
          "https://www.mobiauto.com.br/comprar/carros-usados/brasil/honda/civic/cambio-manual",
          "https://www.mobiauto.com.br/comprar/carros-usados/brasil/honda/civic?cambio=manual"]:
    o = probe(u); h = o.pop("_h", "")
    o.pop("lists", None)
    ids = re.findall(r'"(?:id|adId|listingId|dealId)"\s*:\s*"?(\d{5,})', h)
    o["n_ids"] = len(set(ids)); o["first_ids"] = list(dict.fromkeys(ids))[:6]
    o["manual_words"] = len(re.findall(r"(?i)\bmanual\b", h)); o["auto_words"] = len(re.findall(r"(?i)autom[aá]tic", h))
    out["pages"].append(o); time.sleep(2)
# one ad page per site (first detail link found)
for pg in out["pages"][:4]:
    links = pg.get("detail_links") or []
    if links:
        u = links[0] if links[0].startswith("http") else ("https://napista.com.br" if "napista" in pg["url"] else "https://www.mobiauto.com.br") + links[0]
        o = probe(u); h = o.pop("_h", "")
        o["kind"] = "ad"
        o["desc_like"] = [m[:300] for m in re.findall(r'"(?:description|descricao|observacoes?|comments?)"\s*:\s*"((?:[^"\\]|\\.){30,1500})"', h)][:3]
        o["date_like"] = re.findall(r'"(\w*(?:[Dd]ate|[Cc]reated|[Pp]ublish|[Pp]ublicad)\w*)"\s*:\s*"?([^",}]{6,30})', h)[:12]
        o["gear_like"] = re.findall(r'"(\w*(?:transmission|cambio|gear|transmissao)\w*)"\s*:\s*"?([^",}]{1,30})', h, re.I)[:8]
        o["img_like"] = list(dict.fromkeys(re.findall(r'https://[^"\s]+?\.(?:jpe?g|webp|png)', h)))[:5]
        out["pages"].append(o); time.sleep(2)
json.dump(out, open(os.path.join(D, "br_probe.json"), "w"), indent=1, ensure_ascii=False)
