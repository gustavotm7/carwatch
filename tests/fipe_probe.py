"""One-off probe of the public FIPE API (v2): model names, year codes and a price, to set up CarWatch's matching."""
import json, re, urllib.request, urllib.error

BASES = ["https://fipe.parallelum.com.br/api/v2", "https://fipe.api.br/api/v2"]
out = {"errors": []}


def get(path):
    for b in BASES:
        try:
            req = urllib.request.Request(b + path, headers={"User-Agent": "CarWatch/1.0 (personal used-car price tracker)"})
            with urllib.request.urlopen(req, timeout=30) as r:
                return b, json.loads(r.read().decode())
        except urllib.error.HTTPError as e:
            out["errors"].append(f"{b}{path}: HTTP {e.code} {e.read()[:200]!r}")
        except Exception as e:
            out["errors"].append(f"{b}{path}: {e}")
    return None, None


base, refs = get("/references")
out["base"], out["refs"] = base, (refs or [])[:14]
_, brands = get("/cars/brands")
want = {"HONDA", "TOYOTA", "BMW", "NISSAN"}
out["brands"] = [b for b in (brands or []) if b["name"].split()[0].upper() in want]
pat = re.compile(r"civic|yaris|\bm3\b|\bm5\b|\b3[23][058]|350z|370z|\bz\b", re.I)
for b in out["brands"]:
    _, models = get(f"/cars/brands/{b['code']}/models")
    out.setdefault("models", {})[b["name"]] = [m for m in (models or []) if pat.search(m["name"])]
civ = [m for m in out.get("models", {}).get("Honda", []) if re.search(r"LXR", m["name"])]
if civ:
    hid = next(b["code"] for b in out["brands"] if b["name"].upper().startswith("HONDA"))
    _, yrs = get(f"/cars/brands/{hid}/models/{civ[0]['code']}/years")
    out["civic_model"], out["civic_years"] = civ[0], yrs
    if yrs:
        _, price = get(f"/cars/brands/{hid}/models/{civ[0]['code']}/years/{yrs[0]['code']}")
        out["civic_price"] = price
        if refs and len(refs) > 6:
            _, old = get(f"/cars/brands/{hid}/models/{civ[0]['code']}/years/{yrs[0]['code']}?reference={refs[6]['code']}")
            out["civic_price_6m_ago"] = old
            fc = (price or {}).get("codeFipe")
            if fc:
                _, hist = get(f"/cars/{fc}/years/{yrs[0]['code']}/history")
                out["history_endpoint"] = hist
json.dump(out, open("tests/fipe_probe.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps(out, ensure_ascii=False)[:3000])
