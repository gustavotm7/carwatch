import csv, json, os, re, time, importlib.util
D = os.path.dirname(__file__)
spec = importlib.util.spec_from_file_location("cc", os.path.join(D, "car_collector.py"))
cc = importlib.util.module_from_spec(spec); spec.loader.exec_module(cc)
import requests
cc.DELAY_S = 1.2
s = requests.Session()
s.headers.update({"User-Agent": cc.UA, "Accept-Language": "en-CA,en;q=0.8,fr-CA;q=0.5"})
rows = [r for r in cc.load_rows(os.path.join(D, "car_listings_20261118.csv")) if r["status"] in ("active", "unverified")]
out = {"check": [], "fresh": []}
for r in rows:
    oc, res = cc.check_listing(s, r)
    o = {"id": r["listing_id"][:24], "name": r["name"], "status": r["status"], "outcome": oc}
    if oc == "ok":
        row = dict(r); cc.apply_page_details(row, res, "2026-11-18")
        o.update(price=res.get("price"), km=res.get("km"), year=res.get("year"), trans=res.get("trans_field"),
                 body=res.get("body_raw"), desc=len(res.get("desc") or ""), photo=row["photo_id"][-50:],
                 red=row["red_flags"], green=row["green_flags"], trim=row["trim"])
    else:
        o["reason"] = res
    out["check"].append(o)
for src, url in (("autotrader", "https://www.autotrader.ca/cars/bmw/m3/"), ("kijiji", "https://www.kijiji.ca/b-cars-trucks/canada/nissan-350z/k0c174l0")):
    g = cc.fetch(s, url)
    links = list(dict.fromkeys((cc.make_id(src, m.group(2)), cc.HOSTS[src] + m.group(1)) for m in cc.LINK_RE[src].finditer(g.text)))[:5]
    for lid, u in links:
        a = cc.fetch(s, u); info = cc.parse_listing(a.text)
        slug = u.rsplit("/", 1)[-1] if src != "kijiji" else u.rsplit("/", 2)[-2]
        key = cc.detect_key(info.get("title", ""), slug, info.get("year"))
        o = {"src": src, "url": u[-70:], "title": info.get("title"), "year": info.get("year"), "price": info.get("price"),
             "km": info.get("km"), "trans_field": info.get("trans_field"), "body_raw": info.get("body_raw"),
             "desc": len(info.get("desc") or ""), "image": (info.get("image") or "")[:90], "key": key}
        if key:
            body = cc.body_of(info.get("body_raw", ""), slug + " " + info.get("title", ""))
            fam, gen, name = cc.classify(key, info["year"], body)
            o.update(name=name, body=body, trans=cc.classify_trans(fam, info.get("trans_field", ""), info.get("desc", "")),
                     flags=cc.flags_of(info.get("desc")), trim=cc.trim_of(fam, info.get("title", "") + " " + slug))
        out["fresh"].append(o)
json.dump(out, open(os.path.join(D, "verify.json"), "w"), indent=1, default=str)
