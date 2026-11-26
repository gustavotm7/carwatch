"""One-off: why do so many Brazilian re-checks look gone? Compare saved km, search km and the ad page."""
import csv, json, os, sys, time, requests
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "collector"))
import car_collector as cc
here = os.path.dirname(__file__)
cc.load_config(os.path.join(here, "..", "cars_config_br.json"))
cc.DATA_DIR = os.path.join(here, "..", "data_br")
rows = cc.load_rows(cc.latest_snapshot())
by_id = {r["listing_id"]: r for r in rows}
s = requests.Session(); s.headers.update({"User-Agent": cc.UA, "Accept-Language": "pt-BR,pt;q=0.9,en;q=0.8"})
keep = [x for x in cc.SEARCHES if any(k in x[1] for k in ("honda/civic", "toyota/yaris", "chevrolet/monza", "fiat/bravo"))][:14]
cc.SEARCHES[:] = keep
stats = dict(searches=[])
cc.DEADLINE[0] = time.time() + 240
cc.search_ids(s, set(by_id), stats)
out = {"searches": stats["searches"], "seen": len(cc.SEARCH_SEEN), "diff": []}
n_same = 0
for lid, it in cc.SEARCH_SEEN.items():
    r = by_id.get(lid)
    if not r or r["status"] not in ("active", "unverified"):
        continue
    if it.get("km") and r["km"] and abs(int(it["km"]) - int(r["km"])) > 2000:
        out["diff"].append({"id": lid, "url": r["url"], "row_km": r["km"], "search_km": it["km"], "title": it.get("title"),
                            "row_first": r["first_seen"], "row_note": r["note"][:80]})
    else:
        n_same += 1
out["same"] = n_same
cc.DEADLINE[0] = 0
for d in out["diff"][:12]:
    try:
        outcome, res = cc.check_listing(s, by_id[d["id"]])
        d["page"] = [outcome, res if outcome != "ok" else {k: res.get(k) for k in ("price", "km", "year")}]
    except Exception as e:
        d["page"] = ["error", str(e)[:150]]
json.dump(out, open(os.path.join(here, "br_results.json"), "w"), indent=1, ensure_ascii=False, default=str)
