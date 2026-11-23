"""One-off: do NaPista / Mobiauto photos load when the page asking for them is the CarWatch site?"""
import json, os, requests
D = os.path.dirname(__file__)
out = []
for u in json.load(open(os.path.join(D, "photo_urls.json"))):
    for ref in ("https://gustavotm7.github.io/carwatch/", None):
        h = {"User-Agent": "Mozilla/5.0 (Linux; Android 14) AppleWebKit/537.36 Chrome/131.0 Mobile Safari/537.36", "Accept": "image/avif,image/webp,image/*,*/*"}
        if ref: h["Referer"] = ref
        try:
            r = requests.get(u, headers=h, timeout=30)
            out.append({"url": u[-60:], "referer": bool(ref), "status": r.status_code, "type": r.headers.get("content-type"), "bytes": len(r.content)})
        except Exception as e:
            out.append({"url": u[-60:], "error": str(e)[:100]})
json.dump(out, open(os.path.join(D, "photo_check.json"), "w"), indent=1)
