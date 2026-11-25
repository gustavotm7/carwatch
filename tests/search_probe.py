"""One-off check (run by .github/workflows/test-search.yml, removed afterwards): opens candidate search pages for
models being added, prints what their ads are, and lists the filter parameters a site's search page uses."""
import collections
import json
import re
import sys

sys.path.insert(0, "collector")
import requests                                   # noqa: E402
import car_collector as cc                        # noqa: E402

sess = requests.Session()
sess.headers.update({"User-Agent": cc.UA, "Accept-Language": "en-CA,en;q=0.8,fr-CA;q=0.5"})
cfg = json.load(open("tests/search_probe_urls.json"))


def probe(source, url, open_n):
    r = cc.fetch(sess, url)
    print(f"\n== {source} {getattr(r, 'status_code', 'no answer')} {url}")
    if r is None or r.status_code != 200:
        return
    if r.url != url:
        print("   redirected to", r.url)
    if source in cc.JSON_SEARCH:
        items = cc.br_search_items(source, r.text)
        print(f"   {len(items)} ads on page 1")
        for it in items[:5]:
            print("    ", it.get("year"), "|", it.get("gear"), "|", it.get("price"), "|", (it.get("title") or "")[:80])
        return
    links = []
    for m in cc.LINK_RE[source].finditer(r.text):
        u = cc.HOSTS[source] + m.group(1)
        if u not in links:
            links.append(u)
    print(f"   {len(links)} ads on page 1 ({r.text.count('/offers/')} '/offers/' and {r.text.count('vehicleTransmission')} 'vehicleTransmission' in the page)")
    if source == "kijiji":
        print("   addresses:", [u.split("/v-cars-trucks/", 1)[-1][:60] for u in links[:20]])
    for u in links[:open_n]:
        rr = cc.fetch(sess, u)
        if rr is None or rr.status_code != 200:
            print("    could not open", u[:100])
            continue
        info = cc.parse_listing(rr.text)
        print("    ", info.get("year"), "|", (info.get("title") or "")[:70], "|", info.get("trans_field"), "|",
              info.get("price"), "|", info.get("km"), "| body:", info.get("body_raw"), "|", u.rsplit("/", 1)[-1][:60])


def params(url, words):
    r = cc.fetch(sess, url)
    if r is None or r.status_code != 200:
        print("\n== parameters: no page", url)
        return
    names = collections.Counter(re.findall(r"[?&]([A-Za-z_][A-Za-z_0-9]{1,24})=", r.text))
    print(f"\n== query parameters in {url}:\n  ", names.most_common(50))
    for w in words:
        hits = list(re.finditer(w, r.text))
        print(f"   [{w}] {len(hits)} times")
        for m in hits[:4]:
            print("     ...", r.text[max(0, m.start() - 140): m.end() + 140].replace("\n", " "), "...")


for source, url, n in cfg["probe"]:
    probe(source, url, n)
for url, words in cfg["params"]:
    params(url, words)
