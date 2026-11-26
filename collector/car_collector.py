#!/usr/bin/env python3
"""
car_collector.py  -  daily used-car watcher + Excel builder  (v2)

What one run does
  1. Starts today's snapshot file  car_listings_YYYYMMDD.csv  (header written first, same file used for
     the whole run) and loads the most recent finished snapshot as yesterday's state.
  2. Re-opens every listing's own page: price change -> prev_price, page gone / replaced by a different
     car -> status 'gone', otherwise last_seen / verified = today. Re-reads the ad text for flags and
     trim, and replaces the photo if the ad now shows a different first photo.
  3. Reads the search pages in SEARCHES (Autotrader, Kijiji, CarPages), opens each listing it has not
     seen before and adds it when it matches one of the cars below with a manual (or conflicting)
     gearbox. Rejected ads are remembered (status 'skip') so they are not opened again every day.
  4. Downloads a small thumbnail per car into .\\photos and embeds it in the workbook (works in any Excel,
     phone previews included). If a download fails, that row falls back to Excel's IMAGE() function.
  5. Fits a fair-price model per car family, then rebuilds used_cars_canada.xlsx (sorted by price).

Usage (normally started by the 'Update now' button in CarWatch.xlsm)
  python car_collector.py                 full run
  python car_collector.py --no-network    only rebuild the workbook from the latest snapshot
  python car_collector.py --selftest      read 3 stored listings + every search page, print what it found

STATUS: the Excel part, the scoring and the bookkeeping are tested offline with mock pages. The page
parsing has NOT been tested against the live sites (they could not be reached from where this was
written). Run --selftest first. A page it cannot read is left as it was; the script only marks a listing
'gone' on a clear signal. Check the sites' terms of use before automating; it waits 2 s between page
requests and makes roughly 250-400 requests a day.
"""
from __future__ import annotations

import argparse
import csv
import unicodedata
import datetime as dt
import glob
import hashlib
import io
import json
import math
import os
import re
import shutil
import sys
import time
from html import unescape

print("CarWatch collector started", dt.datetime.now().strftime("%Y-%m-%d %H:%M"), flush=True)


def _ensure_packages():
    """Install requests / openpyxl / pillow on first use, so no manual pip step is needed."""
    import importlib
    import importlib.util
    import subprocess
    need = [pip for mod, pip in (("requests", "requests"), ("openpyxl", "openpyxl"), ("PIL", "pillow"))
            if importlib.util.find_spec(mod) is None]
    if need:
        print("installing", " ".join(need), "...", flush=True)
        subprocess.call([sys.executable, "-m", "pip", "install", "--disable-pip-version-check", *need])
        importlib.invalidate_caches()


_ensure_packages()

try:
    import requests
except ImportError:  # --no-network still works without it
    requests = None
try:
    from PIL import Image as PILImage
except ImportError:  # without Pillow: no embedded photos, IMAGE() formulas only
    PILImage = None

from openpyxl import Workbook
from openpyxl.chart import Reference, ScatterChart, Series
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.properties import PageSetupProperties

# --------------------------------------------------------------------------------------------
# Settings
# --------------------------------------------------------------------------------------------
HERE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = HERE                                    # snapshots, workbook, log (set with --data-dir)
SNAPSHOT_FMT = "car_listings_{:%Y%m%d}.csv"      # one snapshot per day, written from the start of the run
XLSX_NAME = "cars_data.xlsx"                      # CarWatch.xlsm loads this file (Update now button / on open)
PHOTO_DIR = os.path.join(HERE, "photos")         # moved under --data-dir when given
DELAY_S = 2.0                                      # pause between page requests
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) personal-car-watch/2.0"
MAX_NEW_PER_RUN = 80                               # unseen ads opened per run
PAGES_PER_SEARCH = 3                               # result pages read per search
PAGES_JSON = 20                                    # NaPista / Mobiauto: their search data is read directly (stops early)
SEARCH_SEEN = {}                                   # today's search data per ad id (NaPista / Mobiauto)
AUTO_FOUND = {}                                    # automatics found today for families that allow them
ALLOW_AUTO = set()                                 # families kept with automatics too (Brazil: FB Civic, Yaris sedan)
SPLITS = {}                                        # family -> [(title regex, sub-generation)]: e.g. FA5 Civic Si apart
GEN_BODY = {}                                      # family -> the only body it came in ('body' in the config)
GONE_RECHECK_DAYS = 7                              # keep re-checking 'gone' ads this long (they sometimes return)

# Searches come from cars_config.json (per car, per site). Extra searches can be added here as (source, url).
EXTRA_SEARCHES = []
SEARCHES = []                                       # filled by load_config()

# Score settings: written to the Summary tab every day. Change them HERE to make a change stick.
SCORE = {
    "base": 5.0,                       # score of a car priced exactly at its fair price
    "gearbox": {"\u2714 Manual": 0.0, "? Unconfirmed": -0.5, "\u26A0 Conflict": -1.0, "Automatic": 0.0},
    "red": {"Rebuilt title": -3.0, "Salvage title": -4.0, "Accident mentioned": -1.0, "Sold as-is": -1.0,
            "Needs work": -1.5, "Modified": -0.5, "Price hidden": -1.0, "Gearbox problem": -1.0,
            "Clutch slipping": -0.5},
    "green": {"Rod bearings done": 1.0, "Subframe reinforced": 0.5, "VANOS/timing done": 0.5,
              "Clutch done": 0.5, "Gearbox serviced": 0.5, "Service records": 0.5, "No accidents stated": 0.5,
              "One owner": 0.5, "Winter stored": 0.5},
    "green_cap": 2.0,
}

MIN_PRICE = 1500                     # below this the ad is not showing a real price ($1 = "call us" / sale pending)
MAX_PRICE = {"CA": 1_000_000, "BR": 8_000_000}   # above this it's a typo ($1,699,000 for a 2004 STI)
COLS = ["listing_id", "source", "family", "gen", "name", "year", "body", "price", "km", "trans",
        "place", "url", "photo_id", "note", "first_seen", "last_seen", "prev_price", "status", "verified",
        "trim", "red_flags", "green_flags", "flags_read", "photo_alt", "desc", "listed", "price_hist", "page_read"]
# listed: date the ad went up, as the site states it (Autotrader / Kijiji). price_hist: "date:price;date:price"
# with one entry per price change, starting with the first price CarWatch saw.

UUID = r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}"
UUID_RE = re.compile(UUID)
PHOTO_RE = re.compile(r"listing-images/(" + UUID + r")_(" + UUID + r")\.(jpg|jpeg|png|webp)", re.I)
PROV = r"(?:AB|BC|MB|NB|NL|NS|NT|NU|ON|PE|QC|SK|YT)"
PHOTO_BASE = "https://prod.pictures.autoscout24.net/listing-images/"

LINK_RE = {
    "autotrader": re.compile(r"(?:https?://www\.autotrader\.ca)?(/offers/[A-Za-z0-9_\-.%+]+-(" + UUID + r"))"),
    "kijiji": re.compile(r"(?:https?://www\.kijiji\.ca)?(/v-cars-trucks/[A-Za-z0-9_\-.%/]+?/(m?\d{6,12}))(?=[\"'?#\s])"),
    "carpages": re.compile(r"(?:https?://www\.carpages\.ca)?(/used-cars/[a-z\-]+/[a-z\-]+/(\d{4}-[a-z0-9\-]+-\d{5,10}))"),
    "napista": re.compile(r"(?:https?://napista\.com\.br)?(/anuncios/(" + UUID + r"))"),
    "mobiauto": re.compile(r"(?:https?://www\.mobiauto\.com\.br)?(/comprar/carros/[a-z0-9\-]+/[a-z0-9\-]+/[a-z0-9\-]+/"
                           r"\d{4}/[a-z0-9\-]+/detalhes/(\d{5,10}))"),
}
HOSTS = {"autotrader": "https://www.autotrader.ca", "kijiji": "https://www.kijiji.ca",
         "carpages": "https://www.carpages.ca", "napista": "https://napista.com.br",
         "mobiauto": "https://www.mobiauto.com.br"}
JSON_SEARCH = {"napista", "mobiauto"}     # search pages carry the ads as data: unwanted ones are skipped unopened
COUNTRY = {"country": "CA", "currency": "CAD", "links": []}   # from the config file (Brazil: BR / BRL / OLX links)


def source_of(url):
    for s in HOSTS:
        if s in url:
            return s
    return "other"


def make_id(source, key):
    return key if source == "autotrader" else f"{source}:{key}"


def id_key(row):
    """The part of the listing id that must still appear in the ad's address after redirects."""
    lid = row["listing_id"]
    if row["source"] == "carpages":
        return lid.rsplit("-", 1)[-1]
    return lid.split(":")[-1]


# --------------------------------------------------------------------------------------------
# Car rules
# --------------------------------------------------------------------------------------------
MODELS = []            # enabled models from cars_config.json
FAMILY_STYLE = {}      # family -> (fill, text, short name, description)
TRIMS = {}             # family -> ([(regex, label)], multi)
NEEDS_TEXT_CONFIRM = set()   # families that also came with SMG/DCT: the data field alone is not enough
COMPS = {}             # family -> sold-price settings (Bring a Trailer page ids, title regex, years, default body)


def load_config(path=None):
    """Read cars_config.json (next to the script, or one folder up) and fill the car rules."""
    global MODELS, SEARCHES
    cands = [path] if path else [os.path.join(HERE, "cars_config.json"), os.path.join(HERE, "..", "cars_config.json")]
    cfg_path = next((c for c in cands if c and os.path.exists(c)), None)
    if not cfg_path:
        raise FileNotFoundError("cars_config.json not found (put it next to car_collector.py)")
    with open(cfg_path, encoding="utf-8") as f:
        cfg = json.load(f)
    MODELS = [m for m in cfg["models"] if m.get("enabled", True)]
    COUNTRY.update(country=cfg.get("country", "CA"), currency=cfg.get("currency", "CAD"),
                   links=[{"model": m["model"], "families": [g["family"] for g in m["generations"]] +
                                                        [sp["family"] for g in m["generations"] for sp in g.get("split", [])],
                           "links": [{"site": s, "url": u} for s, u in m.get("links", [])],
                           "links_auto": [{"site": s, "url": u} for s, u in m.get("links_auto", [])]}
                          for m in MODELS if m.get("links") or m.get("links_auto")])
    FAMILY_STYLE.clear(); TRIMS.clear(); NEEDS_TEXT_CONFIRM.clear(); COMPS.clear(); ALLOW_AUTO.clear(); SPLITS.clear(); GEN_BODY.clear()
    searches = []
    for m in MODELS:
        gens = []
        for g in m["generations"]:
            gens.append(g)
            if g.get("fipe"):                            # FIPE matching rule; split families use their parent's
                FIPE_CFG[g["family"]] = g["fipe"]
                for sp in g.get("split", []):
                    FIPE_CFG[sp["family"]] = g["fipe"]
            for sp in g.get("split", []):               # a trim priced on its own (own chip, own fair price)
                sg = {k: v for k, v in g.items() if k != "split"}
                sg.update(sp)
                SPLITS.setdefault(g["family"], []).append((sp["match"], sg))
                gens.append(sg)
        for g in gens:
            FAMILY_STYLE.setdefault(g["family"], (g.get("color", "#888888").lstrip("#").upper(), "10151C",
                                                  g.get("short", g["name"]), g.get("desc", g["name"])))
            if g.get("trims"):
                TRIMS.setdefault(g["family"], ([(rx, lab) for rx, lab in g["trims"]], bool(g.get("trim_multi"))))
            if g.get("smg_risk"):
                NEEDS_TEXT_CONFIRM.add(g["family"])
            if g.get("allow_auto"):
                ALLOW_AUTO.add(g["family"])
            if g.get("body"):
                GEN_BODY[g["family"]] = g["body"]
            if g.get("comps"):
                c = COMPS.setdefault(g["family"], {**g["comps"], "years": list(g["years"])})
                c["years"] = [min(c["years"][0], g["years"][0]), max(c["years"][1], g["years"][1])]
        for src, urls in (m.get("searches") or {}).items():
            for u in urls:
                if (src, u) not in searches:
                    searches.append((src, u))
    SEARCHES = searches + [s for s in EXTRA_SEARCHES if s not in searches]
    return cfg_path


def _norm(text):
    return (text or "").lower().replace("-", " ").replace("_", " ")


def detect_key(title: str, slug: str, year, ignore_make=False):
    """-> the matching model (dict) from cars_config.json, or None."""
    if year is None:
        return None
    t = _norm(title + " " + slug)
    for m in MODELS:
        if m.get("make") and not ignore_make and m["make"].lower() not in t:
            continue
        if not all(re.search(rx, t) for rx in m.get("match", [])):
            continue
        if any(re.search(rx, t) for rx in m.get("none", [])):
            continue
        if gen_for(m, year, t):
            return m
    return None


def fix_body(model, body, text=""):
    """Bodies the car never came in (a 'Sedan' 370Z) fall back to the model's own list."""
    allowed = model.get("bodies") or []
    if not allowed or body in allowed:
        return body
    if "conv" in allowed and re.search(r"roadster|convertible|cabrio|spyder|spider", _norm(text)):
        return "conv"
    return allowed[0] if len(allowed) == 1 or body in ("unk", "sedan", "wagon", "hatch") else body


def gen_for(model, year, text="", body=None):
    """The generation an ad belongs to: by year, unless the ad names a generation's chassis code ('code' in the
    config) within a year of its range: a JDM 'S13 Silvia' first registered in 1988, an FD RX-7 sold as a 1992.
    A generation with its own 'bodies' takes only those bodies once the body is known (Z3 M Coupe vs Roadster)."""
    t = _norm(text) if text else ""
    fits = lambda g: ((body is None or not g.get("bodies") or body in g["bodies"])
                      and (not g.get("need") or bool(t and re.search(g["need"], t))))   # 'need': a word the ad must have (Si coupe)
    if t:
        for g in model["generations"]:
            if (g.get("code") and g["years"][0] - 1 <= year <= g["years"][1] + 1 and re.search(g["code"], t)
                    and fits(g)):
                return g
    for g in model["generations"]:
        if g["years"][0] <= year <= g["years"][1] and fits(g):
            return g
    return None


def classify(model, year, body, text=""):
    """-> (family, generation/chassis, display name). ValueError: no generation for that year and body."""
    g = gen_for(model, year, text, body)
    if g is None:
        raise ValueError(f"{model['model']} {year} {body}")
    name = (g.get("name_by_body") or {}).get(body, g["name"])
    return g["family"], name.split(" ")[0], name


def classify_trans(family: str, field: str, desc: str) -> str:
    """m = manual confirmed, c = field and text disagree, u = unconfirmed, a = automatic/unknown (skipped)."""
    f, d = (field or "").lower(), (desc or "").lower()
    d = re.sub(r"manual (e|do|da|de) (chave|propriet\w*|dono|instru\w*|revis\w*)|manuais|piloto autom\u00e1tico|"
               r"ar(-| )condicionado autom\u00e1tico|climatizador autom\u00e1tico|far(o|\u00f3)is autom\u00e1ticos", " ", d)
    pt_auto = bool(re.search(r"c[a\u00e2]mbio:?\s*(autom|cvt)|autom[a\u00e1]tic[oa] de \d marchas|transmiss[a\u00e3]o autom", d))
    pt_man = bool(re.search(r"c[a\u00e2]mbio:?\s*(manual|mec)|manual de [456] marchas|transmiss[a\u00e3]o manual", d))
    if pt_auto and not pt_man:
        return "a"                                      # Brazil: the seller says automatic, the data field is often wrong
    auto_w = r"auto|smg|dct|dkg|steptronic|sequential|dual|cvt|tiptronic|\bpdk\b|\bsst\b"   # 'Autom\u00e1tico/Automatizado' too
    field_a = bool(re.search(auto_w, f))
    field_m = bool(re.search(r"manual|manuel|stick|6[- ]?spd|6mt", f)) and not field_a
    text_m = bool(re.search(r"6[- ]?speed manual|\b6[- ]?spd\b|\b6mt\b|\bmanual\b|\bmanuel+e?\b|\bstick\b|three[- ]pedal", d)) or pt_man
    text_a = bool(re.search(r"\bsmg\b|steptronic|\bautomatic\b|\bautomatique\b|\bdct\b|\bdkg\b|tiptronic|paddle"
                            r"|\bpdk\b|\btc[ -]?sst\b|\bsst\b|twin[ -]clutch"     # Porsche PDK, Evo X MR's TC-SST
                            r"|c\u00e2mbio autom|cambio autom|\bautom\u00e1tic[oa]\b|\bautomatizad[oa]\b|\bcvt\b", d))
    if field_m and text_a:
        return "c"
    if field_m:
        return "m" if (family not in NEEDS_TEXT_CONFIRM or text_m) else "u"
    if field_a and text_m:
        return "c"
    if text_m and not text_a and not field_a:
        return "u"
    return "a"


def body_of(raw: str, slug: str = "") -> str:
    first = _body_words(slug)                           # title / address words are more reliable than the field
    return first if first != "unk" else _body_words(raw)


def _body_words(t: str) -> str:
    t = (t or "").lower()
    if re.search(r"convertible|cabrio|roadster|cabriolet|convers[i\u00ed]vel", t):
        return "conv"
    if re.search(r"wagon|touring wgn|estate|wgn|sports? wagon|perua", t) or (re.search(r"touring", t) and "bmw" in t):
        return "wagon"
    if re.search(r"coupe|coup\u00e9|\bcpe\b|cup\u00ea", t):
        return "coupe"
    if "sedan" in t or "berline" in t or re.search(r"\bsed[a\u00e3]\b|\bsed\.", t):
        return "sedan"
    if re.search(r"hatch|liftback|fastback|\bhb\b|\b[35] ?(dr|door|portes|portas)\b|\b[35]p\b", t):
        return "hatch"                                  # FK8 Civic Type R, GR STI; a car that never came so falls back
    return "unk"


def trim_of(family, text):
    t = _norm(text)
    rules, multi = TRIMS.get(family, ([], False))
    found = []
    for rx, label in rules:
        if re.search(rx, t) and label not in found:
            found.append(label)
            if not multi:
                break
    return " ".join(found[:2])


FLAG_RULES = {
    "Rebuilt title": r"rebuilt (title|status|vehicle)|(title|status|brand(ed)?)\s*[:\-]?\s*rebuilt|reconstruit",
    "Salvage title": r"\bsalvage\b|irr[e\u00e9]cup[e\u00e9]rable",
    "Accident mentioned": r"(minor|small|light|previous|prior|front|rear|side)[- ](end )?(accident|collision|damage)"
                          r"|accident (history|reported|claim|repaired)|carfax shows|claims? (of|totall?ing) \$?\d"
                          r"|passagem (por|em) leil[a\u00e3]o|sinistrad|recuperad[oa] de sinistro|chassi remarcado"
                          r"|indenizad[oa] (pela|por) seguradora",
    "Sold as-is": r"\bas[- ]is\b|tel quel|not (safety )?certified|without (a )?safety|unsafetied|vendo no estado|no estado em que se encontra",
    "Needs work": r"needs (some )?(work|tlc|repair|a clutch|clutch|engine|transmission)|check engine light"
                  r"|engine (knock|noise)|does not run|not running|project car|mechanic.?s special"
                  r"|precisa de (reparo|reforma|manuten)|para reforma|n[a\u00e3]o liga|motor (batendo|fundido)|para ret[i\u00ed]fica",
    "Modified": r"\btuned\b|\bmodified\b|\bmods\b|coilovers?|\bheaders\b|supercharg|turbo kit|\bstage ?[123]\b"
                r"|aftermarket (exhaust|intake|suspension)|\bcat[- ]?back\b"
                r"|rebaixad|\bremap|turbinad|preparad[oa] |escape esportivo|suspens[a\u00e3]o (de )?rosca",
    "Rod bearings done": r"rod bearings? (were |have been |was )?(done|replaced|serviced|upgraded)"
                         r"|(new|replaced|fresh|upgraded) rod bearings|rod bearing (service|job|replacement)"
                         r"|coussinets (de bielle )?(remplac|fait)",
    "Subframe reinforced": r"subframe (reinforc|repair|plates?|kit|fixed|done)|reinforced (rear )?subframe|reinforcement plates",
    "VANOS/timing done": r"vanos (rebuilt|done|replaced|serviced|seals)|new (timing chain|chain guides|vanos)"
                         r"|(timing chain|chain guides?) (were |was |have been )?(done|replaced|serviced)",
    "Clutch done": r"new clutch|clutch (was |has been )?(replaced|done|installed)|recent(ly replaced)? clutch"
                   r"|nouvel embrayage|embrayage (neuf|remplac)|embreagem (nova|trocada|rec[e\u00e9]m trocada)",
    "Service records": r"service (records|history)|maintenance (records|history)|records available|full history"
                       r"|dealer (maintained|serviced)|carnet d.entretien|historique d.entretien"
                       r"|revis(?:[o\u00f5]es|ado) (?:\w+ ){0,3}concession[a\u00e1]ria|todas as revis[o\u00f5]es|manual e chave reserva",
    "No accidents stated": r"no accidents?|accident[- ]free|never been in an accident|sans accident|jamais accident"
                           r"|aucun accident|clean carfax|nunca batid|sem batida|sem sinistro|sem passagem por leil",
    "One owner": r"\bone owner\b|\b1 owner\b|single owner|1 propri[e\u00e9]taire|un (seul )?propri[e\u00e9]taire"
                 r"|[u\u00fa]nico dono|1 dono\b",
    "Winter stored": r"never seen (a )?winter|no winters?|winter stored|stored (for the |in )?winters?"
                     r"|summer (car|only|driven)|jamais (roul[e\u00e9]e? l.)?hiver|remis[e\u00e9]e? l.hiver",
    # gearbox: service and faults (any gearbox), clutch (manuals only, see GEAR_ONLY)
    "Gearbox serviced": r"(troca|trocad[oa]|substitu\w*) (do |de )?([o\u00f3]leo|fluido) (do |de |da )?(c[a\u00e2]mbio|cvt|transmiss)"
                        r"|([o\u00f3]leo|fluido) (do |de |da )?(c[a\u00e2]mbio|cvt|transmiss\w*)( cvt| autom\w*)? (trocad|novo|substitu)"
                        r"|c[a\u00e2]mbio (revisado|retificado|reformado|refeito|novo)"
                        r"|transmission (fluid|oil|service) (was |has been )?(done|changed|replaced|flushed)"
                        r"|(new|rebuilt) (transmission|gearbox)|(transmission|gearbox) (was |has been )?(rebuilt|replaced)",
    "Gearbox problem": r"c[a\u00e2]mbio (dando |com |d[a\u00e1] )?(trancos?|solavancos?)"
                       r"|c[a\u00e2]mbio (patinando|arranhando|com (defeito|problema|barulho|ru[i\u00ed]do))"
                       r"|marchas? (escapando|arranhando)|(?<!sem )(?<!nenhum )(defeito|problema) no c[a\u00e2]mbio"
                       r"|transmission (slipping|slips|problem|issue|needs)|(slipping|harsh) (shifts?|shifting)"
                       r"|grinds? (into|in) (\w+ )?gear",
    "Clutch slipping": r"embreagem (patinando|alta|no fim|ruim|batendo)|clutch (is )?slipping|slipping clutch",
}
GEAR_ONLY = {"Clutch done": "manual", "Clutch slipping": "manual"}      # no clutch pedal on an automatic
GEAR_FLAGS = ("Gearbox serviced", "Gearbox problem", "Clutch slipping")  # added later: also read from stored ad text


def gear_flags(rows):
    """Gearbox-aware flags: adds the gearbox flags to ads read before they existed, and drops clutch flags
    from automatics."""
    for r in rows:
        red = [f for f in (r.get("red_flags") or "").split("; ") if f]
        green = [f for f in (r.get("green_flags") or "").split("; ") if f]
        if r.get("flags_read") and r.get("desc"):
            nr, ng = flags_of((r.get("note") or "") + "\n" + r["desc"])
            red += [f for f in nr.split("; ") if f in GEAR_FLAGS and f not in red]
            green += [f for f in ng.split("; ") if f in GEAR_FLAGS and f not in green]
        if r.get("trans") == "a":
            red = [f for f in red if GEAR_ONLY.get(f) != "manual"]
            green = [f for f in green if GEAR_ONLY.get(f) != "manual"]
        r["red_flags"], r["green_flags"] = "; ".join(red), "; ".join(green)


def flags_of(text):
    t = (text or "").lower()
    red = [k for k in SCORE["red"] if k in FLAG_RULES and re.search(FLAG_RULES[k], t)]       # 'Price hidden' has no
    green = [k for k in SCORE["green"] if k in FLAG_RULES and re.search(FLAG_RULES[k], t)]   # text rule: set elsewhere
    return "; ".join(red), "; ".join(green)


# --------------------------------------------------------------------------------------------
# Snapshot files (CSV)
# --------------------------------------------------------------------------------------------
def marker_of(path):
    return os.path.splitext(path)[0] + ".ok"


def latest_snapshot():
    files = sorted(glob.glob(os.path.join(DATA_DIR, "car_listings_*.csv")))
    if not files:
        return None
    done = [f for f in files if os.path.exists(marker_of(f))]
    if done:
        return done[-1]

    def n_rows(f):
        with open(f, newline="", encoding="utf-8") as fh:
            return sum(1 for _ in fh)
    return max(files, key=lambda f: (n_rows(f), f))   # no finished marker anywhere: take the fullest file


def load_rows(path):
    if not path:
        return []
    with open(path, newline="", encoding="utf-8") as f:
        rows = [{c: r.get(c, "") or "" for c in COLS} for r in csv.DictReader(f)]
    # v2 could not read the year on Autotrader ad pages and wrongly rejected their ads: give those a second look
    return [r for r in rows if not (r["status"] == "skip" and r["source"] == "autotrader"
                                    and not re.search(r"\[v[345]\]$", r["note"]))]


SKIP_TAG = "[v5]"                                    # bumped when the car list grows: older rejections get a second look


def recheck_old_skips(rows):
    """Ads rejected before the current SKIP_TAG get one second look when the rules changed in their favour:
    'not one of the cars on the list' whose address names a car tracked now (a Kijiji address with a year must fit a
    generation: a 1987 Supra is an A70 now; one without a year, as on Autotrader, must name a model marked "recheck"),
    and Kijiji ads rejected as automatic whose title says manual (the title counts for the gearbox now)."""
    keep, n = [], 0
    for r in rows:
        if (r["status"] == "skip" and r["note"].startswith("automatic") and not r["note"].endswith(SKIP_TAG)
                and r["source"] == "kijiji"
                and re.search(r"\bmanu[ae]l|\b[56] ?(mt|spd|speed)\b|\bstick\b", _norm(r["url"].rsplit("/", 2)[-2]))):
            n += 1                                       # '2013 Mazda Miata GS manual': the title says manual now counts
            continue
        if r["status"] == "skip" and r["note"].startswith("not one of the cars") and not r["note"].endswith(SKIP_TAG):
            slug = r["url"].rsplit("/", 2)[-2] if r["source"] == "kijiji" else r["url"].rsplit("/", 1)[-1]
            t = _norm(slug)
            y = re.search(r"\b(19[5-9]\d|20[0-3]\d)\b", t)
            if y:
                again = detect_key(slug, "", int(y.group(1)), ignore_make=True) is not None
            else:
                again = any(m.get("recheck") and all(re.search(rx, t) for rx in m.get("match", []))
                            and not any(re.search(rx, t) for rx in m.get("none", [])) for m in MODELS)
            if again:
                n += 1
                continue
        keep.append(r)
    return keep, n


class Snapshot:
    """Today's CSV: created (header only) when the run starts, then written row by row."""

    def __init__(self, path):
        self.path = path
        if os.path.exists(path):
            shutil.copy2(path, path + ".bak")           # same-day rerun: keep a copy of what was there
        if os.path.exists(marker_of(path)):
            os.remove(marker_of(path))                  # not "finished" until the run completes
        self.fh = open(path, "w", newline="", encoding="utf-8")
        self.w = csv.DictWriter(self.fh, fieldnames=COLS, extrasaction="ignore")
        self.w.writeheader()
        self.fh.flush()

    def add(self, row):
        self.w.writerow(row)
        self.fh.flush()

    def finish(self):
        self.fh.close()
        with open(marker_of(self.path), "w") as f:
            f.write(dt.datetime.now().isoformat(timespec="seconds") + "\n")


# --------------------------------------------------------------------------------------------
# Web access and page parsing
# --------------------------------------------------------------------------------------------
_last_request = [0.0]


DEADLINE = [0.0]                                   # --max-minutes: past this time, no more pages are opened


def time_up():
    return bool(DEADLINE[0]) and time.time() > DEADLINE[0]


def fetch(sess, url, delay=None):
    """GET with a pause between requests; returns a Response or None."""
    wait = (DELAY_S if delay is None else delay) - (time.time() - _last_request[0])
    if wait > 0:
        time.sleep(wait)
    err = None
    for attempt in range(2):
        try:
            r = sess.get(url, timeout=30, allow_redirects=True)
            _last_request[0] = time.time()
            return r
        except Exception as e:  # network error
            _last_request[0] = time.time()
            err = e
            time.sleep(3)
    print("   request failed:", url[:90], "-", err)
    return None


def to_int(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return int(round(v))
    s = str(v).strip().replace(",", "")
    m = re.fullmatch(r"\$?\s*(\d+)\.\d+", s)
    if m:
        return int(m.group(1))
    d = re.sub(r"[^\d]", "", s)
    return int(d) if d else None


def _walk(o):
    if isinstance(o, dict):
        yield o
        for v in o.values():
            yield from _walk(v)
    elif isinstance(o, list):
        for v in o:
            yield from _walk(v)


def html_text(html):
    h = re.sub(r"(?is)<(script|style|noscript)[^>]*>.*?</\1>", " ", html)
    h = re.sub(r"(?s)<[^>]+>", "\n", h)
    h = unescape(h)
    h = re.sub(r"[ \t\r\f\v\u00a0]+", " ", h)
    return re.sub(r"\n\s*", "\n", h).strip()


def parse_listing(html):
    """Best-effort read of one listing page. Returns a dict (maybe with few keys); never raises."""
    info = {}
    try:
        for m in re.finditer(r'(?is)<script[^>]+type=["\']application/ld\+json["\'][^>]*>(.*?)</script>', html):
            try:
                data = json.loads(m.group(1))
            except ValueError:
                continue
            for o in _walk(data):
                if "price" in o and ("priceCurrency" in o or "Offer" in str(o.get("@type"))):
                    info.setdefault("price", to_int(o["price"]))
                if "mileageFromOdometer" in o:
                    mo = o["mileageFromOdometer"]
                    val = to_int(mo.get("value") if isinstance(mo, dict) else mo)
                    if val is not None and isinstance(mo, dict) and str(mo.get("unitCode", "")).upper() == "SMI":
                        val = int(val * 1.609)
                    info.setdefault("km", val)
                if "vehicleTransmission" in o:
                    info.setdefault("trans_field", str(o["vehicleTransmission"]))
                if "bodyType" in o:
                    info.setdefault("body_raw", str(o["bodyType"]))
                for k in ("vehicleModelDate", "modelDate", "productionDate"):
                    if k in o and to_int(str(o[k])[:4]):
                        info.setdefault("year", to_int(str(o[k])[:4]))
                if o.get("addressLocality"):
                    info.setdefault("place", f"{o['addressLocality']}, {o.get('addressRegion', '')}".strip(", "))
                if o.get("name") and any(t in str(o.get("@type")) for t in ("Car", "Vehicle", "Product")):
                    info.setdefault("title", str(o["name"]))
                    if isinstance(o.get("description"), str) and len(o["description"]) > 40:
                        info.setdefault("desc", o["description"])
                    img = o.get("image")
                    if isinstance(img, list) and img:
                        img = img[0]
                    if isinstance(img, dict):
                        img = img.get("url") or img.get("contentUrl")
                    if isinstance(img, str) and img.startswith("http"):
                        info.setdefault("image", img)
    except Exception:
        pass

    try:                                                # Autotrader: full ad data in the Next.js data block
        m = re.search(r'(?is)<script[^>]+id=["\']__NEXT_DATA__["\'][^>]*>(.*?)</script>', html)
        ld = json.loads(m.group(1))["props"]["pageProps"].get("listingDetails") if m else None
        if isinstance(ld, dict):
            veh = ld.get("vehicle") or {}
            raw = veh.get("rawData") or {}
            if ld.get("description"):
                d0 = re.sub(r"(?i)<br\s*/?>|</li>|</p>", "\n", str(ld["description"]))
                info["desc"] = unescape(re.sub(r"<[^>]+>", " ", d0)).strip()
            if veh.get("modelYear"):
                info["year"] = to_int(veh["modelYear"])
            if veh.get("transmissionType"):
                info["trans_field"] = str(veh["transmissionType"])
            bt = ((raw.get("bodyType") or {}).get("raw") if isinstance(raw.get("bodyType"), dict) else None) or veh.get("bodyType")
            if bt:
                info["body_raw"] = str(bt)
            imgs = ld.get("images") or []
            if imgs and isinstance(imgs[0], str):
                info["image"] = imgs[0]
            if re.match(r"\d{4}-\d{2}-\d{2}", str(ld.get("createdTimestampWithOffset") or "")):
                info["listed"] = str(ld["createdTimestampWithOffset"])[:10]
    except Exception:
        pass
    try:                                                # NaPista ('offer') / Mobiauto ('deal'): ad data in the page
        m = re.search(r'(?is)<script[^>]+id=["\']__NEXT_DATA__["\'][^>]*>(.*?)</script>', html)
        pp = json.loads(m.group(1))["props"]["pageProps"] if m else {}
        info.update(br_ad(pp))
    except Exception:
        pass
    if "listed" not in info:                            # Kijiji: date the ad was posted (sortingDate moves on bumps)
        m = re.search(r'"activationDate"\s*:\s*"(\d{4}-\d{2}-\d{2})', html)
        if m:
            info["listed"] = m.group(1)
    page_title = re.search(r"(?is)<title[^>]*>(.*?)</title>", html)
    page_title = unescape(re.sub(r"\s+", " ", page_title.group(1))).strip() if page_title else ""
    if page_title and not re.search(r"\b(19[89]\d|20[0-3]\d)\b", info.get("title") or ""):
        info["title"] = page_title.split("|")[0].strip() or info.get("title")   # Autotrader: "2008 BMW M3 | $ ..."
    og = re.search(r'<meta[^>]+property=["\']og:image["\'][^>]+content=["\']([^"\']+)', html, re.I) or \
        re.search(r'<meta[^>]+content=["\']([^"\']+)["\'][^>]+property=["\']og:image', html, re.I)
    if og and re.search(r"\.(jpe?g|png|webp)|/images/", og.group(1), re.I):
        info["image"] = unescape(og.group(1))         # Kijiji: JSON-LD 'image' can be the ad's own address

    text = html_text(html)
    info["text"] = text
    if info.get("price") is None:
        m = re.search(r'(?:product:price:amount|og:price:amount)["\'][^>]*content=["\']([\d.,]+)', html, re.I)
        if m:
            info["price"] = to_int(m.group(1))
    if info.get("price") is None:
        m = re.search(r"\$\s?(\d{1,3}(?:,\d{3})+|\d{4,6})\b", text)
        if m:
            info["price"] = to_int(m.group(1))
    if info.get("km") is None:
        m = re.search(r"(?:Mileage|Kilom[e\u00e8]trage|Odometer)\s*:?\s*([\d][\d,\s.]{2,9}?)\s*km", text, re.I)
        if m:
            info["km"] = to_int(m.group(1))
    if info.get("trans_field") is None:
        m = re.search(r"Transmission\s*:?\s*([A-Za-z][A-Za-z0-9\- /()]{2,30})", text)
        if m:
            info["trans_field"] = m.group(1).strip()
    if info.get("body_raw") is None:
        m = re.search(r"Body (?:type|style)\s*:?\s*([A-Za-z][A-Za-z\- ]{2,25})", text, re.I)
        if m:
            info["body_raw"] = m.group(1).strip()
    if info.get("title") is None:
        m = re.search(r"(?is)<h1[^>]*>(.*?)</h1>", html) or re.search(r"(?is)<title[^>]*>(.*?)</title>", html)
        if m:
            info["title"] = re.sub(r"\s+", " ", unescape(re.sub(r"<[^>]+>", " ", m.group(1)))).strip()
    if info.get("year") is None and info.get("title"):
        m = re.search(r"\b(19[89]\d|20[0-3]\d)\b", info["title"])
        if m:
            info["year"] = int(m.group(1))
    if info.get("place") is None:
        m = re.search(r"([A-Z][\w.'\- ]{2,30}),\s?(" + PROV + r")\b", text[:1500])
        if m:
            info["place"] = f"{m.group(1).strip()}, {m.group(2)}"
    if not info.get("desc"):
        dm = re.search(r"(?:Vehicle Description|Description)\s*\n(.{40,3000})", text, re.S)
        if dm:
            info["desc"] = dm.group(1)
    if not info.get("desc"):                            # description stored as JSON in the page (Kijiji, CarPages)
        cands = [m.group(1) for m in re.finditer(r'"description"\s*:\s*"((?:[^"\\]|\\.){40,5000})"', html)]
        if cands:
            try:
                info["desc"] = json.loads('"' + max(cands, key=len) + '"')
            except ValueError:
                info["desc"] = max(cands, key=len)
    if not info.get("image"):
        m = re.search(r'<meta[^>]+property=["\']og:image["\'][^>]+content=["\']([^"\']+)', html, re.I) or \
            re.search(r'<meta[^>]+content=["\']([^"\']+)["\'][^>]+property=["\']og:image', html, re.I)
        if m:
            info["image"] = unescape(m.group(1))
    return info


def _place(city, uf):
    city = (city or "").strip()
    if city.isupper():
        city = " ".join(w if i == 0 or w not in ("de", "da", "do", "das", "dos", "e") else w
                        for i, w in enumerate(city.title().split()))
        city = re.sub(r"(?<=\s)(De|Da|Do|Das|Dos|E)(?=\s)", lambda m: m.group(1).lower(), city)
    return ", ".join(x for x in (city, (uf or "").strip()) if x)


def br_ad(pp):
    """Ad fields from a NaPista ('offer') or Mobiauto ('deal') page. Prices/km are numbers in the data."""
    info = {}
    off, deal = pp.get("offer"), pp.get("deal")
    if isinstance(off, dict) and isinstance(off.get("offer"), dict):
        off = off["offer"]                              # some pages nest the ad one level down
    if isinstance(off, dict) and off.get("id"):
        loc = off.get("location") or {}
        info.update(title=f'{off.get("modelYear", "")} {off.get("versionName", "")}'.strip(),
                    price=to_int(off.get("price")), km=to_int(off.get("mileage")), year=to_int(off.get("modelYear")),
                    trans_field=str((off.get("transmissionType") or {}).get("name") or ""),
                    place=_place(loc.get("city"), loc.get("uf")), status=str(off.get("status") or ""))
        if off.get("description"):
            info["desc"] = str(off["description"]).replace("\\n", "\n")
        if re.match(r"\d{4}-\d{2}-\d{2}", str(off.get("createdDate") or "")):
            info["listed"] = str(off["createdDate"])[:10]
        ph = sorted(off.get("photos") or [], key=lambda x: x.get("order", 0))
        if ph and str(ph[0].get("url", "")).startswith("https://"):
            info["image"] = ph[0]["url"]
        info["body_raw"] = str(off.get("versionName") or "")
    elif isinstance(deal, dict) and deal.get("id"):
        info.update(title=" ".join(str(x) for x in (deal.get("modelYear"), deal.get("makeName"), deal.get("modelName"),
                                                     deal.get("trimName"), deal.get("bodystyleName")) if x),
                    price=to_int(deal.get("price")), km=to_int(deal.get("km")), year=to_int(deal.get("modelYear")),
                    trans_field=str(deal.get("transmissionName") or ""), body_raw=str(deal.get("bodystyleName") or ""),
                    place=_place(deal.get("dealerCity"), deal.get("dealerState")))
        if deal.get("comments"):
            info["desc"] = str(deal["comments"])
        imgs = pp.get("images") or []
        if imgs and str(imgs[0].get("imageSrc", "")).startswith("https://"):
            info["image"] = imgs[0]["imageSrc"]
    return {k: v for k, v in info.items() if v not in (None, "")}


MOBI_IMG = "https://image1.mobiauto.com.br/images/api/images/v1.0/{}/transform/fl_progressive,f_webp,q_70,w_1080"


def br_search_items(source, html):
    """Ads in a NaPista / Mobiauto search page (main results only), with what a card needs."""
    m = re.search(r'(?is)<script[^>]+id=["\']__NEXT_DATA__["\'][^>]*>(.*?)</script>', html)
    if not m:
        return []
    try:
        pp = json.loads(m.group(1))["props"]["pageProps"]
    except (ValueError, KeyError):
        return []
    out = []
    if source == "napista":
        for o in (pp.get("searchResult") or {}).get("offers") or []:
            loc = o.get("location") or {}
            ph = sorted(o.get("photos") or [], key=lambda x: x.get("order", 0))
            out.append({"key": str(o.get("id")), "title": f'{o.get("modelYear", "")} {o.get("versionName", "")}',
                        "year": to_int(o.get("modelYear")), "gear": str(o.get("versionName") or ""),
                        "price": to_int(o.get("price")), "km": to_int(o.get("mileage")),
                        "place": _place(loc.get("city"), loc.get("uf")), "body_raw": str(o.get("versionName") or ""),
                        "photo": ph[0].get("url", "") if ph else "",
                        "listed": str(o.get("createdDate") or "")[:10] if re.match(r"\d{4}-", str(o.get("createdDate") or "")) else ""})
    else:
        for o in (pp.get("deals") or {}).get("results") or []:
            tr = o.get("trim") or {}
            body = (tr.get("bodystyle") or {}).get("name") or ""
            title = " ".join(str(x) for x in ((tr.get("model") or {}).get("year"), (tr.get("make") or {}).get("name"),
                                                (tr.get("model") or {}).get("name"), tr.get("name"), body) if x)
            loc = (o.get("dealer") or {}).get("location") or o.get("location") or {}
            img = (o.get("images") or [{}])[0].get("imageId")
            out.append({"key": str(o.get("id")), "title": title,
                        "year": to_int((tr.get("model") or {}).get("year") or tr.get("productionYear")),
                        "gear": str((tr.get("transmission") or {}).get("name") or ""),
                        "price": to_int(o.get("price")), "km": to_int(o.get("km")),
                        "place": _place(loc.get("city"), loc.get("state") or loc.get("uf")), "body_raw": body,
                        "photo": MOBI_IMG.format(img) if img else "", "listed": ""})
    return out


AUTO_IN_TITLE = re.compile(r"\baut\b|\baut\.|\bauto\b|autom[a\u00e1]tic|automatizad|\bcvt\b|tiptronic|steptronic|\bdct\b|\bsmg\b", re.I)


def is_auto_gear(gear):
    return bool(AUTO_IN_TITLE.search(gear or "")) and not re.search(r"manual|\bmec", gear or "", re.I)


def resync_families(rows):
    """A generation renamed or re-cut in the config (NSX 91-01 became NA1 / NA2): saved ads move to the new one."""
    n = 0
    for r in rows:
        if not r["family"] or r["family"] in FAMILY_STYLE or not str(r["year"]).isdigit():
            continue
        slug = r["url"].rsplit("/", 2)[-2] if r["source"] == "kijiji" else r["url"].rsplit("/", 1)[-1]
        key = detect_key(r["note"], slug, int(r["year"]))
        if not key:
            continue
        try:
            r["family"], r["gen"], r["name"] = classify(key, int(r["year"]), r["body"] or "unk", f'{r["note"]} {slug}')
        except ValueError:
            continue
        n += 1
    return n


def family_of(key, year, title):
    try:
        return classify(key, year, fix_body(key, body_of("", title), title), title)[0]
    except Exception:
        return ""


def photo_of(row, html, info):
    """Autotrader: '<listing>_<image>[.ext]' (first photo of THIS listing). Others: full image URL."""
    if row["source"] == "autotrader":
        for m in PHOTO_RE.finditer(html):
            if m.group(1) == row["listing_id"]:
                ext = m.group(3).lower()
                return f"{m.group(1)}_{m.group(2)}" + ("" if ext in ("jpg", "jpeg") else "." + ext)
        return ""
    img = info.get("image") or ""
    return img if img.startswith("https://") else ""


GONE_TEXT = re.compile(r"no longer available|has been sold|listing (?:has )?(?:expired|been removed)|"
                       r"ad (?:is no longer|has been deleted)|n.est plus disponible|annonce (?:a expir|est expir)", re.I)


def check_listing(sess, row):
    """-> ('ok', info) | ('gone', reason) | ('fail', reason)"""
    r = fetch(sess, row["url"])
    if r is None:
        return "fail", "no response"
    if r.status_code in (404, 410):
        return "gone", f"HTTP {r.status_code}"
    if r.status_code != 200:
        return "fail", f"HTTP {r.status_code}"
    final = r.url.split("?")[0]
    if row["source"] == "autotrader":
        if row["listing_id"] not in final:
            return "gone", "redirected to a search page (ad removed)"
    elif row["source"] in ("kijiji", "carpages", "napista") and id_key(row) not in final:
        return "gone", "redirected away from the ad"
    elif row["source"] == "mobiauto" and f"/detalhes/{id_key(row)}" not in final:
        return "gone", "redirected away from the ad"
    info = parse_listing(r.text)
    if info.get("status") and info["status"].upper() not in ("PUBLISHED", "ACTIVE", "ONLINE"):
        return "gone", f"ad status {info['status']}"
    if GONE_TEXT.search(info.get("text", "")[:6000]) and info.get("price") is None:
        return "gone", "page says the listing is no longer available"
    if info.get("price") is None:
        return "fail", "could not read the price from the page"
    if info.get("km") and row["km"] and abs(info["km"] - int(row["km"])) > 2000:
        return "gone", f"odometer now {info['km']} km: a different car on the same address"
    info["_html"] = r.text
    return "ok", info


def apply_page_details(row, info, today):
    """Flags, trim and photo from a freshly read page (only overwrite when the page gave something)."""
    desc = info.get("desc") or ""
    if len(desc) >= 40:
        red, green = flags_of((info.get("title") or "") + "\n" + desc)
        row["red_flags"], row["green_flags"], row["flags_read"] = red, green, today
        row["desc"] = re.sub(r"\s+", " ", desc).strip()[:700]
    t = trim_of(row["family"], (info.get("title") or "") + " " + row["url"].rsplit("/", 1)[-1])
    if t:
        row["trim"] = t
    pid = photo_of(row, info.get("_html", ""), info)
    if pid and pid != row["photo_id"]:
        row["photo_id"], row["photo_alt"] = pid, ""
    ls = info.get("listed")
    if ls and "2000-01-01" <= ls <= today:
        row["listed"] = min(ls, row["listed"]) if row.get("listed") else ls


def hist_points(text):
    pts = []
    for part in (text or "").split(";"):
        dd, _, pp = part.partition(":")
        if dd and pp.strip().isdigit():
            pts.append((dd.strip(), int(pp)))
    return pts


def hist_add(row, day, price):
    """Add a price to the row's history when it differs from the last one (keeps the latest 40 changes)."""
    if price in (None, "") or int(price) < MIN_PRICE:
        return
    pts = hist_points(row.get("price_hist"))
    if not pts or pts[-1][1] != int(price):
        pts.append((day, int(price)))
    row["price_hist"] = ";".join(f"{dd}:{pp}" for dd, pp in pts[-40:])


def backfill_hist(rows):
    """One-off: rebuild price histories from the saved daily snapshots for rows that have none yet."""
    todo = {r["listing_id"]: r for r in rows if not r.get("price_hist") and r["price"] and r["status"] != "skip"}
    if not todo:
        return 0
    seen = {}
    for f in sorted(glob.glob(os.path.join(DATA_DIR, "car_listings_*.csv"))):
        m = re.search(r"(\d{4})(\d{2})(\d{2})\.csv$", f)
        if not m:
            continue
        day = "-".join(m.groups())
        with open(f, newline="", encoding="utf-8") as fh:
            for r in csv.DictReader(fh):
                lid, pr = r.get("listing_id"), (r.get("price") or "").strip()
                if lid in todo and pr.isdigit() and int(pr) >= MIN_PRICE and r.get("status") in ("active", "unverified", "gone", ""):
                    seen.setdefault(lid, []).append((day, int(pr)))
    for lid, row in todo.items():
        pts = []
        for day, pr in sorted(seen.get(lid, [])):
            if not pts or pts[-1][1] != pr:
                pts.append((day, pr))
        if not pts:
            if int(row["price"]) < MIN_PRICE:
                continue
            pts = [(row["first_seen"] or "", int(row["price"]))]
        elif row["first_seen"] and row["first_seen"] < pts[0][0]:
            pts[0] = (row["first_seen"], pts[0][1])
        if pts[-1][1] != int(row["price"]) and int(row["price"]) >= MIN_PRICE:
            pts.append((row["last_seen"] or pts[-1][0], int(row["price"])))
        row["price_hist"] = ";".join(f"{dd}:{pp}" for dd, pp in pts if dd)
    return len(todo)


# --------------------------------------------------------------------------------------------
# Daily steps
# --------------------------------------------------------------------------------------------
SKIP_SOURCES = set()


def refresh_existing(sess, rows, today, stats):
    cand_gone = []
    for row in rows:
        if row["source"] in SKIP_SOURCES:
            continue
        st = row["status"]
        if st == "gone":
            last = d(row["last_seen"])
            if not last or (d(today) - last).days > GONE_RECHECK_DAYS:
                continue
        elif st not in ("active", "unverified"):
            continue
        seen = SEARCH_SEEN.get(row["listing_id"])
        if seen and seen.get("price") and not (seen.get("km") and row["km"] and abs(int(seen["km"]) - int(row["km"])) > 2000):
            outcome, res = "ok", {"price": seen["price"], "km": seen["km"] if (seen.get("km") or 0) >= 500 else None,
                                  "_search": True}
            stats["seen_in_search"] = stats.get("seen_in_search", 0) + 1
        else:
            if time_up():                                # out of time: this ad keeps its status until the next run
                stats["late"] = stats.get("late", 0) + 1
                continue
            outcome, res = check_listing(sess, row)
            stats["checked"] += 1
        if outcome == "ok":
            hidden = res["price"] < MIN_PRICE
            if hidden:                                  # keep the last real price; flag it below
                real = [p for _, p in hist_points(row.get("price_hist")) if p >= MIN_PRICE]
                res["price"] = next((int(v) for v in (row["price"], row["prev_price"]) if v and int(v) >= MIN_PRICE),
                                    real[-1] if real else res["price"])
            if row["price"] and str(res["price"]) != str(row["price"]) and st == "active":
                row["prev_price"] = row["price"]
                stats["price_changes"] += 1
            if not row.get("price_hist") and row["price"]:
                hist_add(row, row["first_seen"] or today, row["price"])
            row["price"] = str(res["price"])
            hist_add(row, today, row["price"])
            if res.get("km"):
                row["km"] = str(res["km"])
            row["status"] = "active"
            row["last_seen"] = row["verified"] = today
            if not res.get("_search"):
                apply_page_details(row, res, today)
            if hidden:
                red = [f for f in (row.get("red_flags") or "").split("; ") if f and f != "Price hidden"]
                row["red_flags"] = "; ".join(red + ["Price hidden"])
            if st == "gone":
                stats["back"] += 1
        elif outcome == "gone":
            if st != "gone":
                cand_gone.append((row, res))
        else:
            stats["failed"] += 1
            print("   not read:", row["listing_id"][:20], row["name"], "-", res)
    # Safety valve: a block page or a site redesign must not wipe the whole list.
    limit = max(5, int(0.35 * max(1, stats["checked"])))
    if len(cand_gone) > limit:
        print(f"!! {len(cand_gone)} listings look gone out of {stats['checked']} checked - more than "
              f"expected, so none were marked gone. The site may be blocking or has changed.")
        stats["failed"] += len(cand_gone)
        return
    for row, reason in cand_gone:
        row["status"] = "gone"
        row["note"] = (row["note"] + f"; gone {today}: {reason}").strip("; ")
        stats["gone"] += 1


def page_url(source, base, i):
    if i == 0:
        return base
    if source == "kijiji":
        return re.sub(r"/(k0c\d+l\d+)", rf"/page-{i + 1}/\1", base)
    if source == "napista":
        return f"{base}{'&' if '?' in base else '?'}pn={i + 1}"
    if source == "mobiauto":                             # .../civic/pagina-2  or  .../civic/pagina-2/s/cambio-manual
        head, sep_s, tail = base.partition("/s/")
        return f"{head}/pagina-{i + 1}" + (f"/s/{tail}" if sep_s else "")
    sep = "&" if "?" in base else "?"
    return f"{base}{sep}page={i + 1}"


def kijiji_slug_may_match(url):
    """Kijiji addresses carry the ad's title ('.../2023-bmw-x1-xdrive28i-gr-sup.../1744574968'). Keyword searches
    return many other cars, so an address with a year that names none of the tracked cars is not opened."""
    slug = url.rsplit("/", 2)[-2]
    m = re.search(r"\b(19[5-9]\d|20[0-3]\d)\b", _norm(slug))
    return not m or detect_key(slug, "", int(m.group(1)), ignore_make=True) is not None


def search_ids(sess, known, stats):
    found = {}
    SEARCH_SEEN.clear()
    AUTO_FOUND.clear()
    for source, base in SEARCHES:
        rx = LINK_RE[source]
        seen_here, status = set(), None
        for i in range(PAGES_JSON if source in JSON_SEARCH else PAGES_PER_SEARCH):
            if time_up():
                break
            r = fetch(sess, page_url(source, base, i))
            if (i == 0 and r is not None and source in JSON_SEARCH
                    and len(r.url.split("?")[0].rstrip("/")) < len(base.split("?")[0].rstrip("/"))):
                status = "ignored by site"               # Mobiauto sends an unknown model to its all-cars page
                break
            if r is None or r.status_code != 200:
                if i == 0:
                    status = f"HTTP {getattr(r, 'status_code', 'none')}"
                break
            page_ids = {}
            for m in rx.finditer(r.text):
                page_ids[make_id(source, m.group(2))] = HOSTS[source] + m.group(1)
            if source in JSON_SEARCH:
                items = br_search_items(source, r.text)
                fresh = {make_id(source, it["key"]) for it in items} - seen_here
                if not fresh:
                    break                                # last page (or paging not supported)
                seen_here |= fresh
                for it in items:
                    lid = make_id(source, it["key"])
                    url = page_ids.get(lid) or (f"{HOSTS[source]}/anuncios/{it['key']}" if source == "napista" else None)
                    key = detect_key(it["title"], "", it["year"]) if url else None
                    fam = family_of(key, it["year"], it["title"]) if key else ""
                    auto = is_auto_gear(it["gear"])
                    if not key or (auto and fam not in ALLOW_AUTO):
                        stats["prefiltered"] = stats.get("prefiltered", 0) + 1     # not opened
                        continue
                    it.update(url=url, source=source, family=fam)
                    SEARCH_SEEN[lid] = it
                    if lid in known or lid in found or lid in AUTO_FOUND:
                        continue
                    if auto:
                        AUTO_FOUND[lid] = (it, key)       # added from the search data, page read later
                    else:
                        found[lid] = (source, url, len(stats["searches"]))
                continue
            fresh = set(page_ids) - seen_here
            if not fresh:
                break                                    # paging not supported this way / last page
            seen_here |= fresh
            for lid in fresh:
                if lid in known:
                    continue
                if source == "kijiji" and not kijiji_slug_may_match(page_ids[lid]):
                    stats["prefiltered"] = stats.get("prefiltered", 0) + 1     # 2023 X1 under 'bmw-1m': not opened
                    continue
                found.setdefault(lid, (source, page_ids[lid], len(stats["searches"])))
        stats["searches"].append((source, base, status or len(seen_here)))
    return found


def discover_new(sess, rows, today, stats, max_new, found=None):
    known = {r["listing_id"] for r in rows}
    if found is None:
        found = search_ids(sess, known, stats)
    stats["candidates"] = len(found)
    added, skipped = [], []
    queues = {}
    for lid, (source, url, n) in found.items():
        queues.setdefault(n, []).append((lid, source, url))
    order = []
    while any(queues.values()):                          # one ad from each search in turn: the cap is shared fairly
        for n in sorted(queues):
            if queues[n]:
                order.append(queues[n].pop(0))
    for lid, source, url in order[:max_new]:
        if time_up():
            break
        r = fetch(sess, url)
        if r is None or r.status_code != 200:
            continue
        info = parse_listing(r.text)
        slug = url.rsplit("/", 1)[-1] if source != "kijiji" else url.rsplit("/", 2)[-2]
        key = detect_key(info.get("title", ""), slug, info.get("year"))
        reason = ""
        if not key:
            reason = "not one of the cars on the list"
        elif info.get("price") is None or info.get("km") is None or info["price"] < MIN_PRICE:
            reason = "price or km not readable"
        elif info["km"] < 500 or (dt.date.today().year - (info.get("year") or 0) >= 5 and info["km"] < 5000):
            reason = f"km not real ({info['km']} km on a {info.get('year')})"
        elif info["price"] > MAX_PRICE.get(COUNTRY["country"], 10 ** 7):
            reason = f"price not real ({info['price']})"
        if not reason and any(re.search(rx, _norm(info.get("body_raw", ""))) for rx in key.get("none_body", [])):
            reason = f"body not on the list ({info.get('body_raw')})"   # a hatchback Yaris, a coupe Civic
        if not reason:
            body = fix_body(key, body_of(info.get("body_raw", ""), slug + " " + info.get("title", "")),
                            slug + " " + info.get("title", ""))
            try:
                fam, gen, name = classify(key, info["year"], body, f'{info.get("title", "")} {slug}')
            except ValueError:                           # a 2012 335i sedan is an F30, not an E90
                reason = f"body not on the list for that year ({body})"
        if not reason:
            trans = classify_trans(fam, info.get("trans_field", ""),     # the title counts too: '... 6-SPEED MANUAL'
                                   f'{info.get("title", "")} {slug.replace("-", " ")}. {info.get("desc", "")}')
            if trans == "a" and fam not in ALLOW_AUTO:
                reason = "automatic / gearbox not stated"
        if reason:
            if reason != "price or km not readable":      # remember, so it is not opened again tomorrow
                skipped.append({**{c: "" for c in COLS}, "listing_id": lid, "source": source, "url": url,
                                "note": reason + " " + SKIP_TAG, "first_seen": today, "last_seen": today, "status": "skip"})
            continue
        row = {**{c: "" for c in COLS},
               "listing_id": lid, "source": source, "family": fam, "gen": gen, "name": name,
               "year": str(info["year"]), "body": body, "price": str(info["price"]), "km": str(info["km"]),
               "trans": trans, "place": info.get("place", ""), "url": url,
               "note": (info.get("title") or "")[:120], "first_seen": today, "last_seen": today,
               "status": "active", "verified": today}
        info["_html"] = r.text
        apply_page_details(row, info, today)
        row["page_read"] = today
        hist_add(row, today, row["price"])
        added.append(row)
    opened = min(len(order), max_new)
    # automatics (FB Civic, Yaris sedan) straight from the search data: no page opened today
    n_auto = 0
    for lid, (it, key) in AUTO_FOUND.items():
        if lid in known or not it.get("price") or not it.get("km") or it["price"] < MIN_PRICE:
            continue
        if it["km"] < 500 or (dt.date.today().year - (it.get("year") or 0) >= 5 and it["km"] < 5000):
            continue
        body = fix_body(key, body_of(it.get("body_raw", ""), it["title"]), it["title"])
        try:
            fam, gen, name = classify(key, it["year"], body, it["title"])
        except ValueError:
            continue
        row = {**{c: "" for c in COLS},
               "listing_id": lid, "source": it["source"], "family": fam, "gen": gen, "name": name,
               "year": str(it["year"]), "body": body, "price": str(it["price"]), "km": str(it["km"]), "trans": "a",
               "place": it.get("place", ""), "url": it["url"], "note": it["title"][:120], "photo_id": it.get("photo", ""),
               "first_seen": today, "last_seen": today, "status": "active", "verified": today}
        if it.get("listed") and "2000-01-01" <= it["listed"] <= today:
            row["listed"] = it["listed"]
        row["trim"] = trim_of(fam, it["title"]) or ""
        hist_add(row, today, row["price"])
        added.append(row)
        n_auto += 1
    stats["new_auto"] = n_auto
    # what is left of today's page budget: read ads added from search data (description, flags, true gearbox)
    budget = max(0, max_new - opened)
    todo = [r for r in rows + added if r["status"] == "active" and r["source"] in JSON_SEARCH
            and not r.get("page_read") and not r.get("desc")]
    todo.sort(key=lambda r: r["first_seen"], reverse=True)
    n_read = 0
    for row in todo[:budget]:
        if time_up():
            break
        r = fetch(sess, row["url"])
        if r is None or r.status_code != 200:
            continue
        info = parse_listing(r.text)
        info["_html"] = r.text
        apply_page_details(row, info, today)
        row["page_read"] = today
        t2 = classify_trans(row["family"], info.get("trans_field", ""), f'{row["note"]}. {info.get("desc", "")}')
        if row["trans"] == "a" and t2 != "a":
            row["trans"] = t2                               # the page says manual after all
        n_read += 1
    stats["details_read"] = n_read
    stats["new"] = len(added)
    stats["skipped"] = len(skipped)
    return added + skipped


# --------------------------------------------------------------------------------------------
# Photos
# --------------------------------------------------------------------------------------------
def photo_urls(row):
    """Candidate image addresses, small JPEG rendition first (JPEG works in every Excel)."""
    pid = row["photo_id"]
    out = []
    for p in (pid, row.get("photo_alt", "")):
        if not p:
            continue
        if p.startswith("http"):
            out.append(p)
            continue
        stem, ext = (p.rsplit(".", 1) + ["jpg"])[:2] if "." in p else (p, "jpg")
        out.append(f"{PHOTO_BASE}{stem}.{ext}/250x188.jpg")
        out.append(f"{PHOTO_BASE}{stem}.{ext}/250x188.webp")
    return out


def photo_file(url):
    return os.path.join(PHOTO_DIR, hashlib.sha1(url.encode()).hexdigest()[:16] + ".jpg")


def download_photos(sess, rows, stats):
    """Save a 240x180 JPEG thumbnail per active car into .\\photos (cached by address)."""
    if PILImage is None:
        print("   Pillow not installed: photos stay as IMAGE() formulas (pip install pillow)")
        return
    os.makedirs(PHOTO_DIR, exist_ok=True)
    for row in rows:
        if row["status"] != "active":
            continue
        for url in photo_urls(row):
            path = photo_file(url)
            if os.path.exists(path):
                break
            r = fetch(sess, url, delay=0.3)
            if r is None or r.status_code != 200 or len(r.content) < 500:
                continue
            try:
                im = PILImage.open(io.BytesIO(r.content)).convert("RGB")
                im.thumbnail((240, 180))
                im.save(path, "JPEG", quality=82)
                stats["photos_new"] += 1
                break
            except Exception:
                continue


def local_photo(row):
    for url in photo_urls(row):
        p = photo_file(url)
        if os.path.exists(p):
            return p
    return None


def photo_formula(row):
    urls = [u for u in photo_urls(row) if u.endswith(".jpg") or not u.startswith(PHOTO_BASE)][:2]
    if not urls:
        return None
    f = f'_xlfn.IMAGE("{urls[-1]}","photo",0)'
    for u in reversed(urls[:-1]):
        f = f'IFERROR(_xlfn.IMAGE("{u}","photo",0),{f})'
    return "=" + f


# --------------------------------------------------------------------------------------------
# Fair-price model  (log price = family level + km slope + year slope + convertible effect)
# --------------------------------------------------------------------------------------------
def _solve(A, b):
    n = len(A)
    M = [row[:] + [b[i]] for i, row in enumerate(A)]
    for c in range(n):
        p = max(range(c, n), key=lambda r: abs(M[r][c]))
        if abs(M[p][c]) < 1e-12:
            continue
        M[c], M[p] = M[p], M[c]
        for r in range(n):
            if r != c and M[r][c]:
                f = M[r][c] / M[c][c]
                for k in range(c, n + 1):
                    M[r][k] -= f * M[c][k]
    return [M[i][n] / M[i][i] if abs(M[i][i]) > 1e-12 else 0.0 for i in range(n)]


def _ols(X, y, ridge=1e-6):
    k = len(X[0])
    A = [[sum(x[i] * x[j] for x in X) + (ridge if i == j else 0) for j in range(k)] for i in range(k)]
    b = [sum(x[i] * yy for x, yy in zip(X, y)) for i in range(k)]
    return _solve(A, b)


def fit_fair_price(active, adj=None):
    """Returns ({family: (a, b_km, b_year, n)}, b_conv, resid_sd). km in units of 10,000 km, year-2000.
    adj: {state: log price level}; prices are divided by their state's level before fitting (Brazil)."""
    adj = adj or {}

    def ly(r):
        return math.log(float(r["price"])) - adj.get(area_of(r), 0.0) - VADJ.get((mkey(r), version_of(r)), 0.0)
    data = [r for r in active if r["price"] and r["km"] and r["year"]
            and not re.search(r"Rebuilt|Salvage", r.get("red_flags", ""))]
    if len(data) < 8:
        return {}, 0.0, 0.2
    fams = sorted({mkey(r) for r in data})
    # the convertible effect can only be measured when some model is listed both as convertible and fixed roof
    # (Brazil: one convertible, so it would just copy that model's level)
    conv_ok = any(len({r["body"] == "conv" for r in data if mkey(r) == f}) == 2 for f in fams)

    def feats(r):
        return float(r["km"]) / 1e4, float(r["year"]) - 2000, 1.0 if r["body"] == "conv" and conv_ok else 0.0

    use = data
    for _ in range(2):                                 # fit, drop strong outliers, refit
        X = [[1.0 if mkey(r) == f else 0.0 for f in fams] + list(feats(r)) for r in use]
        y = [ly(r) for r in use]
        beta = _ols(X, y)
        res = [yy - sum(a * b for a, b in zip(x, beta)) for x, yy in zip(X, y)]
        sd = (sum(e * e for e in res) / max(1, len(res) - len(beta))) ** 0.5
        keep = [r for r, e in zip(use, res) if abs(e) <= 2.5 * sd]
        if len(keep) == len(use):
            break
        use = keep
    b_km, b_yr, b_conv = beta[len(fams):]
    model = {}
    K = 15.0                                            # shrink family slopes toward the pooled slopes
    for f in fams:
        fr_fit = [r for r in use if mkey(r) == f]
        bk, by = b_km, b_yr
        n_fit = len(fr_fit)
        if n_fit >= 12:
            Xf = [[1.0, feats(r)[0], feats(r)[1]] for r in fr_fit]
            yf = [ly(r) - b_conv * feats(r)[2] for r in fr_fit]
            _, bk_f, by_f = _ols(Xf, yf, ridge=1e-3)
            bk = (n_fit * bk_f + K * b_km) / (n_fit + K)
            by = (n_fit * by_f + K * b_yr) / (n_fit + K)
        fr = [r for r in data if mkey(r) == f]       # level = MEDIAN over every car of the family (robust)
        lv = sorted(ly(r) - bk * feats(r)[0] - by * feats(r)[1] - b_conv * feats(r)[2]
                    for r in fr)
        m = len(lv)
        a = lv[m // 2] if m % 2 else (lv[m // 2 - 1] + lv[m // 2]) / 2
        model[f] = (a, bk, by, m)
    return model, b_conv, sd


GEARBOX_LABEL = {"m": "\u2714 Manual", "c": "\u26A0 Conflict", "u": "? Unconfirmed", "a": "Automatic"}


def apply_split(r):
    """Move a car to its sub-generation (e.g. FA5 Civic -> FA5 Civic Si) when its title or trim matches, and give
    a generation that only came one way its body (Brazil's FB Civic Si: coupe; MR2 Spyder: roadster)."""
    for rx, sg in SPLITS.get(r["family"], []):
        if re.search(rx, f'{r.get("note", "")} {r.get("trim", "")}', re.I):
            r["family"], r["gen"], r["name"] = sg["family"], sg["name"].split(" ")[0], sg["name"]
            break
    if r["family"] in GEN_BODY:
        r["body"] = GEN_BODY[r["family"]]


def mkey(r):
    """Fair-price group: automatics are priced against automatics only."""
    return r["family"] + ("|auto" if r.get("trans") == "a" else "")


# --------------------------------------------------------------------------------------------
# Regional price levels (Brazil): how far each state's asking prices sit from the national fair price,
# for the same model, km and year. The app uses them for a local fair price around a chosen town.
# --------------------------------------------------------------------------------------------
REGION_OF = {uf: g for g, s in {"N": "AC AM AP PA RO RR TO", "NE": "AL BA CE MA PB PE PI RN SE",
                                "CO": "DF GO MS MT", "SE": "ES MG RJ SP", "S": "PR RS SC"}.items() for uf in s.split()}
K_AREA = 10.0          # a place with few cars leans on the wider area as if it had this many cars at its level
AREA_CLIP = 0.25       # one car moves a price level by at most 25% (typos, odd ads)
AREAS = {}             # filled by fit_with_areas(): state and region levels, towns table


def area_of(r):
    if COUNTRY["country"] != "BR":
        return None
    uf = (r.get("place") or "").rpartition(",")[2].strip().upper()
    return uf if uf in REGION_OF else None


K_VER = 6.0           # a version with few cars leans on the model's average as if it had this many at 0
VADJ = {}             # (model key, version) -> log price level against the model's average version
VERSIONS = {}         # export: {model key: {version: [level, cars]}}


def version_of(r):
    v = (r.get("trim") or "").strip().upper()
    return v or None


def version_levels(gaps):
    """gaps: [(row, log gap against the model without versions)]. Per model, each version's level, shrunk toward
    the model's average; the car-weighted mean over known versions is 0, so an ad without a version gets 0."""
    by = {}
    for r, e in gaps:
        v = version_of(r)
        if v:
            s = by.setdefault(mkey(r), {}).setdefault(v, [0.0, 0])
            s[0] += e
            s[1] += 1
    adj, exp_ = {}, {}
    for k, vs in by.items():
        raw = {v: s / (n + K_VER) for v, (s, n) in vs.items() if n >= 2}
        if len(raw) < 2:
            continue
        tot = sum(vs[v][1] for v in raw)
        mean = sum(raw[v] * vs[v][1] for v in raw) / tot
        for v in raw:
            adj[(k, v)] = raw[v] - mean
            exp_.setdefault(k, {})[v] = [round(raw[v] - mean, 4), vs[v][1]]
    return adj, exp_


def log_fair_base(model, b_conv, r):
    """Log of the model's fair price for this year, km and body, before version and region; or None."""
    m = model.get(mkey(r))
    try:
        km, year = float(r["km"]), float(r["year"])
    except (TypeError, ValueError):
        return None
    if not m or not km:
        return None
    a, bk, by, _ = m
    return a + bk * km / 1e4 + by * (year - 2000) + b_conv * (r["body"] == "conv")


def log_fair(model, b_conv, r):
    """Log of the national fair price (with the version's level), or None."""
    b = log_fair_base(model, b_conv, r)
    return None if b is None else b + VADJ.get((mkey(r), version_of(r)), 0.0)


def area_levels(gaps, normalize=True):
    """gaps: [(state, log gap)]. Region level = sum / (n + K): few cars -> near 0 (national).
    State level shrinks the same way toward its region. normalize: the median well-measured state = 0."""
    st, rg = {}, {}
    for uf, e in gaps:
        for d, k in ((st, uf), (rg, REGION_OF[uf])):
            s = d.setdefault(k, [0.0, 0])
            s[0] += e
            s[1] += 1
    reg = {g: rg.get(g, [0.0, 0])[0] / (rg.get(g, [0.0, 0])[1] + K_AREA) for g in sorted(set(REGION_OF.values()))}
    lv = {uf: (st.get(uf, [0.0, 0])[0] + K_AREA * reg[REGION_OF[uf]]) / (st.get(uf, [0.0, 0])[1] + K_AREA)
          for uf in REGION_OF}
    off = 0.0
    well = sorted(lv[uf] for uf in lv if st.get(uf, [0, 0])[1] >= 10)
    if normalize and well:
        m = len(well)
        off = well[m // 2] if m % 2 else (well[m // 2 - 1] + well[m // 2]) / 2
    return ({uf: v - off for uf, v in lv.items()}, {g: v - off for g, v in reg.items()},
            {uf: st.get(uf, [0, 0])[1] for uf in lv}, {g: rg.get(g, [0, 0])[1] for g in reg})


def fit_with_areas(active):
    """Fair-price model. Prices are first corrected for their version (LXR, EXR...: each version's level against the
    model's average) and, in Brazil, for their state's price level, so the national fair price is the typical
    state's, not pulled up by Rio or set by Sao Paulo's many ads. A few rounds of fitting settle the three together."""
    global AREAS, VADJ, VERSIONS
    VADJ, VERSIONS, AREAS = {}, {}, {}
    model, b_conv, sd = fit_fair_price(active)
    br = COUNTRY["country"] == "BR"
    if br:
        load_geo()
    usable = [r for r in active if r["price"] and r["year"] and not re.search(r"Rebuilt|Salvage", r.get("red_flags", ""))]

    def clip(x):
        return max(-AREA_CLIP, min(AREA_CLIP, x))
    lv = {}
    for _ in range(3):
        base = []
        for r in usable:
            lb = log_fair_base(model, b_conv, r)
            if lb is not None:
                base.append((r, math.log(float(r["price"])) - lb))
        if br:
            lv = area_levels([(area_of(r), clip(e - VADJ.get((mkey(r), version_of(r)), 0.0))) for r, e in base if area_of(r)])[0]
        VADJ, VERSIONS = version_levels([(r, clip(e - lv.get(area_of(r), 0.0))) for r, e in base])
        model, b_conv, sd = fit_fair_price(active, lv)
    if not br:
        return model, b_conv, sd
    usable = [r for r in usable if area_of(r)]

    def gaps():
        out = []
        for r in usable:
            lf = log_fair(model, b_conv, r)
            if lf is not None:
                out.append((r, clip(math.log(float(r["price"])) - lf)))
        return out
    g = gaps()
    lv, reg, n_st, n_rg = area_levels([(area_of(r), e) for r, e in g], normalize=False)
    towns = {}
    for r, e in g:
        geo = geo_of(r["place"])
        if geo:
            s = towns.setdefault(geo, [0, 0.0])
            s[0] += 1
            s[1] += e
    AREAS = {"k": K_AREA, "uf": {uf: [round(v, 4), n_st[uf]] for uf, v in lv.items()},
             "reg": {k: [round(v, 4), n_rg[k]] for k, v in reg.items()}, "region_of": REGION_OF,
             "towns": [[la, lo, n, round(s, 3)] for (la, lo), (n, s) in sorted(towns.items())]}
    return model, b_conv, sd


def score_rows(active):
    """Same maths as the workbook: fair price, % under fair, confidence and the 0-10 score, per listing id."""
    model, b_conv, sd = fit_with_areas(active)
    ppp = max(0.03, round(2 * sd / 5, 2))

    def clamp(x, lo, hi):
        return max(lo, min(hi, x))
    good, great = clamp(0.5 * sd, 0.03, 0.07), clamp(1.25 * sd, 0.08, 0.15)
    deal = {"good": round(good, 3), "great": round(great, 3),          # % under (or over) the fair price for each label
            "sus": round(max(great + 0.07, clamp(2.2 * sd, 0.15, 0.30)), 3)}

    def price_part(vs):                                  # a price far under the market earns nothing: too good to be true
        s = deal["sus"]
        return vs if vs < s else max(0.0, s - (vs - s) * 3)
    out = {}
    for r in active:
        o = {"fair": None, "vs_fair": None, "conf": None, "score": None}
        m = model.get(mkey(r))
        try:
            price, km, year = float(r["price"]), float(r["km"] or 0), float(r["year"])
        except ValueError:
            out[r["listing_id"]] = o
            continue
        value_pts = 0.0
        if m:
            a, bk, by, n = m
            fair = round(math.exp(a + bk * km / 1e4 + by * (year - 2000) + b_conv * (r["body"] == "conv")
                                  + VADJ.get((mkey(r), version_of(r)), 0.0)), -2)
            conf = n / (n + 5)
            vs = 1 - price / fair if fair else None
            o.update(fair=int(fair), vs_fair=round(vs, 4) if vs is not None else None, conf=round(conf, 2))
            if vs is not None:
                value_pts = price_part(vs) / ppp * conf
        red = [f for f in (r.get("red_flags") or "").split("; ") if f]
        green = [f for f in (r.get("green_flags") or "").split("; ") if f]
        pts = (SCORE["base"] + value_pts + SCORE["gearbox"].get(GEARBOX_LABEL.get(r["trans"], ""), 0.0)
               + sum(SCORE["red"].get(f, 0.0) for f in red)
               + min(SCORE["green_cap"], sum(SCORE["green"].get(f, 0.0) for f in green)))
        o["score"] = round(max(0.0, min(10.0, pts)), 1)
        o["sr"] = round(pts, 3)                          # before the 0-10 limits: the app moves it for a local price
        out[r["listing_id"]] = o
    return out, {"sd": round(sd, 3), "pct_per_point": ppp, "deal": deal, "versions": VERSIONS,
                 "families": {f: {"n": v[3]} for f, v in model.items()}}


def photo_big(row):
    """Large photo address for the web app (first = main, second = fallback)."""
    out = []
    for p in (row.get("photo_id"), row.get("photo_alt")):
        if not p:
            continue
        if p.startswith("http"):
            out.append(re.sub(r",w_\d+$", ",w_1080", p) if "mobiauto.com.br" in p else p)   # Mobiauto: sharper size
        else:
            stem, ext = (p.rsplit(".", 1) if "." in p else (p, "jpg"))
            out.append(f"{PHOTO_BASE}{stem}.{ext}/1280x960.webp")
    return out


GEO = {}                                           # "city|UF" (no accents, lower case) -> (lat, lon); Brazil only


def geo_key(city, uf=""):
    s = unicodedata.normalize("NFKD", city or "").encode("ascii", "ignore").decode().lower()
    s = re.sub(r"[^a-z0-9]+", " ", s).strip()
    return f"{s}|{(uf or '').strip().upper()}"


def load_geo():
    """Town coordinates for the location filter (collector/geo_br.csv, IBGE municipalities)."""
    if GEO or COUNTRY["country"] != "BR":
        return
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "geo_br.csv")
    if os.path.exists(path):
        with open(path, newline="", encoding="utf-8") as f:
            for r in csv.DictReader(f):
                GEO[geo_key(r["city"], r["uf"])] = (float(r["lat"]), float(r["lon"]))


def geo_of(place):
    """'Sao Jose dos Campos, SP' -> (lat, lon) or None."""
    if not place or not GEO:
        return None
    city, _, uf = place.rpartition(",")
    if not city:
        return None
    return GEO.get(geo_key(city, uf)) or GEO.get(geo_key(re.sub(r"\s+(de|do|da|dos|das)\s+\S+$", "", city.strip()), uf))


def listing_days(r, until):
    """(days on the market, 'site' or 'tracked'): from the site's own listing date when known."""
    start = d(r.get("listed")) or d(r["first_seen"])
    if not start or not until:
        return None, None
    return max(0, (until - start).days), ("site" if r.get("listed") else "tracked")


# --------------------------------------------------------------------------------------------
# FIPE (Brazil): the national reference price of each version and model year, from the public FIPE API
# (fipe.parallelum.com.br, its own copy of the FIPE table; free: 500 requests a day, 3 months of history).
# Cached in data_br/fipe_cache.json, so a month costs only the new prices.
# --------------------------------------------------------------------------------------------
FIPE_BASES = ["https://fipe.parallelum.com.br/api/v2", "https://fipe.api.br/api/v2"]
FIPE_CFG = {}          # family -> {"brand": "Honda", "match": regex on the FIPE model name, "versions": bool}
FIPE = {}              # the cache


def fipe_get(sess, path, st):
    if st["left"] <= 0 or time_up():
        return None
    st["left"] -= 1
    time.sleep(0.2)
    for b in FIPE_BASES:
        try:
            r = sess.get(b + path, timeout=30, headers={"User-Agent": "CarWatch/1.0 (personal used-car price tracker)"})
        except Exception:
            continue
        if r.status_code == 200:
            return r.json()
        if r.status_code in (402, 404):
            return None
    return None


def fipe_price(s):
    """'R$ 71.094,00' -> 71094"""
    s = re.sub(r"[^\d,]", "", s or "")
    return int(float(s.replace(",", "."))) if s else None


def fipe_key(family, version, auto, year, body=""):
    return f"{family}|{version or ''}|{'a' if auto else 'm'}|{year}" + ("|conv" if body == "conv" else "")


def fipe_candidates(cfg, models, version, auto, body=""):
    """FIPE models for a family, version, gearbox and body. Best first: the same body, then the shortest name
    (the base version: 'XL Sedan' before 'XL Live Sedan'). A convertible's FIPE entry never prices a closed car."""
    rx = re.compile(cfg["match"], re.I)
    arx = re.compile(cfg["auto"], re.I) if cfg.get("auto") else None     # brands that mark automatics their own way (330iA)
    out = []
    for m in models:
        n = m["name"]
        if not rx.search(n):
            continue
        if cfg.get("versions", True) and not re.search(r"(?<![A-Za-z])" + re.escape(version) + r"(?![A-Za-z])", n, re.I):
            continue
        aut = re.search(r"\bAut\b|\bCVT\b", n, re.I) or (arx and arx.search(n))
        mec = re.search(r"\bMec\b", n, re.I)
        if (auto and mec and not aut) or (not auto and aut and not mec):
            continue
        cab = re.search(r"Cabrio|Conv|Roadster", n, re.I)
        if cab and body != "conv":
            continue
        pen = (0 if cab else 1) if body == "conv" else \
              2 if (body == "sedan" and re.search(r"Coup|\b2p\b", n, re.I)) or (body == "coupe" and re.search(r"Sedan|Sed\.|\b4p\b", n, re.I)) else 0
        out.append((pen, len(n), m))
    return [m for _, _, m in sorted(out, key=lambda x: (x[0], x[1]))]


def load_fipe(path):
    global FIPE
    try:
        with open(path, encoding="utf-8") as f:
            FIPE = json.load(f)
    except (OSError, ValueError):
        FIPE = {}


def update_fipe(sess, rows, path, budget=300):
    """Fetch what's missing for the versions and years on sale; at most `budget` requests a run."""
    load_fipe(path)
    st = {"left": budget}
    refs = fipe_get(sess, "/references", st)
    if not refs:
        print("FIPE: no answer")
        return
    cur, month = str(refs[0]["code"]), refs[0]["month"]
    if FIPE.get("ref") != cur:                          # a new month: refresh the lists too
        FIPE.update(ref=cur, month=month, models={}, years={})
    for k in ("prices", "map", "models", "years"):
        FIPE.setdefault(k, {})
    targets = set()
    for r in rows:
        cfg = FIPE_CFG.get(r["family"])
        if not cfg or r["status"] != "active" or not r["year"]:
            continue
        v = version_of(r) if cfg.get("versions", True) else None
        if cfg.get("versions", True) and not v:
            continue
        targets.add((r["family"], v, r.get("trans") == "a", int(float(r["year"])), r.get("body") or ""))
    if not FIPE.get("brands"):                          # an empty answer is asked again next run
        FIPE["brands"] = {b["name"].split()[0].upper(): b["code"] for b in (fipe_get(sess, "/cars/brands", st) or [])}
    done = 0
    for fam, v, auto, year, body in sorted(targets, key=lambda x: (x[0], x[1] or "", x[2], x[3], x[4])):
        cfg = FIPE_CFG[fam]
        bcode = FIPE["brands"].get(cfg["brand"].upper())
        if not bcode:
            continue
        if cfg["brand"] not in FIPE["models"]:
            ms = fipe_get(sess, f"/cars/brands/{bcode}/models", st)
            if ms is None:
                break
            FIPE["models"][cfg["brand"]] = ms
        key = fipe_key(fam, v, auto, year, body)
        for m in fipe_candidates(cfg, FIPE["models"][cfg["brand"]], v or "", auto, body):
            if m["code"] not in FIPE["years"]:
                ys = fipe_get(sess, f"/cars/brands/{bcode}/models/{m['code']}/years", st)
                if ys is None:
                    break
                FIPE["years"][m["code"]] = ys
            yc = next((y["code"] for y in FIPE["years"][m["code"]] if str(y["code"]).startswith(f"{year}-")), None)
            if not yc:
                continue
            pk = f"{m['code']}|{yc}"
            pr = FIPE["prices"].setdefault(pk, {"name": m["name"], "hist": {}})
            if pr.get("ref") != cur:
                if not pr.get("codeFipe"):
                    one = fipe_get(sess, f"/cars/brands/{bcode}/models/{m['code']}/years/{yc}", st)
                    if not one:
                        break
                    pr["codeFipe"] = one.get("codeFipe")
                    pr["hist"][cur] = [month, fipe_price(one.get("price"))]
                h = fipe_get(sess, f"/cars/{pr['codeFipe']}/years/{yc}/history", st) if pr.get("codeFipe") else None
                for x in (h or {}).get("priceHistory", []):
                    if fipe_price(x.get("price")):
                        pr["hist"][str(x["reference"])] = [x["month"], fipe_price(x["price"])]
                if pr["hist"].get(cur):
                    pr["ref"] = cur
                    done += 1
            if pr["hist"]:
                FIPE["map"][key] = pk
            break
        if st["left"] <= 0:
            break
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(FIPE, f, ensure_ascii=False, separators=(",", ":"))
    print(f"FIPE {month}: {done} prices updated, {len(FIPE['map'])} versions matched, {budget - st['left']} requests")


def fipe_export(active):
    """For the app: the month, and each version/year on sale with its FIPE history [[month, price], ...] (oldest first)."""
    series = {}
    for r in active:
        _, key = fipe_of(r)
        if key and key not in series:
            pr = FIPE["prices"][FIPE["map"][key]]
            series[key] = {"name": pr["name"], "hist": [pr["hist"][k] for k in sorted(pr["hist"], key=int)]}
    return {"month": FIPE.get("month"), "series": series}


def fipe_of(r):
    """(latest FIPE price, map key) for a listing, or (None, None)."""
    cfg = FIPE_CFG.get(r.get("family"))
    if not cfg or not FIPE.get("map") or not r.get("year"):
        return None, None
    v = version_of(r) if cfg.get("versions", True) else None
    key = fipe_key(r["family"], v, r.get("trans") == "a", int(float(r["year"])), r.get("body") or "")
    pk = FIPE["map"].get(key)
    pr = FIPE.get("prices", {}).get(pk) if pk else None
    if not pr or not pr.get("hist"):
        return None, None
    last = pr["hist"][max(pr["hist"], key=int)]
    return last[1], key


def _med(v):
    v = sorted(x for x in v if x is not None)
    if not v:
        return None
    m = len(v)
    return v[m // 2] if m % 2 else (v[m // 2 - 1] + v[m // 2]) / 2


def market_stats(rows, model, b_conv, today):
    """Per fair-price group (model, + '|auto' for automatics), for the app's Prices tab:
    model: [level, km slope, year slope, cars] of the national fair-price fit (log price);
    weeks: [day, ads on sale, median price / fair price, new that week, gone that week] from the first tracked day;
    gone: ads that left in the last 90 days (often sold): count, median days listed, median last price / fair,
          share that had cut their price, median total cut."""
    live = [r for r in rows if r["status"] in ("active", "gone") and r["price"] and d(r["first_seen"])]
    if not live:
        return {}
    start = min(d(r["first_seen"]) for r in live)
    days = [start + dt.timedelta(days=7 * k) for k in range((today - start).days // 7 + 1)]
    if days[-1] < today:
        days.append(today)
    area = (AREAS or {}).get("uf", {})

    def idx(r, price):                                   # asking price / fair price (national, state-corrected)
        lf = log_fair(model, b_conv, r)
        if lf is None:
            return None
        return price / math.exp(lf + (area.get(area_of(r), [0.0])[0] if area else 0.0))

    def price_on(r, day):
        pts = [(d(x), p) for x, p in hist_points(r.get("price_hist"))]
        before = [p for x, p in pts if x and x <= day]
        return before[-1] if before else (pts[0][1] if pts else int(r["price"]))

    out = {}
    for key in sorted({mkey(r) for r in live}):
        rs = [r for r in live if mkey(r) == key]
        weeks, prev = [], None
        for day in days:
            on = [r for r in rs if d(r["first_seen"]) <= day and (r["status"] == "active" or d(r["last_seen"]) >= day)]
            ix = [idx(r, price_on(r, day)) for r in on]
            new = sum(1 for r in rs if prev and prev < d(r["first_seen"]) <= day)
            gone = sum(1 for r in rs if r["status"] == "gone" and prev and prev < d(r["last_seen"]) <= day)
            m = _med(ix)
            weeks.append([day.isoformat(), len(on), round(m, 4) if m else None, new, gone])
            prev = day
        left = [r for r in rs if r["status"] == "gone" and d(r["last_seen"]) and (today - d(r["last_seen"])).days <= 90]
        cuts = []
        for r in left:
            pts = [p for _, p in hist_points(r.get("price_hist"))]
            cuts.append(pts[-1] / pts[0] - 1 if len(pts) > 1 and min(pts[1:]) < pts[0] else None)
        m = model.get(key)
        gi, gd = _med([idx(r, int(r["price"])) for r in left]), _med([listing_days(r, d(r["last_seen"]))[0] for r in left])
        gc = _med([c for c in cuts if c is not None])
        out[key] = {"model": [round(m[0], 5), round(m[1], 5), round(m[2], 5), m[3]] if m else None, "weeks": weeks,
                    "gone": {"n": len(left), "days": gd, "idx": round(gi, 4) if gi else None,
                             "cut_share": round(sum(c is not None for c in cuts) / len(left), 3) if left else None,
                             "cut": round(gc, 4) if gc is not None else None}}
    return out


def _near(a, b):
    """Same town: the geo points are within 40 km, or (no points) the place text is the same."""
    if a["geo"] and b["geo"]:
        (la1, lo1), (la2, lo2) = a["geo"], b["geo"]
        x = math.radians(lo2 - lo1) * math.cos(math.radians((la1 + la2) / 2))
        return 6371 * math.hypot(x, math.radians(la2 - la1)) < 40
    return bool(a["place"]) and _norm(a["place"]).split(",")[0].strip() == _norm(b["place"]).split(",")[0].strip()


def merge_same_car(cars):
    """One car listed on two sites (a dealer on Autotrader and Kijiji) becomes one entry: the cheaper ad leads and
    the other is kept under 'also'. Same generation, year, gearbox and town, and either the same price with km
    within 1.5 % (300 km), or a price within 5 % with the exact same, non-round km. Closest pairs are matched first."""
    by = {}
    for i, c in enumerate(cars):
        by.setdefault((c["family"], c["year"], c["trans"]), []).append(i)
    pairs = []
    for ix in by.values():
        for a in ix:
            for b in ix:
                x, y = cars[a], cars[b]
                if a >= b or x["source"] == y["source"] or not _near(x, y):
                    continue
                if x["km"] and y["km"]:
                    dk = abs(x["km"] - y["km"])
                    if x["price"] == y["price"]:
                        ok = dk <= max(300, 0.015 * max(x["km"], y["km"]))
                    else:
                        ok = dk == 0 and x["km"] % 1000 != 0 and abs(x["price"] - y["price"]) <= 0.05 * max(x["price"], y["price"])
                else:
                    ok = not x["km"] and not y["km"] and x["price"] == y["price"]
                if ok:
                    pairs.append((abs(x["price"] - y["price"]) / max(x["price"], 1), abs((x["km"] or 0) - (y["km"] or 0)), a, b))
    lead, taken = {}, set()
    for _, _, a, b in sorted(pairs):
        if a in taken or b in taken:
            continue
        x, y = cars[a], cars[b]
        keep, drop = (a, b) if (x["price"], x["first_seen"]) <= (y["price"], y["first_seen"]) else (b, a)
        taken |= {a, b}
        lead[keep] = drop
    out, dropped = [], set(lead.values())
    for i, c in enumerate(cars):
        if i in dropped:
            continue
        if i in lead:
            o = cars[lead[i]]
            c["also"] = [{"id": o["id"], "source": o["source"], "url": o["url"], "price": o["price"]}]
            if not c["photos"] and o["photos"]:
                c["photos"] = o["photos"]
            if not c.get("desc") and o.get("desc"):
                c["desc"] = o["desc"]
        out.append(c)
    return out


def export_json(rows, path, run_info=None, stats=None, comps=None):
    """cars.json for the web app: every active car with its cues, score, flags, price history and sold comps."""
    today = dt.date.today()
    load_geo()
    cap = MAX_PRICE.get(COUNTRY["country"], 10 ** 7)
    active = [r for r in rows if r["status"] == "active" and r["price"] and int(r["price"]) <= cap]
    scores, model_info = score_rows(active)
    model, b_conv, _ = fit_with_areas(active)
    comps = comps or []
    comp_idx, comp_list = {}, []

    def comp_ref(c):
        if c["id"] not in comp_idx:
            comp_idx[c["id"]] = len(comp_list)
            comp_list.append({"t": c["title"], "sold": c["sold"], "km": _i(c["km"]) if c["km"] else None,
                              "cad": _i(c["cad"]), "usd": _i(c["price"]) if c["cur"] == "USD" else None,
                              "url": c["url"], "year": _i(c["year"]), "body": c["body"]})
        return comp_idx[c["id"]]
    names, fams = {}, {}
    for m in MODELS:
        brand = m.get("brand") or m["model"]                # the app groups its model chips by brand
        for g in m["generations"]:
            fams.setdefault(g["family"], {"family": g["family"], "label": g.get("short", g["name"]),
                                          "color": g.get("color", "#888888"), "desc": g.get("desc", ""),
                                          "model": m["model"], "brand": brand})
            nb = g.get("name_by_body") or {}              # bodies it leaves out keep the generation's own name
            bodies = [g["body"]] if g.get("body") else g.get("bodies") or m.get("bodies") or []
            pairs = [(v, k) for k, v in nb.items() if k in bodies] + [(g["name"], b) for b in bodies if b not in nb]
            for nm, b in pairs or [(g["name"], None)]:
                n = names.setdefault(nm, {"name": nm, "family": g["family"], "color": g.get("color", "#888888"),
                                          "years": list(g["years"]), "bodies": [], "brand": brand, "model": m["model"],
                                          **({"daily": True} if g.get("daily") else {})})   # daily drivers: their own group in the app
                n["years"] = [min(n["years"][0], g["years"][0]), max(n["years"][1], g["years"][1])]
                if b and b not in n["bodies"]:
                    n["bodies"].append(b)                   # the ad reader's default body
            for _, sg in SPLITS.get(g["family"], []):
                fams.setdefault(sg["family"], {"family": sg["family"], "label": sg.get("short", sg["name"]),
                                               "color": sg.get("color", "#888888"), "desc": sg.get("desc", ""),
                                               "model": m["model"], "brand": brand})
                names.setdefault(sg["name"], {"name": sg["name"], "family": sg["family"],
                                              "color": sg.get("color", "#888888"), "years": list(sg["years"]),
                                              "bodies": [sg["body"]] if sg.get("body") else list(bodies), "brand": brand,
                                              "model": m["model"], **({"daily": True} if sg.get("daily") else {})})
    cars = []
    for r in active:
        sc = scores.get(r["listing_id"], {})
        days, days_src = listing_days(r, today)
        hist = hist_points(r.get("price_hist"))
        est, near = match_comps(r, comps, model, b_conv, today) if comps else (None, [])
        sold = {"est": est, "vs": round(1 - int(r["price"]) / est, 4) if est else None,
                "refs": [[comp_ref(c), adj] for c, adj in near[:5]]} if near else None
        cars.append({
            "id": r["listing_id"], "url": r["url"], "source": r["source"], "name": r["name"], "family": r["family"],
            "year": int(r["year"]) if r["year"] else None, "body": r["body"], "trans": r["trans"],
            "price": int(r["price"]), "prev_price": int(r["prev_price"]) if r["prev_price"] else None,
            "km": int(r["km"]) if r["km"] else None, "place": r["place"], "trim": r.get("trim", ""),
            "red": [f for f in (r.get("red_flags") or "").split("; ") if f],
            "green": [f for f in (r.get("green_flags") or "").split("; ") if f],
            "flags_read": bool(r.get("flags_read")), "note": r["note"], "desc": r.get("desc", ""),
            "photos": photo_big(r),
            "first_seen": r["first_seen"], "last_seen": r["last_seen"], "verified": r["verified"],
            "listed": r.get("listed") or None, "days": days, "days_src": days_src,
            **(lambda fp: {"fipe": fp[0], "fipe_k": fp[1]} if fp[0] else {})(fipe_of(r)),
            "geo": geo_of(r["place"]),
            "hist": [[dd, pp] for dd, pp in hist] if len(hist) > 1 else None, "sold": sold, **sc})
    firsts = sorted(r["first_seen"] for r in rows if r["first_seen"] and r["status"] != "skip")
    n_before = len(cars)
    cars = merge_same_car(cars)
    if stats is not None:
        stats["merged"] = n_before - len(cars)
    print(f"Same car on two sites: {n_before - len(cars)} ads folded into another site's listing")
    autos = [c for c in cars if c["trans"] == "a"]
    cars = [c for c in cars if c["trans"] != "a"]
    auto_path = re.sub(r"\.json$", "", path) + "_auto.json"
    data = {
        "updated": dt.datetime.now(dt.timezone.utc).isoformat(timespec="minutes"),
        "country": COUNTRY["country"], "currency": COUNTRY["currency"], "links": COUNTRY["links"],
        "baseline": firsts[0] if firsts else "",
        "repo": os.environ.get("GITHUB_REPOSITORY", ""),
        "families": [fams[f] for f in fams if any(c["family"] == f for c in cars)],
        "names": [n for n in names.values() if any(c["name"] == n["name"] for c in cars)],
        "names_all": list(names.values()),            # every tracked model, with or without ads: the ad reader needs them
        "score": {**SCORE, **model_info}, "run": run_info or [],
        "areas": AREAS or None,
        "market": market_stats(rows, model, b_conv, today), "b_conv": round(b_conv, 5),
        "fipe": fipe_export(active) if FIPE.get("map") else None,
        "searches": [{"source": s, "url": u, "ads": n} for s, u, n in (stats or {}).get("searches", [])],
        "cars": cars,
        "auto_file": os.path.basename(auto_path) if autos else None, "auto_count": len(autos),
        "comps": comp_list,
        "sold_bands": sold_bands(comps, today) if comps else {},
        "sold_hist": sold_history(comps) if comps else {},
        "gone": [{"id": r["listing_id"], "name": r["name"], "family": r["family"], "year": r["year"], "km": r["km"],
                  "body": r["body"], "trans": r["trans"], "price": r["price"], "url": r["url"],
                  "last_seen": r["last_seen"], "days": listing_days(r, d(r["last_seen"]))[0]}
                 for r in sorted((r for r in rows if r["status"] == "gone" and d(r["last_seen"])
                                  and (today - d(r["last_seen"])).days <= 90),
                                 key=lambda r: r["last_seen"], reverse=True)[:150]],
    }
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, separators=(",", ":"))
    if autos or os.path.exists(auto_path):
        for c in autos:
            c["desc"] = c["desc"][:400]                  # keeps the automatics file light on a phone
        adata = {**data, "cars": autos, "auto_file": None, "auto_count": len(autos), "gone": [],
                 "families": [fams[f] for f in fams if any(c["family"] == f for c in autos)],
                 "names": [n for n in names.values() if any(c["name"] == n["name"] for c in autos)]}
        with open(auto_path, "w", encoding="utf-8") as f:
            json.dump(adata, f, ensure_ascii=False, separators=(",", ":"))


def prune_snapshots(keep):
    files = sorted(glob.glob(os.path.join(DATA_DIR, "car_listings_*.csv")))
    for f in files[:-keep]:
        for p in (f, marker_of(f), f + ".bak"):
            if os.path.exists(p):
                os.remove(p)


# --------------------------------------------------------------------------------------------
# Sold prices: Bring a Trailer auction results (US), in CAD at the Bank of Canada rate of the sale day
# --------------------------------------------------------------------------------------------
BAT_API = ("https://bringatrailer.com/wp-json/bringatrailer/1.0/data/listings-filter"
           "?page={page}&per_page=60&get_items=1&get_stats=0&sort=td&state=sold&base_filter[keyword_pages][]={pid}")
FX_API = "https://www.bankofcanada.ca/valet/observations/FXUSDCAD/json?start_date={start}"
BROWSER_UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36"
COMPS_FILE = "comps_bat.csv"          # in --data-dir; kept between runs, newest results added every day
COMPS_STATE = "comps_state.json"      # newest result already read, per results page
FX_FILE = "fx_usdcad.csv"
COMPS_KEEP_YEARS = 3                  # results older than this are dropped
COMPS_MAX_PAGES = 25                  # per results page and run (60 results each)
COMP_COLS = ["id", "family", "title", "year", "body", "km", "trans", "price", "cur", "cad", "sold", "url"]
COMP_SKIP = re.compile(r"\b(race|track|rally|drift|project|parts|shell|widebody|wide-body|replica|tribute|"
                       r"conversion|supercharged|turbocharged|twin-turbo|rebuilt|salvage|engine|wheels|seats)\b|"
                       r"-(powered|swapped)\b|\bswap", re.I)
COMP_AUTO = re.compile(r"\bSMG\b|dual-clutch|\bDCT\b|\b(?:four|five|six|seven|4|5|6|7)-speed (?:automatic|auto\b|steptronic)|"
                       r"automatic (?:transmission|transaxle|gearbox)|steptronic|sequential manual|paddle-shift", re.I)
COMPS_VERSION = 2                     # bump when the reading rules change: the stored results are read again
COMP_MANUAL = re.compile(r"\b(four|five|six|4|5|6)-speed manual|manual (transmission|gearbox)", re.I)


def load_csv_rows(path, cols):
    if not os.path.exists(path):
        return []
    with open(path, newline="", encoding="utf-8") as f:
        return [{c: r.get(c, "") or "" for c in cols} for r in csv.DictReader(f)]


def save_csv_rows(path, cols, rows):
    tmp = path + ".tmp"
    with open(tmp, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)
    os.replace(tmp, path)


def update_fx(sess, start):
    """USD->CAD daily rates since `start` (Bank of Canada), cached in fx_usdcad.csv. Returns [(date, rate)]."""
    path = os.path.join(DATA_DIR, FX_FILE)
    have = {r["date"]: float(r["rate"]) for r in load_csv_rows(path, ["date", "rate"]) if r["rate"]}
    since = max(have) if have else start
    if sess is not None and since < (dt.date.today() - dt.timedelta(days=1)).isoformat():
        r = fetch(sess, FX_API.format(start=since), delay=0.5)
        try:
            for o in r.json()["observations"]:
                if o.get("FXUSDCAD", {}).get("v"):
                    have[o["d"]] = float(o["FXUSDCAD"]["v"])
        except Exception as e:
            print("   exchange rate not updated:", getattr(r, "status_code", "no response"), str(e)[:80])
        save_csv_rows(path, ["date", "rate"], [{"date": k, "rate": v} for k, v in sorted(have.items())])
    return sorted(have.items())


def fx_on(series, day):
    """Rate of the last business day on or before `day` (or the nearest one)."""
    if not series:
        return None
    lo, hi = 0, len(series) - 1
    if day < series[0][0]:
        return series[0][1]
    while lo < hi:
        mid = (lo + hi + 1) // 2
        if series[mid][0] <= day:
            lo = mid
        else:
            hi = mid - 1
    return series[lo][1]


def comp_km(title, text):
    """Odometer in km from a Bring a Trailer title ('40k-Mile ...') or summary ('... now shows 62k miles')."""
    if re.search(r"\bTMU\b|true mileage (is )?unknown", title + " " + text, re.I):
        return None
    m = re.search(r"\b(\d{1,3}(?:,\d{3})*|\d+(?:\.\d)?)(k)?-(Mile|Kilometer|km)\b", title, re.I)
    if not m:
        m = re.search(r"\b(?:shows|showing|indicates|reads|displays|registers|reflects|has|with)\s+(?:just\s+|only\s+|"
                      r"approximately\s+|roughly\s+|about\s+|under\s+|fewer than\s+|less than\s+|over\s+|nearly\s+)?"
                      r"(\d{1,3}(?:,\d{3})+|\d+(?:\.\d)?)(k)?\s*(miles|kilometers|km)\b", text, re.I)
    if not m:
        return None
    v = float(m.group(1).replace(",", "")) * (1000 if m.group(2) else 1)
    if m.group(3).lower().startswith("mile"):
        v *= 1.609
    return int(round(v, -2)) if 100 <= v <= 900000 else None


def comp_from_item(it, family, cfg, fx):
    """One auction result -> comps row, or None when it is not a comparable car."""
    title = unescape(re.sub(r"<[^>]+>", " ", str(it.get("title") or ""))).strip()
    text = unescape(re.sub(r"<[^>]+>", " ", str(it.get("excerpt") or "")))
    ym = re.search(r"\b(19[5-9]\d|20[0-3]\d)\b", title)
    if not ym or not re.search(cfg["match"], title, re.I):
        return None
    year = int(ym.group(1))
    y0, y1 = cfg["years"]
    if not (y0 <= year <= y1) or COMP_SKIP.search(title):
        return None
    sold_text = str(it.get("sold_text") or "")
    if not sold_text.lower().startswith("sold for"):
        return None                                     # "Bid to": reserve not met, no sale
    price = to_int(it.get("current_bid"))
    cur = str(it.get("currency") or "").upper() or \
        (re.search(r"\b(USD|CAD|EUR|GBP|AUD|CHF|JPY)\b", sold_text) or [None, "USD"])[1]
    ts = it.get("sold_text_timestamp") or it.get("timestamp_end")
    if not price or not ts or cur not in ("USD", "CAD"):
        return None
    sold = dt.datetime.fromtimestamp(int(ts), dt.timezone.utc).date().isoformat()
    rate = 1.0 if cur == "CAD" else fx_on(fx, sold)
    if rate is None:
        return None
    if re.search(r"\b[456]-Speed\b", title) or cfg.get("all_manual"):
        trans = "m"
    elif COMP_AUTO.search(title + " " + text[:400]):
        trans = "a"
    elif COMP_MANUAL.search(text[:400]) or cfg.get("all_manual"):
        trans = "m"
    else:
        trans = "u"
    tl = title.lower()
    body = ("conv" if re.search(r"convertible|cabrio|roadster", tl) else "wagon" if re.search(r"wagon|touring", tl)
            else "sedan" if "sedan" in tl else "coupe" if "coupe" in tl else "")
    if not body:
        body = next((b for rx, b in cfg.get("body_by_title", []) if re.search(rx, title, re.I)), cfg.get("body", ""))
    return {"id": str(it.get("id") or it.get("url")), "family": family, "title": title[:110], "year": year, "body": body,
            "km": comp_km(title, text) or "", "trans": trans, "price": price, "cur": cur,
            "cad": int(round(price * rate, -2)), "sold": sold, "url": str(it.get("url") or "")}


def update_comps(sess, stats):
    """Read new Bring a Trailer results for every car with 'comps' in the config. Returns all stored comps."""
    path = os.path.join(DATA_DIR, COMPS_FILE)
    comps = {r["id"]: r for r in load_csv_rows(path, COMP_COLS)}
    state_path = os.path.join(DATA_DIR, COMPS_STATE)
    try:
        with open(state_path, encoding="utf-8") as f:
            state = json.load(f)
    except (OSError, ValueError):
        state = {}
    if state.get("_version") != COMPS_VERSION:          # new reading rules: read every result again
        state, comps = {"_version": COMPS_VERSION}, {}
    km_samples = []
    cutoff = dt.datetime.now(dt.timezone.utc) - dt.timedelta(days=365 * COMPS_KEEP_YEARS)
    fx = update_fx(sess, (cutoff.date() - dt.timedelta(days=10)).isoformat())
    stats.setdefault("comps_new", 0)
    stats.setdefault("comps_pages", 0)
    if sess is not None and COMPS:
        hdr = {"User-Agent": BROWSER_UA, "Accept": "application/json"}
        for family, cfg in COMPS.items():
            for pid in cfg.get("bat", []):
                key = f"{family}:{pid}"
                newest = float(state.get(key, 0))
                stop_ts = max(cutoff.timestamp(), newest - 3 * 86400)   # small overlap: late results get updated
                top = newest
                for page in range(1, COMPS_MAX_PAGES + 1):
                    url = BAT_API.format(page=page, pid=pid)
                    wait = 1.5 - (time.time() - _last_request[0])
                    if wait > 0:
                        time.sleep(wait)
                    try:
                        r = sess.get(url, headers=hdr, timeout=30)
                        _last_request[0] = time.time()
                        items = r.json().get("items") or []
                    except Exception as e:
                        print(f"   sold prices: {family} page {page} not read ({str(e)[:60]})")
                        break
                    stats["comps_pages"] += 1
                    if not items:
                        break
                    oldest = None
                    for it in items:
                        ts = float(it.get("timestamp_end") or 0)
                        oldest = ts if oldest is None else min(oldest, ts)
                        top = max(top, ts)
                        if ts < stop_ts:
                            continue
                        row = comp_from_item(it, family, cfg, fx)
                        if row and not row["km"] and len(km_samples) < 15:
                            km_samples.append(f'{row["title"][:60]} | {unescape(str(it.get("excerpt") or ""))[:260]}')
                        if row:
                            if row["id"] not in comps:
                                stats["comps_new"] += 1
                            comps[row["id"]] = row
                    if oldest is not None and oldest < stop_ts:
                        break
                state[key] = top
        with open(state_path, "w", encoding="utf-8") as f:
            json.dump(state, f, indent=1)
        for s in km_samples:
            print("   km not found:", s)
    keep_from = cutoff.date().isoformat()
    rows = sorted((r for r in comps.values() if r["sold"] >= keep_from and r["family"] in COMPS),
                  key=lambda r: (r["family"], r["sold"]), reverse=True)
    save_csv_rows(path, COMP_COLS, rows)
    stats["comps_total"] = len(rows)
    return rows


def _i(v):
    return int(float(v))


def _median(v):
    v = sorted(v)
    n = len(v)
    return None if not n else v[n // 2] if n % 2 else (v[n // 2 - 1] + v[n // 2]) / 2


def _quantile(v, q):
    v = sorted(v)
    if not v:
        return None
    i = (len(v) - 1) * q
    lo = int(i)
    return v[lo] + (v[min(lo + 1, len(v) - 1)] - v[lo]) * (i - lo)


def match_comps(car, comps, model, b_conv, today):
    """Closest manual sales of the same car; each one moved to this car's km and year with the fair-price slopes.
    -> (estimate or None, [(comp, adjusted CAD)])"""
    fam_m = model.get(car["family"])
    bk, by = (fam_m[1], fam_m[2]) if fam_m else (0.0, 0.0)
    try:
        year, km = int(car["year"]), int(car["km"] or 0)
    except ValueError:
        return None, []
    conv = car["body"] == "conv"
    pool = [c for c in comps if c["family"] == car["family"] and c["trans"] == "m" and (c["body"] == "conv") == conv
            and abs(_i(c["year"]) - year) <= 3]
    if not conv and sum(1 for c in pool if c["body"] == car["body"]) >= 3:
        pool = [c for c in pool if c["body"] == car["body"]]
    recent = [c for c in pool if (today - d(c["sold"])).days <= 730]
    if len(recent) >= 4:
        pool = recent

    def dist(c):
        months = (today - d(c["sold"])).days / 30.4
        dk = abs(_i(c["km"]) - km) / 40000 if c["km"] and km else 1.5
        return abs(_i(c["year"]) - year) + dk + months / 24

    best = sorted(pool, key=dist)[:6]
    out = []
    for c in best:
        shift = (by * (year - _i(c["year"])) + (bk * (km - _i(c["km"])) / 1e4 if c["km"] and km else 0.0)
                 + b_conv * (conv - (c["body"] == "conv")))
        adj = _i(c["cad"]) * math.exp(max(-0.4, min(0.4, shift)))      # at most -33% / +50%: no wild extrapolation
        out.append((c, int(round(adj, -2))))
    est = _median([a for _, a in out]) if len(out) >= 3 else None
    return (int(round(est, -2)) if est else None), out


def sold_bands(comps, today):
    """Middle half of manual sales in the last 2 years, per family and convertible / not."""
    groups = {}
    for c in comps:
        if c["trans"] == "m" and (today - d(c["sold"])).days <= 730:
            groups.setdefault(f'{c["family"]}|{"conv" if c["body"] == "conv" else "fixed"}', []).append(_i(c["cad"]))
    return {k: {"n": len(v), "p25": int(round(_quantile(v, .25), -2)), "med": int(round(_median(v), -2)),
                "p75": int(round(_quantile(v, .75), -2))} for k, v in groups.items() if len(v) >= 3}


def sold_history(comps):
    """Sold prices over time per generation (Bring a Trailer, in CAD): [half-year, median, p25, p75, sales, manual share]
    for halves with at least 3 sales, oldest first."""
    groups = {}
    for c in comps:
        sd = d(c["sold"])
        if not sd or not _i(c["cad"]):
            continue
        groups.setdefault(c["family"], {}).setdefault(f"{sd.year}-H{1 if sd.month <= 6 else 2}", []).append(c)
    out = {}
    for fam, halves in groups.items():
        rows = []
        for h in sorted(halves):
            v = [_i(c["cad"]) for c in halves[h]]
            if len(v) >= 3:
                rows.append([h, int(round(_median(v), -2)), int(round(_quantile(v, .25), -2)), int(round(_quantile(v, .75), -2)),
                             len(v), round(sum(c["trans"] == "m" for c in halves[h]) / len(v), 2)])
        if len(rows) >= 2:
            out[fam] = rows
    return out


# --------------------------------------------------------------------------------------------
# Excel
# --------------------------------------------------------------------------------------------
FONT = "Arial"
BODY_STYLE = {   # body -> (label, fill, text)
    "sedan": ("\u25AC Sedan", "E1E5EA", "2B3440"), "coupe": ("\u25D7 Coupe", "DDD8F2", "2F2A5C"),
    "conv": ("\u2600 Convertible", "FCE9C8", "6B4A00"), "wagon": ("\u25AD Wagon", "E8DCC8", "5A4326"),
    "hatch": ("\u25E3 Hatchback", "D9EFE7", "1F5A45"),
    "unk": ("? Not stated", "EDEDED", "666666"),
}
TRANS_STYLE = {   # trans -> (label, fill, text)
    "m": ("\u2714 Manual", "CDEFD9", "0B5E2A"), "c": ("\u26A0 Conflict", "FDE7B0", "7A4B00"),
    "u": ("? Unconfirmed", "E4E4E4", "555555"),
}
TIERS = [("under $15k", "7ED6B6", "0B3B2E"), ("$15k - 25k", "3FBF94", "0B3B2E"), ("$25k - 35k", "1D9E75", "FFFFFF"),
         ("$35k - 50k", "0F7A60", "FFFFFF"), ("$50k and up", "0B5E49", "FFFFFF")]
HEADERS = ["Photo", "Car", "Trim", "Score", "Year", "Body", "Gearbox", "Price (CAD)", "Fair price",
           "vs fair", "Km", "Location", "Red flags", "Good signs", "Ad detail", "Prev price", "Price change",
           "First seen", "New?", "Last seen", "Verified", "Source", "Link", "Family key", "Rank key", "URL"]
COL = {h: i + 1 for i, h in enumerate(HEADERS)}
L = {h: get_column_letter(i + 1) for i, h in enumerate(HEADERS)}
SRC_LABEL = {"autotrader": "Autotrader", "kijiji": "Kijiji", "carpages": "CarPages", "cargurus": "CarGurus"}
NOT_READ = "not read yet"
PX = 9525                                              # EMU per pixel


def _f(bold=False, color="000000", size=10, italic=False, underline=None):
    return Font(name=FONT, bold=bold, color=color, size=size, italic=italic, underline=underline)


def _fill(hexc):
    return PatternFill("solid", start_color=hexc, end_color=hexc)


def d(s):
    try:
        return dt.date.fromisoformat(str(s)[:10])
    except (TypeError, ValueError):
        return None


def embed_photo(ws, path, row_idx):
    """Picture inside cell A<row>, set to move and size with the cell (so filters hide it with the row)."""
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
    img = XLImage(path)
    w, h = img.width, img.height
    scale = min(124 / w, 92 / h)
    w, h = int(w * scale), int(h * scale)
    img.width, img.height = w, h
    x0, y0 = (136 - w) // 2, (101 - h) // 2
    img.anchor = TwoCellAnchor(editAs="twoCell",
                               _from=AnchorMarker(col=0, colOff=x0 * PX, row=row_idx - 1, rowOff=y0 * PX),
                               to=AnchorMarker(col=0, colOff=(x0 + w) * PX, row=row_idx - 1, rowOff=(y0 + h) * PX))
    ws.add_image(img)


def build_xlsx(rows, path, today, run_info=None):
    active = sorted((r for r in rows if r["status"] == "active" and r["price"]), key=lambda r: int(r["price"]))
    gone = sorted((r for r in rows if r["status"] == "gone"), key=lambda r: r["last_seen"], reverse=True)
    firsts = [d(r["first_seen"]) for r in rows if d(r["first_seen"]) and r["status"] != "skip"]
    baseline = min(firsts) if firsts else today
    model, b_conv, sd = fit_fair_price(active)
    fams = [f for f in FAMILY_STYLE if any(r["family"] == f for r in active)]
    pct_per_point = max(0.03, round(2 * sd / 5, 2))

    wb = Workbook()
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    ws = wb.active
    ws.title = "Cars"
    sm = wb.create_sheet("Summary")
    thin = Side(style="thin", color="D9D9D9")
    first, last = 5, 4 + max(1, len(active))
    rng = lambda col: f"{L[col]}{first}:{L[col]}{last}"
    absr = lambda col: f"Cars!${L[col]}${first}:${L[col]}${last}"

    # ================= Summary layout (built first: the Cars formulas point at these cells) =================
    sm.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFGHIJKLMNO", [20, 11, 10, 11, 11, 11, 16, 11, 3, 12, 11, 12, 12, 6, 11]):
        sm.column_dimensions[col].width = w
    sm["A1"] = "Summary"
    sm["A1"].font = _f(True, "1F2937", 14)
    sm["A2"] = (f"As of {today:%Y-%m-%d}. The tables follow the Cars tab; the fair-price model (grey, right) and the "
                f"score settings (blue) are refreshed by car_collector.py on every run.")
    sm["A2"].font = _f(italic=True, color="666666", size=9)

    def header(row, col0, names, fill="1F2937", size=10):
        for i, h in enumerate(names):
            c = sm.cell(row=row, column=col0 + i, value=h)
            c.font, c.fill = _f(True, "FFFFFF", size), _fill(fill)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    header(4, 1, ["Model", "Listings", "Manual \u2714", "Lowest", "Median", "Highest", "Median km", "Best score"])
    header(4, 10, ["Family key", "Level (ln $)", "per 10,000 km", "per model year", "n", "Confidence"], "6B7280", 9)
    sm.row_dimensions[4].height = 26
    m0 = 5
    m1 = m0 + max(1, len(fams)) - 1
    tot = m1 + 1
    for k, f in enumerate(fams):
        x = m0 + k
        fill, txt, short, _ = FAMILY_STYLE[f]
        key = f"$J{x}"
        vals = [short, f"=COUNTIF({absr('Family key')},{key})",
                f'=COUNTIFS({absr("Family key")},{key},{absr("Gearbox")},"\u2714 Manual")',
                f"=_xlfn.MINIFS({absr('Price (CAD)')},{absr('Family key')},{key})", None,
                f"=_xlfn.MAXIFS({absr('Price (CAD)')},{absr('Family key')},{key})", None,
                f"=_xlfn.MAXIFS({absr('Score')},{absr('Family key')},{key})"]
        for col, val in enumerate(vals, 1):
            c = sm.cell(row=x, column=col, value=val)
            c.font, c.border = _f(), Border(bottom=thin)
            c.alignment = Alignment(horizontal="center")
        sm[f"E{x}"] = ArrayFormula(f"E{x}", f"=MEDIAN(IF({absr('Family key')}={key},{absr('Price (CAD)')}))")
        sm[f"G{x}"] = ArrayFormula(f"G{x}", f"=MEDIAN(IF({absr('Family key')}={key},{absr('Km')}))")
        c = sm.cell(row=x, column=1)
        c.font, c.fill = _f(True, txt), _fill(fill)
        for col, fmt in ((4, "$#,##0"), (5, "$#,##0"), (6, "$#,##0"), (7, "#,##0"), (8, "0.0")):
            sm.cell(row=x, column=col).number_format = fmt
        a, bk, by, n = model.get(f, (math.log(20000), 0.0, 0.0, 0))
        for col, val in zip(range(10, 15), (f, round(a, 5), round(bk, 5), round(by, 5), n)):
            sm.cell(row=x, column=col, value=val).font = _f(color="0000FF", size=9)
        c = sm.cell(row=x, column=15, value=f"=N{x}/(N{x}+5)")
        c.font, c.number_format = _f(size=9), "0%"
    vals = ["All", f"=SUM(B{m0}:B{m1})", f"=SUM(C{m0}:C{m1})", f"=MIN({absr('Price (CAD)')})",
            f"=MEDIAN({absr('Price (CAD)')})", f"=MAX({absr('Price (CAD)')})", f"=MEDIAN({absr('Km')})",
            f"=MAX({absr('Score')})"]
    for col, val in enumerate(vals, 1):
        c = sm.cell(row=tot, column=col, value=val)
        c.font, c.alignment = _f(True), Alignment(horizontal="center")
    for col, fmt in ((4, "$#,##0"), (5, "$#,##0"), (6, "$#,##0"), (7, "#,##0"), (8, "0.0")):
        sm.cell(row=tot, column=col).number_format = fmt
    sm.cell(row=tot, column=10, value="Convertible").font = _f(size=9)
    sm.cell(row=tot, column=11, value=round(b_conv, 5)).font = _f(color="0000FF", size=9)
    CONV = f"Summary!$K${tot}"
    c = sm.cell(row=tot + 1, column=10, value=f"Fitted {today:%Y-%m-%d} from these listings; typical error \u00b1{sd * 100:.0f}%. How it works: Legend tab.")
    c.font = _f(italic=True, color="666666", size=8)

    # -- top 10 by score (A:H, under the model table)
    T0 = tot + 3
    sm.cell(row=T0, column=1, value="Top 10 by score").font = _f(True, "1F2937", 12)
    header(T0 + 1, 1, ["#", "Car", "Year", "Price", "Score", "vs fair", "Location", "Link"])
    for k in range(min(10, len(active))):
        x = T0 + 2 + k
        idx = f"MATCH(LARGE({absr('Rank key')},{k + 1}),{absr('Rank key')},0)"
        vals = [k + 1, f"=INDEX({absr('Car')},{idx})", f"=INDEX({absr('Year')},{idx})",
                f"=INDEX({absr('Price (CAD)')},{idx})", f"=INDEX({absr('Score')},{idx})",
                f"=INDEX({absr('vs fair')},{idx})", f"=INDEX({absr('Location')},{idx})",
                f'=HYPERLINK(INDEX({absr("URL")},{idx}),"Open ad \u2197")']
        for col, val in enumerate(vals, 1):
            c = sm.cell(row=x, column=col, value=val)
            c.font = _f(color="0563C1", underline="single") if col == 8 else _f(bold=(col == 5))
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = Border(bottom=thin)
        sm.cell(row=x, column=4).number_format = "$#,##0"
        sm.cell(row=x, column=5).number_format = "0.0"
        sm.cell(row=x, column=6).number_format = '+0%;-0%;0%'
    sm.conditional_formatting.add(f"F{T0 + 2}:F{T0 + 11}", FormulaRule(
        formula=[f"AND(ISNUMBER(F{T0 + 2}),F{T0 + 2}>0)"], font=Font(name=FONT, color="0B7A3E")))
    sm.conditional_formatting.add(f"F{T0 + 2}:F{T0 + 11}", FormulaRule(
        formula=[f"AND(ISNUMBER(F{T0 + 2}),F{T0 + 2}<0)"], font=Font(name=FONT, color="C00000")))
    sm.conditional_formatting.add(f"E{T0 + 2}:E{T0 + 11}", ColorScaleRule(
        start_type="num", start_value=0, start_color="F8696B", mid_type="num", mid_value=5, mid_color="FFEB84",
        end_type="num", end_value=10, end_color="63BE7B"))

    # -- score settings (A:B, under the top 10)
    S0 = T0 + 14
    sm.cell(row=S0, column=1, value="Score settings").font = _f(True, "1F2937", 12)
    sm.cell(row=S0 + 1, column=1, value=(
        "Score = base + (% under fair \u00f7 % per point) x model confidence + gearbox + red flags + good signs (capped), 0-10. "
        "Edits here last until the next daily run; change SCORE in car_collector.py to keep them.")).font = _f(italic=True, color="666666", size=8)
    settings = [("Score at fair price", SCORE["base"], "0.0"),
                ("% under fair per point", pct_per_point, "0%"),
                ("Good-signs cap", SCORE["green_cap"], "0.0")]
    for i, (k, val, fmt) in enumerate(settings):
        sm.cell(row=S0 + 2 + i, column=1, value=k).font = _f()
        c = sm.cell(row=S0 + 2 + i, column=2, value=val)
        c.font, c.number_format = _f(color="0000FF"), fmt
    sm.cell(row=S0 + 3, column=3, value=f"\u00b1{2 * sd * 100:.0f}% from fair \u2248 \u00b15 points").font = _f(italic=True, color="666666", size=8)
    BASE, PPP, CAP = f"Summary!$B${S0 + 2}", f"Summary!$B${S0 + 3}", f"Summary!$B${S0 + 4}"
    rowp = S0 + 6
    ranges = {}
    for title, items in (("Gearbox", SCORE["gearbox"]), ("Red flag", SCORE["red"]), ("Good sign", SCORE["green"])):
        header(rowp, 1, [title, "Points"], "6B7280", 9)
        for i, (k, val) in enumerate(items.items()):
            sm.cell(row=rowp + 1 + i, column=1, value=k).font = _f(size=9)
            c = sm.cell(row=rowp + 1 + i, column=2, value=val)
            c.font, c.number_format = _f(color="0000FF", size=9), "+0.0;-0.0;0"
        ranges[title] = (f"Summary!$A${rowp + 1}:$A${rowp + len(items)}", f"Summary!$B${rowp + 1}:$B${rowp + len(items)}")
        rowp += len(items) + 2
    GB_N, GB_P = ranges["Gearbox"]
    RED_N, RED_P = ranges["Red flag"]
    GRN_N, GRN_P = ranges["Good sign"]
    LR = rowp + 1                                      # "Last run" block goes here

    # ================= Cars =================
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws["A1"] = "Used cars in Canada \u2013 manual BMW M / E46 / E91 and Nissan Z, sorted by price"
    ws["A1"].font = _f(True, "1F2937", 14)
    ws["A2"], ws["B2"], ws["C2"], ws["D2"] = "As of", today, "Baseline", baseline
    for c in ("A2", "C2"):
        ws[c].font = _f(True, "555555")
    for c in ("B2", "D2"):
        ws[c].number_format = "yyyy-mm-dd"
        ws[c].font = _f(color="0000FF")
        ws[c].alignment = Alignment(horizontal="left")
    ws["F2"] = f'="Showing "&SUBTOTAL(103,B{first}:B{last})&" of "&COUNTA(B{first}:B{last})&" verified listings"'
    ws["F2"].font = _f(True, "1F2937")
    ws["A3"] = ("Sorted by price; sort by Score with the filter arrow. Score 0-10: price vs fair price for that model, "
                "km and year, plus gearbox and flags (Summary tab). Colours and symbols: Legend tab.")
    ws["A3"].font = _f(italic=True, color="666666", size=9)

    for i, h in enumerate(HEADERS, 1):
        c = ws.cell(row=4, column=i, value=h)
        c.font = _f(True, "FFFFFF", 10)
        c.fill = _fill("1F2937")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[4].height = 30

    n_embedded = n_formula = 0
    for n, r in enumerate(active):
        x = first + n
        ws.row_dimensions[x].height = 76
        fam_fill, fam_txt = FAMILY_STYLE.get(r["family"], ("666666", "FFFFFF", "", ""))[:2]
        body_lab, bf, bt = BODY_STYLE.get(r["body"], BODY_STYLE["unk"])
        tr_lab, tf, tt = TRANS_STYLE.get(r["trans"], TRANS_STYLE["u"])
        read = bool(r.get("flags_read"))
        v = {
            "Car": r["name"], "Trim": r.get("trim", ""), "Year": int(r["year"]) if r["year"] else None,
            "Body": body_lab, "Gearbox": tr_lab, "Price (CAD)": int(r["price"]),
            "Km": int(r["km"]) if r["km"] else None, "Location": r["place"],
            "Red flags": (r.get("red_flags") or "") if read else NOT_READ,
            "Good signs": (r.get("green_flags") or "") if read else NOT_READ,
            "Ad detail": r["note"], "Prev price": int(r["prev_price"]) if r["prev_price"] else None,
            "First seen": d(r["first_seen"]), "Last seen": d(r["last_seen"]), "Verified": d(r["verified"]),
            "Source": SRC_LABEL.get(r["source"], r["source"]), "Link": "Open ad \u2197",
            "Family key": r["family"], "URL": r["url"],
        }
        X = lambda h: f"{L[h]}{x}"
        mrow = f"MATCH({X('Family key')},Summary!$J${m0}:$J${m1},0)"
        v["Fair price"] = (f"=IFERROR(ROUND(EXP(INDEX(Summary!$K${m0}:$K${m1},{mrow})"
                           f"+INDEX(Summary!$L${m0}:$L${m1},{mrow})*{X('Km')}/10000"
                           f"+INDEX(Summary!$M${m0}:$M${m1},{mrow})*({X('Year')}-2000)"
                           f"+{CONV}*ISNUMBER(SEARCH(\"Convertible\",{X('Body')}))),-2),\"\")")
        v["vs fair"] = f'=IF({X("Fair price")}="","",1-{X("Price (CAD)")}/{X("Fair price")})'
        v["Score"] = (f"=ROUND(MAX(0,MIN(10,{BASE}+IF({X('vs fair')}=\"\",0,{X('vs fair')}/{PPP}"
                      f"*INDEX(Summary!$O${m0}:$O${m1},{mrow}))"
                      f"+IFERROR(INDEX({GB_P},MATCH({X('Gearbox')},{GB_N},0)),0)"
                      f"+SUMPRODUCT(ISNUMBER(SEARCH({RED_N},{X('Red flags')}))*{RED_P})"
                      f"+MIN({CAP},SUMPRODUCT(ISNUMBER(SEARCH({GRN_N},{X('Good signs')}))*{GRN_P})))),1)")
        v["Price change"] = f'=IF({X("Prev price")}="","",{X("Price (CAD)")}-{X("Prev price")})'
        v["New?"] = f'=IF(AND({X("First seen")}>$D$2,{X("First seen")}>=$B$2-2),"NEW","")'
        v["Rank key"] = f"={X('Score')}+(100000-ROW())/1000000"
        pic = local_photo(r)
        if pic and PILImage is not None:
            try:
                embed_photo(ws, pic, x)
                n_embedded += 1
            except Exception:
                pic = None
        if not pic:
            pf = photo_formula(r)
            if pf:
                v["Photo"] = pf
                n_formula += 1
        for h, val in v.items():
            c = ws.cell(row=x, column=COL[h], value=val)
            c.font = _f()
            c.alignment = Alignment(vertical="center", horizontal="center",
                                    wrap_text=h in ("Location", "Ad detail", "Red flags", "Good signs", "Trim"))
        for col in range(1, len(HEADERS) + 1):
            ws.cell(row=x, column=col).border = Border(bottom=thin)
        c = ws.cell(row=x, column=COL["Car"])
        c.fill, c.font = _fill(fam_fill), _f(True, fam_txt, 11)
        c = ws.cell(row=x, column=COL["Trim"])
        c.font = _f(size=9, color="333333")
        c = ws.cell(row=x, column=COL["Body"])
        c.fill, c.font = _fill(bf), _f(True, bt, 9)
        c = ws.cell(row=x, column=COL["Gearbox"])
        c.fill, c.font = _fill(tf), _f(True, tt, 9)
        ws.cell(row=x, column=COL["Year"]).font = _f(True)
        c = ws.cell(row=x, column=COL["Score"])
        c.number_format, c.font = "0.0", _f(True, size=13)
        c = ws.cell(row=x, column=COL["Price (CAD)"])
        c.number_format, c.font = "$#,##0", _f(True, size=11)
        ws.cell(row=x, column=COL["Fair price"]).number_format = "$#,##0"
        ws.cell(row=x, column=COL["Fair price"]).font = _f(color="555555")
        c = ws.cell(row=x, column=COL["vs fair"])
        c.number_format, c.font = '+0%" under";-0%" over";"at fair"', _f(True)
        c = ws.cell(row=x, column=COL["Km"])
        c.number_format, c.alignment = "#,##0", Alignment(vertical="center", horizontal="right")
        for h, colr, fill in (("Red flags", "9C0006", "FFE1E1"), ("Good signs", "0B5E2A", "E3F4E8")):
            c = ws.cell(row=x, column=COL[h])
            if c.value == NOT_READ:
                c.font = _f(italic=True, color="999999", size=8)
            elif c.value:
                c.font, c.fill = _f(True, colr, 9), _fill(fill)
        ws.cell(row=x, column=COL["Prev price"]).number_format = "$#,##0"
        ws.cell(row=x, column=COL["Price change"]).number_format = '\u25B2 $#,##0;\u25BC $#,##0;"-"'
        for h in ("First seen", "Last seen", "Verified"):
            ws.cell(row=x, column=COL[h]).number_format = "yyyy-mm-dd"
        c = ws.cell(row=x, column=COL["Ad detail"])
        c.font, c.alignment = _f(color="666666", size=9), Alignment(vertical="center", wrap_text=True)
        ws.cell(row=x, column=COL["New?"]).font = _f(True, "0B7A3E")
        c = ws.cell(row=x, column=COL["Link"])
        c.hyperlink, c.font = r["url"], _f(color="0563C1", underline="single")

    ws.conditional_formatting.add(rng("Year"), ColorScaleRule(start_type="min", start_color="F2F2F2",
                                                              end_type="max", end_color="A6A6A6"))
    ws.conditional_formatting.add(rng("Km"), DataBarRule(start_type="num", start_value=0, end_type="num",
                                                         end_value=250000, color="E8A33D", showValue=True))
    ws.conditional_formatting.add(rng("Score"), ColorScaleRule(start_type="num", start_value=0, start_color="F8696B",
                                                               mid_type="num", mid_value=5, mid_color="FFEB84",
                                                               end_type="num", end_value=10, end_color="63BE7B"))
    bounds = [("lessThan", ["15000"]), ("between", ["15000", "24999"]), ("between", ["25000", "34999"]),
              ("between", ["35000", "49999"]), ("greaterThanOrEqual", ["50000"])]
    for (op, fm), (_, fill, txt) in zip(bounds, TIERS):
        ws.conditional_formatting.add(rng("Price (CAD)"), CellIsRule(operator=op, formula=fm, fill=_fill(fill),
                                                                     font=Font(name=FONT, bold=True, color=txt, size=11)))
    GREEN_F, RED_F = Font(name=FONT, bold=True, color="0B7A3E"), Font(name=FONT, bold=True, color="C00000")
    for h, sign in (("vs fair", 1), ("Price change", -1)):
        c0 = f"{L[h]}{first}"
        ws.conditional_formatting.add(rng(h), FormulaRule(formula=[f"AND(ISNUMBER({c0}),{c0}*{sign}>0)"], font=GREEN_F))
        ws.conditional_formatting.add(rng(h), FormulaRule(formula=[f"AND(ISNUMBER({c0}),{c0}*{sign}<0)"], font=RED_F))
    ws.conditional_formatting.add(rng("Verified"), FormulaRule(formula=[f"${L['Verified']}{first}<$B$2-3"],
                                                               font=Font(name=FONT, color="C00000", italic=True)))
    widths = {"Photo": 18.5, "Car": 12, "Trim": 12, "Score": 10.5, "Year": 7, "Body": 14, "Gearbox": 14,
              "Price (CAD)": 12, "Fair price": 11, "vs fair": 11, "Km": 15, "Location": 18, "Red flags": 17,
              "Good signs": 19, "Ad detail": 34, "Prev price": 10, "Price change": 11, "First seen": 11,
              "New?": 6, "Last seen": 11, "Verified": 11, "Source": 11, "Link": 10, "Family key": 10,
              "Rank key": 10, "URL": 30}
    for h, w in widths.items():
        ws.column_dimensions[L[h]].width = w
    for h in ("Family key", "Rank key", "URL"):
        ws.column_dimensions[L[h]].hidden = True
    ws.freeze_panes = "C5"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.print_title_rows = "4:4"
    if active:
        ws.auto_filter.ref = f"A4:{get_column_letter(len(HEADERS))}{last}"

    # chart data on a hidden sheet: one km/price column pair per family (values, rewritten every run)
    cd = wb.create_sheet("ChartData")
    pts = {f: [(int(r["km"]), int(r["price"])) for r in active if r["family"] == f and r["km"]] for f in fams}
    for i, f in enumerate(fams):
        cd.cell(row=1, column=2 * i + 1, value=f"{FAMILY_STYLE[f][2]} km")
        cd.cell(row=1, column=2 * i + 2, value=FAMILY_STYLE[f][2])
        for n, (km, pr) in enumerate(pts[f]):
            cd.cell(row=2 + n, column=2 * i + 1, value=km)
            cd.cell(row=2 + n, column=2 * i + 2, value=pr)
    cd.sheet_state = "hidden"
    if active:
        ch = ScatterChart()
        prices = sorted(int(r["price"]) for r in active)
        cap = 120000
        above = sum(1 for p in prices if p > cap)
        if above > 3:
            cap, above = int(prices[-1] * 1.05), 0
        ch.title = "Price vs km" + (f" ({above} car{'s' if above > 1 else ''} above ${cap // 1000}k not shown)" if above else "")
        ch.style = 13
        ch.x_axis.title = "Km"
        ch.y_axis.title = "Price (CAD)"
        ch.y_axis.scaling.min = 0
        ch.y_axis.scaling.max = cap
        ch.y_axis.majorUnit = 20000 if cap <= 160000 else 50000
        ch.x_axis.scaling.min = 0
        ch.y_axis.number_format = "$#,##0"
        ch.x_axis.number_format = "#,##0"
        for i, f in enumerate(fams):
            xs = Reference(cd, min_col=2 * i + 1, min_row=2, max_row=1 + max(1, len(pts[f])))
            ys = Reference(cd, min_col=2 * i + 2, min_row=1, max_row=1 + max(1, len(pts[f])))
            s = Series(ys, xs, title_from_data=True)
            s.marker.symbol = "circle"
            s.marker.size = 7
            s.marker.graphicalProperties.solidFill = FAMILY_STYLE[f][0]
            s.marker.graphicalProperties.line.solidFill = FAMILY_STYLE[f][0]
            s.graphicalProperties.line.noFill = True
            ch.series.append(s)
        ch.roundedCorners = False
        ch.legend.position = "r"
        ch.legend.overlay = False
        ch.title.overlay = False
        ch.x_axis.title.overlay = False
        ch.y_axis.title.overlay = False
        ch.x_axis.delete = False
        ch.y_axis.delete = False
        ch.width, ch.height = 19, 11
        sm.add_chart(ch, f"J{T0}")

    # last run
    lr = LR
    sm.cell(row=lr, column=1, value="Last run").font = _f(True, "1F2937", 12)
    lines = [f"Workbook built {dt.datetime.now():%Y-%m-%d %H:%M}"]
    if run_info:
        lines += run_info
    lines.append(f"Photos: {n_embedded} embedded, {n_formula} as IMAGE() formulas (need internet + Microsoft 365)")
    for i, txt in enumerate(lines):
        c = sm.cell(row=lr + 1 + i, column=1, value=txt)
        c.font = _f(size=9)

    # ================= Legend =================
    lg = wb.create_sheet("Legend")
    lg.sheet_view.showGridLines = False
    lg.column_dimensions["A"].width = 24
    lg.column_dimensions["B"].width = 90
    r = 1
    lg.cell(row=r, column=1, value="How to read the Cars tab").font = _f(True, "1F2937", 14)
    r += 2

    def head(txt):
        nonlocal r
        lg.cell(row=r, column=1, value=txt).font = _f(True, "1F2937", 11)
        r += 1

    def swatch(label, fill, txt, note):
        nonlocal r
        c = lg.cell(row=r, column=1, value=label)
        if fill:
            c.fill = _fill(fill)
        c.font = _f(True, txt, 10)
        c.alignment = Alignment(horizontal="center")
        lg.cell(row=r, column=2, value=note).font = _f(size=10)
        lg.cell(row=r, column=2).alignment = Alignment(wrap_text=True, vertical="center")
        r += 1

    head("Score (0-10)")
    swatch("8.0", "63BE7B", "000000", "Priced well under the fair price for that model, km and year, with no red flags")
    swatch("5.0", "FFEB84", "000000", "About the fair price, nothing special either way")
    swatch("2.0", "F8696B", "000000", "Over-priced and/or red flags (rebuilt title, needs work, gearbox doubt)")
    swatch("Fair price", "", "555555", "What the model expects for that car (family, km, year, convertible or not). "
           "Estimated from these listings only, so it is a guide to compare cars, not a valuation")
    swatch("vs fair", "", "0B7A3E", "+12% under = listed 12% below the fair price; -8% over = 8% above")
    swatch("Confidence", "", "555555", "How much the price part of the score counts: n/(n+5) for a family of n cars "
           "(3 cars = 38%, 26 cars = 84%). Few cars = rough fair price, so it moves the score less")
    swatch("How it is fitted", "", "555555", "ln(price) = family level + km slope + year slope + convertible effect. "
           "Slopes are shared across families (own slopes for families with 12+ cars); the family level is the "
           "median, so one collector-grade car does not drag the rest. Rebuilt/salvage titles are left out")
    r += 1
    head("Flags (read from the ad text)")
    swatch("Red flags", "FFE1E1", "9C0006", "Rebuilt / salvage title, accident mentioned, sold as-is (not safetied), needs work, modified")
    swatch("Good signs", "E3F4E8", "0B5E2A", "Rod bearings, subframe, VANOS/timing or clutch done, service records, no accidents stated, one owner, winter stored")
    swatch(NOT_READ, "", "999999", "The ad text has not been read yet; the next daily run fills it in. Counts as no flags in the score")
    swatch("(blank)", "", "000000", "Ad was read and none of the phrases were found. Flags only catch what the seller wrote - always ask")
    r += 1
    head("Car (generation + model, one colour per family)")
    for fam, (fill, txt, short, desc) in FAMILY_STYLE.items():
        swatch(short, fill, txt, desc)
    r += 1
    head("Body")
    for k in ("sedan", "coupe", "conv", "wagon", "hatch"):
        lab, fill, txt = BODY_STYLE[k]
        swatch(lab, fill, txt, "")
    r += 1
    head("Gearbox")
    swatch(TRANS_STYLE["m"][0], TRANS_STYLE["m"][1], TRANS_STYLE["m"][2], "Manual confirmed by the listing's data and its text")
    swatch(TRANS_STYLE["c"][0], TRANS_STYLE["c"][1], TRANS_STYLE["c"][2],
           "The data field and the seller's text disagree (SMG/auto vs manual). Ask the seller before you drive out")
    swatch(TRANS_STYLE["u"][0], TRANS_STYLE["u"][1], TRANS_STYLE["u"][2],
           "Field says manual but the text never confirms it (common on E46/E60/E9X M cars that also came with SMG/DCT)")
    r += 1
    head("Price tiers (colour of the price cell)")
    for lab, fill, txt in TIERS:
        swatch(lab, fill, txt, "")
    r += 1
    head("Other cues")
    for k, v in (("Year", "Grey ramp: lighter = older, darker = newer"),
                 ("Km", "Orange bar: full width = 250,000 km"),
                 ("Price change", "Green \u25BC = price dropped since the previous check, red \u25B2 = went up"),
                 ("New?", "NEW = first seen in the last 3 days (not shown for the baseline day)"),
                 ("Verified", "Last day the ad's own page was opened and matched. Red italics = older than 3 days"),
                 ("Photo", "Embedded picture once the daily script has downloaded it; before that an IMAGE() formula "
                           "(Microsoft 365, needs internet)")):
        lg.cell(row=r, column=1, value=k).font = _f(True)
        lg.cell(row=r, column=2, value=v).font = _f()
        r += 1

    # ================= Gone =================
    gs = wb.create_sheet("Gone")
    gs.sheet_view.showGridLines = False
    gh = ["Car", "Year", "Last price", "Km", "Location", "First seen", "Last seen", "Days listed", "Link"]
    for i, h in enumerate(gh, 1):
        c = gs.cell(row=1, column=i, value=h)
        c.font, c.fill = _f(True, "FFFFFF"), _fill("1F2937")
        c.alignment = Alignment(horizontal="center")
    for n, rr in enumerate(gone):
        x = 2 + n
        vals = [rr["name"], int(rr["year"]) if rr["year"] else None, int(rr["price"]) if rr["price"] else None,
                int(rr["km"]) if rr["km"] else None, rr["place"], d(rr["first_seen"]), d(rr["last_seen"]),
                f"=G{x}-F{x}+1", "Open ad \u2197"]
        for i, val in enumerate(vals, 1):
            gs.cell(row=x, column=i, value=val).font = _f()
        gs.cell(row=x, column=3).number_format = "$#,##0"
        gs.cell(row=x, column=4).number_format = "#,##0"
        gs.cell(row=x, column=6).number_format = gs.cell(row=x, column=7).number_format = "yyyy-mm-dd"
        gs.cell(row=x, column=9).hyperlink = rr["url"]
        gs.cell(row=x, column=9).font = _f(color="0563C1", underline="single")
    if not gone:
        gs.cell(row=3, column=1, value="Nothing yet. Listings the daily run finds sold or removed are moved here.").font = _f(italic=True, color="666666")
    for i, w in enumerate([14, 7, 11, 10, 22, 11, 11, 11, 10], 1):
        gs.column_dimensions[get_column_letter(i)].width = w

    # ================= README =================
    rd = wb.create_sheet("README")
    rd.sheet_view.showGridLines = False
    rd.column_dimensions["A"].width = 110
    lines = [
        ("How this workbook stays up to date", True),
        ("Click 'Update now' at the top of the Cars tab. It runs car_collector.py in the background (5-15 min, a minimised window shows progress) and reloads the tabs when it finishes. You can keep working meanwhile.", False),
        ("'Test sites' opens a window that shows what the script can read from each site (the old --selftest).", False),
        ("If a scheduled daily run has produced newer data, CarWatch.xlsm loads it automatically when you open it.", False),
        ("The script writes cars_data.xlsx and a dated CSV snapshot (car_listings_YYYYMMDD.csv) in the same folder; CarWatch.xlsm copies the tabs from cars_data.xlsx.", False),
        ("Missing Python packages (requests, openpyxl, pillow) are installed automatically on the first run.", False),
        ("Photos: downloaded into the photos folder and embedded, so they show in any Excel or phone preview. Cars without a downloaded photo use IMAGE() (Microsoft 365 + internet).", False),
        ("", False),
        ("Optional: daily run without opening Excel (Command Prompt)", True),
        ("schtasks /Create /SC DAILY /ST 07:30 /TN CarWatch /TR \"python C:\\CarWatch\\car_collector.py\"", False),
        ("", False),
        ("Rules the daily run follows", True),
        ("Searches Autotrader, Kijiji and CarPages (list at the top of the script). A listing is added only if its own page opened and showed a matching car with a manual (or conflicting) gearbox.", False),
        ("Ads it rejects (automatic, other models) are remembered so the daily budget goes to ads it has not seen.", False),
        ("A listing is marked gone only on a clear signal (404/410, redirect to another listing, different odometer, 'no longer available'). A page it cannot read is kept as it was.", False),
        ("If more than about a third of the checks look gone at once, nothing is marked gone (likely a block page or a site change).", False),
        ("Flags and trim are re-read from each ad every day. The score settings live in SCORE at the top of the script.", False),
        ("", False),
        ("Known limits", True),
        ("The live-page parsing has not been tested against autotrader.ca, Kijiji or CarPages from where this was built. Run --selftest first; the Summary tab's 'Last run' lists how many ads each search returned.", False),
        ("Flags come from the seller's own words, so a missing red flag is not a clean history. Get a Carfax and an inspection.", False),
        ("The fair price is a statistical estimate from these listings, not an appraisal; families with fewer than 5 cars are rough.", False),
        ("Check each site's terms of use before automating requests. The script waits 2 seconds between page requests.", False),
        ("Prices are as listed (before tax). Km and CAD throughout.", False),
    ]
    for i, (txt, bold) in enumerate(lines, 1):
        c = rd.cell(row=i, column=1, value=txt)
        c.font = _f(bold, "1F2937" if bold else "000000", 12 if bold else 10)
        c.alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(path)
    return len(active), len(gone), n_embedded, n_formula


def write_workbook(rows, today, run_info=None):
    path = os.path.join(DATA_DIR, XLSX_NAME)
    try:
        return path, build_xlsx(rows, path, today, run_info)
    except PermissionError:
        alt = os.path.join(DATA_DIR, f"cars_data_{today:%Y%m%d}.xlsx")
        print(f"!! {XLSX_NAME} is open in Excel - wrote {os.path.basename(alt)} instead")
        return alt, build_xlsx(rows, alt, today, run_info)


# --------------------------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------------------------
def selftest(sess, rows):
    by_src = {}
    for r in rows:
        if r["status"] == "active":
            by_src.setdefault(r["source"], []).append(r)
    sample = by_src.get("autotrader", [])[:2] + [v[0] for k, v in by_src.items() if k != "autotrader"][:2]
    for row in sample:
        print("\n", row["source"], row["name"], row["year"], row["place"],
              "\n  stored :", row["price"], "$", row["km"], "km", row["trans"])
        outcome, res = check_listing(sess, row)
        if outcome == "ok":
            print("  read   :", res.get("price"), "$", res.get("km"), "km | gearbox field:", res.get("trans_field"),
                  "| body:", res.get("body_raw"), "| place:", res.get("place"))
            print("  text   :", len(res.get("desc") or ""), "chars of description | flags:", flags_of(res.get("desc")))
            print("  photo  :", photo_of(row, res["_html"], res) or "not found")
        else:
            print("  ->", outcome, res)
    print("\nSearch pages (ads found on page 1):")
    for source, base in SEARCHES:
        r = fetch(sess, base)
        n = len({m.group(2) for m in LINK_RE[source].finditer(r.text)}) if r is not None and r.status_code == 200 else None
        print(f"  {n if n is not None else 'ERR':>4}  {base}")


def main(argv=None):
    ap = argparse.ArgumentParser(description="Daily used-car watcher and Excel builder")
    ap.add_argument("--no-network", action="store_true", help="rebuild the workbook from the latest snapshot only")
    ap.add_argument("--selftest", action="store_true", help="read stored listings + search pages and print what was parsed")
    ap.add_argument("--max-new", type=int, default=MAX_NEW_PER_RUN)
    ap.add_argument("--max-minutes", type=float, default=0,
                    help="stop opening pages after this many minutes (the run still saves and exports; the rest waits)")
    ap.add_argument("--config", help="path of cars_config.json (default: next to the script, or one folder up)")
    ap.add_argument("--data-dir", help="folder for snapshots, workbook and log (default: the script's folder)")
    ap.add_argument("--export-json", help="also write the web app's data file here (e.g. site/data/cars.json)")
    ap.add_argument("--no-xlsx", action="store_true", help="do not build the Excel workbook")
    ap.add_argument("--no-photos", action="store_true", help="do not download photo thumbnails")
    ap.add_argument("--skip-sources", default="", help="comma list of sites not to visit, e.g. carpages")
    ap.add_argument("--keep-snapshots", type=int, default=0, help="delete snapshots older than the newest N (0 = keep all)")
    ap.add_argument("--no-comps", action="store_true", help="do not read sold prices (Bring a Trailer)")
    ap.add_argument("--no-fipe", action="store_true", help="do not fetch FIPE reference prices (Brazil)")
    ap.add_argument("--done-file", help="write 'ok ...' or 'error: ...' here when the run ends (used by CarWatch.xlsm)")
    args = ap.parse_args(argv)
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    status = "error: stopped before finishing"
    if args.done_file:
        with open(os.path.join(HERE, args.done_file), "w", encoding="utf-8") as f:
            f.write("running\n")
    try:
        status = run(args) or "ok"
    except SystemExit as e:
        status = f"error: {e}"
        raise
    except Exception as e:
        import traceback
        traceback.print_exc()
        status = f"error: {type(e).__name__}: {e}"
    finally:
        if args.done_file:
            with open(os.path.join(HERE, args.done_file), "w", encoding="utf-8") as f:
                f.write(status + "\n")
        print(status, flush=True)
    if status.startswith("error"):
        sys.exit(1)


def run(args):
    global DATA_DIR, PHOTO_DIR
    if args.data_dir:
        DATA_DIR = os.path.abspath(args.data_dir)
        PHOTO_DIR = os.path.join(DATA_DIR, "photos")
        os.makedirs(DATA_DIR, exist_ok=True)
    print("config:", load_config(args.config), flush=True)
    if getattr(args, "max_minutes", 0):
        DEADLINE[0] = time.time() + args.max_minutes * 60
    skip = {x.strip() for x in (args.skip_sources or "").split(",") if x.strip()}
    if skip:
        SEARCHES[:] = [x for x in SEARCHES if x[0] not in skip]
        SKIP_SOURCES.update(skip)
    today = dt.date.today()
    prev_path = latest_snapshot()
    rows = load_rows(prev_path)
    print(f"previous snapshot: {os.path.basename(prev_path) if prev_path else 'none'} ({len(rows)} rows)")
    for r in rows:
        if r["status"] == "active" and r["price"].isdigit() and int(r["price"]) < MIN_PRICE:
            if r["prev_price"].isdigit() and int(r["prev_price"]) >= MIN_PRICE:
                r["price"], r["prev_price"] = r["prev_price"], ""
                red = [f for f in (r.get("red_flags") or "").split("; ") if f and f != "Price hidden"]
                r["red_flags"] = "; ".join(red + ["Price hidden"])
            else:
                r["status"] = "unverified"            # no real price known: keep it out of the list until read again
    # automatics are now kept for some families: forget earlier 'automatic' rejections on those sites
    rows = [r for r in rows if not (r["status"] == "skip" and r["source"] in JSON_SEARCH and r["note"].startswith("automatic"))]
    rows, n_again = recheck_old_skips(rows)
    if n_again:
        print(f"{n_again} earlier rejected ads get a second look (they may be one of the newly added cars)", flush=True)
    n_moved = resync_families(rows)
    if n_moved:
        print(f"{n_moved} saved ads moved to a renamed or re-cut generation", flush=True)
    for r in rows:
        apply_split(r)
    n_hist = backfill_hist(rows)
    if n_hist:
        print(f"price history rebuilt from saved snapshots for {n_hist} listings")

    sess = None
    if not args.no_network:
        if requests is None:
            sys.exit("The 'requests' package is missing: pip install requests")
        sess = requests.Session()
        sess.headers.update({"User-Agent": UA, "Accept-Language": "en-CA,en;q=0.8,fr-CA;q=0.5"})
    if args.selftest:
        selftest(sess, rows)
        return "ok selftest"

    snap = Snapshot(os.path.join(DATA_DIR, SNAPSHOT_FMT.format(today)))       # created before any network call
    stats = dict(checked=0, failed=0, gone=0, back=0, price_changes=0, candidates=0, new=0, skipped=0,
                 photos_new=0, searches=[])
    new_rows = []
    if sess is not None:
        found = search_ids(sess, {r["listing_id"] for r in rows}, stats)    # searches first: their data
        refresh_existing(sess, rows, today.isoformat(), stats)               # confirms ads without opening them
        new_rows = discover_new(sess, rows, today.isoformat(), stats, args.max_new, found)
        for r in rows + new_rows:
            apply_split(r)
        if not args.no_photos:
            download_photos(sess, rows + new_rows, stats)
    gear_flags(rows + new_rows)
    if COUNTRY["country"] == "BR" and FIPE_CFG:
        fpath = os.path.join(DATA_DIR, "fipe_cache.json")
        if sess is not None and not args.no_fipe:
            try:
                update_fipe(sess, rows + new_rows, fpath)
            except Exception as e:                       # never let the reference prices break the run
                print(f"FIPE update failed: {e}")
        load_fipe(fpath)
    for r in rows + new_rows:
        snap.add(r)
    snap.finish()
    all_rows = rows + new_rows
    comps = []
    if args.export_json and not args.no_comps:
        try:
            comps = update_comps(sess, stats)
        except Exception as e:                           # sold prices are extra: never stop the run for them
            print("!! sold prices not updated:", type(e).__name__, e)

    run_info = []
    if sess is not None:
        run_info.append(f"Checked {stats['checked']} ads ({stats['failed']} unreadable), {stats['price_changes']} price "
                        f"changes, {stats['gone']} newly gone, {stats['back']} back")
        run_info.append(f"Search: {stats['candidates']} unseen ads, {stats['new']} added, {stats['skipped']} rejected "
                        f"(remembered), {stats['photos_new']} new photos"
                        + (f", {stats['prefiltered']} other cars in the results not opened" if stats.get("prefiltered") else ""))
        if time_up():
            run_info.append(f"Out of time: {stats.get('late', 0)} ads not re-checked; the next run carries on")
        if stats.get("new_auto") or stats.get("details_read") or stats.get("seen_in_search"):
            run_info.append(f"Brazil: {stats.get('seen_in_search', 0)} ads confirmed from search pages (not opened), "
                            f"{stats.get('new_auto', 0)} automatics added from search data, "
                            f"{stats.get('details_read', 0)} ads read for their description")
        if "comps_total" in stats:
            run_info.append(f"Sold prices: {stats['comps_new']} new Bring a Trailer results "
                            f"({stats['comps_pages']} pages read), {stats['comps_total']} stored")
        for source, base, n in stats["searches"]:
            if n == 0 and source in JSON_SEARCH:          # NaPista/Mobiauto pages that exist: simply none for sale today
                run_info.append(f"   {n!s:>8} ads  {base}  (none for sale today)")
                continue
            flag = "  <-- check this search" if (n == 0 or isinstance(n, str)) else ""
            run_info.append(f"   {n!s:>8} ads  {base}{flag}")
    else:
        run_info.append("Rebuilt from the saved snapshot (no web access this run)")
    n_active = sum(1 for r in all_rows if r["status"] == "active")
    n_gone = sum(1 for r in all_rows if r["status"] == "gone")
    if not args.no_xlsx:
        path, (n_active, n_gone, n_emb, n_form) = write_workbook(all_rows, today, run_info)
        print(f"workbook: {os.path.basename(path)} | photos embedded {n_emb} / formula {n_form}")
    if args.export_json:
        export_json(all_rows, args.export_json, run_info, stats if sess is not None else None, comps)
        print("web data:", args.export_json)
    if args.keep_snapshots:
        prune_snapshots(args.keep_snapshots)
    print(f"snapshot: {os.path.basename(snap.path)} | active {n_active}, gone {n_gone}")
    for line in run_info:
        print(line)
    with open(os.path.join(DATA_DIR, "collector_log.txt"), "a", encoding="utf-8") as f:
        f.write(f"{dt.datetime.now():%Y-%m-%d %H:%M} active={n_active} gone={n_gone} "
                f"{ {k: v for k, v in stats.items() if k != 'searches'} } searches={stats['searches']}\n")
    if sess is None:
        return f"ok {n_active} cars (rebuilt from saved data)"
    return (f"ok {n_active} cars: {stats['new']} new, {stats['gone']} sold/removed, "
            f"{stats['price_changes']} price changes")


try:
    load_config()
except FileNotFoundError:
    pass

if __name__ == "__main__":
    main()
