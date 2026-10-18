#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Zakkingskaart Amsterdam
=======================

Haalt alle meetbouten en hun metingen op bij de gemeente Amsterdam
(api.data.amsterdam.nl/v1/meetbouten/), berekent per meetbout hoe snel de
gevel zakt, en maakt daar een interactieve kaart van:

  rood    = zakt snel            (standaard: 2,0 mm/jaar of meer)
  geel    = zakt normaal         (standaard: 1,0 tot 2,0 mm/jaar)
  groen   = zakt nauwelijks      (standaard: minder dan 1,0 mm/jaar)
  groen, gestippeld/omrand = gestabiliseerd: eerst snel, nu nauwelijks
            (vaak een teken van funderingsherstel)
  grijs   = geen recente of te weinig metingen

Alle grenzen zijn op de kaart zelf aan te passen.

Gebruik (zie ook LEESMIJ):
    python3 zakkingskaart.py              # alles ophalen en kaart maken
    python3 zakkingskaart.py --proef      # snelle proefrun (klein deel van de data)
    python3 zakkingskaart.py --vernieuw   # data opnieuw ophalen (cache negeren)
    python3 zakkingskaart.py --site       # website maken in de map 'site'
                                          # (alleen nieuwe metingen ophalen)
    python3 zakkingskaart.py --help       # alle opties

Er zijn geen extra pakketten nodig: alleen Python 3.8 of nieuwer.
Uitvoer (standaard in de map 'zakkingskaart_uitvoer'):
    zakkingskaart.html            -> open in je browser
    meetbouten_analyse.csv        -> alle cijfers per meetbout (Excel, puntkomma)
    cache/                        -> opgehaalde ruwe data (voor snelle herhaling)

Bron: Gemeente Amsterdam, dataset Meetbouten. De kleuren zijn een indicatie op
basis van gevelmetingen, geen oordeel over de fundering van een individueel pand.
"""

import argparse
import base64
import csv
import gzip
import hashlib
import json
import math
import os
import re
import ssl
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
from datetime import date, timedelta

API = "https://api.data.amsterdam.nl/v1/meetbouten/"
USER_AGENT = "zakkingskaart/1.0 (open-source hobbyproject)"

# Velden die we per meting nodig hebben (houdt de download klein).
METING_FIELDS = "identificatie,hoortBijMeetboutId,datum,zakkingCumulatief,typeMeting,publiceerbaar"

# Website-modus: Leaflet wordt zelf gehost (geen verzoeken van bezoekers naar een CDN).
# De bestanden worden bij de eerste run opgehaald en gecontroleerd op de officiele
# SRI-hashes van Leaflet 1.9.4 (zie leafletjs.com/download.html).
LEAFLET_VERSION = "1.9.4"
LEAFLET_SRI = {
    "leaflet.js": "sha256-20nQCchB9co0qIjJZRGuk2/Z9VM+kNiyxNV1lvTlZBo=",
    "leaflet.css": "sha256-p4NxAoJBhIIN+hmNHrzRCf9tD/miZyoHS5obTRR9BMY=",
}
LEAFLET_IMAGES = ("layers.png", "layers-2x.png", "marker-icon.png", "marker-icon-2x.png", "marker-shadow.png")
LEAFLET_CDNS = (  # pakketmap; bestanden staan in dist/, de licentie in de hoofdmap
    "https://unpkg.com/leaflet@%s/" % LEAFLET_VERSION,
    "https://cdn.jsdelivr.net/npm/leaflet@%s/" % LEAFLET_VERSION,
)
STATE_FILE = "zakkingskaart_state.json"

try:  # nette uitvoer op Windows-consoles
    sys.stdout.reconfigure(errors="replace")
except Exception:
    pass


def log(msg=""):
    print(msg, flush=True)


def fmt_int(n):
    return f"{n:,}".replace(",", ".")


# --------------------------------------------------------------------------
# HTTP
# --------------------------------------------------------------------------

class FatalHTTP(Exception):
    def __init__(self, code, url, body=""):
        super().__init__(f"HTTP {code} voor {url}")
        self.code = code
        self.body = body


def http_get_json(url, retries=8, timeout=120, raw_bytes=False):
    delay = 2.0
    last_err = None
    for attempt in range(1, retries + 1):
        req = urllib.request.Request(url, headers={
            "Accept": "*/*" if raw_bytes else "application/hal+json, application/json",
            "User-Agent": USER_AGENT,
        })
        try:
            with urllib.request.urlopen(req, timeout=timeout) as r:
                raw = r.read()
            if raw_bytes:
                return raw
            return json.loads(raw.decode("utf-8"))
        except urllib.error.HTTPError as e:
            body = ""
            try:
                body = e.read().decode("utf-8", "replace")[:500]
            except Exception:
                pass
            if e.code in (429, 500, 502, 503, 504):
                ra = e.headers.get("Retry-After") if e.headers else None
                wait = float(ra) if (ra and ra.replace(".", "", 1).isdigit()) else delay
                log(f"   server antwoordt {e.code}, opnieuw over {wait:.0f} s (poging {attempt}/{retries})")
                time.sleep(wait)
                delay = min(delay * 2, 90)
                last_err = e
                continue
            raise FatalHTTP(e.code, url, body)
        except urllib.error.URLError as e:
            reason = getattr(e, "reason", e)
            if isinstance(reason, ssl.SSLCertVerificationError) or "CERTIFICATE_VERIFY_FAILED" in str(reason):
                log("\nFOUT: Python kan het beveiligingscertificaat niet controleren.")
                log("Op een Mac los je dit op door in Finder naar Programma's > Python 3.x te gaan")
                log("en dubbel te klikken op 'Install Certificates.command'. Draai daarna het script opnieuw.\n")
                sys.exit(2)
            log(f"   verbindingsprobleem ({reason}), opnieuw over {delay:.0f} s (poging {attempt}/{retries})")
            time.sleep(delay)
            delay = min(delay * 2, 90)
            last_err = e
        except (TimeoutError, ConnectionError, json.JSONDecodeError) as e:
            log(f"   onderbroken ({e.__class__.__name__}), opnieuw over {delay:.0f} s (poging {attempt}/{retries})")
            time.sleep(delay)
            delay = min(delay * 2, 90)
            last_err = e
    raise RuntimeError(f"Ophalen mislukt na {retries} pogingen: {url} ({last_err})")


def build_url(collection, params):
    return API + collection + "/?" + urllib.parse.urlencode(params, safe=",[]")


def with_page(url, page):
    parts = urllib.parse.urlsplit(url)
    q = dict(urllib.parse.parse_qsl(parts.query, keep_blank_values=True))
    q["page"] = str(page)
    return urllib.parse.urlunsplit(parts._replace(query=urllib.parse.urlencode(q, safe=",[]")))


def page_number(url):
    q = dict(urllib.parse.parse_qsl(urllib.parse.urlsplit(url).query))
    try:
        return int(q.get("page", "1"))
    except ValueError:
        return 1


def embedded_items(data, collection):
    emb = data.get("_embedded") or {}
    if collection in emb and isinstance(emb[collection], list):
        return emb[collection]
    for v in emb.values():
        if isinstance(v, list):
            return v
    if isinstance(data, list):
        return data
    return []


# --------------------------------------------------------------------------
# Download met cache en hervatten
# --------------------------------------------------------------------------

def slim_meting(rec):
    bout = rec.get("hoortBijMeetboutId")
    if bout in (None, ""):
        hb = rec.get("hoortBijMeetbout")
        if isinstance(hb, dict):
            bout = hb.get("identificatie") or hb.get("id")
        elif isinstance(hb, (str, int)):
            bout = hb
    if bout in (None, ""):
        lk = (rec.get("_links") or {}).get("hoortBijMeetbout")
        if isinstance(lk, dict):
            bout = lk.get("identificatie") or lk.get("title")
    return {
        "i": rec.get("identificatie"),
        "b": None if bout in (None, "") else str(bout),
        "d": rec.get("datum"),
        "z": rec.get("zakkingCumulatief"),
        "t": rec.get("typeMeting"),
        "p": rec.get("publiceerbaar", True),
    }


def download(collection, base_params, cache_dir, refresh, page_size, delay, max_pages, slim=None, key="identificatie",
             name=None):
    """Haalt een hele collectie op (hervatbaar). 'name' is het voorvoegsel van de cachebestanden
    (standaard de naam van de collectie)."""
    os.makedirs(cache_dir, exist_ok=True)
    name = name or collection
    final = os.path.join(cache_dir, f"{name}.jsonl.gz")
    part = os.path.join(cache_dir, f"{name}.part.jsonl")
    statef = os.path.join(cache_dir, f"{name}.state.json")

    if refresh:
        for f in (final, part, statef):
            if os.path.exists(f):
                os.remove(f)

    if os.path.exists(final):
        log(f"-> {collection}: uit cache ({final})")
        return load_jsonl_gz(final)

    params = dict(base_params)
    params["_format"] = "json"
    params["_pageSize"] = str(page_size)

    url = None
    fetched = 0
    pages = 0
    if os.path.exists(statef) and os.path.exists(part):
        try:
            st = json.load(open(statef, encoding="utf-8"))
            url = st.get("next")
            fetched = st.get("count", 0)
            pages = st.get("pages", 0)
            log(f"-> {collection}: hervat download na {fetched:,} records".replace(",", "."))
        except Exception:
            url = None
    if url is None:
        if os.path.exists(part):
            os.remove(part)
        url = build_url(collection, params)
        log(f"-> {collection}: ophalen ...")

    first = True
    t0 = time.time()
    while url:
        try:
            data = http_get_json(url)
        except FatalHTTP as e:
            if first and e.code == 400:
                # Valterug: zonder veldselectie, en/of met kleinere pagina's
                if "_fields" in params:
                    log("   (veldselectie niet geaccepteerd, verder zonder)")
                    params.pop("_fields")
                elif int(params["_pageSize"]) > 1000:
                    log("   (paginagrootte te groot, verder met 1000)")
                    params["_pageSize"] = "1000"
                else:
                    raise
                url = build_url(collection, params)
                continue
            raise
        first = False
        items = embedded_items(data, collection)
        if not items:
            break
        with open(part, "a", encoding="utf-8") as fh:
            for rec in items:
                fh.write(json.dumps(slim(rec) if slim else rec, ensure_ascii=False, separators=(",", ":")) + "\n")
        fetched += len(items)
        pages += 1

        nxt = ((data.get("_links") or {}).get("next") or {})
        nxt = nxt.get("href") if isinstance(nxt, dict) else (nxt if isinstance(nxt, str) else None)
        if not nxt and len(items) >= int(params["_pageSize"]):
            nxt = with_page(url, page_number(url) + 1)
        url = nxt
        if max_pages and pages >= max_pages:
            url = None
        with open(statef, "w", encoding="utf-8") as fh:
            json.dump({"next": url, "count": fetched, "pages": pages}, fh)

        el = time.time() - t0
        log(f"   {fmt_int(fetched):>9} records  ({pages} pagina's, {el / 60:.1f} min)")
        if delay:
            time.sleep(delay)

    # afronden: ontdubbelen en comprimeren
    seen = set()
    out = []
    if os.path.exists(part):
        with open(part, encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if not line:
                    continue
                rec = json.loads(line)
                k = rec.get(key)
                if k is not None:
                    if k in seen:
                        continue
                    seen.add(k)
                out.append(rec)
    write_jsonl_gz(final, out)
    for f in (part, statef):
        if os.path.exists(f):
            os.remove(f)
    log(f"   klaar: {len(out):,} {collection}".replace(",", "."))
    return out


def load_jsonl_gz(path):
    out = []
    with gzip.open(path, "rt", encoding="utf-8") as fh:
        for line in fh:
            line = line.strip()
            if line:
                out.append(json.loads(line))
    return out


def write_jsonl_gz(path, recs):
    """Schrijft atomair: eerst naar een tijdelijk bestand, dan vervangen. Een onderbroken
    run laat zo nooit een half cachebestand achter."""
    tmp = path + ".tmp"
    with gzip.open(tmp, "wt", encoding="utf-8") as fh:
        for rec in recs:
            fh.write(json.dumps(rec, ensure_ascii=False, separators=(",", ":")) + "\n")
    os.replace(tmp, path)


def write_text(path, text):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(text)
    os.replace(tmp, path)


# --------------------------------------------------------------------------
# Website-modus: cache bijhouden tussen runs
# --------------------------------------------------------------------------

def load_state(cache_dir):
    try:
        with open(os.path.join(cache_dir, STATE_FILE), encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return {}


def save_state(cache_dir, state):
    write_text(os.path.join(cache_dir, STATE_FILE), json.dumps(state, indent=2, ensure_ascii=False))


def days_since(iso):
    try:
        return (date.today() - date.fromisoformat(str(iso)[:10])).days
    except (TypeError, ValueError):
        return None


def refresh_collection(collection, base_params, cache_dir, args, max_pages, slim=None, key="identificatie"):
    """Haalt een collectie volledig opnieuw op onder een tijdelijke naam en vervangt de cache
    pas als dat gelukt is. Mislukt het, dan blijft de oude cache staan."""
    tmpname = collection + ".nieuw"
    stale = os.path.join(cache_dir, tmpname + ".jsonl.gz")
    if os.path.exists(stale):  # resultaat van een eerder afgebroken run: niet hergebruiken
        os.remove(stale)
    recs = download(collection, base_params, cache_dir, False, args.paginagrootte, args.pauze, max_pages,
                    slim=slim, key=key, name=tmpname)
    os.replace(stale, os.path.join(cache_dir, collection + ".jsonl.gz"))
    return recs


def update_metingen(cache_dir, args, max_pages, state):
    """Werkt de metingen-cache bij. Normaal alleen de metingen vanaf (laatst bekende datum
    min een terugkijkvenster), met het filter datum[gte]. Volledig opnieuw als er nog geen
    cache is, bij --vernieuw, of als de laatste volledige download te oud is (vangt
    correcties en verwijderingen in oudere metingen op)."""
    final = os.path.join(cache_dir, "metingen.jsonl.gz")
    params = {"_fields": METING_FIELDS}
    last_full = state.get("metingenVolledig")
    if last_full is None and os.path.exists(final):
        last_full = date.fromtimestamp(os.path.getmtime(final)).isoformat()
    age = days_since(last_full)
    reason = None
    if args.vernieuw:
        reason = "--vernieuw"
    elif not os.path.exists(final):
        reason = "nog geen cache"
    elif age is None or age >= args.volledig_na:
        reason = f"laatste volledige download {age} dagen geleden"

    if reason is None:
        old = load_jsonl_gz(final)
        dates = [str(m.get("d") or "")[:10] for m in old if m.get("d")]
        if not dates:
            reason = "cache bevat geen datums"
    if reason is not None:
        log(f"-> metingen: volledig ophalen ({reason})")
        recs = refresh_collection("metingen", params, cache_dir, args, max_pages, slim=slim_meting, key="i")
        state["metingenVolledig"] = date.today().isoformat()
        state["laatsteUpdate"] = {"soort": "volledig", "metingen": len(recs)}
        return recs

    since = (date.fromisoformat(max(dates)) - timedelta(days=args.terugkijk)).isoformat()
    log(f"-> metingen: {fmt_int(len(old))} in cache, laatste datum {max(dates)}; "
        f"nieuwe ophalen vanaf {since} ({args.terugkijk} dagen terug)")
    dparams = dict(params)
    dparams["datum[gte]"] = since
    try:
        delta = download("metingen", dparams, cache_dir, True, args.paginagrootte, args.pauze, max_pages,
                         slim=slim_meting, key="i", name="metingen.delta")
    except FatalHTTP as e:
        log(f"   filter op datum niet geaccepteerd (HTTP {e.code}); valt terug op volledig ophalen")
        recs = refresh_collection("metingen", params, cache_dir, args, max_pages, slim=slim_meting, key="i")
        state["metingenVolledig"] = date.today().isoformat()
        state["laatsteUpdate"] = {"soort": "volledig (filter niet geaccepteerd)", "metingen": len(recs)}
        return recs

    ignored = sum(1 for m in delta if m.get("d") and str(m["d"])[:10] < since)
    if ignored:
        log(f"   let op: {fmt_int(ignored)} metingen van voor {since} ontvangen; het datumfilter lijkt "
            "genegeerd te worden. Het resultaat klopt wel, maar de download is groter dan nodig.")

    merged = {}
    loose = []
    for m in old:
        if m.get("i") is None:
            loose.append(m)
        else:
            merged[m["i"]] = m
    new = changed = 0
    for m in delta:
        k = m.get("i")
        if k is None:
            continue
        if k not in merged:
            new += 1
        elif merged[k] != m:
            changed += 1
        merged[k] = m
    recs = list(merged.values()) + loose
    if new or changed:
        write_jsonl_gz(final, recs)
    delta_file = os.path.join(cache_dir, "metingen.delta.jsonl.gz")
    if os.path.exists(delta_file):
        os.remove(delta_file)
    log(f"   {fmt_int(len(delta))} opgehaald: {fmt_int(new)} nieuw, {fmt_int(changed)} gewijzigd; "
        f"totaal {fmt_int(len(recs))}")
    state["laatsteUpdate"] = {"soort": "incrementeel", "vanaf": since, "opgehaald": len(delta),
                              "nieuw": new, "gewijzigd": changed, "metingen": len(recs)}
    return recs


def sri_of(data):
    return "sha256-" + base64.b64encode(hashlib.sha256(data).digest()).decode("ascii")


def vendor_leaflet(site_dir, cache_dir):
    """Zet Leaflet in site/leaflet/. Haalt de bestanden eenmalig op (bewaard in de cache) en
    controleert leaflet.js en leaflet.css op de officiele hashes."""
    src = os.path.join(cache_dir, "leaflet-" + LEAFLET_VERSION)
    os.makedirs(os.path.join(src, "images"), exist_ok=True)

    def fetch(rel):
        last = None
        for base in LEAFLET_CDNS:
            try:
                return http_get_json(base + rel, retries=3, timeout=60, raw_bytes=True)
            except (FatalHTTP, RuntimeError) as e:
                last = e
        raise RuntimeError(f"Leaflet-bestand {rel} niet op te halen ({last})")

    files = {}
    for fn, want in LEAFLET_SRI.items():
        p = os.path.join(src, fn)
        data = open(p, "rb").read() if os.path.exists(p) else None
        if data is None or sri_of(data) != want:
            data = fetch("dist/" + fn)
            got = sri_of(data)
            if got != want:
                raise RuntimeError(f"{fn} heeft niet de verwachte hash (verwacht {want}, gekregen {got}). "
                                   "Gestopt om geen onbekende code te publiceren.")
            with open(p, "wb") as fh:
                fh.write(data)
        files[fn] = data
    for fn in LEAFLET_IMAGES:
        rel = "images/" + fn
        p = os.path.join(src, "images", fn)
        data = open(p, "rb").read() if os.path.exists(p) else None
        if not data or not data.startswith(b"\x89PNG"):
            data = fetch("dist/" + rel)
            if not data.startswith(b"\x89PNG"):
                raise RuntimeError(f"Leaflet-afbeelding {fn} is geen PNG-bestand.")
            with open(p, "wb") as fh:
                fh.write(data)
        files[rel] = data
    # De BSD-2-licentie van Leaflet vereist dat de licentietekst wordt meegeleverd.
    p = os.path.join(src, "LICENSE")
    data = open(p, "rb").read() if os.path.exists(p) else b""
    if b"Redistribution" not in data:
        data = fetch("LICENSE")
        if b"Redistribution" not in data or b"Agafonkin" not in data:
            raise RuntimeError("De licentietekst van Leaflet ziet er niet uit zoals verwacht.")
        with open(p, "wb") as fh:
            fh.write(data)
    files["LICENSE"] = data

    dst = os.path.join(site_dir, "leaflet")
    os.makedirs(os.path.join(dst, "images"), exist_ok=True)
    for rel, data in files.items():
        if rel == "LICENSE":
            rel = "LICENSE.txt"  # met extensie, zodat browsers het als tekst tonen
        with open(os.path.join(dst, rel), "wb") as fh:
            fh.write(data)


# --------------------------------------------------------------------------
# Meetbouten interpreteren
# --------------------------------------------------------------------------

def rd_to_wgs84(x, y):
    """Rijksdriehoekscoordinaten -> WGS84 (benadering, nauwkeurig tot ~1 m)."""
    dx = (x - 155000.0) * 1e-5
    dy = (y - 463000.0) * 1e-5
    som_n = (3235.65389 * dy - 32.58297 * dx ** 2 - 0.24750 * dy ** 2 - 0.84978 * dx ** 2 * dy
             - 0.06550 * dy ** 3 - 0.01709 * dx ** 2 * dy ** 2 - 0.00738 * dx + 0.00530 * dx ** 4
             - 0.00039 * dx ** 2 * dy ** 3 + 0.00033 * dx ** 4 * dy - 0.00012 * dx * dy)
    som_e = (5260.52916 * dx + 105.94684 * dx * dy + 2.45656 * dx * dy ** 2 - 0.81885 * dx ** 3
             + 0.05594 * dx * dy ** 3 - 0.05607 * dx ** 3 * dy + 0.01199 * dy - 0.00256 * dx ** 3 * dy ** 2
             + 0.00128 * dx * dy ** 4 + 0.00022 * dy ** 2 - 0.00022 * dx ** 2 + 0.00026 * dx ** 5)
    return 52.15517440 + som_n / 3600.0, 5.38720621 + som_e / 3600.0


def parse_point(geom):
    if geom is None:
        return None
    coords = None
    if isinstance(geom, dict):
        coords = geom.get("coordinates")
    elif isinstance(geom, str):
        m = re.search(r"POINT\s*\(\s*([-\d.]+)\s+([-\d.]+)", geom, re.I)
        if m:
            coords = [float(m.group(1)), float(m.group(2))]
    elif isinstance(geom, (list, tuple)) and len(geom) >= 2:
        coords = geom
    if not coords or len(coords) < 2:
        return None
    a, b = float(coords[0]), float(coords[1])
    if 0 < a < 300000 and 300000 < b < 650000:  # RD
        return rd_to_wgs84(a, b)
    if 3 < a < 8 and 50 < b < 54:  # lon, lat
        return b, a
    if 50 < a < 54 and 3 < b < 8:  # lat, lon
        return a, b
    return None


def _has_letters(s):
    return isinstance(s, str) and re.search(r"[A-Za-z]", s) is not None


def find_address(rec):
    cands = []
    for k, v in rec.items():
        kl = k.lower()
        if "adres" in kl or "nummeraanduiding" in kl:
            if isinstance(v, str):
                cands.append((0 if ("nabij" in kl or "adres" in kl) else 1, v))
            elif isinstance(v, dict):
                for kk in ("omschrijving", "title", "naam", "adres"):
                    if _has_letters(v.get(kk)):
                        cands.append((1, v.get(kk)))
    links = rec.get("_links") or {}
    for k, v in links.items():
        kl = k.lower()
        if ("adres" in kl or "nummeraanduiding" in kl) and isinstance(v, dict) and _has_letters(v.get("title")):
            cands.append((2, v["title"]))
    cands = [c for c in cands if _has_letters(c[1]) and c[1].strip() not in ("-", "")]
    if not cands:
        return ""
    cands.sort(key=lambda c: c[0])
    return re.sub(r"\s+", " ", cands[0][1]).strip()


def find_relation(rec, name):
    """Geeft (identificatie, leesbare code) van een relatie zoals ligtInBouwblok."""
    ident, title = None, None
    v = rec.get(name)
    if isinstance(v, dict):
        ident = v.get("identificatie") or v.get("id")
        title = v.get("title") or v.get("code") or v.get("naam")
    elif isinstance(v, (str, int)):
        ident = v
    if ident is None:
        ident = rec.get(name + "Id")
    lk = (rec.get("_links") or {}).get(name)
    if isinstance(lk, dict):
        ident = ident or lk.get("identificatie")
        title = title or lk.get("title")
    return (str(ident) if ident not in (None, "") else None), (str(title) if title not in (None, "") else None)


def street_of(address):
    if not address:
        return ""
    m = re.match(r"^\s*(.+?)\s+\d", address)
    return (m.group(1) if m else address).strip()


def year_frac(iso):
    try:
        d = date.fromisoformat(str(iso)[:10])
    except ValueError:
        return None
    return d.year + (d.timetuple().tm_yday - 1) / 365.25


def series_for(metingen):
    """Sorteert metingen, neemt het laatste bruikbare meetsegment (na een nieuwe nulmeting)."""
    pts = []
    for m in metingen:
        if m.get("p") is False:
            continue
        t = year_frac(m.get("d"))
        z = m.get("z")
        if t is None or z is None:
            continue
        try:
            z = float(z)
        except (TypeError, ValueError):
            continue
        pts.append((t, z, (m.get("t") or "").upper(), str(m.get("d"))[:10]))
    pts.sort(key=lambda p: p[0])
    # dubbele datums: laatste telt
    dedup = {}
    for p in pts:
        dedup[round(p[0], 4)] = p
    pts = [dedup[k] for k in sorted(dedup)]
    # splits bij een nieuwe nulmeting ('N') na de eerste meting
    segs, cur = [], []
    for p in pts:
        if p[2] == "N" and cur:
            segs.append(cur)
            cur = []
        cur.append(p)
    if cur:
        segs.append(cur)
    for seg in reversed(segs):
        if len(seg) >= 2:
            return seg
    return segs[-1] if segs else []


def analyse(pts, min_span):
    n = len(pts)
    res = {"n": n, "recent": None, "rfrom": None, "early": None, "efrom": None, "total": None,
           "cum": None, "first": None, "last": None, "lastY": None, "firstY": None}
    if n == 0:
        return res
    res["first"], res["last"] = pts[0][3], pts[-1][3]
    res["firstY"], res["lastY"] = pts[0][0], pts[-1][0]
    res["cum"] = pts[-1][1] - pts[0][1]
    if n < 2:
        return res
    t1, c1 = pts[-1][0], pts[-1][1]
    t0, c0 = pts[0][0], pts[0][1]
    if t1 - t0 >= 1:
        res["total"] = (c1 - c0) / (t1 - t0)
    cand = [p for p in pts[:-1] if t1 - p[0] >= min_span]
    if cand:
        p = cand[-1]  # meest recente meting die minstens min_span jaar terug ligt
        res["recent"] = (c1 - p[1]) / (t1 - p[0])
        res["rfrom"] = p[0]
        q = pts[0]
        if p[0] - q[0] >= min_span:
            res["early"] = (p[1] - q[1]) / (p[0] - q[0])
            res["efrom"] = q[0]
    return res


def classify(r, now_y, red, yellow, stab_early, stab_now, stale_years, status=""):
    if r["recent"] is None:
        return "geen recente data"
    if (now_y - r["lastY"]) > stale_years or "verval" in (status or "").lower():
        return "geen recente data"
    if r["early"] is not None and r["early"] >= stab_early and r["recent"] < stab_now:
        return "gestabiliseerd"
    if r["recent"] >= red:
        return "zakt snel"
    if r["recent"] >= yellow:
        return "zakt normaal"
    return "zakt nauwelijks"


def order_segments(idxs, latlon, gap_m=80.0):
    """Zet de meetbouten van een straat(deel) op volgorde langs de gevel en knipt bij gaten."""
    if len(idxs) < 2:
        return [idxs]
    lat0 = sum(latlon[i][0] for i in idxs) / len(idxs)
    lon0 = sum(latlon[i][1] for i in idxs) / len(idxs)
    kx = 111320.0 * math.cos(math.radians(lat0))
    ky = 110540.0
    pts = {i: ((latlon[i][1] - lon0) * kx, (latlon[i][0] - lat0) * ky) for i in idxs}
    sxx = sum(p[0] ** 2 for p in pts.values())
    syy = sum(p[1] ** 2 for p in pts.values())
    sxy = sum(p[0] * p[1] for p in pts.values())
    ang = 0.5 * math.atan2(2 * sxy, sxx - syy)
    ux, uy = math.cos(ang), math.sin(ang)
    ordered = sorted(idxs, key=lambda i: pts[i][0] * ux + pts[i][1] * uy)
    segs, cur = [], [ordered[0]]
    for a, b in zip(ordered, ordered[1:]):
        d = math.hypot(pts[a][0] - pts[b][0], pts[a][1] - pts[b][1])
        if d > gap_m:
            segs.append(cur)
            cur = []
        cur.append(b)
    segs.append(cur)
    return segs


# --------------------------------------------------------------------------
# Uitvoer
# --------------------------------------------------------------------------

def r2(v, d=2):
    return None if v is None else round(v, d)


def write_csv(path, rows, args, now_y):
    def num(v, d=2):
        return "" if v is None else f"{v:.{d}f}".replace(".", ",")
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(["meetbout", "adres", "straat", "bouwblok", "status", "lat", "lon",
                    "aantal_metingen", "eerste_meting", "laatste_meting",
                    "recent_mm_per_jaar", "recent_vanaf_jaar", "eerder_mm_per_jaar", "eerder_vanaf_jaar",
                    "gemiddeld_mm_per_jaar", "totaal_gezakt_mm", "klasse"])
        for b in rows:
            r = b["r"]
            w.writerow([b["id"], b["addr"], b["street"], b["blk_code"] or "", b["status"],
                        f"{b['lat']:.6f}", f"{b['lon']:.6f}", r["n"], r["first"] or "", r["last"] or "",
                        num(r["recent"]), "" if r["rfrom"] is None else int(r["rfrom"]),
                        num(r["early"]), "" if r["efrom"] is None else int(r["efrom"]),
                        num(r["total"]), num(r["cum"], 1),
                        classify(r, now_y, args.rood, args.geel, 1.5, 0.5, args.max_ouderdom, b["status"])])


def short_hash(text):
    return hashlib.sha256(text.encode("utf-8")).hexdigest()[:10]


def write_site(out_dir, cache_dir, Bdata, G, meta, rows, args, now_y):
    """Website-modus: een lichte index.html die data.json ophaalt; de meetreeksen voor de
    grafiekjes staan apart in reeksen.json en worden na het tekenen van de kaart geladen.
    Bestandsnamen krijgen een ?v=<hash> zodat browsers na een update de nieuwe data pakken."""
    try:
        vendor_leaflet(out_dir, cache_dir)
    except RuntimeError as e:
        log(f"\nFOUT: {e}")
        sys.exit(1)

    meta = dict(meta)
    meta["site"] = True
    core = [b[:17] + [None] for b in Bdata]
    if not args.zonder_grafieken:
        reeksen = json.dumps([b[17] for b in Bdata], ensure_ascii=False, separators=(",", ":"))
        meta["seriesUrl"] = "reeksen.json?v=" + short_hash(reeksen)
        write_text(os.path.join(out_dir, "reeksen.json"), reeksen)
    elif os.path.exists(os.path.join(out_dir, "reeksen.json")):
        os.remove(os.path.join(out_dir, "reeksen.json"))
    data = json.dumps({"b": core, "g": G, "meta": meta}, ensure_ascii=False, separators=(",", ":"))
    write_text(os.path.join(out_dir, "data.json"), data)

    boot = SITE_BOOT.replace("__DATAURL__", "data.json?v=" + short_hash(data))

    # zoekmachines: canonical adres, Open Graph, robots.txt en sitemap.xml
    site_url = args.site_url.strip()
    if not site_url.endswith("/"):
        site_url += "/"
    esc = lambda s: (s.replace("&", "&amp;").replace('"', "&quot;").replace("<", "&lt;"))
    descr = ("Onafhankelijke kaart van gevelzakking in Amsterdam, op basis van de openbare meetbouten van de "
             "gemeente. Zoek je straat of adres en zie hoe snel gevels de laatste jaren zakken.")
    head = [
        f'<link rel="canonical" href="{esc(site_url)}">',
        '<meta property="og:type" content="website">',
        '<meta property="og:locale" content="nl_NL">',
        '<meta property="og:title" content="Zakkingskaart Amsterdam">',
        f'<meta property="og:description" content="{esc(descr)}">',
        f'<meta property="og:url" content="{esc(site_url)}">',
    ]
    gv = re.sub(r"[^A-Za-z0-9_\-]", "", args.google_verificatie.strip())
    if gv:
        head.append(f'<meta name="google-site-verification" content="{gv}">')
    write_text(os.path.join(out_dir, "robots.txt"), f"User-agent: *\nAllow: /\nSitemap: {site_url}sitemap.xml\n")
    write_text(os.path.join(out_dir, "sitemap.xml"),
               '<?xml version="1.0" encoding="UTF-8"?>\n'
               '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n'
               f'  <url><loc>{esc(site_url)}</loc><lastmod>{meta["generated"]}</lastmod>'
               '<changefreq>weekly</changefreq></url>\n</urlset>\n')

    page = (HTML_TEMPLATE
            .replace("__HEAD_EXTRA__", "\n".join(head))
            .replace("__LEAFLET_CSS__", "leaflet/leaflet.css?v=" + LEAFLET_VERSION)
            .replace("__LEAFLET_JS__", "leaflet/leaflet.js?v=" + LEAFLET_VERSION)
            .replace("__BOOT__", boot))
    html_path = os.path.join(out_dir, "index.html")
    write_text(html_path, page)
    write_text(os.path.join(out_dir, ".nojekyll"), "")
    write_csv(os.path.join(out_dir, "meetbouten_analyse.csv"), rows, args, now_y)
    return html_path


# Opstartcode voor de website: haalt data.json op en start dan de kaart.
SITE_BOOT = r"""(function(){
var el=document.getElementById("loading");el.hidden=false;
fetch("__DATAURL__").then(function(r){if(!r.ok)throw new Error("HTTP "+r.status);return r.json();})
.then(function(D){el.hidden=true;zkStart(D);})
.catch(function(e){el.className="err";el.textContent="De kaartgegevens konden niet worden geladen ("+(e&&e.message||e)+"). Ververs de pagina of probeer het later nog eens.";el.hidden=false;});
})();"""


def main():
    global API
    ap = argparse.ArgumentParser(description="Maak een zakkingskaart van alle meetbouten in Amsterdam.")
    ap.add_argument("--site", action="store_true",
                    help="website maken: index.html + data.json + reeksen.json in de map 'site', "
                         "met een blijvende cache in de map 'cache' en alleen nieuwe metingen ophalen")
    ap.add_argument("--uitvoer", default=None,
                    help="map voor de uitvoer (standaard: zakkingskaart_uitvoer, of 'site' met --site)")
    ap.add_argument("--cache", default=None,
                    help="map voor de cache (standaard: <uitvoer>/cache, of 'cache' met --site; "
                         "bij --site nooit binnen de sitemap)")
    ap.add_argument("--terugkijk", type=int, default=730,
                    help="--site: nieuwe metingen ophalen vanaf de laatst bekende datum min dit aantal dagen, "
                         "zodat laat ingevoerde metingen ook meekomen (standaard 730)")
    ap.add_argument("--volledig-na", type=int, default=90,
                    help="--site: alle metingen opnieuw ophalen als de laatste volledige download "
                         "zoveel dagen oud is; vangt correcties op (standaard 90)")
    ap.add_argument("--forceer", action="store_true",
                    help="--site: ook publiceren als er veel minder meetbouten of metingen zijn dan de vorige keer")
    ap.add_argument("--site-url", default=os.environ.get("ZK_SITE_URL") or "https://zakkingskaart.github.io/",
                    help="--site: het publieke adres van de site, voor zoekmachines (canonical, sitemap). "
                         "Standaard https://zakkingskaart.github.io/ (ook via ZK_SITE_URL)")
    ap.add_argument("--google-verificatie", default=os.environ.get("ZK_GOOGLE_VERIFICATIE", ""),
                    help="--site: code van Google Search Console (meta-tag google-site-verification), "
                         "ook via ZK_GOOGLE_VERIFICATIE")
    ap.add_argument("--api", default=API, help=argparse.SUPPRESS)  # alleen voor testen
    ap.add_argument("--vernieuw", action="store_true", help="data opnieuw ophalen en de cache negeren")
    ap.add_argument("--proef", action="store_true", help="snelle proefrun: alleen de eerste paar pagina's ophalen")
    ap.add_argument("--paginagrootte", type=int, default=2000, help="records per verzoek (standaard 2000)")
    ap.add_argument("--pauze", type=float, default=0.2, help="pauze tussen verzoeken in seconden (standaard 0,2)")
    ap.add_argument("--min-periode", type=float, default=4.0,
                    help="minimale meetperiode in jaren voor een snelheid (standaard 4)")
    ap.add_argument("--rood", type=float, default=2.0, help="standaard grens 'zakt snel' in mm/jaar (op de kaart aanpasbaar)")
    ap.add_argument("--geel", type=float, default=1.0, help="standaard grens 'zakt normaal' in mm/jaar (op de kaart aanpasbaar)")
    ap.add_argument("--max-ouderdom", type=float, default=10.0,
                    help="laatste meting ouder dan dit aantal jaren telt als 'geen recente data' (standaard 10)")
    ap.add_argument("--bbox", default=None,
                    help="alleen dit gebied op de kaart: zuid,west,noord,oost in graden, bv. 52.33,4.85,52.38,4.96")
    ap.add_argument("--zonder-grafieken", action="store_true", help="geen meetreeksen in de kaart (kleiner bestand)")
    ap.add_argument("--niet-openen", action="store_true", help="kaart niet automatisch in de browser openen")
    ap.add_argument("--zonder-teller", action="store_true",
                    help="--site: geen anonieme bezoekersteller (standaard telt de site anoniem via de feedback-ontvanger)")
    ap.add_argument("--feedback-url", default=os.environ.get("ZK_FEEDBACK_URL", ""),
                    help="webadres van de feedback-ontvanger (Google Apps Script, eindigt op /exec; zie "
                         "feedback/Code.gs). Zonder adres geen feedbackknop. Ook via omgevingsvariabele ZK_FEEDBACK_URL")
    args = ap.parse_args()

    API =args.api.rstrip("/") + "/"
    out_dir = os.path.abspath(args.uitvoer or ("site" if args.site else "zakkingskaart_uitvoer"))
    if args.cache:
        cache_dir = os.path.abspath(args.cache)
    elif args.site:
        cache_dir = os.path.abspath("cache_proef" if args.proef else "cache")
    else:
        cache_dir = os.path.join(out_dir, "cache_proef" if args.proef else "cache")
    if args.site and (cache_dir == out_dir or cache_dir.startswith(out_dir + os.sep)):
        log("FOUT: bij --site mag de cache niet in de sitemap staan (die wordt gepubliceerd). Kies een andere --cache.")
        sys.exit(2)
    os.makedirs(out_dir, exist_ok=True)
    os.makedirs(cache_dir, exist_ok=True)
    max_pages = 3 if args.proef else 0
    state = load_state(cache_dir) if args.site else {}

    log("Zakkingskaart Amsterdam" + (" (website)" if args.site else ""))
    log("=======================")
    log(f"Uitvoer: {out_dir}")
    if args.site:
        log(f"Cache:   {cache_dir}")
    if args.proef:
        log("PROEFRUN: alleen de eerste pagina's; veel bouten zullen grijs zijn.")
    log("")

    try:
        if args.site:
            # meetbouten: klein genoeg om elke keer volledig te verversen (status en adres kunnen wijzigen)
            log("-> meetbouten: verversen")
            bouten_raw = refresh_collection("meetbouten", {}, cache_dir, args, max_pages)
            metingen = update_metingen(cache_dir, args, max_pages, state)
        else:
            bouten_raw = download("meetbouten", {}, cache_dir, args.vernieuw, args.paginagrootte, args.pauze, max_pages)
            metingen = download("metingen", {"_fields": METING_FIELDS}, cache_dir, args.vernieuw, args.paginagrootte,
                                args.pauze, max_pages, slim=slim_meting, key="i")
    except KeyboardInterrupt:
        log("\nGestopt. Draai het script opnieuw om de download te hervatten.")
        sys.exit(1)
    except FatalHTTP as e:
        log(f"\nFOUT: {e}\n{e.body}")
        sys.exit(1)
    except RuntimeError as e:
        log(f"\nFOUT: {e}")
        sys.exit(1)

    log("")
    log("Analyseren ...")
    by_bout = {}
    for m in metingen:
        if m.get("b"):
            by_bout.setdefault(m["b"], []).append(m)

    bbox = None
    if args.bbox:
        try:
            s, w, n, e = [float(v) for v in args.bbox.split(",")]
            bbox = (s, w, n, e)
        except ValueError:
            log("Let op: --bbox niet begrepen, hele stad wordt getoond.")

    today = date.today()
    now_y = year_frac(today.isoformat())
    rows = []
    no_geom = 0
    for rec in bouten_raw:
        bid = rec.get("identificatie")
        if bid is None:
            continue
        bid = str(bid)
        ll = parse_point(rec.get("geometrie") or rec.get("geometry"))
        if not ll or not all(math.isfinite(v) for v in ll):
            no_geom += 1
            continue
        if bbox and not (bbox[0] <= ll[0] <= bbox[2] and bbox[1] <= ll[1] <= bbox[3]):
            continue
        addr = find_address(rec)
        blk_id, blk_code = find_relation(rec, "ligtInBouwblok")
        status = rec.get("statusOmschrijving") or ""
        if not status and rec.get("statusCode") is not None:
            status = {1: "Actueel", 2: "Niet te meten", 3: "Vervallen"}.get(rec.get("statusCode"), str(rec.get("statusCode")))
        pts = series_for(by_bout.get(bid, []))
        r = analyse(pts, args.min_periode)
        rows.append({"id": bid, "lat": ll[0], "lon": ll[1], "addr": addr, "street": street_of(addr),
                     "blk_id": blk_id, "blk_code": blk_code or blk_id, "status": status, "r": r, "pts": pts})

    if not rows:
        log("Geen meetbouten gevonden om te tonen. Controleer de --bbox of draai met --vernieuw.")
        sys.exit(1)
    with_addr = sum(1 for b in rows if b["street"])
    if with_addr < 0.5 * len(rows):
        log(f"Let op: maar {fmt_int(with_addr)} van de {fmt_int(len(rows))} meetbouten hebben een adres;")
        log("straatkleuren zijn dan beperkt. De losse meetbouten worden wel allemaal getoond.")

    # straatdelen: per straat en bouwblok, op volgorde langs de gevel
    latlon = [(b["lat"], b["lon"]) for b in rows]
    groups = {}
    for i, b in enumerate(rows):
        if not b["street"]:
            continue
        groups.setdefault((b["street"], b["blk_id"] or b["blk_code"] or ""), []).append(i)
    G = []
    for (street, blk), idxs in sorted(groups.items()):
        segs = order_segments(idxs, latlon)
        G.append([street, rows[idxs[0]]["blk_code"] or "", segs])

    # compacte data voor de kaart
    Bdata = []
    for b in rows:
        r = b["r"]
        ser = None
        if not args.zonder_grafieken and len(b["pts"]) >= 2:
            ser = []
            for p in b["pts"]:
                ser.append(round(p[0] - 1900, 2))
                ser.append(round(p[1], 1))
        Bdata.append([b["id"], round(b["lat"], 6), round(b["lon"], 6), b["addr"], b["street"], b["blk_code"] or "",
                      b["status"], r["n"], r["first"] or "", r["last"] or "", r2(r["lastY"], 3),
                      r2(r["recent"]), r2(r["rfrom"], 2), r2(r["early"]), r2(r["efrom"], 2), r2(r["total"]),
                      r2(r["cum"], 1), ser])

    meta = {
        "generated": today.isoformat(),
        "nowYear": round(now_y, 3),
        "api": API,
        "count": len(rows),
        "withSeries": sum(1 for b in rows if b["r"]["recent"] is not None),
        "defaults": {"red": args.rood, "yellow": args.geel, "stabEarly": 1.5, "stabNow": 0.5,
                     "staleYears": args.max_ouderdom},
        "proef": bool(args.proef),
        "minSpan": args.min_periode,
        "feedbackUrl": args.feedback_url.strip(),
        # anonieme teller: zelfde Apps Script; alleen in website-modus
        "statsUrl": "" if (args.zonder_teller or not args.site) else args.feedback_url.strip(),
    }
    if args.site:
        # Veiligheidscheck: niet publiceren als de gemeente (tijdelijk) veel minder levert.
        prev = state.get("vorigePublicatie") or {}
        problems = []
        if prev.get("meetbouten") and len(rows) < 0.9 * prev["meetbouten"]:
            problems.append(f"meetbouten: {fmt_int(len(rows))}, vorige keer {fmt_int(prev['meetbouten'])}")
        if prev.get("metingen") and len(metingen) < 0.95 * prev["metingen"]:
            problems.append(f"metingen: {fmt_int(len(metingen))}, vorige keer {fmt_int(prev['metingen'])}")
        if problems and not args.forceer and not args.proef:
            log("\nGESTOPT: veel minder data dan bij de vorige publicatie (" + "; ".join(problems) + ").")
            log("De bestaande site blijft staan. Klopt het toch, draai dan opnieuw met --forceer.")
            save_state(cache_dir, state)
            sys.exit(3)
        html_path = write_site(out_dir, cache_dir, Bdata, G, meta, rows, args, now_y)
        csv_path = os.path.join(out_dir, "meetbouten_analyse.csv")
        if not args.proef:
            state["vorigePublicatie"] = {"datum": today.isoformat(), "meetbouten": len(rows),
                                         "metingen": len(metingen)}
        save_state(cache_dir, state)
    else:
        payload = json.dumps({"b": Bdata, "g": G, "meta": meta}, ensure_ascii=False, separators=(",", ":"))
        payload = payload.replace("</", "<\\/")
        html_path = os.path.join(out_dir, "zakkingskaart.html")
        page = (HTML_TEMPLATE
                .replace("__HEAD_EXTRA__", "")
                .replace("__LEAFLET_CSS__", "https://unpkg.com/leaflet@%s/dist/leaflet.css" % LEAFLET_VERSION)
                .replace("__LEAFLET_JS__", "https://unpkg.com/leaflet@%s/dist/leaflet.js" % LEAFLET_VERSION)
                .replace("__BOOT__", "zkStart(__DATA__);")
                .replace("__DATA__", payload))
        with open(html_path, "w", encoding="utf-8") as fh:
            fh.write(page)
        csv_path = os.path.join(out_dir, "meetbouten_analyse.csv")
        write_csv(csv_path, rows, args, now_y)

    counts = {}
    for b in rows:
        k = classify(b["r"], now_y, args.rood, args.geel, 1.5, 0.5, args.max_ouderdom, b["status"])
        counts[k] = counts.get(k, 0) + 1
    log("")
    log(f"Meetbouten op de kaart: {len(rows):,}".replace(",", "."))
    if no_geom:
        log(f"(zonder locatie overgeslagen: {no_geom})")
    for k in ("zakt snel", "zakt normaal", "zakt nauwelijks", "gestabiliseerd", "geen recente data"):
        log(f"  {k:<20} {counts.get(k, 0):>7,}".replace(",", "."))
    log("")
    log(f"Controle: {fmt_int(sum(1 for b in rows if b['street']))} met adres, "
        f"{fmt_int(sum(1 for b in rows if b['r']['n'] > 0))} met metingen, "
        f"{fmt_int(len(G))} straatdelen. Voorbeelden:")
    for b in rows[:3]:
        rr = b["r"]
        rec_txt = "-" if rr["recent"] is None else ("%.2f" % rr["recent"]).replace(".", ",")
        log(f"  {b['id']} | {b['addr'] or '(geen adres)'} | blok {b['blk_code'] or '-'} | "
            f"{b['lat']:.5f}, {b['lon']:.5f} | {rr['n']} metingen | recent {rec_txt} mm/j")
    log("")
    if args.site:
        def mb(fn):
            return os.path.getsize(os.path.join(out_dir, fn)) / 1e6
        log(f"Website: {out_dir}")
        log(f"  index.html {mb('index.html'):.2f} MB, data.json {mb('data.json'):.1f} MB"
            + (f", reeksen.json {mb('reeksen.json'):.1f} MB" if os.path.exists(os.path.join(out_dir, 'reeksen.json')) else ""))
        log("Lokaal bekijken (de pagina laadt data.json, dus dubbelklikken werkt niet):")
        log(f"  python3 -m http.server 8000 --directory \"{out_dir}\"   en open http://localhost:8000")
        return
    size_mb = os.path.getsize(html_path) / 1e6
    log(f"Kaart:  {html_path}  ({size_mb:.1f} MB)")
    log(f"Cijfers: {csv_path}")
    if not args.niet_openen:
        try:
            webbrowser.open("file://" + urllib.request.pathname2url(html_path))
        except Exception:
            pass


HTML_TEMPLATE = r"""<!doctype html>
<html lang="nl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Zakkingskaart Amsterdam: hoe snel zakken de gevels in jouw straat?</title>
<meta name="description" content="Onafhankelijke kaart van gevelzakking in Amsterdam, op basis van de openbare meetbouten van de gemeente. Zoek je straat of adres en zie hoe snel gevels de laatste jaren zakken.">
__HEAD_EXTRA__
<link rel="stylesheet" href="__LEAFLET_CSS__">
<style>
:root{--ink:#1c1c1a;--muted:#6b6a64;--panel:rgba(255,255,255,.96);--line:#e3e1d9;
  --rood:#d7301f;--geel:#e6a700;--groen:#1a9850;--stab:#1a9850;--oud:#9a9a9a;--ring:#08306b;--accent:#24507a}
*{box-sizing:border-box}
html,body{height:100%;margin:0;font:13px/1.45 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;color:var(--ink);background:#eceae4}
#map{position:absolute;inset:0}
.panel{position:absolute;z-index:1000;background:var(--panel);border:1px solid var(--line);border-radius:10px;box-shadow:0 6px 20px rgba(0,0,0,.12)}
#search{top:12px;left:58px;padding:7px;display:flex;gap:6px;width:min(380px,calc(100vw - 76px))}
#search input{flex:1;min-width:0;padding:7px 10px;border:1px solid var(--line);border-radius:7px;font:inherit;background:#fff;color:var(--ink)}
#search button,.btn{padding:7px 12px;border:0;border-radius:7px;background:var(--accent);color:#fff;font:inherit;font-weight:600;cursor:pointer}
#search button:focus-visible,.btn:focus-visible,#legend button:focus-visible{outline:2px solid var(--ink);outline-offset:2px}
#msg{position:absolute;top:58px;left:58px;z-index:1000;background:#fff3cd;border:1px solid #f0d58a;padding:6px 10px;border-radius:7px}
#loading{position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);z-index:2500;background:#fff;border:1px solid var(--line);border-radius:10px;padding:12px 16px;box-shadow:0 6px 20px rgba(0,0,0,.12);font-weight:600;max-width:calc(100vw - 32px);text-align:center}
#loading.err{background:#fde8e6;color:#7a1c12;border-color:#e9a79f;font-weight:500}
#loading[hidden]{display:none}
#legend{top:12px;right:12px;width:300px;max-height:calc(100% - 90px);overflow:auto;padding:12px 14px}
#legend .head{display:flex;align-items:center;justify-content:space-between;gap:8px}
#legend h1{font-size:15px;margin:0}
#legend .toggle{border:1px solid var(--line);background:#fff;border-radius:6px;width:28px;height:26px;cursor:pointer;font:inherit}
.rows{display:grid;gap:6px;margin:10px 0 6px}
.row{display:grid;grid-template-columns:18px 22px 1fr auto;align-items:center;gap:6px;cursor:pointer}
.row input{margin:0}
.sw{display:inline-block;width:18px;height:6px;border-radius:3px}
.sw.dot{width:11px;height:11px;border-radius:50%}
.sw.stab{background:var(--stab);box-shadow:0 0 0 2px var(--ring)}
.row .lbl{line-height:1.25}
.row .lbl small{display:block;color:var(--muted)}
.row .cnt{color:var(--muted);font-variant-numeric:tabular-nums;font-size:12px}
.note{color:var(--muted);font-size:12px;margin:6px 0}
details{border-top:1px solid var(--line);margin-top:8px;padding-top:8px}
summary{cursor:pointer;font-weight:600}
.set{display:grid;grid-template-columns:1fr 74px;gap:6px 8px;align-items:center;margin-top:8px}
.set input[type=number]{width:74px;padding:4px 6px;border:1px solid var(--line);border-radius:6px;font:inherit}
.chk{display:flex;gap:6px;align-items:center;margin-top:6px}
.src{font-size:11px;color:var(--muted);margin-top:10px;border-top:1px solid var(--line);padding-top:8px}
.pop{min-width:240px}
.pop .ttl{font-weight:700;font-size:14px}
.pop .sub{color:var(--muted);font-size:12px;margin-bottom:6px}
.pill{display:inline-flex;align-items:center;gap:6px;font-weight:600;font-size:12px;margin:2px 0 6px}
.pill i{width:10px;height:10px;border-radius:50%;display:inline-block}
table.kv{border-collapse:collapse;font-size:12px;font-variant-numeric:tabular-nums}
table.kv th{text-align:left;font-weight:500;color:var(--muted);padding:1px 10px 1px 0;white-space:nowrap}
table.kv td{padding:1px 0}
.m{color:var(--muted)}
.list{max-height:220px;overflow:auto;margin-top:6px;border-top:1px solid var(--line)}
.list table{border-collapse:collapse;width:100%;font-size:12px;font-variant-numeric:tabular-nums}
.list td{padding:3px 4px;border-bottom:1px solid #f0eee8}
.list tr{cursor:pointer}
.list tr:hover{background:#f4f2ec}
.list td.n{text-align:right;white-space:nowrap}
.list .sw{width:9px;height:9px;border-radius:50%}
.chart{display:block;margin-top:6px}
a.raw{display:inline-block;margin-top:4px;font-size:12px;color:var(--accent)}
.leaflet-popup-content{margin:12px 14px}
@media (max-width:640px){
  #legend{width:calc(100vw - 24px);top:auto;bottom:24px;max-height:55%}
  #legend.collapsed .body{display:none}
}
#legend.collapsed .body{display:none}
#modal,#fbModal,#aboutModal{position:fixed;inset:0;z-index:4000;background:rgba(20,20,18,.5);display:flex;align-items:center;justify-content:center;padding:16px}
#modal[hidden],#fbModal[hidden],#aboutModal[hidden]{display:none}
#legend .indep{margin:4px 0 0;font-size:12px;color:var(--muted)}
#legend .indep a,#aboutModal a{color:var(--accent)}
#aboutModal .mbox{max-width:640px}
#aBody{overflow:auto;padding:4px 18px 18px;font-size:13.5px;line-height:1.55}
#aBody h3{font-size:14px;margin:16px 0 4px}
#aBody p{margin:0 0 8px}
#aClose{border:1px solid var(--line);background:#fff;border-radius:8px;width:34px;height:34px;font-size:20px;line-height:1;cursor:pointer;flex:none}
#fbBtn{border:1px solid var(--line);background:#fff;border-radius:6px;height:26px;padding:0 9px;cursor:pointer;font:inherit;font-weight:600;color:var(--accent);white-space:nowrap}
#fbBtn:focus-visible,#fbModal button:focus-visible{outline:2px solid var(--ink);outline-offset:2px}
#legend .head .btns{display:flex;gap:6px;align-items:center}
#fbModal .mbox{max-width:460px}
.stars{display:flex;gap:4px;margin:4px 0 12px;border:0;padding:0}
.stars legend{font-weight:600;margin-bottom:6px;padding:0}
.stars input{position:absolute;opacity:0;width:1px;height:1px}
.stars label{font-size:30px;line-height:1;color:#cfccc2;cursor:pointer;padding:0 2px}
.stars label.on{color:#e6a700}
.stars input:focus-visible+label{outline:2px solid var(--ink);border-radius:4px}
.fbf label.t{display:block;font-weight:600;margin:10px 0 4px}
.fbf textarea{width:100%;padding:8px 10px;border:1px solid var(--line);border-radius:8px;font:inherit}
.fbf textarea{min-height:64px;resize:vertical}
.fbf .hp{position:absolute;left:-9999px;width:1px;height:1px;overflow:hidden}
.fbf .row2{display:flex;justify-content:space-between;align-items:center;gap:10px;margin-top:12px}
.fbf .cnt2{color:var(--muted);font-size:12px}
#fbStatus{margin-top:10px;font-size:13px}
#fbStatus.ok{color:#11733a} #fbStatus.err{color:#a3261a}
.mbox{background:#fff;border-radius:12px;width:100%;max-width:980px;max-height:92vh;display:flex;flex-direction:column;box-shadow:0 20px 60px rgba(0,0,0,.3)}
.mhead{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;padding:16px 18px 10px;border-bottom:1px solid var(--line)}
.mhead h2{margin:0;font-size:17px}
#mClose{border:1px solid var(--line);background:#fff;border-radius:8px;width:34px;height:34px;font-size:20px;line-height:1;cursor:pointer;flex:none}
#mBody{overflow:auto;padding:14px 18px 18px}
#mBody h3{font-size:14px;margin:18px 0 6px}
.mt-wrap{overflow-x:auto;border:1px solid var(--line);border-radius:8px}
table.mt{border-collapse:collapse;width:100%;font-size:12.5px;font-variant-numeric:tabular-nums}
table.mt th{position:sticky;top:0;background:#f6f5f1;text-align:left;font-weight:600;padding:7px 9px;border-bottom:1px solid var(--line);white-space:nowrap}
table.mt td{padding:6px 9px;border-bottom:1px solid #f0eee8;white-space:nowrap}
table.mt td.n,table.mt th.n{text-align:right}
table.mt tr.rec td{background:#fbf6e6}
.mnote{color:var(--muted);font-size:12px;margin:6px 0 0}
.msum{display:flex;flex-wrap:wrap;gap:6px 18px;font-size:13px;margin:0 0 10px}
.msum b{font-variant-numeric:tabular-nums}
pre.rawjson{background:#f6f5f1;border:1px solid var(--line);border-radius:8px;padding:10px 12px;font:12px/1.45 ui-monospace,Menlo,Consolas,monospace;max-height:45vh;overflow:auto;white-space:pre;margin:0}
.warn{background:#fff3cd;border:1px solid #f0d58a;border-radius:8px;padding:8px 10px;font-size:12.5px;margin-bottom:10px}
</style>
</head>
<body>
<div id="map" role="application" aria-label="Kaart met meetbouten"></div>
<form id="search" class="panel" autocomplete="off">
  <input id="q" list="streets" placeholder="Zoek een straat, adres of meetbout-nummer" aria-label="Zoeken">
  <button type="submit">Zoek</button>
  <datalist id="streets"></datalist>
</form>
<div id="msg" hidden></div>
<div id="loading" role="status" aria-live="polite" hidden>Kaartgegevens laden …</div>
<div id="fbModal" hidden role="dialog" aria-modal="true" aria-labelledby="fbTitle">
  <div class="mbox"><div class="mhead"><div><h2 id="fbTitle">Wat vind je van deze kaart?</h2><div class="sub">Je feedback helpt de kaart te verbeteren.</div></div><button id="fbClose" class="xbtn" type="button" aria-label="Sluiten" style="border:1px solid var(--line);background:#fff;border-radius:8px;width:34px;height:34px;font-size:20px;cursor:pointer">&times;</button></div>
  <div id="fbBody" style="padding:14px 18px 18px">
    <form class="fbf" id="fbForm" novalidate>
      <fieldset class="stars" id="fbStars"><legend>Jouw beoordeling</legend></fieldset>
      <label class="t" for="fbMsg">Toelichting <span class="m">(optioneel)</span></label>
      <textarea id="fbMsg" rows="3" maxlength="280" placeholder="Wat vind je goed, of wat kan beter?"></textarea>
      <div class="hp" aria-hidden="true"><label for="fbHp">Laat dit veld leeg</label><input id="fbHp" type="text" tabindex="-1" autocomplete="off"></div>
      <p class="mnote">We vragen geen naam of e-mailadres. Je beoordeling en toelichting komen als e-mail binnen bij de beheerder van deze kaart; zet er dus geen persoonlijke gegevens in. De kaart gebruikt geen cookies en geen tracking.</p>
      <div class="row2"><span class="cnt2" id="fbCnt">0 / 280</span><button class="btn" id="fbSend" type="submit">Versturen</button></div>
      <div id="fbStatus" role="status" aria-live="polite"></div>
    </form>
  </div></div>
</div>
<div id="aboutModal" hidden role="dialog" aria-modal="true" aria-labelledby="aTitle">
  <div class="mbox"><div class="mhead"><div><h2 id="aTitle">Over deze kaart</h2></div><button id="aClose" type="button" aria-label="Sluiten">&times;</button></div>
  <div id="aBody">
    <h3>Onafhankelijk initiatief</h3>
    <p>Zakkingskaart is een onafhankelijk, niet-commercieel initiatief. Het is geen website van de gemeente Amsterdam en wordt niet door de gemeente gemaakt, gecontroleerd of goedgekeurd.</p>
    <h3>Wat de kaart laat zien</h3>
    <p>De kaart gebruikt de openbare meetboutgegevens van de gemeente Amsterdam en wordt wekelijks automatisch bijgewerkt. Meetbouten zijn meetpunten in gevels waarvan de gemeente periodiek de hoogte meet. Per meetbout laat de kaart zien hoe snel die de laatste jaren is gezakt (de recente snelheid), of het zakken duidelijk is afgenomen (&ldquo;gestabiliseerd&rdquo;, vaak na funderingsherstel) en of er recente metingen zijn. De kaart van de gemeente kleurt op de gemiddelde snelheid sinds de eerste meting; daardoor kunnen de kleuren verschillen.</p>
    <h3>Disclaimer</h3>
    <p>De kleuren zijn een berekende indicatie op basis van metingen aan gevels. Ze zijn geen oordeel over de fundering, de staat of de waarde van een pand, en kunnen onvolledig of verouderd zijn. Aan deze kaart kunnen geen rechten worden ontleend. Neem je een beslissing over een pand, laat dan een funderingsonderzoek doen door een deskundige en raadpleeg de officiële gegevens van de gemeente.</p>
    <h3>Privacy</h3>
    <p>Deze site gebruikt geen cookies en volgt geen bezoekers. We tellen alleen anoniem hoe vaak de kaart per dag wordt geopend, vanaf mobiel of computer, via welke soort bron (bijvoorbeeld een zoekmachine of een andere website) en hoe vaak er wordt gezocht en geklikt. Er worden daarbij geen IP-adressen of andere kenmerken van jou opgeslagen. Je kaartinstellingen worden alleen in je eigen browser bewaard. De site wordt gehost door GitHub, de achtergrondkaart komt van PDOK en &ldquo;Alle meetgegevens bekijken&rdquo; haalt gegevens op bij de gemeente; die partijen zien daarbij, zoals bij elke website, je IP-adres. Feedback komt als e-mail binnen bij de beheerder; er wordt geen naam of e-mailadres gevraagd.</p>
    <h3>Bronnen</h3>
    <p>Meetbouten: Gemeente Amsterdam (api.data.amsterdam.nl). Achtergrondkaart: PDOK / Kadaster (CC BY 4.0). Kaartbibliotheek: Leaflet (BSD-2). Broncode en rekenmethode: <a href="https://github.com/zakkingskaart/zakkingskaart.github.io" target="_blank" rel="noopener">GitHub</a>.</p>
  </div></div>
</div>
<div id="modal" hidden role="dialog" aria-modal="true" aria-labelledby="mTitle">
  <div class="mbox"><div class="mhead"><div><h2 id="mTitle"></h2><div class="sub" id="mSub"></div></div><button id="mClose" type="button" aria-label="Sluiten">&times;</button></div>
  <div id="mBody"></div></div>
</div>
<aside id="legend" class="panel">
  <div class="head"><h1>Zakkingskaart Amsterdam</h1><div class="btns"><button id="fbBtn" type="button" hidden>&#9733; Feedback</button><button class="toggle" id="toggle" aria-expanded="true" title="In- of uitklappen">–</button></div></div>
  <p class="indep">Onafhankelijk initiatief, los van de gemeente Amsterdam · <a href="#over" id="aboutLink">Over &amp; disclaimer</a></p>
  <div class="body">
    <div class="rows" id="rows"></div>
    <p class="note" id="viewnote">Aantallen: meetbouten in beeld.</p>
    <p class="note">Lijnen = straatdeel (per bouwblok, mediaan van de bouten). Klik op een lijn of punt voor details.</p>
    <details id="settings">
      <summary>Grenzen aanpassen</summary>
      <div class="set">
        <label for="sRed">Zakt snel vanaf (mm/jaar)</label><input id="sRed" type="number" step="0.1" min="0">
        <label for="sYellow">Zakt normaal vanaf (mm/jaar)</label><input id="sYellow" type="number" step="0.1" min="0">
        <label for="sStabE">Gestabiliseerd: eerder minstens (mm/jaar)</label><input id="sStabE" type="number" step="0.1" min="0">
        <label for="sStabN">... en nu minder dan (mm/jaar)</label><input id="sStabN" type="number" step="0.1" min="0">
        <label for="sStale">Laatste meting hooguit (jaar oud)</label><input id="sStale" type="number" step="1" min="1">
      </div>
      <label class="chk"><input id="cStale" type="checkbox"> Oudere metingen tóch kleuren</label>
      <label class="chk"><input id="cDots" type="checkbox"> Losse meetbouten tonen</label>
      <p class="note"><button class="btn" id="reset" type="button">Standaard herstellen</button></p>
    </details>
    <div class="src" id="src"></div>
  </div>
</aside>
<script src="__LEAFLET_JS__"></script>
<script>
window.__showErr=function(t){var d=document.getElementById("err");if(!d){d=document.createElement("div");d.id="err";
d.style.cssText="position:absolute;left:12px;right:12px;bottom:40px;z-index:3000;background:#fde8e6;color:#7a1c12;border:1px solid #e9a79f;border-radius:8px;padding:10px 12px;font:12px/1.4 ui-monospace,Menlo,Consolas,monospace;white-space:pre-wrap;max-height:40%;overflow:auto";
document.body.appendChild(d);d.textContent="Er ging iets mis bij het tekenen van de meetbouten. Stuur deze melding door:\n\n";}d.textContent+=t+"\n";};
window.addEventListener("error",function(e){window.__showErr((e.message||"fout")+" (regel "+e.lineno+":"+e.colno+")");});
// "Over & disclaimer": werkt los van de kaart, ook als de data niet laadt. Deelbaar als #over.
(function(){var m=document.getElementById("aboutModal"),prev=null;
function open(e){if(e)e.preventDefault();prev=document.activeElement;m.hidden=false;document.getElementById("aClose").focus();}
function close(){m.hidden=true;if(location.hash==="#over"&&history.replaceState)history.replaceState(null,"",location.pathname+location.search);if(prev&&prev.focus)prev.focus();}
document.getElementById("aboutLink").addEventListener("click",open);
document.getElementById("aClose").addEventListener("click",close);
m.addEventListener("click",function(e){if(e.target===m)close();});
document.addEventListener("keydown",function(e){if(e.key==="Escape"&&!m.hidden)close();});
if(location.hash==="#over")open();})();
</script>
<script>
function zkStart(D){
"use strict";
try{
if(typeof L==="undefined"){document.body.insertAdjacentHTML("beforeend",'<p style="position:absolute;top:80px;left:20px;z-index:2000;background:#fff;padding:12px;border-radius:8px">De kaartbibliotheek kon niet worden geladen. Controleer je internetverbinding en open het bestand opnieuw.</p>');return;}
const B=D.b,G=D.g,M=D.meta;
let serState="ok",curPop=null;   // meetreeksen: "ok" | "laden" | "fout"

// ---------- anonieme teller ----------
// Meldt alleen: kaart geopend, mobiel/desktop, herkomst, en hoe vaak er is gezocht en geklikt.
// Geen cookies, geen opslag in de browser, geen kenmerken per bezoeker.
const TEL=(function(){
  const url=(M.statsUrl||"").trim(),geen={tel(){}};
  if(!/^https:\/\//i.test(url))return geen;
  if(/^(localhost|127\.|0\.0\.0\.0|\[::1\])/.test(location.hostname)||location.protocol==="file:")return geen;
  if(navigator.webdriver||/bot|crawl|spider|slurp|preview|lighthouse/i.test(navigator.userAgent||""))return geen;
  const ev={zoek:0,punt:0,meet:0};
  function send(o){try{const body=JSON.stringify(Object.assign({soort:"stat"},o));
    if(navigator.sendBeacon&&navigator.sendBeacon(url,new Blob([body],{type:"text/plain;charset=utf-8"})))return;
    fetch(url,{method:"POST",mode:"no-cors",keepalive:true,headers:{"Content-Type":"text/plain;charset=utf-8"},body:body}).catch(()=>{});}catch(e){}}
  let bron="direct";
  try{const r=document.referrer?new URL(document.referrer).hostname.replace(/^www\./,"").toLowerCase():"";
    const eigen=location.hostname.replace(/^www\./,"").toLowerCase();
    if(!r)bron="direct";
    else if(r===eigen||r==="zakkingskaart.github.io"||r==="zakkingskaart.nl")bron="intern";
    else if(/(^|\.)google\./.test(r))bron="google";
    else if(/(^|\.)(bing|duckduckgo|ecosia|yahoo|startpage|qwant|search\.brave)\./.test(r))bron="zoekmachine";
    else if(/(^|\.)(facebook|instagram|twitter|x|t|linkedin|reddit|whatsapp|threads|bsky|mastodon|tiktok)\.(com|co|app|net|social|nl)$/.test(r)||/^(lm|l|m)\.facebook\.com$/.test(r))bron="social";
    else bron="site:"+r;}catch(e){}
  send({bezoek:1,apparaat:matchMedia("(max-width:640px)").matches?"mobiel":"desktop",bron:bron});
  function flush(){if(ev.zoek||ev.punt||ev.meet){send({zoek:ev.zoek,punt:ev.punt,meet:ev.meet});ev.zoek=ev.punt=ev.meet=0;}}
  document.addEventListener("visibilitychange",()=>{if(document.visibilityState==="hidden")flush();});
  window.addEventListener("pagehide",flush);
  return {tel(k){if(k in ev)ev[k]++;}};
})();
const I={id:0,lat:1,lon:2,addr:3,street:4,blk:5,status:6,n:7,first:8,last:9,lastY:10,recent:11,rfrom:12,early:13,efrom:14,total:15,cum:16,ser:17};
const KL=["rood","geel","groen","stab","oud"];
const COL={rood:"#d7301f",geel:"#e6a700",groen:"#1a9850",stab:"#1a9850",oud:"#9a9a9a"};
const RING="#08306b";
const DEF={red:M.defaults.red,yellow:M.defaults.yellow,stabEarly:M.defaults.stabEarly,stabNow:M.defaults.stabNow,
  staleYears:M.defaults.staleYears,colorStale:false,showDots:true,show:{rood:true,geel:true,groen:true,stab:true,oud:true},base:"pdokgrijs"};
let S=JSON.parse(JSON.stringify(DEF));
try{const s=JSON.parse(localStorage.getItem("zakkingskaart.v1")||"null");if(s){S=Object.assign(S,s);S.show=Object.assign({},DEF.show,s.show||{});}}catch(e){}
function save(){try{localStorage.setItem("zakkingskaart.v1",JSON.stringify(S));}catch(e){}}

const nf=(v,d=2)=>v==null?"–":Number(v).toLocaleString("nl-NL",{minimumFractionDigits:d,maximumFractionDigits:d});
const esc=s=>String(s==null?"":s).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const fmtDate=iso=>iso?iso.slice(8,10)+"-"+iso.slice(5,7)+"-"+iso.slice(0,4):"–";
function median(a){const s=a.slice().sort((x,y)=>x-y),m=s.length>>1;return s.length%2?s[m]:(s[m-1]+s[m])/2;}
function label(c){
  if(c==="rood")return["Zakt snel","vanaf "+nf(S.red,1)+" mm/jaar"];
  if(c==="geel")return["Zakt normaal",nf(S.yellow,1)+" tot "+nf(S.red,1)+" mm/jaar"];
  if(c==="groen")return["Zakt nauwelijks","minder dan "+nf(S.yellow,1)+" mm/jaar"];
  if(c==="stab")return["Gestabiliseerd","eerder ≥ "+nf(S.stabEarly,1)+", nu < "+nf(S.stabNow,1)+" mm/jaar; vaak hersteld"];
  return["Geen recente data","laatste meting > "+S.staleYears+" jaar oud, vervallen of te kort"];
}
function cls(b){
  const r=b[I.recent];
  if(r==null)return"oud";
  const stale=(M.nowYear-b[I.lastY])>S.staleYears||/verval/i.test(b[I.status]||"");
  if(stale&&!S.colorStale)return"oud";
  const e=b[I.early];
  if(e!=null&&e>=S.stabEarly&&r<S.stabNow)return"stab";
  if(r>=S.red)return"rood";
  if(r>=S.yellow)return"geel";
  return"groen";
}

// ---------- kaart ----------
const map=L.map("map",{preferCanvas:true,zoomSnap:0.5,minZoom:10}).setView([52.36,4.9],13);
const rend=L.canvas({padding:0.4,tolerance:6});
const PDOK="https://service.pdok.nl/brt/achtergrondkaart/wmts/v2_0/";
const pdokOpt={maxNativeZoom:19,maxZoom:21,attribution:"Achtergrond: <a href=\"https://www.pdok.nl\" target=\"_blank\" rel=\"noopener\">PDOK</a> / Kadaster (BRT)"};
const bases={
  pdokgrijs:L.tileLayer(PDOK+"grijs/EPSG:3857/{z}/{x}/{y}.png",pdokOpt),
  pdokpastel:L.tileLayer(PDOK+"pastel/EPSG:3857/{z}/{x}/{y}.png",pdokOpt),
  pdokstandaard:L.tileLayer(PDOK+"standaard/EPSG:3857/{z}/{x}/{y}.png",pdokOpt),
  luchtfoto:L.tileLayer("https://service.pdok.nl/hwh/luchtfotorgb/wmts/v1_0/Actueel_orthoHR/EPSG:3857/{z}/{x}/{y}.jpeg",{maxNativeZoom:19,maxZoom:21,attribution:"Luchtfoto: PDOK / Beeldmateriaal Nederland"}),
  geen:L.layerGroup()
};
(bases[S.base]||bases.pdokgrijs).addTo(map);
L.control.layers({"Grijs (PDOK)":bases.pdokgrijs,"Pastel (PDOK)":bases.pdokpastel,"Standaard (PDOK)":bases.pdokstandaard,"Luchtfoto (PDOK)":bases.luchtfoto,"Geen achtergrond":bases.geen},null,{position:"bottomleft"}).addTo(map);
map.on("baselayerchange",e=>{for(const k in bases)if(bases[k]===e.layer)S.base=k;save();});
// Laadt de achtergrond niet? Dan een melding; de meetbouten werken ook zonder achtergrond.
for(const k in bases){const t=bases[k];if(!t.on||k==="geen")continue;let ok=0,bad=0,shown=false;
  t.on("tileload",()=>ok++);
  t.on("tileerror",()=>{bad++;if(!shown&&bad>=6&&ok===0){shown=true;const el=document.getElementById("msg");
    el.textContent="De achtergrondkaart laadt niet. Kies linksonder een andere achtergrond; de meetbouten werken ook zonder.";el.hidden=false;setTimeout(()=>el.hidden=true,8000);}});}
L.control.scale({imperial:false,position:"bottomleft"}).addTo(map);
map.attributionControl.addAttribution("Meetbouten: Gemeente Amsterdam");

const lineLayer=L.layerGroup().addTo(map), dotLayer=L.layerGroup().addTo(map);
const inLine=new Uint8Array(B.length);
const lines=[];
G.forEach((g,gi)=>g[2].forEach(seg=>{
  if(seg.length<2)return;
  seg.forEach(i=>inLine[i]=1);
  const pl=L.polyline(seg.map(i=>[B[i][I.lat],B[i][I.lon]]),{renderer:rend,weight:4,opacity:.9,lineCap:"round",lineJoin:"round"});
  pl._zkG=gi;pl._zkVis=false;
  pl.on("click",e=>{L.DomEvent.stopPropagation(e);openList(groupTitle(gi),G[gi][2].flat(),e.latlng);});
  lines.push(pl);
}));
const dots=B.map((b,i)=>{
  const m=L.circleMarker([b[I.lat],b[I.lon]],{renderer:rend,radius:3,weight:1,color:"#fff",opacity:1,fillOpacity:1});
  m._zkI=i;m._zkVis=false;
  m.on("click",e=>{L.DomEvent.stopPropagation(e);openBout(i);});
  return m;
});
const bounds=L.latLngBounds(B.filter(b=>isFinite(b[I.lat])&&isFinite(b[I.lon])).map(b=>[b[I.lat],b[I.lon]]));
if(bounds.isValid())map.fitBounds(bounds,{padding:[20,20],maxZoom:15});

let C=[],GC=[];
function groupClass(gi){
  const idx=G[gi][2].flat();let vals=[],stab=0;
  for(const i of idx){const c=C[i];if(c!=="oud"){vals.push(B[i][I.recent]);if(c==="stab")stab++;}}
  if(!vals.length)return["oud",null,0];
  const med=median(vals);
  const c=stab*2>=vals.length?"stab":med>=S.red?"rood":med>=S.yellow?"geel":"groen";
  return[c,med,vals.length];
}
function sizes(){
  const z=map.getZoom();
  return{r:z>=18?8:z>=17?6.5:z>=16?5:z>=15?4:z>=14?3:z>=13?2.3:1.7,
         w:z>=18?10:z>=17?8:z>=16?6.5:z>=15?5:z>=14?4:z>=13?3:2.2,
         sw:z>=16?1.5:z>=15?1:z>=14?0.6:0};
}
function styleAll(){
  C=B.map(cls);GC=G.map((g,gi)=>groupClass(gi));
  const sz=sizes();let added=false;
  for(const pl of lines){
    const c=GC[pl._zkG][0],vis=!!S.show[c];
    pl.setStyle({color:COL[c],weight:sz.w,opacity:c==="oud"?.5:.9,dashArray:c==="stab"?(sz.w*0.3)+" "+(sz.w*1.6):null});
    if(vis&&!pl._zkVis){lineLayer.addLayer(pl);pl._zkVis=true;added=true;}
    else if(!vis&&pl._zkVis){lineLayer.removeLayer(pl);pl._zkVis=false;}
  }
  for(const m of dots){
    const i=m._zkI,c=C[i],vis=!!S.show[c]&&(S.showDots||!inLine[i]);
    const st=c==="stab";
    m.setStyle({radius:st?sz.r+0.5:sz.r,fillColor:COL[c],color:st?RING:"#ffffff",weight:st?Math.max(1.2,sz.sw+0.6):sz.sw,stroke:st||sz.sw>0,fillOpacity:c==="oud"?.75:1});
    if(vis&&!m._zkVis){dotLayer.addLayer(m);m._zkVis=true;}
    else if(!vis&&m._zkVis){dotLayer.removeLayer(m);m._zkVis=false;}
    else if(vis&&added&&m.bringToFront){m.bringToFront();}
  }
  counts();
}
function resize(){
  const sz=sizes();
  for(const pl of lines)if(pl._zkVis){const c=GC[pl._zkG][0];pl.setStyle({weight:sz.w,dashArray:c==="stab"?(sz.w*0.3)+" "+(sz.w*1.6):null});}
  for(const m of dots)if(m._zkVis){const st=C[m._zkI]==="stab";m.setStyle({radius:st?sz.r+0.5:sz.r,weight:st?Math.max(1.2,sz.sw+0.6):sz.sw,stroke:st||sz.sw>0});}
}
map.on("zoomend",resize);
map.on("moveend",counts);

// ---------- legenda ----------
const rowsEl=document.getElementById("rows");
function buildRows(){
  rowsEl.innerHTML=KL.map(k=>{const[l,s]=label(k);
    const sw=k==="stab"?'<i class="sw dot stab"></i>':'<i class="sw" style="background:'+COL[k]+'"></i>';
    return '<label class="row"><input type="checkbox" data-k="'+k+'"'+(S.show[k]?" checked":"")+'>'+sw+
      '<span class="lbl">'+l+'<small>'+esc(s)+'</small></span><span class="cnt" id="cnt-'+k+'">–</span></label>';}).join("");
  rowsEl.querySelectorAll("input").forEach(inp=>inp.addEventListener("change",()=>{S.show[inp.dataset.k]=inp.checked;save();styleAll();}));
}
function counts(){
  const bb=map.getBounds();const n={rood:0,geel:0,groen:0,stab:0,oud:0};
  for(let i=0;i<B.length;i++){const b=B[i];if(bb.contains([b[I.lat],b[I.lon]]))n[C[i]]++;}
  for(const k of KL){const el=document.getElementById("cnt-"+k);if(el)el.textContent=n[k].toLocaleString("nl-NL");}
}
const fld={sRed:"red",sYellow:"yellow",sStabE:"stabEarly",sStabN:"stabNow",sStale:"staleYears"};
function fillSettings(){
  for(const id in fld)document.getElementById(id).value=S[fld[id]];
  document.getElementById("cStale").checked=!!S.colorStale;
  document.getElementById("cDots").checked=!!S.showDots;
}
for(const id in fld)document.getElementById(id).addEventListener("change",e=>{
  const v=parseFloat(String(e.target.value).replace(",","."));
  if(!isFinite(v)||v<0){fillSettings();return;}
  S[fld[id]]=v;
  if(S.yellow>S.red){if(id==="sRed")S.yellow=S.red;else S.red=S.yellow;}
  save();fillSettings();buildRows();styleAll();
});
document.getElementById("cStale").addEventListener("change",e=>{S.colorStale=e.target.checked;save();styleAll();});
document.getElementById("cDots").addEventListener("change",e=>{S.showDots=e.target.checked;save();styleAll();});
document.getElementById("reset").addEventListener("click",()=>{const base=S.base;S=JSON.parse(JSON.stringify(DEF));S.base=base;save();fillSettings();buildRows();styleAll();});
const legend=document.getElementById("legend"),tg=document.getElementById("toggle");
tg.addEventListener("click",()=>{const c=legend.classList.toggle("collapsed");tg.textContent=c?"+":"–";tg.setAttribute("aria-expanded",String(!c));});
if(window.innerWidth<640){legend.classList.add("collapsed");tg.textContent="+";tg.setAttribute("aria-expanded","false");}
L.DomEvent.disableClickPropagation(legend);L.DomEvent.disableScrollPropagation(legend);
L.DomEvent.disableClickPropagation(document.getElementById("search"));
document.getElementById("src").innerHTML="Bron: Gemeente Amsterdam, meetbouten (api.data.amsterdam.nl), opgehaald "+fmtDate(M.generated)+". "+
  M.count.toLocaleString("nl-NL")+" meetbouten, waarvan "+M.withSeries.toLocaleString("nl-NL")+" met een meetreeks van minstens "+nf(M.minSpan,0)+" jaar. "+
  "Snelheid = verschil tussen de laatste meting en de meest recente meting die minstens "+nf(M.minSpan,0)+" jaar eerder ligt. "+
  "Indicatie op basis van gevelmetingen, geen funderingsoordeel."+(M.proef?" <b>Proefrun: onvolledige data.</b>":"");

// ---------- popups ----------
function chart(b,c){
  const s=b[I.ser];if(!s||s.length<4)return"";
  const W=290,H=130,l=38,r=10,t=10,bt=22;const ts=[],cs=[];
  for(let k=0;k<s.length;k+=2){ts.push(s[k]+1900);cs.push(s[k+1]);}
  const t0=Math.floor(ts[0]),t1=Math.ceil(ts[ts.length-1])||t0+1;
  const cmin=Math.min(0,...cs),cmax=Math.max(1,...cs);
  const X=v=>l+(v-t0)/((t1-t0)||1)*(W-l-r),Y=v=>t+(v-cmin)/((cmax-cmin)||1)*(H-t-bt);
  const all=ts.map((v,k)=>X(v).toFixed(1)+","+Y(cs[k]).toFixed(1)).join(" ");
  let rec="";const rf=b[I.rfrom];
  if(rf!=null){const pts=ts.map((v,k)=>[v,cs[k]]).filter(p=>p[0]>=rf-1e-6);rec='<polyline points="'+pts.map(p=>X(p[0]).toFixed(1)+","+Y(p[1]).toFixed(1)).join(" ")+'" fill="none" stroke="'+COL[c]+'" stroke-width="3" stroke-linecap="round"/>';}
  const dotsS=ts.map((v,k)=>'<circle cx="'+X(v).toFixed(1)+'" cy="'+Y(cs[k]).toFixed(1)+'" r="2.6" fill="#24507a"/>').join("");
  return '<svg class="chart" width="'+W+'" height="'+H+'" viewBox="0 0 '+W+' '+H+'" role="img" aria-label="Cumulatieve zakking">'+
    '<line x1="'+l+'" x2="'+(W-r)+'" y1="'+Y(0)+'" y2="'+Y(0)+'" stroke="#c9c6bc"/>'+
    '<line x1="'+l+'" x2="'+(W-r)+'" y1="'+Y(cmax)+'" y2="'+Y(cmax)+'" stroke="#eeece6"/>'+
    '<text x="'+(l-5)+'" y="'+(Y(0)+4)+'" text-anchor="end" font-size="10" fill="#6b6a64">0</text>'+
    '<text x="'+(l-5)+'" y="'+(Y(cmax)+4)+'" text-anchor="end" font-size="10" fill="#6b6a64">'+nf(cmax,0)+' mm</text>'+
    '<text x="'+l+'" y="'+(H-6)+'" font-size="10" fill="#6b6a64">'+t0+'</text>'+
    '<text x="'+(W-r)+'" y="'+(H-6)+'" text-anchor="end" font-size="10" fill="#6b6a64">'+t1+'</text>'+
    '<text x="'+((l+W-r)/2)+'" y="'+(H-6)+'" text-anchor="middle" font-size="10" fill="#6b6a64">zakking ↓</text>'+
    '<polyline points="'+all+'" fill="none" stroke="#24507a" stroke-width="1.5"/>'+rec+dotsS+'</svg>';
}
function pill(c){const[l]=label(c);return '<div class="pill"><i style="background:'+COL[c]+(c==="stab"?";box-shadow:0 0 0 2px "+RING:"")+'"></i>'+l+'</div>';}
function openBout(i){
  if(!curPop||curPop.i!==i)TEL.tel("punt");
  const b=B[i],c=C[i];
  const y=v=>v==null?"":Math.floor(v);
  const age=b[I.lastY]!=null?Math.max(0,Math.round(M.nowYear-b[I.lastY])):null;
  const html='<div class="pop"><div class="ttl">'+esc(b[I.addr]||"Adres onbekend")+'</div>'+
    '<div class="sub">Meetbout '+esc(b[I.id])+(b[I.blk]?" · blok "+esc(b[I.blk]):"")+(b[I.status]?" · "+esc(b[I.status]):"")+'</div>'+pill(c)+
    '<table class="kv">'+
    '<tr><th>Recent</th><td>'+(b[I.recent]==null?"te korte meetreeks":nf(b[I.recent])+" mm/jaar <span class=m>"+y(b[I.rfrom])+"–"+y(b[I.lastY])+"</span>")+'</td></tr>'+
    '<tr><th>Eerder</th><td>'+(b[I.early]==null?"–":nf(b[I.early])+" mm/jaar <span class=m>"+y(b[I.efrom])+"–"+y(b[I.rfrom])+"</span>")+'</td></tr>'+
    '<tr><th>Gemiddeld</th><td>'+(b[I.total]==null?"–":nf(b[I.total])+" mm/jaar")+'</td></tr>'+
    '<tr><th>Totaal gezakt</th><td>'+(b[I.cum]==null?"–":nf(b[I.cum],1)+" mm")+'</td></tr>'+
    '<tr><th>Metingen</th><td>'+b[I.n]+(b[I.last]?" · laatste "+fmtDate(b[I.last])+(age!=null?" ("+age+" jaar geleden)":""):"")+'</td></tr>'+
    '</table>'+(b[I.ser]?chart(b,c):serState==="laden"?'<p class="mnote">Grafiek wordt geladen …</p>':serState==="fout"&&b[I.n]>1?'<p class="mnote">Grafiek kon niet worden geladen.</p>':"")+
    '<a class="raw" href="#" data-meas="'+i+'">Alle meetgegevens bekijken</a></div>';
  const p=L.popup({maxWidth:340,autoPanPadding:[20,20]}).setLatLng([b[I.lat],b[I.lon]]).setContent(html).openOn(map);
  curPop={i:i,p:p};
  const el=p.getElement&&p.getElement();
  if(el){const a=el.querySelector("a[data-meas]");if(a)a.addEventListener("click",ev=>{ev.preventDefault();openMeas(i);});}
}
function houseNo(a){const m=String(a||"").match(/\s(\d+)/);return m?parseInt(m[1],10):1e9;}
function groupTitle(gi){return (G[gi][0]||"Zonder adres")+(G[gi][1]?" · blok "+G[gi][1]:"");}
function openList(title,idx,latlng){
  idx=idx.slice().sort((a,b)=>houseNo(B[a][I.addr])-houseNo(B[b][I.addr])||String(B[a][I.addr]).localeCompare(String(B[b][I.addr]),"nl"));
  const v=idx.filter(i=>C[i]!=="oud").map(i=>B[i][I.recent]);const med=v.length?median(v):null;
  const n={rood:0,geel:0,groen:0,stab:0,oud:0};idx.forEach(i=>n[C[i]]++);
  const mix=KL.filter(k=>n[k]).map(k=>'<span style="white-space:nowrap"><i class="sw" style="display:inline-block;width:9px;height:9px;border-radius:50%;background:'+COL[k]+'"></i> '+n[k]+"</span>").join(" &nbsp;");
  const rows=idx.map(i=>'<tr data-i="'+i+'"><td><i class="sw" style="background:'+COL[C[i]]+'"></i></td><td>'+esc(B[i][I.addr]||B[i][I.id])+'</td><td class="n">'+(B[i][I.recent]==null?"–":nf(B[i][I.recent]))+'</td><td class="n m">'+esc((B[i][I.last]||"").slice(0,4))+"</td></tr>").join("");
  const html='<div class="pop"><div class="ttl">'+esc(title)+'</div><div class="sub">'+idx.length+" meetbouten · mediaan recent "+(med==null?"–":nf(med)+" mm/jaar")+'</div>'+
    '<div>'+mix+'</div><div class="list"><table><tbody>'+rows+'</tbody></table></div><div class="sub" style="margin-top:4px">mm/jaar recent · jaar laatste meting. Klik een regel voor details.</div></div>';
  const p=L.popup({maxWidth:340,autoPanPadding:[20,20]}).setLatLng(latlng).setContent(html).openOn(map);
  const el=p.getElement&&p.getElement();
  if(el)el.querySelectorAll("tr[data-i]").forEach(tr=>tr.addEventListener("click",()=>{const i=+tr.dataset.i;map.setView([B[i][I.lat],B[i][I.lon]],Math.max(map.getZoom(),18));setTimeout(()=>openBout(i),250);}));
}

// ---------- meetgegevens (tabel + ruwe data) ----------
const modal=document.getElementById("modal"),mTitle=document.getElementById("mTitle"),mSub=document.getElementById("mSub"),mBody=document.getElementById("mBody");
let lastFocus=null;
function closeMeas(){modal.hidden=true;mBody.innerHTML="";if(lastFocus&&lastFocus.focus)lastFocus.focus();}
document.getElementById("mClose").addEventListener("click",closeMeas);
modal.addEventListener("click",e=>{if(e.target===modal)closeMeas();});
document.addEventListener("keydown",e=>{if(e.key==="Escape"&&!modal.hidden)closeMeas();});
const SOORT={N:"Nulmeting",H:"Herhalingsmeting"};
function measUrl(b,size){return M.api+"metingen/?hoortBijMeetbout.identificatie="+encodeURIComponent(b[I.id])+"&_format=json&_pageSize="+size;}
function summaryHtml(b){
  const y=v=>v==null?"":Math.floor(v);
  return '<div class="msum">'+
    '<span>Recent: <b>'+(b[I.recent]==null?"–":nf(b[I.recent])+" mm/jaar")+'</b> '+(b[I.rfrom]!=null?'<span class="m">'+y(b[I.rfrom])+"–"+y(b[I.lastY])+"</span>":"")+'</span>'+
    '<span>Eerder: <b>'+(b[I.early]==null?"–":nf(b[I.early])+" mm/jaar")+'</b> '+(b[I.efrom]!=null?'<span class="m">'+y(b[I.efrom])+"–"+y(b[I.rfrom])+"</span>":"")+'</span>'+
    '<span>Totaal gezakt: <b>'+(b[I.cum]==null?"–":nf(b[I.cum],1)+" mm")+'</b></span>'+
    '<span>Klasse: <b>'+label(C[b._i!=null?b._i:0])[0]+'</b></span></div>';
}
function yearFracOf(iso){const d=new Date(iso+"T00:00:00Z");if(isNaN(d))return null;const y=d.getUTCFullYear();return y+(Date.UTC(y,d.getUTCMonth(),d.getUTCDate())-Date.UTC(y,0,1))/864e5/365.25;}
function rawSection(url,data,err){
  const link='<a href="'+esc(url)+'" target="_blank" rel="noopener">Open bij de gemeente ↗</a>';
  if(err)return '<h3>Ruwe gegevens van de gemeente</h3><p class="mnote">Kon de ruwe gegevens niet rechtstreeks ophalen ('+esc(err)+'). '+link+'</p>';
  return '<h3>Ruwe gegevens van de gemeente</h3><p class="mnote" style="margin:0 0 6px">Precies zoals de gemeente ze levert (API, JSON). '+link+'</p><pre class="rawjson" id="rawjson"></pre>';
}
function openMeas(i){
  TEL.tel("meet");
  const b=B[i];lastFocus=document.activeElement;
  mTitle.textContent=b[I.addr]||"Meetbout "+b[I.id];
  mSub.textContent="Meetbout "+b[I.id]+(b[I.blk]?" · blok "+b[I.blk]:"")+(b[I.status]?" · "+b[I.status]:"");
  mBody.innerHTML='<p class="mnote">Meetgegevens ophalen bij de gemeente …</p>';
  modal.hidden=false;document.getElementById("mClose").focus();
  const url=measUrl(b,1000);
  fetch(url,{headers:{Accept:"application/hal+json, application/json"}}).then(r=>{if(!r.ok)throw new Error("HTTP "+r.status);return r.json();})
  .then(data=>{
    if(modal.hidden)return;
    const emb=data&&data._embedded||{};let items=emb.metingen||Object.values(emb).find(Array.isArray)||[];
    items=items.slice().sort((a,c)=>String(a.datum).localeCompare(String(c.datum)));
    const rf=b[I.rfrom];
    const rows=items.map(m=>{const t=yearFracOf(String(m.datum||"").slice(0,10));const rec=rf!=null&&t!=null&&t>=rf-0.006;
      const s=m.typeMeting||"";
      return '<tr'+(rec?' class="rec"':'')+'><td>'+fmtDate(String(m.datum||"").slice(0,10))+'</td><td>'+esc(SOORT[s]||s||"–")+
        '</td><td class="n">'+(m.hoogteTovNap==null?"–":nf(m.hoogteTovNap,4))+'</td><td class="n">'+(m.zakking==null?"–":nf(m.zakking,1))+
        '</td><td class="n">'+(m.zakkingCumulatief==null?"–":nf(m.zakkingCumulatief,1))+'</td><td class="n">'+(m.aantalDagen==null?"–":Number(m.aantalDagen).toLocaleString("nl-NL"))+
        '</td><td class="n">'+(m.zakkingssnelheid==null?"–":nf(m.zakkingssnelheid))+'</td><td>'+esc(m.isGemetenDoor||"–")+'</td><td>'+esc(m.wijzeVanInwinnenOmschrijving||"–")+'</td></tr>';}).join("");
    mBody.innerHTML=summaryHtml(Object.assign(b.slice(),{_i:i}))+
      '<div class="mt-wrap"><table class="mt"><thead><tr><th>Datum</th><th>Soort meting</th><th class="n">Hoogte t.o.v. NAP (m)</th><th class="n">Zakking sinds vorige (mm)</th><th class="n">Totaal gezakt (mm)</th><th class="n">Dagen sinds vorige</th><th class="n">Gem. snelheid sinds start (mm/jaar)</th><th>Gemeten door</th><th>Methode</th></tr></thead><tbody>'+
      (rows||'<tr><td colspan="9">Geen metingen gevonden.</td></tr>')+'</tbody></table></div>'+
      '<p class="mnote">'+items.length+' metingen. Geel gemarkeerd: de periode waarover de recente snelheid is berekend. Zakking is positief als het punt lager is komen te liggen.</p>'+
      rawSection(url,data,null);
    const pre=document.getElementById("rawjson");if(pre)pre.textContent=JSON.stringify(data,null,2);
  })
  .catch(err=>{
    if(modal.hidden)return;
    const s=b[I.ser]||[];const rows=[];
    for(let k=0;k<s.length;k+=2){const t=s[k]+1900;const y=Math.floor(t);const d=new Date(Date.UTC(y,0,1)+(t-y)*365.25*864e5);
      rows.push('<tr><td>± '+d.toLocaleDateString("nl-NL",{month:"short",year:"numeric",timeZone:"UTC"})+'</td><td class="n">'+nf(s[k+1],1)+'</td></tr>');}
    mBody.innerHTML='<div class="warn">De actuele gegevens konden niet bij de gemeente worden opgehaald. Hieronder de metingen zoals ze bij het maken van deze kaart zijn opgeslagen (datum bij benadering).</div>'+
      summaryHtml(Object.assign(b.slice(),{_i:i}))+
      '<div class="mt-wrap"><table class="mt"><thead><tr><th>Datum</th><th class="n">Totaal gezakt (mm)</th></tr></thead><tbody>'+(rows.join("")||'<tr><td colspan="2">Geen metingen opgeslagen.</td></tr>')+'</tbody></table></div>'+
      rawSection(url,null,String(err&&err.message||err));
  });
}

// ---------- zoeken ----------
const byStreet=new Map();const nameOf=new Map();
B.forEach((b,i)=>{const s=b[I.street];if(!s)return;const k=s.toLowerCase();if(!byStreet.has(k)){byStreet.set(k,[]);nameOf.set(k,s);}byStreet.get(k).push(i);});
const streetKeys=[...byStreet.keys()].sort((a,b)=>a.localeCompare(b,"nl"));
document.getElementById("streets").innerHTML=streetKeys.map(k=>'<option value="'+esc(nameOf.get(k))+'">').join("");
const msg=document.getElementById("msg");
function say(t){msg.textContent=t;msg.hidden=false;clearTimeout(say._t);say._t=setTimeout(()=>msg.hidden=true,3500);}
document.getElementById("search").addEventListener("submit",e=>{
  e.preventDefault();
  const q=document.getElementById("q").value.trim();if(!q)return;
  TEL.tel("zoek");
  if(/^\d{5,}$/.test(q)){const i=B.findIndex(b=>String(b[I.id])===q);if(i>=0){map.setView([B[i][I.lat],B[i][I.lon]],18);setTimeout(()=>openBout(i),300);return;}}
  const base=q.toLowerCase().replace(/\s+\d.*$/,"").trim();
  let key=byStreet.has(base)?base:(streetKeys.find(k=>k.startsWith(base))||streetKeys.find(k=>k.includes(base)));
  if(!key){say("Niets gevonden voor “"+q+"”.");return;}
  const idx=byStreet.get(key);
  const num=q.match(/\s(\d+)/);
  if(num){const n=parseInt(num[1],10);let best=null,bd=1e9;for(const i of idx){const h=houseNo(B[i][I.addr]);if(Math.abs(h-n)<bd){bd=Math.abs(h-n);best=i;}}
    if(best!=null&&bd<=4){map.setView([B[best][I.lat],B[best][I.lon]],18);setTimeout(()=>openBout(best),300);return;}}
  const bb=L.latLngBounds(idx.map(i=>[B[i][I.lat],B[i][I.lon]]));
  map.fitBounds(bb,{maxZoom:17,padding:[60,60]});
  setTimeout(()=>openList(nameOf.get(key),idx,bb.getCenter()),400);
});

// ---------- feedback ----------
(function(){
  const url=(M.feedbackUrl||"").trim(),btn=document.getElementById("fbBtn");
  if(!/^https:\/\//i.test(url))return;   // geen formulierdienst ingesteld: geen knop
  btn.hidden=false;
  const fm=document.getElementById("fbModal"),form=document.getElementById("fbForm"),st=document.getElementById("fbStatus"),
        msgEl=document.getElementById("fbMsg"),send=document.getElementById("fbSend"),stars=document.getElementById("fbStars");
  let rating=0,prev=null;
  for(let k=1;k<=5;k++){stars.insertAdjacentHTML("beforeend",'<input type="radio" name="fbr" id="fbr'+k+'" value="'+k+'"><label for="fbr'+k+'" title="'+k+(k===1?" ster":" sterren")+'" aria-label="'+k+(k===1?" ster":" sterren")+'">&#9733;</label>');}
  const labs=[...stars.querySelectorAll("label")];
  const paint=n=>labs.forEach((l,ix)=>l.classList.toggle("on",ix<n));
  stars.querySelectorAll("input").forEach(inp=>inp.addEventListener("change",()=>{rating=+inp.value;paint(rating);if(st.className==="err"){st.textContent="";st.className="";}}));
  labs.forEach((l,ix)=>{l.addEventListener("mouseenter",()=>paint(ix+1));l.addEventListener("mouseleave",()=>paint(rating));});
  msgEl.addEventListener("input",()=>{document.getElementById("fbCnt").textContent=msgEl.value.length+" / 280";});
  function open(){prev=document.activeElement;fm.hidden=false;st.textContent="";st.className="";const f=stars.querySelector("input");if(f)f.focus();}
  function close(){fm.hidden=true;if(prev&&prev.focus)prev.focus();}
  btn.addEventListener("click",open);
  document.getElementById("fbClose").addEventListener("click",close);
  fm.addEventListener("click",e=>{if(e.target===fm)close();});
  document.addEventListener("keydown",e=>{if(e.key==="Escape"&&!fm.hidden)close();});
  form.addEventListener("submit",e=>{
    e.preventDefault();
    const text=msgEl.value.trim().slice(0,280);
    if(document.getElementById("fbHp").value){st.className="ok";st.textContent="Bedankt voor je feedback!";return;}
    if(!rating){st.className="err";st.textContent="Kies eerst 1 tot 5 sterren.";return;}
    const body={sterren:rating,tekst:text,versie:M.generated};
    send.disabled=true;st.className="";st.textContent="Versturen …";
    // Google Apps Script beantwoordt een POST met een doorverwijzing; het antwoord lezen hangt af van
    // CORS-headers die we niet in de hand hebben. Daarom 'no-cors' + text/plain: het bericht wordt
    // altijd afgeleverd (geen preflight), alleen het antwoord blijft onleesbaar. Controle gebeurt hierboven.
    fetch(url,{method:"POST",mode:"no-cors",headers:{"Content-Type":"text/plain;charset=utf-8"},body:JSON.stringify(body)})
      .then(()=>{
        st.className="ok";st.textContent="Bedankt voor je feedback!";form.reset();rating=0;paint(0);document.getElementById("fbCnt").textContent="0 / 280";setTimeout(close,1800);})
      .catch(()=>{st.className="err";st.textContent="Versturen lukte niet. Probeer het later nog eens.";})
      .finally(()=>{send.disabled=false;});
  });
  L.DomEvent.disableClickPropagation(btn);
})();

buildRows();fillSettings();styleAll();

// ---------- meetreeksen (website: apart bestand, na de kaart geladen) ----------
if(M.seriesUrl){
  serState="laden";
  fetch(M.seriesUrl).then(r=>{if(!r.ok)throw new Error("HTTP "+r.status);return r.json();})
  .then(arr=>{
    if(!Array.isArray(arr)||arr.length!==B.length)throw new Error("onverwacht formaat");
    for(let i=0;i<B.length;i++)B[i][I.ser]=arr[i];
    serState="ok";
    if(curPop&&map.hasLayer(curPop.p))openBout(curPop.i);
  })
  .catch(()=>{serState="fout";if(curPop&&map.hasLayer(curPop.p))openBout(curPop.i);});
}
}catch(err){window.__showErr(String(err&&err.stack||err));}
}
</script>
<script>__BOOT__</script>
</body>
</html>
"""

if __name__ == "__main__":
    main()
