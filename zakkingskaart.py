#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Zakkingskaart
=============

Haalt alle meetbouten en hun metingen op bij de gemeente Amsterdam
(api.data.amsterdam.nl/v1/meetbouten/), berekent per meetbout hoe snel de
gevel zakt, en maakt daar een interactieve kaart van. In website-modus komen
daar de NAP-peilmerken van Rijkswaterstaat voor heel Nederland bij.

  rood    = zakt snel            (standaard: 2,0 mm/jaar of meer)
  geel    = zakt matig         (standaard: 1,0 tot 2,0 mm/jaar)
  groen   = zakt nauwelijks      (standaard: minder dan 1,0 mm/jaar)
  groen, gestippeld/omrand = gestabiliseerd: eerst snel, nu nauwelijks
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

Bronnen: Gemeente Amsterdam, dataset Meetbouten (CC0); Rijkswaterstaat, NAP-peilmerken
(CC0). De kaart geeft deze openbare gegevens eenvoudiger weer; ze geeft geen oordeel
of advies over gebouwen, de grond of de waarde ervan.
"""

import argparse
import base64
import csv
import gzip
import hashlib
import http.client
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

# NAP-peilmerken van Rijkswaterstaat (heel Nederland, CC0). MapServer-laag 0 = actuele punten,
# 1 = historische punten, tabel 4 = alle gemeten hoogtes per peilmerk.
NAP_SERVICE = "https://geo.rijkswaterstaat.nl/arcgis/rest/services/GDR/nap/MapServer"
NAP_PUNTLAGEN = (0, 1)
NAP_HISTORIE = 4
NAP_PAGINA = 2000

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
        return "zakt matig"
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


# --------------------------------------------------------------------------
# NAP-peilmerken (Rijkswaterstaat)
# --------------------------------------------------------------------------

NAP_TIMEOUT = 45   # seconden; een pagina duurt normaal een paar seconden


class NapVerbinding:
    """Eén blijvende verbinding (keep-alive) voor alle verzoeken aan Rijkswaterstaat. Hun
    beveiliging remt adressen af die veel nieuwe verbindingen openen; met één verbinding
    voor honderden pagina's blijft dat onder de radar."""

    def __init__(self, base):
        u = urllib.parse.urlsplit(base)
        self.https = u.scheme == "https"
        self.host = u.netloc
        self.conn = None
        self.nieuw = 0          # aantal geopende verbindingen (voor het log)

    def sluit(self):
        if self.conn is not None:
            try:
                self.conn.close()
            except Exception:
                pass
        self.conn = None

    def get_json(self, url, retries=4):
        u = urllib.parse.urlsplit(url)
        path = u.path + ("?" + u.query if u.query else "")
        delay, last_err = 2.0, None
        for attempt in range(1, retries + 1):
            try:
                if self.conn is None:
                    if self.https:
                        self.conn = http.client.HTTPSConnection(self.host, timeout=NAP_TIMEOUT,
                                                                context=ssl.create_default_context())
                    else:
                        self.conn = http.client.HTTPConnection(self.host, timeout=NAP_TIMEOUT)
                    self.nieuw += 1
                self.conn.request("GET", path, headers={"User-Agent": USER_AGENT, "Accept": "application/json",
                                                        "Connection": "keep-alive"})
                r = self.conn.getresponse()
                body = r.read()
                if r.will_close:
                    self.sluit()
                if r.status in (429, 500, 502, 503, 504):
                    ra = r.getheader("Retry-After") or ""
                    wait = float(ra) if ra.replace(".", "", 1).isdigit() else delay
                    log(f"   server antwoordt {r.status}, opnieuw over {wait:.0f} s (poging {attempt}/{retries})")
                    last_err = f"HTTP {r.status}"
                    time.sleep(wait)
                    delay = min(delay * 2, 90)
                    continue
                if r.status >= 400:
                    raise FatalHTTP(r.status, url, body[:500].decode("utf-8", "replace"))
                return json.loads(body.decode("utf-8"))
            except (OSError, http.client.HTTPException, ValueError) as e:
                self.sluit()
                last_err = e
                log(f"   verbindingsprobleem ({e.__class__.__name__}: {e}), opnieuw over {delay:.0f} s "
                    f"(poging {attempt}/{retries})")
                time.sleep(delay)
                delay = min(delay * 2, 90)
        raise RuntimeError(f"Ophalen mislukt na {retries} pogingen: {url} ({last_err})")


class NapOnderbroken(RuntimeError):
    """Ophalen bij Rijkswaterstaat gestopt; wat binnen is, staat in de voortgangsbestand."""


def arcgis_query_all(url, params, delay, max_pages=0, label="", after=0, on_page=None, geduld=(), deadline=None,
                     get=None):
    """Haalt records op uit een ArcGIS REST-query, in pagina's van NAP_PAGINA. Bladert op objectid
    (where objectid > laatste) in plaats van met een steeds grotere offset: dat is lichter voor
    de server. Na elke pagina wordt on_page(features, laatste_objectid) aangeroepen. Weigert de
    server een tijd lang, dan wordt telkens langer gewacht (geduld, in seconden)."""
    where0 = params.get("where") or "1=1"
    fields = params.get("outFields", "*")
    if fields != "*" and "objectid" not in fields.split(","):
        fields += ",objectid"
    get = get or (lambda u: http_get_json(u, retries=4))
    total, pages, t0 = 0, 0, time.time()
    while True:
        if deadline and time.time() > deadline:
            raise NapOnderbroken("tijdslimiet bereikt")
        q = dict(params)
        q.update({"where": f"({where0}) AND objectid > {int(after)}", "outFields": fields,
                  "orderByFields": "objectid", "f": "json", "resultOffset": "0",
                  "resultRecordCount": str(NAP_PAGINA)})
        full = url + "/query?" + urllib.parse.urlencode(q)
        data, fout = None, None
        for wacht in (0,) + tuple(geduld):
            if wacht:
                if deadline and time.time() + wacht > deadline:
                    break
                log(f"   Rijkswaterstaat weigert even; {wacht / 60:.0f} min wachten en dan verder")
                time.sleep(wacht)
            try:
                data = get(full)
                break
            except RuntimeError as e:
                fout = e
        if data is None:
            raise NapOnderbroken(str(fout or "tijdslimiet bereikt"))
        if isinstance(data, dict) and data.get("error"):
            err = data["error"]
            raise RuntimeError(f"ArcGIS-fout {err.get('code')}: {err.get('message')}")
        feats = (data or {}).get("features") or []
        if feats:
            ids = [(f.get("attributes") or {}).get("objectid") for f in feats]
            if any(not isinstance(i, int) for i in ids):
                raise RuntimeError("ArcGIS-antwoord zonder objectid")
            after = max(ids)
        total += len(feats)
        pages += 1
        if on_page:
            on_page(feats, after)
        if pages % 25 == 0 or not feats:
            log(f"   {label}: {fmt_int(total)} records erbij ({(time.time() - t0) / 60:.0f} min)")
        if not feats or (len(feats) < NAP_PAGINA and not data.get("exceededTransferLimit")):
            return
        if (max_pages and pages >= max_pages) or pages > 2000:
            return
        if delay:
            time.sleep(delay)


def nap_punt(f):
    a = f.get("attributes") or {}
    g = f.get("geometry") or {}
    lat, lon = g.get("y"), g.get("x")
    if lat is None or lon is None:
        try:
            lat, lon = rd_to_wgs84(float(a.get("x_rd")), float(a.get("y_rd")))
        except (TypeError, ValueError):
            return None
    if not a.get("puntnummer"):
        return None
    return {"p": a["puntnummer"].strip(), "lat": round(lat, 6), "lon": round(lon, 6),
            "s": (a.get("status") or "").strip(), "o": re.sub(r"\s+", " ", a.get("omschrijving") or "").strip()}


def nap_meting(f):
    a = f.get("attributes") or {}
    if a.get("puntnummer") and a.get("projectdatum") is not None and a.get("hoogte") is not None:
        return {"p": a["puntnummer"].strip(), "j": a.get("project_id") or "",
                "d": a["projectdatum"], "h": a["hoogte"]}
    return None


def fetch_nap(args, max_pages, vg, bewaar):
    """Haalt de puntlagen en de historietabel op. vg is de voortgang (per bron de laatste objectid
    en of hij klaar is, plus de records tot nu toe); bewaar() schrijft die naar de cache, zodat
    een volgende run verder kan waar deze stopte."""
    base = args.nap_service.rstrip("/")
    geduld = tuple(float(x) for x in str(args.nap_geduld).split(",") if x.strip())
    deadline = time.time() + args.nap_max_minuten * 60 if args.nap_max_minuten else None
    verb = NapVerbinding(base)
    bronnen = [(str(l), {"where": "1=1", "outFields": "puntnummer,x_rd,y_rd,status,omschrijving",
                         "returnGeometry": "true", "outSR": "4326"}, "punten", nap_punt, f"peilmerken (laag {l})")
               for l in NAP_PUNTLAGEN]
    bronnen.append((str(NAP_HISTORIE), {"where": "1=1", "outFields": "puntnummer,project_id,projectdatum,hoogte",
                                        "returnGeometry": "false"}, "hist", nap_meting, "NAP-historie"))
    for key, params, doel, conv, label in bronnen:
        b = vg["bronnen"].setdefault(key, {"na": 0, "klaar": False})
        if b["klaar"]:
            continue
        if b["na"]:
            log(f"   {label}: verder vanaf objectid {b['na']}")
        teller = [0]

        def on_page(feats, na, b=b, doel=doel, conv=conv):
            for f in feats:
                r = conv(f)
                if r:
                    vg[doel].append(r)
            b["na"] = na
            teller[0] += 1
            if teller[0] % 25 == 0:
                bewaar()

        arcgis_query_all(base + "/" + key, params, args.nap_pauze, max_pages, label,
                         after=b["na"], on_page=on_page, geduld=geduld, deadline=deadline,
                         get=lambda u: verb.get_json(u, retries=3))
        b["klaar"] = True
        bewaar()
        log(f"   {label}: compleet (verbindingen geopend tot nu toe: {verb.nieuw})")
    verb.sluit()
    return vg["punten"], vg["hist"]


def nap_vingerafdruk(args):
    """Korte controle bij Rijkswaterstaat (een handvol verzoeken): aantallen records en de
    nieuwste datums. Verandert dit, dan zijn er nieuwe of gewijzigde gegevens. None als het
    niet lukt."""
    base = args.nap_service.rstrip("/")
    verb = NapVerbinding(base)

    def q(laag, **extra):
        params = {"where": "1=1", "returnGeometry": "false", "f": "json"}
        params.update(extra)
        data = verb.get_json(base + "/" + str(laag) + "/query?" + urllib.parse.urlencode(params), retries=3)
        if not isinstance(data, dict) or data.get("error"):
            raise RuntimeError(f"ArcGIS-fout: {(data or {}).get('error')}")
        return data

    def nieuwste(laag, veld):
        feats = q(laag, outFields=veld, orderByFields=veld + " DESC", resultOffset="0",
                  resultRecordCount="1").get("features") or []
        return (feats[0].get("attributes") or {}).get(veld) if feats else None

    try:
        va = {}
        for laag in NAP_PUNTLAGEN:
            va[f"laag{laag}"] = q(laag, returnCountOnly="true").get("count")
            va[f"laag{laag}Mutatie"] = nieuwste(laag, "datum_mutatie")
        va["metingen"] = q(NAP_HISTORIE, returnCountOnly="true").get("count")
        va["nieuwsteMeting"] = nieuwste(NAP_HISTORIE, "projectdatum")
        if va["metingen"] is None:
            raise RuntimeError("geen aantal ontvangen")
        return va
    except (RuntimeError, FatalHTTP) as e:
        log(f"   controle bij Rijkswaterstaat lukte niet ({e})")
        return None
    finally:
        verb.sluit()


def update_nap(cache_dir, args, max_pages, state):
    """Peilmerken + historie, met eigen cache. Om de --nap-controle dagen (21) wordt kort
    gecontroleerd of Rijkswaterstaat nieuwe of gewijzigde gegevens heeft (aantallen en nieuwste
    datums); zo ja, of als de cache ouder is dan --nap-ververs dagen (91), wordt alles opnieuw
    opgehaald. Peilmerken worden maar eens per 5-15 jaar gemeten. Lukt het ophalen niet in
    één keer, dan wordt de voortgang bewaard en gaat de volgende run verder. Tot het compleet
    is, wordt de oude cache gebruikt; zonder cache wordt de laag overgeslagen."""
    fp = os.path.join(cache_dir, "nap_punten.jsonl.gz")
    fh = os.path.join(cache_dir, "nap_historie.jsonl.gz")
    fv = os.path.join(cache_dir, "nap_voortgang.json.gz")
    have = os.path.exists(fp) and os.path.exists(fh)
    age = days_since(state.get("napOpgehaald"))
    bezig = os.path.exists(fv) and not args.vernieuw
    if have and not bezig and not args.vernieuw and age is not None and age < args.nap_ververs:
        laatste = state.get("napControle") or state.get("napOpgehaald")
        ctrl_age = days_since(laatste)
        if args.nap_controle and (ctrl_age is None or ctrl_age >= args.nap_controle or not state.get("napVingerafdruk")):
            log(f"-> NAP-peilmerken: korte controle bij Rijkswaterstaat (vorige: {laatste})")
            va = nap_vingerafdruk(args)
            if va is not None:
                state["napControle"] = date.today().isoformat()
                oud = state.get("napVingerafdruk")
                if not oud:
                    state["napVingerafdruk"] = va
                    log("   uitgangspunt vastgelegd; vanaf nu worden wijzigingen opgemerkt")
                elif va != oud:
                    verschil = ", ".join(f"{k}: {oud.get(k)} -> {va.get(k)}" for k in sorted(va) if va.get(k) != oud.get(k))
                    log(f"   nieuwe of gewijzigde gegevens ({verschil}); alles wordt opnieuw opgehaald")
                    age = None
                else:
                    log("   geen wijzigingen")
        if age is not None:
            log(f"-> NAP-peilmerken: uit cache (opgehaald {state.get('napOpgehaald')})")
            return load_jsonl_gz(fp), load_jsonl_gz(fh)
    vg = None
    if os.path.exists(fv) and not args.vernieuw:
        try:
            with gzip.open(fv, "rt", encoding="utf-8") as fhd:
                vg = json.load(fhd)
            if vg.get("versie") != 1 or (days_since(vg.get("gestart")) or 0) > 60:
                vg = None
        except (OSError, ValueError):
            vg = None
    if vg:
        log(f"-> NAP-peilmerken: verder met ophalen (begonnen {vg['gestart']}, "
            f"{fmt_int(len(vg['punten']))} peilmerken en {fmt_int(len(vg['hist']))} metingen binnen)")
    else:
        log("-> NAP-peilmerken: ophalen bij Rijkswaterstaat")
        vg = {"versie": 1, "gestart": date.today().isoformat(), "bronnen": {}, "punten": [], "hist": [],
              "vingerafdruk": nap_vingerafdruk(args) if args.nap_controle else None}

    def bewaar():
        tmp = fv + ".tmp"
        with gzip.open(tmp, "wt", encoding="utf-8") as fhd:
            json.dump(vg, fhd, separators=(",", ":"))
        os.replace(tmp, fv)

    try:
        punten, hist = fetch_nap(args, max_pages, vg, bewaar)
        prev = state.get("napAantallen") or {}
        if prev and not args.proef and (len(punten) < 0.9 * prev.get("punten", 0) or len(hist) < 0.9 * prev.get("metingen", 0)):
            os.remove(fv)
            raise RuntimeError(f"veel minder data dan vorige keer ({len(punten)} punten, {len(hist)} metingen; "
                               f"was {prev.get('punten')}, {prev.get('metingen')})")
        if not punten or not hist:
            os.remove(fv)
            raise RuntimeError("geen peilmerken of metingen ontvangen")
        write_jsonl_gz(fp, punten)
        write_jsonl_gz(fh, hist)
        if os.path.exists(fv):
            os.remove(fv)
        state["napOpgehaald"] = date.today().isoformat()
        state["napAantallen"] = {"punten": len(punten), "metingen": len(hist)}
        state["napControle"] = date.today().isoformat()
        if vg.get("vingerafdruk"):
            state["napVingerafdruk"] = vg["vingerafdruk"]
        else:
            state.pop("napVingerafdruk", None)   # volgende run legt het uitgangspunt vast
        log(f"   klaar: {fmt_int(len(punten))} peilmerken, {fmt_int(len(hist))} metingen")
        return punten, hist
    except NapOnderbroken as e:
        bewaar()
        log(f"   let op: NAP-peilmerken nog niet compleet ({e}). Binnen: {fmt_int(len(vg['punten']))} peilmerken, "
            f"{fmt_int(len(vg['hist']))} metingen. De volgende run gaat hier verder.")
    except (RuntimeError, FatalHTTP) as e:
        log(f"   let op: NAP-peilmerken niet opgehaald ({e})")
    if have:
        log("   oude NAP-cache wordt gebruikt")
        return load_jsonl_gz(fp), load_jsonl_gz(fh)
    log("   de NAP-laag wordt deze keer overgeslagen")
    return None


def nap_series(meas):
    """Meetreeks van één peilmerk. Bevat de reeks herziene waarden (project-id met '='; dat is
    het huidige NAP-stelsel), dan worden alleen die gebruikt: anders staan oude en herrekende
    hoogtes door elkaar en ontstaan nepsprongen van centimeters."""
    herzien = [m for m in meas if "=" in (m.get("j") or "")]
    use = herzien if herzien else meas
    pts = []
    for m in use:
        try:
            t = date(1970, 1, 1) + timedelta(milliseconds=int(m["d"]))
            h = float(m["h"])
        except (TypeError, ValueError, OverflowError):
            continue
        ty = t.year + (t.timetuple().tm_yday - 1) / 365.25
        # Twee metingen op dezelfde dag (uit verschillende projecten) blijven allebei staan:
        # ze worden hieronder gemiddeld, en verschillen ze te veel, dan is de reeks onzeker.
        pts.append((ty, h, t.isoformat()))
    pts.sort(key=lambda p: (p[0], p[1]))
    # Metingen binnen een jaar horen bij dezelfde meetcampagne: samenvoegen. Verschillen ze
    # meer dan 1 cm, dan is de reeks onbetrouwbaar (dubbele of herrekende waarde).
    serie, onzeker, groep = [], False, []
    for p in pts + [None]:
        if p is not None and groep and p[0] - groep[0][0] < 1.0:
            groep.append(p)
            continue
        if groep:
            hs = [g[1] for g in groep]
            if max(hs) - min(hs) > 0.010:
                onzeker = True
            serie.append((sum(g[0] for g in groep) / len(groep), sum(hs) / len(hs), groep[-1][2]))
        groep = [p] if p is not None else []
    return serie, bool(herzien), onzeker


def nap_analyse(serie, min_span=8.0):
    """Snelheden in mm/jaar (positief = zakt). 'recent': laatste meting t.o.v. de meest recente
    meting minstens min_span jaar eerder. 'totaal': kleinste-kwadratenlijn over de hele reeks.
    'sprong': een stap van meer dan 25 mm/jaar tussen metingen die minstens 3 jaar uit elkaar
    liggen; dat is voor een gebouw onwaarschijnlijk en wijst op een stelsel- of meetfout."""
    n = len(serie)
    r = {"n": n, "recent": None, "rfrom": None, "total": None, "first": None, "last": None, "lastY": None, "sprong": False}
    if not n:
        return r
    r["first"], r["last"], r["lastY"] = serie[0][2], serie[-1][2], serie[-1][0]
    if n >= 2 and serie[-1][0] - serie[0][0] >= min_span:
        ts = [p[0] for p in serie]
        hs = [p[1] for p in serie]
        mt, mh = sum(ts) / n, sum(hs) / n
        den = sum((t - mt) ** 2 for t in ts)
        if den > 0:
            r["total"] = -1000.0 * sum((t - mt) * (h - mh) for t, h in zip(ts, hs)) / den
        t1, h1 = serie[-1][0], serie[-1][1]
        cand = [p for p in serie[:-1] if t1 - p[0] >= min_span]
        if cand:
            p = cand[-1]
            r["recent"] = -1000.0 * (h1 - p[1]) / (t1 - p[0])
            r["rfrom"] = p[0]
    for a, b in zip(serie, serie[1:]):
        dt = b[0] - a[0]
        if dt >= 1 and abs(b[1] - a[1]) * 1000.0 / dt > 25:
            r["sprong"] = True
    if r["recent"] is not None and r["recent"] < -2.0:  # meer dan 2 mm/jaar rijzen: onwaarschijnlijk
        r["sprong"] = True
    return r


def lees_verbergen(pad):
    """Meetpunten die op verzoek niet op de kaart komen: per regel een meetboutnummer,
    een peilmerknummer of een adres zoals op de kaart (bijvoorbeeld 'Teststraat 2'; dan
    verdwijnen alle meetbouten op dat adres). Alles na # is commentaar. Komt uit het bestand
    --verbergen en/of de omgevingsvariabele ZK_VERBERGEN (op GitHub een niet-openbare
    repository-variabele, zodat niet zichtbaar is welke adressen om verwijdering vroegen)."""
    regels = []
    if pad and os.path.exists(pad):
        with open(pad, encoding="utf-8") as fh:
            regels += fh.read().splitlines()
    regels += re.split(r"[\n;]", os.environ.get("ZK_VERBERGEN", ""))
    ids = set()
    for line in regels:
        t = re.sub(r"\s+", " ", line.split("#", 1)[0]).strip()
        if t:
            ids.add(t.upper())
    return ids


def build_nap(punten, hist, args):
    """Maakt de compacte gegevens voor de kaart: punten (lui geladen) en reeksen (bij de eerste klik)."""
    by_p = {}
    for m in hist:
        by_p.setdefault(m["p"], []).append(m)
    rows, reeksen, seen = [], [], set()
    stats = {"punten": 0, "metReeks": 0, "herzien": 0, "sprong": 0}
    verborgen = getattr(args, "_verbergen", set())
    for pt in punten:
        if pt["p"] in seen or pt["p"].upper() in verborgen:
            continue
        seen.add(pt["p"])
        serie, herzien, onzeker = nap_series(by_p.get(pt["p"], []))
        r = nap_analyse(serie)
        r["sprong"] = r["sprong"] or onzeker
        stats["punten"] += 1
        stats["metReeks"] += 1 if r["recent"] is not None else 0
        stats["herzien"] += 1 if herzien else 0
        stats["sprong"] += 1 if r["sprong"] else 0
        rows.append([pt["p"], pt["lat"], pt["lon"], pt["o"], pt["s"], r["n"], r["first"] or "", r["last"] or "",
                     r2(r["lastY"], 3), r2(r["recent"]), r2(r["rfrom"], 2), r2(r["total"]), 1 if r["sprong"] else 0])
        ser = None
        if len(serie) >= 2:
            h0 = serie[0][1]
            ser = []
            for t, h, _ in serie:
                ser.append(round(t - 1900, 2))
                ser.append(round((h0 - h) * 1000.0, 1))  # zakking in mm sinds de eerste meting
        reeksen.append(ser)
    log(f"   NAP: {fmt_int(stats['punten'])} peilmerken, {fmt_int(stats['metReeks'])} met een reeks van 8+ jaar, "
        f"{fmt_int(stats['herzien'])} met herziene waarden, {fmt_int(stats['sprong'])} met een verdachte sprong")
    return rows, reeksen


# --------------------------------------------------------------------------
# Pagina's per gemeente (NAP-peilmerken) en per Amsterdamse buurt (meetbouten)
# --------------------------------------------------------------------------
# Grenzen: Kadaster, Bestuurlijke Gebieden (CC BY 4.0) en gemeente Amsterdam,
# Kernregistratie gebieden (CC0). Ze worden vereenvoudigd (± 10 m) en in de cache
# bewaard; standaard eens per half jaar opnieuw opgehaald.

GEMEENTEN_URL = "https://api.pdok.nl/kadaster/brk-bestuurlijke-gebieden/ogc/v1/collections/gemeentegebied/items"
GEBIED_TOL = 0.0001  # graden; vereenvoudiging van grenzen (± 7-11 m)
GEBIED_KLASSEN = ("rood", "geel", "groen", "stab", "oud")
BOUT_KLASSE = {"zakt snel": "rood", "zakt matig": "geel", "zakt nauwelijks": "groen",
               "gestabiliseerd": "stab", "geen recente data": "oud"}
GEBIED_KLEUR = {"rood": "#d7301f", "geel": "#e6a700", "groen": "#1a9850", "stab": "#1a9850", "oud": "#9a9a9a"}


def _dp(pts, tol):
    """Douglas-Peucker (zonder recursie) op een lijst (x, y)-punten."""
    n = len(pts)
    if n < 6:
        return pts
    keep = [False] * n
    keep[0] = keep[-1] = True
    stack = [(0, n - 1)]
    tol2 = tol * tol
    while stack:
        a, b = stack.pop()
        ax, ay = pts[a]
        bx, by = pts[b]
        dx, dy = bx - ax, by - ay
        l2 = dx * dx + dy * dy
        best, bi = -1.0, -1
        for i in range(a + 1, b):
            px, py = pts[i]
            if l2 == 0:
                d = (px - ax) ** 2 + (py - ay) ** 2
            else:
                t = ((px - ax) * dx + (py - ay) * dy) / l2
                t = 0.0 if t < 0 else 1.0 if t > 1 else t
                d = (px - ax - t * dx) ** 2 + (py - ay - t * dy) ** 2
            if d > best:
                best, bi = d, i
        if bi > 0 and best > tol2:
            keep[bi] = True
            stack.append((a, bi))
            stack.append((bi, b))
    out = [pts[i] for i in range(n) if keep[i]]
    return out if len(out) >= 4 else pts


def gebied_polygonen(geom, rd=None):
    """GeoJSON Polygon/MultiPolygon -> [[ring, ...], ...] met (lon, lat)-punten, vereenvoudigd.
    rd=None: zelf herkennen of de coördinaten in RD (meters) staan."""
    if not isinstance(geom, dict):
        return []
    t, c = geom.get("type"), geom.get("coordinates")
    if t == "Polygon":
        polys = [c]
    elif t == "MultiPolygon":
        polys = c
    else:
        return []
    out = []
    for poly in polys or []:
        rings = []
        for ring in poly or []:
            try:
                pts = [(float(p[0]), float(p[1])) for p in ring]
            except (TypeError, ValueError, IndexError):
                continue
            if len(pts) < 4:
                continue
            is_rd = rd if rd is not None else abs(pts[0][0]) > 1000
            if is_rd:
                pts = [(lambda ll: (ll[1], ll[0]))(rd_to_wgs84(x, y)) for x, y in pts]
            if pts[0] != pts[-1]:
                pts.append(pts[0])
            pts = _dp(pts, GEBIED_TOL)
            rings.append([[round(x, 6), round(y, 6)] for x, y in pts])
        if rings:
            out.append(rings)
    return out


class _Ring:
    """Ring met edges per horizontale band, voor snelle punt-in-polygoon."""
    __slots__ = ("ymin", "ymax", "h", "bands")

    def __init__(self, pts):
        ys = [p[1] for p in pts]
        self.ymin, self.ymax = min(ys), max(ys)
        nb = max(1, min(256, len(pts) // 6))
        self.h = (self.ymax - self.ymin) / nb or 1e-12
        self.bands = [[] for _ in range(nb)]
        for i in range(len(pts) - 1):
            x1, y1 = pts[i]
            x2, y2 = pts[i + 1]
            if y1 == y2:
                continue
            lo, hi = (y1, y2) if y1 < y2 else (y2, y1)
            b0 = max(0, int((lo - self.ymin) / self.h))
            b1 = min(nb - 1, int((hi - self.ymin) / self.h))
            for b in range(b0, b1 + 1):
                self.bands[b].append((x1, y1, x2, y2))

    def bevat(self, x, y):
        if y < self.ymin or y > self.ymax:
            return False
        band = self.bands[min(len(self.bands) - 1, int((y - self.ymin) / self.h))]
        binnen = False
        for x1, y1, x2, y2 in band:
            if (y1 > y) != (y2 > y) and x < x1 + (y - y1) * (x2 - x1) / (y2 - y1):
                binnen = not binnen
        return binnen


class GebiedIndex:
    """Zoekt bij een punt (lat, lon) het gebied waarin het ligt."""
    CEL = 0.02

    def __init__(self, gebieden):
        self.gebieden = gebieden
        self.vorm = []
        self.cellen = {}
        for gi, g in enumerate(gebieden):
            polys = []
            xmin = ymin = float("inf")
            xmax = ymax = float("-inf")
            for rings in g["polys"]:
                outer = _Ring(rings[0])
                holes = [_Ring(r) for r in rings[1:]]
                polys.append((outer, holes))
                for x, y in rings[0]:
                    xmin, xmax, ymin, ymax = min(xmin, x), max(xmax, x), min(ymin, y), max(ymax, y)
            self.vorm.append(polys)
            g["bbox"] = [xmin, ymin, xmax, ymax]
            for cx in range(int(math.floor(xmin / self.CEL)), int(math.floor(xmax / self.CEL)) + 1):
                for cy in range(int(math.floor(ymin / self.CEL)), int(math.floor(ymax / self.CEL)) + 1):
                    self.cellen.setdefault((cx, cy), []).append(gi)

    def zoek(self, lat, lon):
        for gi in self.cellen.get((int(math.floor(lon / self.CEL)), int(math.floor(lat / self.CEL))), ()):
            for outer, holes in self.vorm[gi]:
                if outer.bevat(lon, lat) and not any(h.bevat(lon, lat) for h in holes):
                    return gi
        return None


def _eigenschap(props, *namen):
    for n in namen:
        v = props.get(n)
        if v not in (None, ""):
            return str(v).strip()
    return None


def _ogc_alle(url, params, delay=0.2, max_pages=200):
    """Alle features van een OGC API Features-collectie (volgt de 'next'-links)."""
    feats, pages = [], 0
    nxt = url + "?" + urllib.parse.urlencode(params)
    while nxt and pages < max_pages:
        data = http_get_json(nxt, retries=4)
        feats.extend((data or {}).get("features") or [])
        nxt = None
        for lk in (data or {}).get("links") or []:
            if lk.get("rel") == "next" and lk.get("href"):
                nxt = lk["href"]
        pages += 1
        if nxt and delay:
            time.sleep(delay)
    return feats


def _dso_alle(url, collection, max_pages=50):
    """Alle records van een DSO-lijst van de gemeente Amsterdam (volgt _links.next)."""
    items, pages, nxt = [], 0, url
    while nxt and pages < max_pages:
        data = http_get_json(nxt, retries=4)
        items.extend(embedded_items(data or {}, collection))
        lk = ((data or {}).get("_links") or {}).get("next")
        nxt = lk.get("href") if isinstance(lk, dict) else None
        pages += 1
    return items


def _actueel(recs):
    """Per identificatie de geldige, nieuwste versie (tijdreeks-datasets van de gemeente)."""
    vandaag = date.today().isoformat()
    beste = {}
    for r in recs:
        eind = r.get("eindGeldigheid")
        if eind and str(eind)[:10] <= vandaag:
            continue
        ident = str(r.get("identificatie") or r.get("id") or "")
        if not ident:
            continue
        try:
            vn = int(r.get("volgnummer") or 0)
        except (TypeError, ValueError):
            vn = 0
        if ident not in beste or vn >= beste[ident][0]:
            beste[ident] = (vn, r)
    return [v[1] for v in beste.values()]


def haal_gemeenten(args):
    feats = _ogc_alle(args.gemeenten_url, {"f": "json", "limit": "50"})
    per_code = {}
    for f in feats:
        p = f.get("properties") or {}
        naam = _eigenschap(p, "naam", "name", "statnaam", "gemeentenaam")
        code = _eigenschap(p, "code", "statcode", "identificatie", "gemeentecode") or naam
        prov = _eigenschap(p, "ligt_in_provincie_naam", "provincienaam", "provincie")
        polys = gebied_polygonen(f.get("geometry"))
        if not naam or not polys:
            continue
        if code in per_code:  # een gemeente in meerdere delen: samenvoegen
            per_code[code]["polys"].extend(polys)
        else:
            per_code[code] = {"naam": naam, "code": code, "prov": prov, "polys": polys}
    return list(per_code.values())


def haal_ams_buurten(args):
    root = args.api.split("meetbouten/")[0]
    wijken = {}
    for w in _actueel(_dso_alle(root + "gebieden/wijken/?_format=json&_pageSize=500", "wijken")):
        if w.get("naam"):
            wijken[str(w.get("identificatie"))] = str(w["naam"]).strip()
    out = []
    for b in _actueel(_dso_alle(root + "gebieden/buurten/?_format=json&_pageSize=1000", "buurten")):
        naam = (b.get("naam") or "").strip()
        polys = gebied_polygonen(b.get("geometrie"))
        if not naam or not polys:
            continue
        wid, _ = find_relation(b, "ligtInWijk")
        out.append({"naam": naam, "code": str(b.get("code") or b.get("identificatie")),
                    "wijk": wijken.get(wid or "", None), "polys": polys})
    return out


def laad_gebieden(cache_dir, soort, haal, args, state, minimum, maximum):
    """Grenzen uit de cache, of (als die ouder is dan --gebieden-ververs dagen) opnieuw ophalen.
    Lukt ophalen niet, dan de oude cache; zonder cache None (de pagina's worden overgeslagen)."""
    fp = os.path.join(cache_dir, soort + ".json.gz")
    sleutel = "gebieden_" + soort
    oud = None
    if os.path.exists(fp):
        try:
            with gzip.open(fp, "rt", encoding="utf-8") as fh:
                oud = json.load(fh)
        except (OSError, ValueError):
            oud = None
    age = days_since(state.get(sleutel))
    if oud and not args.vernieuw and age is not None and age < args.gebieden_ververs:
        return oud
    log(f"-> {soort}: grenzen ophalen")
    try:
        data = haal(args)
        if not (minimum <= len(data) <= maximum):
            raise RuntimeError(f"onverwacht aantal gebieden: {len(data)}")
        tmp = fp + ".tmp"
        with gzip.open(tmp, "wt", encoding="utf-8") as fh:
            json.dump(data, fh, separators=(",", ":"))
        os.replace(tmp, fp)
        state[sleutel] = date.today().isoformat()
        log(f"   klaar: {len(data)} gebieden")
        return data
    except (RuntimeError, FatalHTTP, OSError, ValueError, KeyError, TypeError) as e:
        log(f"   let op: {soort} niet opgehaald ({e})")
        if oud:
            log("   oude versie uit de cache wordt gebruikt")
            return oud
        log("   de pagina's hiervoor worden deze keer overgeslagen")
        return None


def nap_klasse(r, now_y, rood, geel, stale=20.0):
    """Zelfde indeling als op de kaart (napCls)."""
    if r[9] is None or r[12]:
        return "oud"
    if r[8] is None or now_y - r[8] > stale:
        return "oud"
    return "rood" if r[9] >= rood else "geel" if r[9] >= geel else "groen"


def _slug(s):
    import unicodedata
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii").lower()
    s = re.sub(r"[^a-z0-9]+", "-", s).strip("-")
    return s or "gebied"


def _sorteer(naam):
    """Sorteersleutel voor namen: 's-Gravenhage bij de G, hoofdletters en accenten genegeerd."""
    import unicodedata
    n = unicodedata.normalize("NFKD", naam).encode("ascii", "ignore").decode("ascii").lower()
    return re.sub(r"^('s-|'t |[^a-z0-9]+)", "", n)


_RANGTELWOORD = {"1e": "eerste", "2e": "tweede", "3e": "derde", "4e": "vierde", "5e": "vijfde"}


def _straat_sleutel(naam):
    """Sleutel om schrijfwijzen van dezelfde straat samen te voegen: hoofdletters, spaties,
    leestekens, accenten en '1e'/'Eerste' tellen niet mee."""
    import unicodedata
    t = unicodedata.normalize("NFKD", naam).encode("ascii", "ignore").decode("ascii").lower().strip()
    t = re.sub(r"^([1-5]e)\b", lambda m: _RANGTELWOORD[m.group(1)], t)
    return re.sub(r"[^a-z0-9]", "", t)


def _geen_straatnaam(naam):
    """Omschrijvingen in het adresveld die geen straatnaam zijn ('??Vervallen mei', 'Gevel aan …')."""
    return bool(not re.match(r"^[0-9A-Za-zÀ-ÿ']", naam)
                or re.search(r"\?|\bvervallen\b|\bperceel\b|\bgevel\b|\bzie\b|\bnabij\b|\btegenover\b", naam, re.I)
                or len(naam.split()) > 5)


def _nl(v, d=1):
    return "–" if v is None else f"{v:.{d}f}".replace(".", ",").replace("-", "−")


def _h(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;"))


def _mediaan(vals):
    v = sorted(x for x in vals if x is not None)
    if not v:
        return None
    m = len(v) // 2
    return v[m] if len(v) % 2 else (v[m - 1] + v[m]) / 2


class _Telling:
    """Samenvatting van een verzameling meetpunten (klassen, mediaan, laatste meting)."""

    def __init__(self):
        self.n = 0
        self.k = {k: 0 for k in GEBIED_KLASSEN}
        self.recent = []
        self.laatste = None

    def voeg(self, klasse, recent, last_y):
        self.n += 1
        self.k[klasse] += 1
        if klasse != "oud" and recent is not None:
            self.recent.append(recent)
        if last_y:
            self.laatste = max(self.laatste or 0, int(last_y))

    @property
    def met_recent(self):
        return self.n - self.k["oud"]

    @property
    def mediaan(self):
        return _mediaan(self.recent)


GEBIED_CSS = """:root{--ink:#1c1c1a;--muted:#6b6a64;--line:#e3e1d9;--accent:#24507a;--bg:#f6f5f1}
*{box-sizing:border-box}
body{margin:0;font:16px/1.55 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;color:var(--ink);background:var(--bg)}
a{color:var(--accent)}
header{background:#fff;border-bottom:1px solid var(--line)}
header .in,main,footer .in{max-width:860px;margin:0 auto;padding:0 16px}
header .in{display:flex;align-items:center;justify-content:space-between;gap:12px;padding-top:12px;padding-bottom:12px}
header .logo{font-weight:700;text-decoration:none;color:var(--ink);font-size:18px}
header .kaart{font-weight:600}
nav.kruimel{font-size:14px;color:var(--muted);margin:18px 0 4px}
nav.kruimel a{color:var(--muted)}
h1{font-size:28px;line-height:1.2;margin:6px 0 8px}
h2{font-size:20px;margin:28px 0 8px}
.lead{color:#3b3a36;margin:0 0 16px}
.cijfers{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:10px;margin:16px 0}
.cijfer{background:#fff;border:1px solid var(--line);border-radius:10px;padding:10px 12px}
.cijfer b{display:block;font-size:24px;line-height:1.2}
.cijfer span{color:var(--muted);font-size:14px}
.knop{display:inline-block;background:var(--accent);color:#fff;text-decoration:none;font-weight:600;padding:10px 16px;border-radius:8px;margin:8px 0}
table{width:100%;border-collapse:collapse;background:#fff;border:1px solid var(--line);border-radius:10px;overflow:hidden;font-size:15px}
th,td{text-align:left;padding:7px 10px;border-bottom:1px solid var(--line);vertical-align:top}
th{background:#f1efe8;font-weight:600}
td.n,th.n{text-align:right;white-space:nowrap}
.tabel{overflow-x:auto}
.balk{display:flex;height:14px;border-radius:7px;overflow:hidden;background:#eee;margin:6px 0 10px}
.balk i{display:block;height:100%}
.chip{display:inline-block;width:11px;height:11px;border-radius:50%;margin-right:6px;vertical-align:-1px}
.chip.stab{box-shadow:0 0 0 2px #08306b inset}
.uitleg{font-size:15px;color:#3b3a36}
ul.lijst{columns:2 220px;padding-left:18px}
.filter{margin:12px 0}
.anker{display:block;position:relative;top:-80px;visibility:hidden}
tr:has(.anker:target) td{background:#fff3cd}
.filter input{width:100%;max-width:420px;padding:9px 12px;border:1px solid var(--line);border-radius:8px;font:inherit;background:#fff}
footer{border-top:1px solid var(--line);background:#fff;margin-top:36px}
footer .in{font-size:14px;color:var(--muted);padding-top:14px;padding-bottom:24px}
footer p{margin:0 0 8px}
"""


# Zoekveld dat de tabellen op een pagina al typend filtert (geen cookies, geen externe scripts).
FILTER_HTML = ('<p class="filter"><input type="search" id="filter" placeholder="__PH__" aria-label="__PH__" autocomplete="off"></p>'
               "<script>(function(){var f=document.getElementById('filter');if(!f)return;"
               "function n(t){return t.toLowerCase().normalize('NFD').replace(/[\\u0300-\\u036f]/g,'');}"
               "f.addEventListener('input',function(){var q=n(f.value.trim());"
               "document.querySelectorAll('.tabel.f').forEach(function(t){var z=0;"
               "t.querySelectorAll('tbody tr').forEach(function(r){var ok=!q||n(r.textContent).indexOf(q)>=0;r.hidden=!ok;if(ok)z++;});"
               "t.hidden=!z;var h=t.previousElementSibling;if(h&&h.tagName==='H2'&&h.dataset.p!=='1'){h.hidden=!z;}});});})();</script>")


def _og_beeld(url):
    if not url:
        return []
    return [f'<meta property="og:image" content="{_h(url)}">',
            f'<meta property="og:image:width" content="{DEELBEELD_W}">',
            f'<meta property="og:image:height" content="{DEELBEELD_H}">',
            '<meta property="og:image:alt" content="Kaart van Nederland met gekleurde meetpunten: hoe snel zakken ze?">',
            '<meta name="twitter:card" content="summary_large_image">']


def _pagina(titel, beschrijving, canon, root, css, kruimel, inhoud, datum, og_url="", egms=False):
    bron_extra = "; " + _h(EGMS_BRON) if egms else ""
    k = " › ".join(f'<a href="{_h(u)}">{_h(t)}</a>' if u else _h(t) for t, u in kruimel)
    return f"""<!doctype html>
<html lang="nl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{_h(titel)}</title>
<meta name="description" content="{_h(beschrijving)}">
<link rel="canonical" href="{_h(canon)}">
<meta property="og:type" content="website">
<meta property="og:locale" content="nl_NL">
<meta property="og:title" content="{_h(titel)}">
<meta property="og:description" content="{_h(beschrijving)}">
<meta property="og:url" content="{_h(canon)}">
{chr(10).join(_og_beeld(og_url))}
<link rel="stylesheet" href="{_h(css)}">
</head>
<body>
<header><div class="in"><a class="logo" href="{root}">Zakkingskaart</a><a class="kaart" href="{root}">Naar de kaart</a></div></header>
<main>
<nav class="kruimel" aria-label="Kruimelpad">{k}</nav>
{inhoud}
</main>
<footer><div class="in">
<p>Onafhankelijk initiatief, los van gemeente en Rijkswaterstaat. Deze pagina geeft openbare meetgegevens eenvoudiger weer en geeft geen oordeel of advies over gebouwen, de grond of de waarde ervan. Aan de gegevens kunnen geen rechten worden ontleend. Vragen over de meetgegevens gaan naar de bron: zie <a href="{root}#over">Over &amp; disclaimer</a>.</p>
<p>Bronnen: Rijkswaterstaat, NAP-peilmerken (CC0); gemeente Amsterdam, meetbouten en gebiedsindeling (CC0); Kadaster, Bestuurlijke Gebieden (CC BY 4.0){bron_extra}. Bijgewerkt: {_h(datum)}.</p>
</div></footer>
</body>
</html>
"""


def _verdeling(t, rood, geel, met_stab):
    rij = {"rood": ("Zakt snel", f"{_nl(rood)} mm per jaar of meer"),
           "geel": ("Zakt matig", f"{_nl(geel)} tot {_nl(rood)} mm per jaar"),
           "groen": ("Zakt nauwelijks", f"minder dan {_nl(geel)} mm per jaar"),
           "stab": ("Gestabiliseerd", "zakte eerst flink (1,5 mm per jaar of meer), nu bijna niet meer (minder dan 0,5)"),
           "oud": ("Geen recente of te weinig metingen", "")}
    klassen = [k for k in GEBIED_KLASSEN if met_stab or k != "stab"]
    balk = "".join(f'<i style="width:{100.0 * t.k[k] / t.n:.2f}%;background:{GEBIED_KLEUR[k]}" title="{rij[k][0]}"></i>'
                   for k in klassen if t.n and t.k[k])
    rijen = "".join(
        f'<tr><td><span class="chip{" stab" if k == "stab" else ""}" style="background:{GEBIED_KLEUR[k]}"></span>'
        f'{rij[k][0]}</td><td>{rij[k][1]}</td><td class="n">{fmt_int(t.k[k])}</td>'
        f'<td class="n">{_nl(100.0 * t.k[k] / t.n, 0) if t.n else "–"}%</td></tr>' for k in klassen)
    return (f'<div class="balk" role="img" aria-label="Verdeling over de kleuren">{balk}</div>'
            f'<div class="tabel"><table><thead><tr><th>Kleur op de kaart</th><th>Snelheid de laatste jaren</th>'
            f'<th class="n">Aantal</th><th class="n">Aandeel</th></tr></thead><tbody>{rijen}</tbody></table></div>')


def _samen(ts):
    """Voegt tellingen samen (bijvoorbeeld de buurten van een wijk)."""
    t = _Telling()
    for x in ts:
        t.n += x.n
        for k in x.k:
            t.k[k] += x.k[k]
        t.recent.extend(x.recent)
        if x.laatste:
            t.laatste = max(t.laatste or 0, x.laatste)
    return t


KLASSE_LABEL = {"rood": "Zakt snel", "geel": "Zakt matig", "groen": "Zakt nauwelijks", "stab": "Gestabiliseerd",
                "oud": "Geen recente metingen"}


def _huisnummer(adres):
    m = re.search(r"\s(\d+)\s*([A-Za-z]?)", adres or "")
    return (int(m.group(1)), m.group(2).upper()) if m else (10 ** 9, "")


def _datum_nl(d):
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", str(d or ""))
    return f"{m.group(3)}-{m.group(2)}-{m.group(1)}" if m else (str(d) if d else "–")


def _cijfers(items):
    return '<div class="cijfers">' + "".join(
        f'<div class="cijfer"><b>{w}</b><span>{_h(l)}</span></div>' for w, l in items) + "</div>"


def _kaart_link(root, lat, lon, zoom, van=""):
    return f"{root}?kaart={lat:.5f},{lon:.5f},{int(zoom)}" + (f"&amp;van={van}" if van else "")


def _zoom_voor(bbox, lo=9, hi=16):
    xmin, ymin, xmax, ymax = bbox
    span = max(xmax - xmin, (ymax - ymin) * 1.6, 1e-4)
    return max(lo, min(hi, int(math.log2(360.0 * 3.5 / span))))


def write_gebieden(out_dir, site_url, root_css, args, now_y, today, nap_rows, rows, gemeenten, buurten, og_url="",
                   egms=None):
    """Schrijft /gemeenten/, /gemeente/<naam>/ en /gemeente/amsterdam/<buurt>/.
    Geeft de lijst met paden (voor de sitemap) en de gegevens voor de links vanuit de kaart terug."""
    import shutil
    for d in ("gemeente", "gemeenten"):
        p = os.path.join(out_dir, d)
        if os.path.isdir(p):
            shutil.rmtree(p)
    if not gemeenten:
        return [], None
    datum = today.isoformat()
    paden = []

    # --- gemeenten: peilmerken toewijzen ---
    gidx = GebiedIndex(gemeenten)
    g_tel = [_Telling() for _ in gemeenten]
    g_nap = [[] for _ in gemeenten]
    buiten = 0
    nap_gem = []
    for r in nap_rows or []:
        gi = gidx.zoek(r[1], r[2])
        nap_gem.append(-1 if gi is None else gi)
        if gi is None:
            buiten += 1
            continue
        g_tel[gi].voeg(nap_klasse(r, now_y, args.rood, args.geel), r[9], r[8])
        g_nap[gi].append(r)
    slugs, gezien = [], set()
    for g in gemeenten:
        s = _slug(g["naam"])
        if s in gezien:
            s = s + "-" + _slug(g["code"])
        gezien.add(s)
        slugs.append(s)
    ams = next((i for i, g in enumerate(gemeenten) if g["naam"].lower() == "amsterdam"), None)
    log(f"Pagina's: {len(gemeenten)} gemeenten, {fmt_int(len(nap_rows or []) - buiten)} peilmerken toegewezen"
        + (f" ({buiten} buiten alle gemeenten)" if buiten else ""))

    # --- Amsterdam: meetbouten en peilmerken per buurt ---
    b_tel, b_nap, b_straat, bslugs, a_tel = [], [], [], [], _Telling()
    bout_buurt = [-1] * len(rows)
    bout_k = [None] * len(rows)
    if buurten:
        bidx = GebiedIndex(buurten)
        b_tel = [_Telling() for _ in buurten]
        b_nap = [_Telling() for _ in buurten]
        b_straat = [{} for _ in buurten]
        geen = 0
        for ri, b in enumerate(rows):
            k = BOUT_KLASSE[classify(b["r"], now_y, args.rood, args.geel, 1.5, 0.5, args.max_ouderdom, b["status"])]
            a_tel.voeg(k, b["r"]["recent"], b["r"]["lastY"])
            bi = bidx.zoek(b["lat"], b["lon"])
            bout_buurt[ri] = -1 if bi is None else bi
            bout_k[ri] = k
            if bi is None:
                geen += 1
                continue
            b_tel[bi].voeg(k, b["r"]["recent"], b["r"]["lastY"])
            if b["street"]:
                sl = "" if _geen_straatnaam(b["street"]) else _straat_sleutel(b["street"])
                st = b_straat[bi].setdefault(sl, {"namen": {}, "t": _Telling()})
                st["namen"][b["street"]] = st["namen"].get(b["street"], 0) + 1
                st["t"].voeg(k, b["r"]["recent"], b["r"]["lastY"])
        if ams is not None:
            for r in g_nap[ams]:
                bi = bidx.zoek(r[1], r[2])
                if bi is not None:
                    b_nap[bi].voeg(nap_klasse(r, now_y, args.rood, args.geel), r[9], r[8])
        gezien = set()
        for b in buurten:
            s = _slug(b["naam"])
            if s in gezien or s in ("index",):
                s = s + "-" + _slug(b["code"])
            gezien.add(s)
            bslugs.append(s)
        log(f"         {len(buurten)} Amsterdamse buurten, {fmt_int(len(rows) - geen)} meetbouten toegewezen"
            + (f" ({geen} buiten alle buurten)" if geen else ""))
    g_egms = egms_per_gebied(egms, gemeenten, args.rood, args.geel) if egms else None
    b_egms = egms_per_gebied(egms, buurten, args.rood, args.geel) if (egms and buurten) else None
    release = egms[1].get("release", "") if egms else ""
    toon_buurt = [b_tel[i].n >= 3 or b_nap[i].met_recent >= 1 for i in range(len(buurten or []))]

    # --- Amsterdam: wijken (samengesteld uit de buurten) en straten (over buurtgrenzen heen) ---
    wijken, wslugs, buurt_wijk = [], [], []
    straten_a, bout_straat = [], [-1] * len(rows)
    straat_van_sleutel = {}
    if ams is not None and buurten:
        wmap = {}
        for bi, b in enumerate(buurten):
            w = (b.get("wijk") or "").strip()
            if not w:
                buurt_wijk.append(-1)
                continue
            if w not in wmap:
                wmap[w] = len(wijken)
                wijken.append({"naam": w, "buurten": []})
            wijken[wmap[w]]["buurten"].append(bi)
            buurt_wijk.append(wmap[w])
        gw = set()
        for w in wijken:
            sl = base = _slug(w["naam"])
            n = 2
            while sl in gw:
                sl, n = f"{base}-{n}", n + 1
            gw.add(sl)
            wslugs.append(sl)
        per_sl = {}
        for ri, b in enumerate(rows):
            if not b["street"] or _geen_straatnaam(b["street"]):
                continue
            st = per_sl.setdefault(_straat_sleutel(b["street"]), {"namen": {}, "rij": []})
            st["namen"][b["street"]] = st["namen"].get(b["street"], 0) + 1
            st["rij"].append(ri)
        tmp = []
        for sl, st in per_sl.items():
            naam = max(st["namen"].items(), key=lambda kv: (kv[1], sum(c.isupper() for c in kv[0]), kv[0]))[0]
            tmp.append((naam, sl, st))
        tmp.sort(key=lambda x: _sorteer(x[0]))
        gs = {"index"}
        for naam, sl, st in tmp:
            sg = base = _slug(naam)
            n = 2
            while sg in gs:
                sg, n = f"{base}-{n}", n + 1
            gs.add(sg)
            si = len(straten_a)
            straten_a.append({"naam": naam, "slug": sg, "rij": st["rij"], "namen": st["namen"]})
            straat_van_sleutel[sl] = si
            for ri in st["rij"]:
                bout_straat[ri] = si

    def egms_blok(t, waar):
        if t is None:
            return ""
        if not t.n:
            return (f"<h2>Bodemdaling gemeten vanuit de ruimte</h2><p>Voor {waar} zijn geen satellietmetingen "
                    f"beschikbaar.</p>")
        return (f"<h2>Bodemdaling gemeten vanuit de ruimte</h2>"
                f"<p>Satellieten van de Europese Unie meten de hoogte van de grond. In {waar} gaat het om {fmt_int(t.n)} "
                f"blokken van 100 bij 100 meter. Doorsnee zakt de grond daar {_nl(t.mediaan, 1)} mm per jaar "
                f"(gemeten {_h(release.replace('-', '–'))}).</p>"
                + _verdeling(t, args.rood, args.geel, False).replace("Geen recente of te weinig metingen", "Onzekere meting")
                .replace("<th>Snelheid de laatste jaren</th>", f"<th>Snelheid {_h(release.replace('-', '–'))}</th>")
                + "<p class=\"uitleg\">Een satelliet meet grond, straten en gebouwen samen; per blok van 100 meter is dat "
                  "een gemiddelde. Een huis op palen kan stil staan terwijl de straat ervoor zakt, en andersom. Alleen waar "
                  "iets het radarsignaal goed terugkaatst (daken, wegen) is gemeten; in weilanden en op water niet. "
                  "Bron: Copernicus, het aardobservatieprogramma van de EU (" + _h(EGMS_BRON) + ").</p>")

    def schrijf(rel, html):
        p = os.path.join(out_dir, rel, "index.html")
        os.makedirs(os.path.dirname(p), exist_ok=True)
        write_text(p, html)
        paden.append(rel + "/")

    def napzin(t, waar):
        if not t.n:
            return f"<p>In {waar} liggen geen peilmerken van Rijkswaterstaat.</p>"
        s = (f"<p>In {waar} liggen {fmt_int(t.n)} peilmerken van Rijkswaterstaat. "
             f"Van {fmt_int(t.met_recent)} daarvan is bekend hoe snel ze de laatste jaren zakken")
        if t.met_recent:
            s += f"; doorsnee is dat {_nl(t.mediaan, 2)} mm per jaar"
        return s + ".</p>"

    uitleg_nap = ("<p class=\"uitleg\">Een peilmerk (NAP-peilmerk) is een metalen boutje in een huis, brug of ander bouwwerk. "
                  "Rijkswaterstaat meet eens per 5 tot 15 jaar hoe hoog het ligt. De snelheid van de laatste jaren is berekend "
                  "uit metingen die minstens 8 jaar uit elkaar liggen. Peilmerken die meer dan 20 jaar niet zijn gemeten, of "
                  "waarvan de metingen een onverklaarbare sprong maken, tellen als &lsquo;geen recente of te weinig metingen&rsquo;. "
                  "Een peilmerk zakt mee met de grond van de hele omgeving; het zegt iets over de plek, niet per se over één gebouw. "
                  "<b>Doorsnee</b> is de middelste waarde: de helft zakt sneller, de helft langzamer.</p>")

    # --- pagina per gemeente ---
    per_prov = {}
    for gi, g in enumerate(gemeenten):
        per_prov.setdefault(g.get("prov") or "", []).append(gi)
    for gi, g in enumerate(gemeenten):
        t = g_tel[gi]
        root = "../../"
        naam = g["naam"]
        bb = g["bbox"]
        lat, lon = (bb[1] + bb[3]) / 2, (bb[0] + bb[2]) / 2
        delen = [f'<h1>Zakking in {_h(naam)}</h1>',
                 f'<p class="lead">Hoe snel zakken gebouwen en andere bouwwerken in {_h(naam)}? Deze pagina vat de '
                 f'openbare metingen van Rijkswaterstaat samen'
                 + (", en de meetbouten van de gemeente Amsterdam per buurt" if gi == ams and buurten else "") + ".</p>",
                 _cijfers([(fmt_int(t.n), "peilmerken"),
                           (fmt_int(t.met_recent), "met een recente snelheid"),
                           (_nl(t.mediaan, 2) + (" mm/jaar" if t.mediaan is not None else ""), "doorsnee snelheid"),
                           (str(t.laatste or "–"), "laatste meting")]),
                 f'<a class="knop" href="{_kaart_link(root, lat, lon, _zoom_voor(bb), "gemeente/" + slugs[gi] + "/")}">Bekijk {_h(naam)} op de kaart</a>',
                 "<h2>Peilmerken van Rijkswaterstaat</h2>", napzin(t, _h(naam))]
        if t.n:
            delen.append(_verdeling(t, args.rood, args.geel, False))
        delen.append(uitleg_nap)
        if g_egms:
            delen.append(egms_blok(g_egms[gi], _h(naam)))
        if gi == ams and buurten:
            delen.append("<h2>Meetbouten per buurt</h2>")
            delen.append(f"<p>In Amsterdam meet de gemeente daarnaast de hoogte van {fmt_int(a_tel.n)} meetbouten in gevels. "
                         f"Van {fmt_int(a_tel.met_recent)} daarvan is een recente snelheid bekend; doorsnee is dat "
                         f"{_nl(a_tel.mediaan, 2)} mm per jaar.</p>")
            delen.append(_verdeling(a_tel, args.rood, args.geel, True))
            if straten_a:
                delen.append(f'<p><a class="knop" href="straat/">Zoek je straat: alle {fmt_int(len(straten_a))} straten met meetbouten</a></p>')
            if wijken:
                delen.append("<h2>Per wijk</h2>")
                wr = []
                for wi in sorted(range(len(wijken)), key=lambda i: _sorteer(wijken[i]["naam"])):
                    wt = _samen(b_tel[bi] for bi in wijken[wi]["buurten"])
                    wr.append(f'<tr><td><a href="wijk/{wslugs[wi]}/">{_h(wijken[wi]["naam"])}</a></td>'
                              f'<td class="n">{len(wijken[wi]["buurten"])}</td><td class="n">{fmt_int(wt.n)}</td>'
                              f'<td class="n">{_nl(wt.mediaan, 2)}</td><td class="n">{wt.laatste or "–"}</td></tr>')
                delen.append('<div class="tabel"><table><thead><tr><th>Wijk</th><th class="n">Buurten</th>'
                             '<th class="n">Meetbouten</th><th class="n">Doorsnee (mm/jaar)</th><th class="n">Laatste meting</th>'
                             '</tr></thead><tbody>' + "".join(wr) + "</tbody></table></div>")
                delen.append("<h2>Per buurt</h2>")
            rijen = []
            for bi in sorted((i for i in range(len(buurten)) if toon_buurt[i]),
                             key=lambda i: (_sorteer(buurten[i].get("wijk") or ""), _sorteer(buurten[i]["naam"]))):
                bt = b_tel[bi]
                rijen.append(f'<tr><td><a href="{_h(bslugs[bi])}/">{_h(buurten[bi]["naam"])}</a></td>'
                             + (f'<td><a href="wijk/{wslugs[buurt_wijk[bi]]}/">{_h(buurten[bi].get("wijk") or "")}</a></td>'
                                if buurt_wijk and buurt_wijk[bi] >= 0 else "<td></td>")
                             + f'<td class="n">{fmt_int(bt.n)}</td>'
                             f'<td class="n">{_nl(bt.mediaan, 2)}</td><td class="n">{bt.laatste or "–"}</td></tr>')
            delen.append(FILTER_HTML.replace("__PH__", "Zoek een buurt of wijk"))
            delen.append('<div class="tabel f"><table><thead><tr><th>Buurt</th><th>Wijk</th><th class="n">Meetbouten</th>'
                         '<th class="n">Doorsnee (mm/jaar)</th><th class="n">Laatste meting</th></tr></thead><tbody>'
                         + "".join(rijen) + "</tbody></table></div>")
        andere = [i for i in per_prov.get(g.get("prov") or "", []) if i != gi]
        if g.get("prov") and andere:
            delen.append(f"<h2>Andere gemeenten in {_h(g['prov'])}</h2><ul class=\"lijst\">" + "".join(
                f'<li><a href="../{slugs[i]}/">{_h(gemeenten[i]["naam"])}</a></li>'
                for i in sorted(andere, key=lambda i: _sorteer(gemeenten[i]["naam"]))) + "</ul>")
        delen.append('<p><a href="../../gemeenten/">Alle gemeenten</a></p>')
        besch = (f"Openbare NAP-metingen van Rijkswaterstaat in {naam}: {fmt_int(t.n)} peilmerken, waarvan "
                 f"{fmt_int(t.met_recent)} met een recente zakkingssnelheid"
                 + (f" (doorsnee {_nl(t.mediaan, 2)} mm per jaar)" if t.mediaan is not None else "") + ".")
        schrijf(f"gemeente/{slugs[gi]}", _pagina(
            f"Zakking en bodemdaling in {naam} | Zakkingskaart", besch, f"{site_url}gemeente/{slugs[gi]}/", root,
            root + root_css, [("Zakkingskaart", root), ("Gemeenten", "../../gemeenten/"), (naam, None)],
            "\n".join(delen), datum, og_url, bool(egms)))

    # --- pagina per Amsterdamse buurt ---
    if ams is not None and buurten:
        aslug = slugs[ams]
        for bi, b in enumerate(buurten):
            if not toon_buurt[bi]:
                continue
            t, tn = b_tel[bi], b_nap[bi]
            root = "../../../"
            bb = b["bbox"]
            lat, lon = (bb[1] + bb[3]) / 2, (bb[0] + bb[2]) / 2
            waar = f"{_h(b['naam'])}" + (f" ({_h(b['wijk'])})" if b.get("wijk") else "")
            delen = [f"<h1>Zakking in {_h(b['naam'])}, Amsterdam</h1>",
                     f'<p class="lead">Hoe snel zakken de gevels in de buurt {waar}? Deze pagina vat de openbare '
                     f'meetbouten van de gemeente Amsterdam en de peilmerken van Rijkswaterstaat samen.</p>',
                     _cijfers([(fmt_int(t.n), "meetbouten"),
                               (fmt_int(t.met_recent), "met een recente snelheid"),
                               (_nl(t.mediaan, 2) + (" mm/jaar" if t.mediaan is not None else ""),
                                "doorsnee snelheid"),
                               (str(t.laatste or "–"), "laatste meting")]),
                     f'<a class="knop" href="{_kaart_link(root, lat, lon, _zoom_voor(bb, 14, 16), "gemeente/" + aslug + "/" + bslugs[bi] + "/")}">Bekijk deze buurt op de kaart</a>',
                     "<h2>Meetbouten</h2>"]
            if t.n:
                delen.append(f"<p>In deze buurt meet de gemeente de hoogte van {fmt_int(t.n)} meetbouten in gevels. "
                             f"Van {fmt_int(t.met_recent)} daarvan is een recente zakkingssnelheid bekend"
                             + (f"; doorsnee is dat {_nl(t.mediaan, 2)} mm per jaar" if t.met_recent else "") + ".</p>")
                delen.append(_verdeling(t, args.rood, args.geel, True))
                straten = []
                for sl, st in b_straat[bi].items():
                    # weergavenaam: de meest gebruikte schrijfwijze (bij gelijkspel die met hoofdletters)
                    naam = max(st["namen"].items(), key=lambda kv: (kv[1], sum(c.isupper() for c in kv[0]), kv[0]))[0]
                    ankers = "" if not sl else "".join(sorted({f'<span class="anker" id="straat-{_slug(v)}"></span>' for v in st["namen"]}))
                    si = straat_van_sleutel.get(sl, -1) if sl else -1
                    if si >= 0:
                        ankers += f'<a href="../straat/{straten_a[si]["slug"]}/">'
                    straten.append((naam if sl else "Overige meetbouten (adres niet eenduidig)", st["t"], sl == "", ankers,
                                    "</a>" if si >= 0 else ""))
                straten.sort(key=lambda x: (x[2], _sorteer(x[0])))
                if straten:
                    delen.append("<h2>Per straat</h2>")
                    delen.append('<div class="tabel"><table><thead><tr><th>Straat</th><th class="n">Meetbouten</th>'
                                 '<th class="n">Met recente snelheid</th><th class="n">Doorsnee (mm/jaar)</th>'
                                 '<th class="n">Laatste meting</th></tr></thead><tbody>' + "".join(
                                     f'<tr><td>{ankers}{"<i>" + _h(s) + "</i>" if overig else _h(s)}{eind}</td><td class="n">{fmt_int(st.n)}</td>'
                                     f'<td class="n">{fmt_int(st.met_recent)}</td><td class="n">{_nl(st.mediaan, 2)}</td>'
                                     f'<td class="n">{st.laatste or "–"}</td></tr>' for s, st, overig, ankers, eind in straten)
                                 + "</tbody></table></div>")
            else:
                delen.append("<p>In deze buurt staan geen meetbouten van de gemeente.</p>")
            delen.append("<p class=\"uitleg\">Meetbouten zijn meetpunten in gevels waarvan de gemeente periodiek de hoogte "
                         "meet. De recente snelheid is berekend uit metingen die minstens 4 jaar uit elkaar liggen. "
                         "Een meetbout zegt iets over die ene gevel; de kaart en deze pagina geven alleen de gemeten "
                         "snelheden weer. <b>Doorsnee</b> is de middelste waarde: de helft zakt sneller, de helft langzamer.</p>")
            delen.append("<h2>Peilmerken van Rijkswaterstaat</h2>" + napzin(tn, "deze buurt"))
            if b_egms:
                delen.append(egms_blok(b_egms[bi], "deze buurt"))
            wi = buurt_wijk[bi] if buurt_wijk else -1
            delen.append("<p>" + (f'<a href="../wijk/{wslugs[wi]}/">Wijk {_h(wijken[wi]["naam"])}</a> · ' if wi >= 0 else "")
                         + '<a href="../">Alle buurten in Amsterdam</a> · <a href="../straat/">Alle straten</a> · '
                           '<a href="../../../gemeenten/">Alle gemeenten</a></p>')
            besch = (f"Openbare meetbouten van de gemeente Amsterdam in {b['naam']}: {fmt_int(t.n)} meetbouten, waarvan "
                     f"{fmt_int(t.met_recent)} met een recente zakkingssnelheid"
                     + (f" (doorsnee {_nl(t.mediaan, 2)} mm per jaar)" if t.mediaan is not None else "") + ".")
            schrijf(f"gemeente/{aslug}/{bslugs[bi]}", _pagina(
                f"Zakking in {b['naam']}, Amsterdam | Zakkingskaart", besch,
                f"{site_url}gemeente/{aslug}/{bslugs[bi]}/", root, root + root_css,
                [("Zakkingskaart", root), ("Gemeenten", "../../../gemeenten/"), ("Amsterdam", "../")]
                + ([("Wijk " + wijken[buurt_wijk[bi]]["naam"], f"../wijk/{wslugs[buurt_wijk[bi]]}/")]
                   if buurt_wijk and buurt_wijk[bi] >= 0 else []) + [(b["naam"], None)],
                "\n".join(delen), datum, og_url, bool(egms)))

    if ams is not None and buurten:
        aslug = slugs[ams]
        uitleg_bout = ("<p class=\"uitleg\">Meetbouten zijn meetpunten in gevels waarvan de gemeente periodiek de hoogte meet. "
                       "De snelheid van de laatste jaren is berekend uit metingen die minstens 4 jaar uit elkaar liggen. Een "
                       "meetbout zegt iets over die ene gevel; deze pagina geeft alleen de gemeten snelheden weer. "
                       "<b>Doorsnee</b> is de middelste waarde: de helft zakt sneller, de helft langzamer.</p>")

        # --- pagina per wijk ---
        for wi, w in enumerate(wijken):
            bis = w["buurten"]
            t = _samen(b_tel[bi] for bi in bis)
            tn = _samen(b_nap[bi] for bi in bis)
            root = "../../../../"
            bbs = [buurten[bi]["bbox"] for bi in bis]
            bb = [min(x[0] for x in bbs), min(x[1] for x in bbs), max(x[2] for x in bbs), max(x[3] for x in bbs)]
            lat, lon = (bb[1] + bb[3]) / 2, (bb[0] + bb[2]) / 2
            rel = f"gemeente/{aslug}/wijk/{wslugs[wi]}"
            delen = [f"<h1>Zakking in wijk {_h(w['naam'])}, Amsterdam</h1>",
                     f'<p class="lead">Hoe snel zakken de gevels in de wijk {_h(w["naam"])}? Deze pagina vat de openbare '
                     f'meetbouten van de gemeente Amsterdam en de peilmerken van Rijkswaterstaat samen, voor de '
                     f'{len(bis)} buurten in deze wijk.</p>',
                     _cijfers([(fmt_int(t.n), "meetbouten"), (fmt_int(t.met_recent), "met een recente snelheid"),
                               (_nl(t.mediaan, 2) + (" mm/jaar" if t.mediaan is not None else ""), "doorsnee snelheid"),
                               (str(t.laatste or "–"), "laatste meting")]),
                     f'<a class="knop" href="{_kaart_link(root, lat, lon, _zoom_voor(bb, 13, 16), rel + "/")}">Bekijk deze wijk op de kaart</a>',
                     "<h2>Meetbouten</h2>"]
            if t.n:
                delen.append(_verdeling(t, args.rood, args.geel, True))
            br = []
            for bi in sorted(bis, key=lambda i: _sorteer(buurten[i]["naam"])):
                bt = b_tel[bi]
                nm = (f'<a href="../../{bslugs[bi]}/">{_h(buurten[bi]["naam"])}</a>' if toon_buurt[bi]
                      else _h(buurten[bi]["naam"]))
                br.append(f'<tr><td>{nm}</td><td class="n">{fmt_int(bt.n)}</td><td class="n">{_nl(bt.mediaan, 2)}</td>'
                          f'<td class="n">{bt.laatste or "–"}</td></tr>')
            delen.append("<h2>Per buurt</h2>" + '<div class="tabel"><table><thead><tr><th>Buurt</th><th class="n">Meetbouten</th>'
                         '<th class="n">Doorsnee (mm/jaar)</th><th class="n">Laatste meting</th></tr></thead><tbody>'
                         + "".join(br) + "</tbody></table></div>")
            sis = sorted({bout_straat[ri] for ri in range(len(rows)) if bout_straat[ri] >= 0 and bout_buurt[ri] in set(bis)},
                         key=lambda i: _sorteer(straten_a[i]["naam"]))
            if sis:
                delen.append("<h2>Straten in deze wijk</h2><ul class=\"lijst\">" + "".join(
                    f'<li><a href="../../straat/{straten_a[i]["slug"]}/">{_h(straten_a[i]["naam"])}</a></li>' for i in sis) + "</ul>")
            delen.append(uitleg_bout)
            delen.append("<h2>Peilmerken van Rijkswaterstaat</h2>" + napzin(tn, "deze wijk"))
            if b_egms:
                delen.append(egms_blok(_samen(b_egms[bi] for bi in bis), "deze wijk"))
            delen.append('<p><a href="../../">Amsterdam</a> · <a href="../../straat/">Alle straten</a> · '
                         '<a href="../../../../gemeenten/">Alle gemeenten</a></p>')
            besch = (f"Openbare meetbouten van de gemeente Amsterdam in de wijk {w['naam']}: {fmt_int(t.n)} meetbouten in "
                     f"{len(bis)} buurten" + (f", doorsnee {_nl(t.mediaan, 2)} mm per jaar" if t.mediaan is not None else "") + ".")
            schrijf(rel, _pagina(
                f"Zakking in wijk {w['naam']}, Amsterdam | Zakkingskaart", besch, f"{site_url}{rel}/", root, root + root_css,
                [("Zakkingskaart", root), ("Gemeenten", root + "gemeenten/"), ("Amsterdam", "../../"),
                 ("Wijk " + w["naam"], None)], "\n".join(delen), datum, og_url, bool(egms)))

        # --- pagina per straat ---
        for si, st in enumerate(straten_a):
            rij = sorted(st["rij"], key=lambda ri: (_huisnummer(rows[ri]["addr"]), rows[ri]["addr"] or "", rows[ri]["id"]))
            t = _Telling()
            for ri in rij:
                t.voeg(bout_k[ri] or "oud", rows[ri]["r"]["recent"], rows[ri]["r"]["lastY"])
            root = "../../../../"
            rel = f"gemeente/{aslug}/straat/{st['slug']}"
            la = [rows[ri]["lat"] for ri in rij]
            lo = [rows[ri]["lon"] for ri in rij]
            bb = [min(lo), min(la), max(lo), max(la)]
            lat, lon = (bb[1] + bb[3]) / 2, (bb[0] + bb[2]) / 2
            bcount = {}
            for ri in rij:
                if bout_buurt[ri] >= 0:
                    bcount[bout_buurt[ri]] = bcount.get(bout_buurt[ri], 0) + 1
            bl = []
            for bi, _n in sorted(bcount.items(), key=lambda kv: -kv[1]):
                nm = (f'<a href="../../{bslugs[bi]}/">{_h(buurten[bi]["naam"])}</a>' if toon_buurt[bi] else _h(buurten[bi]["naam"]))
                wi = buurt_wijk[bi]
                bl.append(nm + (f' (wijk <a href="../../wijk/{wslugs[wi]}/">{_h(wijken[wi]["naam"])}</a>)' if wi >= 0 else ""))
            naam = st["naam"]
            van = f"&amp;van={rel}/"
            tr = []
            for ri in rij:
                b = rows[ri]
                k = bout_k[ri] or "oud"
                r = b["r"]
                rec = _nl(r["recent"], 2) if (k != "oud" and r["recent"] is not None) else "–"
                tr.append(f'<tr><td>{_h(b["addr"] or b["id"])}</td>'
                          f'<td><span class="chip{" stab" if k == "stab" else ""}" style="background:{GEBIED_KLEUR[k]}"></span>{KLASSE_LABEL[k]}</td>'
                          f'<td class="n">{rec}</td><td class="n">{_nl(r["total"], 2) if r["total"] is not None else "–"}</td>'
                          f'<td class="n">{_datum_nl(r["last"])}</td>'
                          f'<td><a href="{root}?punt={_h(str(b["id"]))}{van}">kaart</a></td></tr>')
            delen = [f"<h1>Zakking {_h(naam)}, Amsterdam</h1>",
                     f'<p class="lead">De gemeente Amsterdam meet in de {_h(naam)} de hoogte van {fmt_int(t.n)} '
                     f'meetbout{"en" if t.n != 1 else ""} in gevels. Hieronder staat per huisnummer hoe snel die de laatste jaren zakken.</p>',
                     _cijfers([(fmt_int(t.n), "meetbouten"), (fmt_int(t.met_recent), "met een recente snelheid"),
                               (_nl(t.mediaan, 2) + (" mm/jaar" if t.mediaan is not None else ""), "doorsnee snelheid"),
                               (str(t.laatste or "–"), "laatste meting")]),
                     f'<a class="knop" href="{_kaart_link(root, lat, lon, _zoom_voor(bb, 15, 18), rel + "/")}">Bekijk de {_h(naam)} op de kaart</a>']
            if bl:
                delen.append("<p>Ligt in: " + "; ".join(bl) + ".</p>")
            if t.n >= 3:
                delen.append(_verdeling(t, args.rood, args.geel, True))
            delen.append("<h2>Per huisnummer</h2>" + '<div class="tabel"><table><thead><tr><th>Adres</th><th>Kleur op de kaart</th>'
                         '<th class="n">Laatste jaren (mm/jaar)</th><th class="n">Sinds 1e meting (mm/jaar)</th>'
                         '<th class="n">Laatste meting</th><th></th></tr></thead><tbody>' + "".join(tr) + "</tbody></table></div>")
            delen.append(uitleg_bout)
            delen.append('<p><a href="../">Alle straten in Amsterdam</a> · <a href="../../">Amsterdam</a> · '
                         '<a href="../../../../gemeenten/">Alle gemeenten</a></p>')
            besch = (f"Zakking in de {naam} (Amsterdam): {fmt_int(t.n)} meetbouten van de gemeente per huisnummer"
                     + (f", doorsnee {_nl(t.mediaan, 2)} mm per jaar" if t.mediaan is not None else "")
                     + (f", laatste meting {t.laatste}" if t.laatste else "") + ".")
            schrijf(rel, _pagina(
                f"Zakking {naam}, Amsterdam: meetbouten per huisnummer | Zakkingskaart", besch, f"{site_url}{rel}/", root,
                root + root_css, [("Zakkingskaart", root), ("Gemeenten", root + "gemeenten/"), ("Amsterdam", "../../"),
                                  ("Straten", "../"), (naam, None)], "\n".join(delen), datum, og_url, False))

        # --- overzicht van alle straten ---
        if straten_a:
            root = "../../../"
            per_letter = {}
            for si, st in enumerate(straten_a):
                l0 = _sorteer(st["naam"])[:1].upper()
                per_letter.setdefault(l0 if l0.isalpha() else "#", []).append(si)
            delen = ["<h1>Zakking per straat in Amsterdam</h1>",
                     f'<p class="lead">Alle {fmt_int(len(straten_a))} straten in Amsterdam met meetbouten van de gemeente. '
                     'Klik op een straat voor de meetbouten per huisnummer.</p>',
                     FILTER_HTML.replace("__PH__", "Zoek je straat")]
            for l0 in sorted(per_letter):
                delen.append(f"<h2>{_h(l0)}</h2>")
                rr = []
                for si in per_letter[l0]:
                    st = straten_a[si]
                    t = _Telling()
                    for ri in st["rij"]:
                        t.voeg(bout_k[ri] or "oud", rows[ri]["r"]["recent"], rows[ri]["r"]["lastY"])
                    rr.append(f'<tr><td><a href="{st["slug"]}/">{_h(st["naam"])}</a></td><td class="n">{fmt_int(t.n)}</td>'
                              f'<td class="n">{_nl(t.mediaan, 2)}</td><td class="n">{t.laatste or "–"}</td></tr>')
                delen.append('<div class="tabel f"><table><thead><tr><th>Straat</th><th class="n">Meetbouten</th>'
                             '<th class="n">Doorsnee (mm/jaar)</th><th class="n">Laatste meting</th></tr></thead><tbody>'
                             + "".join(rr) + "</tbody></table></div>")
            schrijf(f"gemeente/{aslug}/straat", _pagina(
                "Zakking per straat in Amsterdam | Zakkingskaart",
                f"Alle {len(straten_a)} straten in Amsterdam met meetbouten van de gemeente: zakking per huisnummer.",
                f"{site_url}gemeente/{aslug}/straat/", root, root + root_css,
                [("Zakkingskaart", root), ("Gemeenten", root + "gemeenten/"), ("Amsterdam", "../"), ("Straten", None)],
                "\n".join(delen), datum, og_url, False))

    # --- overzicht van alle gemeenten ---
    root = "../"
    groepen = sorted(per_prov.items(), key=lambda kv: (kv[0] == "", kv[0].lower()))
    delen = ["<h1>Zakking per gemeente</h1>",
             '<p class="lead">Per gemeente een samenvatting van de openbare metingen van Rijkswaterstaat: hoeveel '
             'peilmerken er zijn, van hoeveel bekend is hoe snel ze de laatste jaren zakken, en de doorsnee snelheid '
             '(de middelste waarde: de helft zakt sneller, de helft langzamer). '
             'Voor Amsterdam is er ook een overzicht per buurt, op basis van de meetbouten van de gemeente.</p>',
             FILTER_HTML.replace("__PH__", "Zoek je gemeente")]
    for prov, idxs in groepen:
        if prov:
            delen.append(f"<h2>{_h(prov)}</h2>")
        delen.append('<div class="tabel f"><table><thead><tr><th>Gemeente</th><th class="n">Peilmerken</th>'
                     '<th class="n">Met recente snelheid</th><th class="n">Doorsnee (mm/jaar)</th></tr></thead><tbody>'
                     + "".join(f'<tr><td><a href="../gemeente/{slugs[i]}/">{_h(gemeenten[i]["naam"])}</a></td>'
                               f'<td class="n">{fmt_int(g_tel[i].n)}</td><td class="n">{fmt_int(g_tel[i].met_recent)}</td>'
                               f'<td class="n">{_nl(g_tel[i].mediaan, 2)}</td></tr>'
                               for i in sorted(idxs, key=lambda i: _sorteer(gemeenten[i]["naam"])))
                     + "</tbody></table></div>")
    schrijf("gemeenten", _pagina(
        "Zakking en bodemdaling per gemeente | Zakkingskaart",
        "Per gemeente in Nederland een samenvatting van de openbare NAP-metingen van Rijkswaterstaat: aantal peilmerken "
        "en de recente zakkingssnelheid.", f"{site_url}gemeenten/", root, root + root_css,
        [("Zakkingskaart", root), ("Gemeenten", None)], "\n".join(delen), datum, og_url, bool(egms)))
    write_text(os.path.join(out_dir, root_css.split("?")[0]), GEBIED_CSS)
    log(f"         {len(paden)} pagina's geschreven")
    nav = {"g": [[slugs[i], g["naam"], str(g.get("code") or "")] for i, g in enumerate(gemeenten)],
           "a": -1 if ams is None else ams,
           "b": [[bslugs[i], b["naam"]] if toon_buurt[i] else None for i, b in enumerate(buurten or [])],
           "bout": bout_buurt, "nap": nap_gem,
           "w": [[wslugs[i], w["naam"]] for i, w in enumerate(wijken)],
           "bw": buurt_wijk,
           "s": [[st["slug"], st["naam"]] for st in straten_a], "st": bout_straat}
    return paden, nav


# --------------------------------------------------------------------------
# Voorbeeldplaatje voor gedeelde links (og:image), zonder extra pakketten
# --------------------------------------------------------------------------
DEELBEELD = "deelbeeld.png"
DEELBEELD_W, DEELBEELD_H = 1200, 630


def _png(path, w, h, pix):
    """Schrijft een RGB-PNG. pix is een bytearray van w*h*3."""
    import zlib
    import struct

    def chunk(tag, data):
        c = struct.pack(">I", len(data)) + tag + data
        return c + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
    raw = bytearray()
    stride = w * 3
    for y in range(h):
        raw.append(0)
        raw.extend(pix[y * stride:(y + 1) * stride])
    data = (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(bytes(raw), 9)) + chunk(b"IEND", b""))
    tmp = path + ".tmp"
    with open(tmp, "wb") as fh:
        fh.write(data)
    os.replace(tmp, path)


def _rgb(hexkleur):
    return tuple(int(hexkleur[i:i + 2], 16) for i in (1, 3, 5))


def maak_deelbeeld(path, nap_rows, rows, gemeenten, now_y, args):
    """Kaart van Nederland (gemeenten als ondergrond) met de gekleurde meetpunten:
    NAP-peilmerken en de Amsterdamse meetbouten met een recente snelheid."""
    w, h = DEELBEELD_W, DEELBEELD_H
    bg, land, rand = _rgb("#e4e8ec"), _rgb("#fbfaf7"), _rgb("#cfccc3")
    pix = bytearray(bytes(bg) * (w * h))
    lon0, lon1, lat0, lat1 = 3.25, 7.30, 50.70, 53.60
    kx = math.cos(math.radians(52.2))
    sc = min((h - 40) / (lat1 - lat0), (w - 40) / ((lon1 - lon0) * kx))
    ox = (w - (lon1 - lon0) * kx * sc) / 2
    oy = (h - (lat1 - lat0) * sc) / 2

    def xy(lon, lat):
        return ox + (lon - lon0) * kx * sc, oy + (lat1 - lat) * sc

    def zet(x, y, c):
        if 0 <= x < w and 0 <= y < h:
            i = (y * w + x) * 3
            pix[i:i + 3] = bytes(c)

    # ondergrond: gemeenten gevuld (scanline, even-oneven), met grenzen
    for g in gemeenten or []:
        for rings in g["polys"]:
            pts_r = [[xy(x, y) for x, y in r] for r in rings]
            ys = [p[1] for r in pts_r for p in r]
            for yy in range(max(0, int(min(ys))), min(h, int(max(ys)) + 1)):
                yc = yy + 0.5
                xs = []
                for r in pts_r:
                    for i in range(len(r) - 1):
                        (x1, y1), (x2, y2) = r[i], r[i + 1]
                        if (y1 > yc) != (y2 > yc):
                            xs.append(x1 + (yc - y1) * (x2 - x1) / (y2 - y1))
                xs.sort()
                for a, b in zip(xs[0::2], xs[1::2]):
                    xa, xb = max(0, int(a + 0.5)), min(w, int(b + 0.5))
                    if xb > xa:
                        pix[(yy * w + xa) * 3:(yy * w + xb) * 3] = bytes(land) * (xb - xa)
            for r in pts_r:
                for i in range(len(r) - 1):
                    (x1, y1), (x2, y2) = r[i], r[i + 1]
                    n = int(max(abs(x2 - x1), abs(y2 - y1))) + 1
                    for k in range(n + 1):
                        t = k / n
                        zet(int(x1 + t * (x2 - x1)), int(y1 + t * (y2 - y1)), rand)

    # meetpunten: eerst groen, dan geel, dan rood bovenop
    punten = {"groen": [], "geel": [], "rood": []}
    for r in nap_rows or []:
        k = nap_klasse(r, now_y, args.rood, args.geel)
        if k in punten:
            punten[k].append((r[2], r[1]))
    for b in rows or []:
        k = BOUT_KLASSE.get(classify(b["r"], now_y, args.rood, args.geel, 1.5, 0.5, args.max_ouderdom, b["status"]))
        if k in punten:
            punten[k].append((b["lon"], b["lat"]))
    for k in ("groen", "geel", "rood"):
        c = _rgb(GEBIED_KLEUR[k])
        for lon, lat in punten[k]:
            x, y = xy(lon, lat)
            x, y = int(x), int(y)
            for dx, dy in ((0, 0), (1, 0), (-1, 0), (0, 1), (0, -1), (1, 1), (-1, -1), (1, -1), (-1, 1)):
                zet(x + dx, y + dy, c)
    _png(path, w, h, pix)
    return sum(len(v) for v in punten.values())


# --------------------------------------------------------------------------
# EGMS: bodembeweging gemeten door satellieten (Copernicus, European Ground Motion Service)
# --------------------------------------------------------------------------
# Product: Ortho, verticale component (L3, ORTHO-UP), vakken van 100 x 100 m, per tegel van
# 100 x 100 km in ETRS89-LAEA (EPSG:3035). Ophalen gaat via de EGMS-API met een persoonlijke
# CLMS-sleutel (JSON) in de omgevingsvariabele EGMS_SLEUTEL; zonder sleutel wordt de laag
# overgeslagen (of de cache gebruikt). Nieuwe releases verschijnen eens per jaar.

EGMS_API = "https://egms.land.copernicus.eu/insar-api/archive"
EGMS_BBOX = [[3.25, 50.70], [7.30, 53.60]]
EGMS_BRON = ("Copernicus Land Monitoring Service, European Ground Motion Service (EGMS); "
             "bevat aangepaste Copernicus Sentinel-data")


def _b64url(b):
    return base64.urlsafe_b64encode(b).rstrip(b"=").decode("ascii")


def _der_lees(data, pos):
    """Eén DER-element: geeft (tag, inhoud, volgende positie)."""
    tag = data[pos]
    ln = data[pos + 1]
    pos += 2
    if ln & 0x80:
        n = ln & 0x7F
        ln = int.from_bytes(data[pos:pos + n], "big")
        pos += n
    return tag, data[pos:pos + ln], pos + ln


def _rsa_sleutel(pem):
    """(n, d) uit een PEM-privésleutel (PKCS#8 of PKCS#1), zonder extra pakketten."""
    regels = [r.strip() for r in pem.strip().splitlines() if r.strip() and not r.startswith("-----")]
    der = base64.b64decode("".join(regels))
    tag, seq, _ = _der_lees(der, 0)
    if tag != 0x30:
        raise ValueError("geen geldige sleutel")
    # PKCS#8: SEQUENCE { INTEGER 0, SEQUENCE {algoritme}, OCTET STRING {PKCS#1} }
    t1, v1, p = _der_lees(seq, 0)
    t2, v2, p2 = _der_lees(seq, p)
    if t2 == 0x30:
        t3, inner, _ = _der_lees(seq, p2)
        if t3 != 0x04:
            raise ValueError("onbekend sleutelformaat")
        _, seq, _ = _der_lees(inner, 0)
    vals, p = [], 0
    while p < len(seq) and len(vals) < 4:
        t, v, p = _der_lees(seq, p)
        if t == 0x02:
            vals.append(int.from_bytes(v, "big"))
    if len(vals) < 4:
        raise ValueError("sleutel onvolledig")
    return vals[1], vals[3]  # versie, n, e, d


def _rs256(bericht, n, d):
    """RSASSA-PKCS1-v1_5 met SHA-256."""
    k = (n.bit_length() + 7) // 8
    t = bytes.fromhex("3031300d060960864801650304020105000420") + hashlib.sha256(bericht).digest()
    em = b"\x00\x01" + b"\xff" * (k - len(t) - 3) + b"\x00" + t
    return pow(int.from_bytes(em, "big"), d, n).to_bytes(k, "big")


def egms_toegang(sleutel_json):
    """Wisselt de CLMS-sleutel in voor een kortlevend toegangstoken (JWT-bearer, RS256)."""
    sk = json.loads(sleutel_json)
    n, d = _rsa_sleutel(sk["private_key"])
    nu = int(time.time())
    kop = _b64url(json.dumps({"alg": "RS256", "typ": "JWT"}, separators=(",", ":")).encode())
    claims = _b64url(json.dumps({"iss": sk["client_id"], "sub": sk["user_id"], "aud": sk["token_uri"],
                                 "iat": nu, "exp": nu + 3600}, separators=(",", ":")).encode())
    jwt = kop + "." + claims + "." + _b64url(_rs256((kop + "." + claims).encode(), n, d))
    body = urllib.parse.urlencode({"grant_type": "urn:ietf:params:oauth:grant-type:jwt-bearer",
                                   "assertion": jwt}).encode()
    req = urllib.request.Request(sk["token_uri"], data=body, method="POST", headers={
        "Accept": "application/json", "Content-Type": "application/x-www-form-urlencoded", "User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=60) as r:
        tok = json.loads(r.read().decode("utf-8")).get("access_token")
    if not tok:
        raise RuntimeError("geen toegangstoken ontvangen (klopt EGMS_SLEUTEL?)")
    return tok


def _egms_json(url, token, data=None):
    req = urllib.request.Request(url, data=json.dumps(data).encode() if data is not None else None,
                                 method="POST" if data is not None else "GET",
                                 headers={"Authorization": "Bearer " + token, "Accept": "application/json",
                                          "Content-Type": "application/json", "User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=120) as r:
        return json.loads(r.read().decode("utf-8"))


def laea_naar_wgs84(e, n):
    """ETRS89-LAEA (EPSG:3035) -> (lat, lon) in graden (GRS80, Snyder)."""
    a, e2 = 6378137.0, 0.00669438002290
    ec = math.sqrt(e2)
    lat0, lon0 = math.radians(52.0), math.radians(10.0)

    def q(p):
        s = math.sin(p)
        return (1 - e2) * (s / (1 - e2 * s * s) - 1 / (2 * ec) * math.log((1 - ec * s) / (1 + ec * s)))
    qp = q(math.pi / 2)
    rq = a * math.sqrt(qp / 2)
    b0 = math.asin(q(lat0) / qp)
    dd = a * (math.cos(lat0) / math.sqrt(1 - e2 * math.sin(lat0) ** 2)) / (rq * math.cos(b0))
    x, y = e - 4321000.0, n - 3210000.0
    rho = math.hypot(x / dd, dd * y)
    if rho == 0:
        return 52.0, 10.0
    c = 2 * math.asin(rho / (2 * rq))
    b = math.asin(math.cos(c) * math.sin(b0) + dd * y * math.sin(c) * math.cos(b0) / rho)
    lon = lon0 + math.atan2(x * math.sin(c), dd * rho * math.cos(b0) * math.cos(c) - dd * dd * y * math.sin(b0) * math.sin(c))
    lat = (b + (e2 / 3 + 31 * e2 ** 2 / 180 + 517 * e2 ** 3 / 5040) * math.sin(2 * b)
           + (23 * e2 ** 2 / 360 + 251 * e2 ** 3 / 3780) * math.sin(4 * b) + (761 * e2 ** 3 / 45360) * math.sin(6 * b))
    return math.degrees(lat), math.degrees(lon)


def egms_lees_zip(pad, binnen, uit):
    """Leest het CSV-bestand in een EGMS-ZIP regel voor regel. Bewaart per vak (lat, lon,
    zakking = -mean_velocity, std) in de arrays van 'uit', alleen als binnen(lat, lon)."""
    import zipfile
    import io
    import itertools
    n_in = n_tot = 0
    with zipfile.ZipFile(pad) as z:
        naam = next((x for x in z.namelist() if x.lower().endswith(".csv")), None)
        if not naam:
            raise RuntimeError(f"geen CSV in {os.path.basename(pad)}")
        with z.open(naam) as fh:
            tekst = io.TextIOWrapper(fh, encoding="utf-8", errors="replace", newline="")
            eerste = tekst.readline()
            sep = ";" if eerste.count(";") > eerste.count(",") else ","
            kop = [k.strip().strip('"').lower() for k in eerste.split(sep)]
            if "easting" in kop and "mean_velocity" in kop:
                ie, inn, iv = kop.index("easting"), kop.index("northing"), kop.index("mean_velocity")
                istd = kop.index("mean_velocity_std") if "mean_velocity_std" in kop else None
                regels = tekst
            else:  # geen kop: vaste volgorde volgens de productbeschrijving
                ie, inn, iv, istd = 1, 2, 5, 6
                regels = itertools.chain([eerste], tekst)
            m = max(ie, inn, iv, istd or 0) + 1
            for regel in regels:
                d = regel.split(sep, m)
                if len(d) < m:
                    continue
                try:
                    e, nn, v = float(d[ie]), float(d[inn]), float(d[iv])
                    sd = float(d[istd]) if istd is not None and d[istd].strip() else 0.0
                except ValueError:
                    continue
                n_tot += 1
                if not (-200 < v < 200):
                    continue
                lat, lon = laea_naar_wgs84(e, nn)
                if not binnen(lat, lon):
                    continue
                uit["lat"].append(int(round(lat * 1e6)))
                uit["lon"].append(int(round(lon * 1e6)))
                uit["v"].append(int(round(-v * 10)))
                uit["s"].append(min(32767, int(round(sd * 10))))
                n_in += 1
    return n_in, n_tot


def _egms_leeg():
    from array import array
    return {"lat": array("i"), "lon": array("i"), "v": array("h"), "s": array("h")}


def _egms_schrijf(path, punten, info):
    tmp = path + ".tmp"
    with gzip.open(tmp, "wb") as fh:
        kop = json.dumps(dict(info, n=len(punten["lat"]))).encode()
        fh.write(len(kop).to_bytes(4, "big") + kop)
        for k in ("lat", "lon", "v", "s"):
            fh.write(punten[k].tobytes())
    os.replace(tmp, path)


def _egms_lees(path):
    with gzip.open(path, "rb") as fh:
        ln = int.from_bytes(fh.read(4), "big")
        info = json.loads(fh.read(ln).decode())
        p = _egms_leeg()
        for k in ("lat", "lon", "v", "s"):
            p[k].frombytes(fh.read(p[k].itemsize * info["n"]))
    return p, info


def update_egms(cache_dir, args, state, gemeenten):
    """EGMS-punten uit de cache, of (bij een nieuwe release, eens per --egms-controle dagen
    gecontroleerd) opnieuw ophalen. Geeft (punten, info) of None."""
    fp = os.path.join(cache_dir, "egms.bin.gz")
    oud = None
    if os.path.exists(fp):
        try:
            oud = _egms_lees(fp)
        except (OSError, ValueError, KeyError, EOFError):
            oud = None
    sleutel = os.environ.get("EGMS_SLEUTEL", "").strip()
    if not sleutel:
        if oud:
            log(f"-> EGMS: uit cache (release {oud[1].get('release')}; geen EGMS_SLEUTEL om te controleren)")
        else:
            log("-> EGMS: overgeslagen (geen EGMS_SLEUTEL ingesteld)")
        return oud
    age = days_since(state.get("egmsControle"))
    if oud and not args.vernieuw and age is not None and age < args.egms_controle:
        log(f"-> EGMS: uit cache (release {oud[1].get('release')})")
        return oud
    log("-> EGMS: controleren op een nieuwe release")
    tmpdir = os.path.join(cache_dir, "egms.tmp")
    try:
        token = egms_toegang(sleutel)
        api = args.egms_api.rstrip("/")
        releases = _egms_json(api + "/releases", token)
        # het precieze formaat van deze lijst is niet gedocumenteerd: zoek alle 'jjjj-jjjj'
        namen = sorted(set(re.findall(r"(?<!\d)(\d{4}-\d{4})(?!\d)", json.dumps(releases))),
                       key=lambda r: (r[5:], r[:4]))
        if not namen:
            raise RuntimeError(f"geen releases gevonden ({str(releases)[:200]})")
        nieuwste = namen[-1]
        if oud and oud[1].get("release") == nieuwste and not args.vernieuw:
            state["egmsControle"] = date.today().isoformat()
            log(f"   geen nieuwe release (nog steeds {nieuwste})")
            return oud
        res = _egms_json(api + "/search", token, {"id": None, "bbox": EGMS_BBOX, "levels": ["L3"],
                                                   "releases": [nieuwste], "productType": "ORTHO-UP"})
        hits = [h for h in (res.get("hits") or []) if str(h.get("filename", "")).lower().endswith(".zip")
                and re.search(r"_U_", str(h.get("filename", "")))]
        if not hits:
            raise RuntimeError(f"geen tegels gevonden voor {nieuwste} ({res.get('message') or 'leeg antwoord'})")
        log(f"   release {nieuwste}: {len(hits)} tegels, samen {sum(int(h.get('filesize') or 0) for h in hits) / 1e6:,.0f} MB")
        if gemeenten:
            gidx = GebiedIndex(gemeenten)
            binnen = lambda la, lo: gidx.zoek(la, lo) is not None
        else:
            binnen = lambda la, lo: EGMS_BBOX[0][1] <= la <= EGMS_BBOX[1][1] and EGMS_BBOX[0][0] <= lo <= EGMS_BBOX[1][0]
        os.makedirs(tmpdir, exist_ok=True)
        # verder waar een vorige run stopte (per tegel bewaard in egms_deel.bin.gz)
        fdeel = os.path.join(cache_dir, "egms_deel.bin.gz")
        punten, klaar = _egms_leeg(), []
        if os.path.exists(fdeel) and not args.vernieuw:
            try:
                dp, dinfo = _egms_lees(fdeel)
                if dinfo.get("release") == nieuwste:
                    punten, klaar = dp, list(dinfo.get("klaar") or [])
                    log(f"   verder met ophalen: {len(klaar)} tegels al binnen")
            except (OSError, ValueError, KeyError, EOFError):
                pass
        t_start = time.time()
        for h in sorted(hits, key=lambda x: x["filename"]):
            if h["filename"] in klaar:
                continue
            if args.egms_max_minuten and time.time() - t_start > args.egms_max_minuten * 60:
                raise RuntimeError(f"tijdslimiet bereikt na {len(klaar)} van {len(hits)} tegels; de volgende run gaat verder")
            url = f"{api}/download/{h['filename']}?id={res.get('id')}"
            dl = os.path.join(tmpdir, os.path.basename(h["filename"]))
            t0 = time.time()
            req = urllib.request.Request(url, headers={"Authorization": "Bearer " + token, "User-Agent": USER_AGENT})
            for poging in range(1, 4):
                try:
                    with urllib.request.urlopen(req, timeout=300) as r, open(dl, "wb") as fh:
                        while True:
                            blok = r.read(1 << 20)
                            if not blok:
                                break
                            fh.write(blok)
                    break
                except (urllib.error.URLError, OSError) as e:
                    if poging == 3:
                        raise RuntimeError(f"download {h['filename']} mislukt ({e})")
                    time.sleep(10 * poging)
            n_in, n_tot = egms_lees_zip(dl, binnen, punten)
            os.remove(dl)
            klaar.append(h["filename"])
            _egms_schrijf(fdeel, punten, {"release": nieuwste, "klaar": klaar})
            log(f"   {h['filename']}: {fmt_int(n_in)} van {fmt_int(n_tot)} vakken in Nederland ({(time.time() - t0) / 60:.1f} min)")
        if len(punten["lat"]) < 1000 and not args.proef:
            raise RuntimeError(f"te weinig vakken ontvangen ({len(punten['lat'])})")
        info = {"release": nieuwste, "opgehaald": date.today().isoformat(), "tegels": len(hits)}
        _egms_schrijf(fp, punten, info)
        if os.path.exists(fdeel):
            os.remove(fdeel)
        state["egmsControle"] = date.today().isoformat()
        log(f"   klaar: {fmt_int(len(punten['lat']))} vakken van 100 m (release {nieuwste})")
        return punten, info
    except (RuntimeError, ValueError, KeyError, TypeError, OSError, urllib.error.URLError) as e:
        log(f"   let op: EGMS niet opgehaald ({e.__class__.__name__}: {e})")
        if oud:
            log("   oude EGMS-gegevens uit de cache worden gebruikt")
        return oud
    finally:
        if os.path.isdir(tmpdir):
            import shutil
            shutil.rmtree(tmpdir, ignore_errors=True)


EGMS_TEGEL = 0.1          # graden; detailbestanden per tegel van 0,1 x 0,1 graad
EGMS_ONZEKER = 2.0        # mm/jaar: vakken met een grotere standaardafwijking worden grijs


def write_egms(out_dir, egms):
    """Schrijft egms/overzicht.json (vakken van ~1 en ~5 km, mediaan) en egms/<tx>_<ty>.json
    (vakken van 100 m) voor de kaart. Geeft de meta-gegevens voor de kaart terug."""
    import shutil
    d = os.path.join(out_dir, "egms")
    if os.path.isdir(d):
        shutil.rmtree(d)
    if not egms:
        return None
    p, info = egms
    os.makedirs(d)
    tegels, grof, fijn = {}, {}, {}
    for i in range(len(p["lat"])):
        la, lo, v, sd = p["lat"][i] / 1e6, p["lon"][i] / 1e6, p["v"][i], p["s"][i]
        onz = 1 if sd > EGMS_ONZEKER * 10 else 0
        tegels.setdefault((int(math.floor(lo / EGMS_TEGEL)), int(math.floor(la / EGMS_TEGEL))), []).extend(
            [round(la, 5), round(lo, 5), v, onz])
        if not onz:
            fijn.setdefault((int(math.floor(lo / 0.015)), int(math.floor(la / 0.009))), []).append(v)
            grof.setdefault((int(math.floor(lo / 0.075)), int(math.floor(la / 0.045))), []).append(v)
    for (tx, ty), arr in tegels.items():
        write_text(os.path.join(d, f"{tx}_{ty}.json"), json.dumps(arr, separators=(",", ":")))

    def cellen(dic, dx, dy):
        out = []
        for (cx, cy), vs in dic.items():
            if len(vs) >= 3:
                out.extend([round((cy + 0.5) * dy, 4), round((cx + 0.5) * dx, 4), int(_mediaan(vs)), len(vs)])
        return out
    ov = json.dumps({"r1": cellen(fijn, 0.015, 0.009), "r5": cellen(grof, 0.075, 0.045),
                     "d1": [0.015, 0.009], "d5": [0.075, 0.045], "tegel": EGMS_TEGEL,
                     "tegels": sorted(f"{tx}_{ty}" for tx, ty in tegels)}, separators=(",", ":"))
    write_text(os.path.join(d, "overzicht.json"), ov)
    log(f"EGMS: {fmt_int(len(p['lat']))} vakken, {len(tegels)} detailbestanden, overzicht "
        f"{len(ov) / 1e6:.1f} MB (release {info.get('release')})")
    return {"egmsUrl": "egms/overzicht.json?v=" + short_hash(ov), "egmsRelease": info.get("release"),
            "egmsBron": EGMS_BRON, "egmsTegel": EGMS_TEGEL}


def egms_per_gebied(egms, gebieden, rood, geel):
    """Telling per gebied: zakking in klassen (rood/geel/groen, onzeker = oud)."""
    if not egms or not gebieden:
        return None
    p, _ = egms
    idx = GebiedIndex(gebieden)
    tel = [_Telling() for _ in gebieden]
    for i in range(len(p["lat"])):
        gi = idx.zoek(p["lat"][i] / 1e6, p["lon"][i] / 1e6)
        if gi is None:
            continue
        v = p["v"][i] / 10.0
        k = "oud" if p["s"][i] > EGMS_ONZEKER * 10 else ("rood" if v >= rood else "geel" if v >= geel else "groen")
        tel[gi].voeg(k, v, None)
    return tel


def short_hash(text):
    return hashlib.sha256(text.encode("utf-8")).hexdigest()[:10]


def write_site(out_dir, cache_dir, Bdata, G, meta, rows, args, now_y, nap=None, extra_paden=(), og_url="", nav=None):
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
    for fn in ("nap.json", "nap_reeksen.json"):
        if not nap and os.path.exists(os.path.join(out_dir, fn)):
            os.remove(os.path.join(out_dir, fn))
    if nap:
        nap_rows, nap_reeksen, nap_datum = nap
        nr = json.dumps(nap_reeksen, ensure_ascii=False, separators=(",", ":"))
        write_text(os.path.join(out_dir, "nap_reeksen.json"), nr)
        nj = {"p": nap_rows, "reeksen": "nap_reeksen.json?v=" + short_hash(nr), "opgehaald": nap_datum}
        if nav and len(nav.get("nap") or []) == len(nap_rows):
            nj["g"] = nav["nap"]
        nd = json.dumps(nj, ensure_ascii=False, separators=(",", ":"))
        write_text(os.path.join(out_dir, "nap.json"), nd)
        meta["napUrl"] = "nap.json?v=" + short_hash(nd)
        meta["napCount"] = len(nap_rows)
        meta["napService"] = args.nap_service.rstrip("/")
    if nav and "gemeenten/" in extra_paden:
        meta["gebied"] = {k: nav[k] for k in ("g", "a", "b", "w", "bw", "s")}
        if len(nav.get("bout") or []) == len(Bdata):
            meta["gebied"]["bout"] = nav["bout"]
            meta["gebied"]["st"] = nav["st"]
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
    descr = ("Onafhankelijke kaart van zakkende gebouwen in Nederland, op basis van open data: de NAP-peilmerken van "
             "Rijkswaterstaat en, in Amsterdam, de meetbouten van de gemeente. Zoek je adres en zie hoe snel het zakt.")
    head = [
        f'<link rel="canonical" href="{esc(site_url)}">',
        '<meta property="og:type" content="website">',
        '<meta property="og:locale" content="nl_NL">',
        '<meta property="og:title" content="Zakkingskaart">',
        f'<meta property="og:description" content="{esc(descr)}">',
        f'<meta property="og:url" content="{esc(site_url)}">',
    ] + _og_beeld(og_url)
    gv = re.sub(r"[^A-Za-z0-9_\-]", "", args.google_verificatie.strip())
    if gv:
        head.append(f'<meta name="google-site-verification" content="{gv}">')
    write_text(os.path.join(out_dir, "robots.txt"), f"User-agent: *\nAllow: /\nSitemap: {site_url}sitemap.xml\n")
    write_text(os.path.join(out_dir, "sitemap.xml"),
               '<?xml version="1.0" encoding="UTF-8"?>\n'
               '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n'
               f'  <url><loc>{esc(site_url)}</loc><lastmod>{meta["generated"]}</lastmod>'
               '<changefreq>weekly</changefreq></url>\n'
               + "".join(f'  <url><loc>{esc(site_url + p)}</loc><lastmod>{meta["generated"]}</lastmod></url>\n'
                         for p in extra_paden)
               + '</urlset>\n')

    page = (HTML_TEMPLATE
            .replace("__HEAD_EXTRA__", "\n".join(head))
            .replace("__GEBIEDLINK__", ' · <a href="gemeenten/">Cijfers per gemeente en buurt</a>' if "gemeenten/" in extra_paden else "")
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
    ap.add_argument("--verbergen", default=os.path.join(os.path.dirname(os.path.abspath(__file__)), "verbergen.txt"),
                    help="bestand met adressen, meetbout- of peilmerknummers die op verzoek niet op de kaart komen "
                         "(één per regel; standaard verbergen.txt naast het script; ook via ZK_VERBERGEN)")
    ap.add_argument("--zonder-nap", action="store_true",
                    help="--site: geen laag met NAP-peilmerken van Rijkswaterstaat (heel Nederland)")
    ap.add_argument("--nap-ververs", type=int, default=91,
                    help="--site: NAP-peilmerken hoe dan ook opnieuw ophalen als de cache zoveel dagen oud is "
                         "(standaard 91, eens per kwartaal)")
    ap.add_argument("--nap-controle", type=int, default=21,
                    help="--site: om de zoveel dagen kort bij Rijkswaterstaat controleren of er nieuwe gegevens zijn; "
                         "zo ja, dan wordt alles opnieuw opgehaald (standaard 21; 0 = niet controleren)")
    ap.add_argument("--nap-max-minuten", type=float, default=60,
                    help="--site: hoogstens zoveel minuten per run besteden aan het ophalen van NAP-peilmerken; "
                         "de volgende run gaat verder (standaard 60)")
    ap.add_argument("--nap-pauze", type=float, default=1.0,
                    help="--site: pauze tussen verzoeken aan Rijkswaterstaat in seconden (standaard 1)")
    ap.add_argument("--nap-geduld", default="120,300,600,900", help=argparse.SUPPRESS)  # wachttijden bij weigering
    ap.add_argument("--api", default=API, help=argparse.SUPPRESS)  # alleen voor testen
    ap.add_argument("--nap-service", default=NAP_SERVICE, help=argparse.SUPPRESS)  # alleen voor testen
    ap.add_argument("--vernieuw", action="store_true", help="data opnieuw ophalen en de cache negeren")
    ap.add_argument("--proef", action="store_true", help="snelle proefrun: alleen de eerste paar pagina's ophalen")
    ap.add_argument("--paginagrootte", type=int, default=2000, help="records per verzoek (standaard 2000)")
    ap.add_argument("--pauze", type=float, default=0.2, help="pauze tussen verzoeken in seconden (standaard 0,2)")
    ap.add_argument("--min-periode", type=float, default=4.0,
                    help="minimale meetperiode in jaren voor een snelheid (standaard 4)")
    ap.add_argument("--rood", type=float, default=2.0, help="standaard grens 'zakt snel' in mm/jaar (op de kaart aanpasbaar)")
    ap.add_argument("--geel", type=float, default=1.0, help="standaard grens 'zakt matig' in mm/jaar (op de kaart aanpasbaar)")
    ap.add_argument("--max-ouderdom", type=float, default=10.0,
                    help="laatste meting ouder dan dit aantal jaren telt als 'geen recente data' (standaard 10)")
    ap.add_argument("--bbox", default=None,
                    help="alleen dit gebied op de kaart: zuid,west,noord,oost in graden, bv. 52.33,4.85,52.38,4.96")
    ap.add_argument("--zonder-grafieken", action="store_true", help="geen meetreeksen in de kaart (kleiner bestand)")
    ap.add_argument("--niet-openen", action="store_true", help="kaart niet automatisch in de browser openen")
    ap.add_argument("--zonder-egms", action="store_true",
                    help="--site: geen satellietlaag (EGMS, Copernicus)")
    ap.add_argument("--egms-controle", type=int, default=30,
                    help="--site: om de zoveel dagen kijken of er een nieuwe EGMS-release is (standaard 30)")
    ap.add_argument("--egms-max-minuten", type=float, default=60,
                    help="--site: hoogstens zoveel minuten per run aan het ophalen van EGMS-tegels; de volgende run gaat verder")
    ap.add_argument("--egms-api", default=EGMS_API, help=argparse.SUPPRESS)  # alleen voor testen
    ap.add_argument("--zonder-gebieden", action="store_true",
                    help="--site: geen pagina's per gemeente en per Amsterdamse buurt maken")
    ap.add_argument("--gebieden-ververs", type=int, default=182,
                    help="--site: gemeente- en buurtgrenzen opnieuw ophalen als ze zoveel dagen oud zijn (standaard 182)")
    ap.add_argument("--gemeenten-url", default=GEMEENTEN_URL, help=argparse.SUPPRESS)  # alleen voor testen
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

    log("Zakkingskaart" + (" (website)" if args.site else ""))
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
    args._verbergen = lees_verbergen(args.verbergen)
    if args._verbergen:
        log(f"Verbergen op verzoek: {len(args._verbergen)} regel(s)")   # geen adressen in het (openbare) log
    for rec in bouten_raw:
        bid = rec.get("identificatie")
        if bid is None:
            continue
        bid = str(bid)
        if bid.upper() in args._verbergen:
            continue
        ll = parse_point(rec.get("geometrie") or rec.get("geometry"))
        if not ll or not all(math.isfinite(v) for v in ll):
            no_geom += 1
            continue
        if bbox and not (bbox[0] <= ll[0] <= bbox[2] and bbox[1] <= ll[1] <= bbox[3]):
            continue
        addr = find_address(rec)
        if addr and re.sub(r"\s+", " ", addr).strip().upper() in args._verbergen:
            continue
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
        nap = None
        if not args.zonder_nap:
            got = update_nap(cache_dir, args, max_pages, state)
            if got:
                nap_rows, nap_reeksen = build_nap(got[0], got[1], args)
                nap = (nap_rows, nap_reeksen, state.get("napOpgehaald") or "")
        paden = []
        site_url = args.site_url.strip()
        site_url += "" if site_url.endswith("/") else "/"
        gem = buurten = nav = None
        if not args.zonder_gebieden:
            gem = laad_gebieden(cache_dir, "gemeenten", haal_gemeenten, args, state, 300, 400)
            buurten = laad_gebieden(cache_dir, "ams_buurten", haal_ams_buurten, args, state, 100, 1500) if gem else None
        egms = None if args.zonder_egms else update_egms(cache_dir, args, state, gem)
        try:
            em = write_egms(out_dir, egms)
            if em:
                meta.update(em)
        except Exception as e:  # de satellietlaag mag de kaart nooit tegenhouden
            log(f"   let op: EGMS-laag niet gemaakt ({e.__class__.__name__}: {e})")
            import shutil
            shutil.rmtree(os.path.join(out_dir, "egms"), ignore_errors=True)
            meta.pop("egmsUrl", None)
        og_url = ""
        try:  # voorbeeldplaatje voor gedeelde links
            dp = os.path.join(out_dir, DEELBEELD)
            maak_deelbeeld(dp, nap[0] if nap else [], rows, gem, now_y, args)
            og_url = site_url + DEELBEELD + "?v=" + hashlib.sha256(open(dp, "rb").read()).hexdigest()[:10]
        except Exception as e:
            log(f"   let op: voorbeeldplaatje niet gemaakt ({e.__class__.__name__}: {e})")
        if not args.zonder_gebieden:
            if gem:
                try:
                    paden, nav = write_gebieden(out_dir, site_url, "gebied.css?v=" + short_hash(GEBIED_CSS), args, now_y,
                                           today, nap[0] if nap else [], rows, gem, buurten, og_url,
                                           egms if meta.get("egmsUrl") else None)
                except Exception as e:  # pagina's mogen de kaart nooit tegenhouden
                    log(f"   let op: pagina's per gemeente niet gemaakt ({e.__class__.__name__}: {e})")
                    paden, nav = [], None
                    import shutil
                    for d in ("gemeente", "gemeenten"):
                        shutil.rmtree(os.path.join(out_dir, d), ignore_errors=True)
        html_path = write_site(out_dir, cache_dir, Bdata, G, meta, rows, args, now_y, nap, paden, og_url, nav)
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
                .replace("__GEBIEDLINK__", "")
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
    for k in ("zakt snel", "zakt matig", "zakt nauwelijks", "gestabiliseerd", "geen recente data"):
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
<title>Zakkingskaart: hoe snel zakken de gebouwen in jouw buurt?</title>
<meta name="description" content="Onafhankelijke kaart van zakkende gebouwen in Nederland, op basis van open data: de NAP-peilmerken van Rijkswaterstaat en, in Amsterdam, de meetbouten van de gemeente. Zoek je adres en zie hoe snel het de laatste jaren zakt.">
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
#msg a{font-weight:600}
#terug{position:absolute;top:58px;left:58px;z-index:1000;display:flex;align-items:center;gap:4px;padding:4px 6px 4px 10px;font-size:14px}
#terug[hidden]{display:none}
#terug a{font-weight:600;text-decoration:none}
#terug button{border:0;background:none;font-size:18px;line-height:1;cursor:pointer;color:var(--muted);padding:0 4px}
body.terug #msg{top:96px}
body.terug #rond{top:98px;max-height:calc(100% - 128px)}
.pop .meer{margin:6px 0 0;font-size:13px}
.pop .meer a,#rond .meer a{font-weight:600}
#rond .meer{margin:8px 0 0;font-size:13px}
#sheet{position:fixed;left:0;right:0;bottom:0;z-index:1500;background:#fff;border-radius:14px 14px 0 0;box-shadow:0 -6px 24px rgba(0,0,0,.18);max-height:62vh;display:flex;flex-direction:column}
#sheet[hidden]{display:none}
#sheet .greep{width:40px;height:4px;border-radius:2px;background:#ccc;margin:8px auto 2px;flex:none}
#sheet .dicht{position:absolute;top:4px;right:8px;border:0;background:none;font-size:26px;line-height:1;padding:4px 8px;color:var(--muted);cursor:pointer}
#sheet .body{overflow:auto;padding:0 14px 16px;-webkit-overflow-scrolling:touch}
#sheet .leaflet-popup{position:static!important;transform:none!important;left:auto!important;top:auto!important;bottom:auto!important;margin:0!important;opacity:1!important}
#sheet .leaflet-popup-content-wrapper{box-shadow:none!important;border-radius:0!important;padding:0!important;background:none!important}
#sheet .leaflet-popup-content{margin:4px 0 0!important;width:auto!important;max-width:none!important}
#sheet .leaflet-popup-tip-container,#sheet .leaflet-popup-close-button{display:none!important}
body.sheet .leaflet-control-attribution{display:none}
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
#rond{top:60px;left:58px;width:min(380px,calc(100vw - 76px));padding:10px 12px;max-height:calc(100% - 90px);overflow:auto}
#rond[hidden]{display:none}
#rond .rt{display:flex;justify-content:space-between;gap:8px;align-items:flex-start}
#rond h2{font-size:14px;margin:0 0 2px;line-height:1.3}
#rond .x{border:0;background:none;font-size:20px;line-height:1;cursor:pointer;color:var(--muted);padding:0 2px}
#rond .rs{margin:8px 0 0}
#rond .rs b{font-size:13px}
#rond .rbar{display:flex;height:8px;border-radius:4px;overflow:hidden;background:#eee;margin:4px 0 2px}
#rond .rbar i{display:block;height:100%}
#rond .rk{font-size:12px;color:var(--muted)}
.pop .pand,#rond .pand{margin:6px 0 0;font-size:13px}
.pop .pand .m,#rond .pand .m{color:var(--muted);font-size:12px}
.pop .omg{margin:6px 0 0;font-size:12px;color:#3b3a36;background:#f4f2ec;border-radius:6px;padding:5px 7px}
@media (max-width:640px){
  #rond{left:12px;width:calc(100vw - 24px);max-height:45%}
  #legend{width:calc(100vw - 24px);top:auto;bottom:24px;max-height:55%}
  #legend.collapsed .body{display:none}
}
#legend.collapsed .body{display:none}
#modal,#fbModal,#aboutModal{position:fixed;inset:0;z-index:4000;background:rgba(20,20,18,.5);display:flex;align-items:center;justify-content:center;padding:16px}
#modal[hidden],#fbModal[hidden],#aboutModal[hidden]{display:none}
#intro{position:fixed;inset:0;z-index:4500;background:rgba(20,20,18,.55);display:flex;align-items:center;justify-content:center;padding:16px}
#intro[hidden]{display:none}
#aboutModal{z-index:4600}
#intro .mbox{max-width:520px}
#intro .ibody{padding:4px 22px 20px;overflow:auto}
#intro ul{margin:6px 0 14px;padding-left:20px}
#intro li{margin:0 0 9px;line-height:1.45}
#intro .iakk{margin:0 0 14px;line-height:1.45}
#intro .iact{display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap}
#intro button{background:var(--accent);color:#fff;border:0;border-radius:8px;padding:10px 16px;font:inherit;font-weight:600;cursor:pointer}
#intro a{color:var(--accent)}
.pop .disc{margin:6px 0 0;font-size:11.5px;color:#77746c}
.pop .disc a{color:inherit}
#legend .indep{margin:4px 0 0;font-size:12px;color:var(--muted)}
#legend .indep a,#aboutModal a{color:var(--accent)}
#aBody ul{margin:4px 0 10px;padding-left:20px}
#aBody li{margin:0 0 6px}
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
table.mt tr.oud td{color:#9a978f}
.mnote{color:var(--muted);font-size:12px;margin:6px 0 0}
.msum{display:flex;flex-wrap:wrap;gap:6px 18px;font-size:13px;margin:0 0 10px}
.msum b{font-variant-numeric:tabular-nums}
pre.rawjson{background:#f6f5f1;border:1px solid var(--line);border-radius:8px;padding:10px 12px;font:12px/1.45 ui-monospace,Menlo,Consolas,monospace;max-height:45vh;overflow:auto;white-space:pre;margin:0}
.warn{background:#fff3cd;border:1px solid #f0d58a;border-radius:8px;padding:8px 10px;font-size:12.5px;margin-bottom:10px}
.naprow{border-top:1px solid var(--line);margin-top:8px;padding-top:8px}
.naprow[hidden]{display:none}
.sw.ring{width:12px;height:12px;border-radius:50%;background:#fff;box-shadow:inset 0 0 0 2.5px var(--accent)}
#napnote{margin:4px 0 0}
#napnote:empty{display:none}
</style>
</head>
<body>
<div id="map" role="application" aria-label="Kaart met meetbouten en NAP-peilmerken"></div>
<form id="search" class="panel" autocomplete="off">
  <input id="q" list="streets" placeholder="Zoek een adres, straat of plaats in Nederland" aria-label="Zoeken">
  <button type="submit">Zoek</button>
  <datalist id="streets"></datalist>
</form>
<div id="msg" hidden></div>
<div id="terug" class="panel" hidden></div>
<div id="sheet" hidden role="dialog" aria-label="Details"><button type="button" class="dicht" aria-label="Sluiten">×</button><div class="greep"></div><div class="body"></div></div>
<div id="rond" class="panel" hidden role="region" aria-label="Meetpunten rond het gezochte adres"></div>
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
<div id="intro" hidden role="dialog" aria-modal="true" aria-labelledby="iTitle">
  <div class="mbox"><div class="mhead"><div><h2 id="iTitle">Voordat je de kaart opent</h2><div class="sub">Onafhankelijk hobbyproject, niet van de gemeente of Rijkswaterstaat</div></div></div>
  <div class="ibody"><ul>
    <li>Deze kaart maakt openbare meetgegevens van de gemeente Amsterdam en Rijkswaterstaat makkelijker te bekijken, en rekent ze om naar hoe snel een meetpunt zakt. Meer niet: de kaart geeft <b>geen oordeel of advies</b> over gebouwen, de grond of de waarde ervan.</li>
    <li>Een meetpunt zit in één gevel of bouwwerk en zegt niet alles over de omgeving. Metingen kunnen oud, onvolledig of onjuist zijn.</li>
    <li>Vragen over de gegevens, of klopt er iets niet? Die gaan naar de bron: de gemeente Amsterdam of Rijkswaterstaat. Wat zij aanpassen, komt vanzelf op deze kaart.</li>
  </ul>
  <p class="iakk">Door verder te gaan, ga je ermee akkoord dat je deze kaart gebruikt zoals hij is, en dat je er <b>geen rechten aan kunt ontlenen</b>.</p>
  <div class="iact"><a href="#over" id="iMore">Over &amp; disclaimer</a><button id="iOk" type="button">Akkoord, naar de kaart</button></div></div></div>
</div>
<div id="aboutModal" hidden role="dialog" aria-modal="true" aria-labelledby="aTitle">
  <div class="mbox"><div class="mhead"><div><h2 id="aTitle">Over deze kaart</h2></div><button id="aClose" type="button" aria-label="Sluiten">&times;</button></div>
  <div id="aBody">
    <h3>Onafhankelijk initiatief</h3>
    <p>Zakkingskaart is een onafhankelijk, niet-commercieel initiatief. Het is geen website van een gemeente of van Rijkswaterstaat en wordt niet door hen gemaakt, gecontroleerd of goedgekeurd.</p>
    <h3>Wat de kaart laat zien</h3>
    <p>De kaart laat zien hoe snel gebouwen en de grond in Nederland zakken. Ze gebruikt drie openbare bronnen: de peilmerken van Rijkswaterstaat (heel Nederland), de meetbouten van de gemeente Amsterdam (veel meer meetpunten, alleen in Amsterdam) en satellietmetingen van de Europese Unie (heel Nederland).</p>
    <h3>Begrippen</h3>
    <ul>
      <li><b>mm per jaar:</b> hoeveel millimeter een punt per jaar lager komt te liggen. 1 mm per jaar is 1 centimeter in 10 jaar.</li>
      <li><b>Meetbout:</b> een boutje in de gevel van een gebouw. De gemeente Amsterdam meet regelmatig hoe hoog het zit.</li>
      <li><b>Peilmerk:</b> een metalen boutje in een gebouw of brug, waarvan Rijkswaterstaat eens per 5 tot 15 jaar de hoogte meet. De hoogte wordt gemeten ten opzichte van het NAP (Normaal Amsterdams Peil), de standaardhoogte van Nederland.</li>
      <li><b>Satellietmeting:</b> satellieten meten met radar de hoogte van grond, straten en gebouwen samen, per blok van 100 bij 100 meter.</li>
      <li><b>Doorsnee (of doorgaans):</b> de middelste waarde van een groep meetpunten: de helft zakt sneller, de helft langzamer. Zo bepaalt één uitschieter niet het beeld.</li>
      <li><b>Bouwjaar en adressen:</b> uit de BAG, het landelijke register van gebouwen en adressen. Bij heel oude panden is het bouwjaar soms een schatting. Het aantal adressen laat zien of een pand bijvoorbeeld een appartementengebouw is.</li>
      <li><b>Laatste jaren:</b> de snelheid berekend uit de laatste meting en een meting van minstens 4 jaar eerder (bij peilmerken minstens 8 jaar).</li>
    </ul>
    <h3>Amsterdam: meetbouten</h3>
    <p>In Amsterdam gebruikt de kaart de openbare meetboutgegevens van de gemeente en wordt ze wekelijks automatisch bijgewerkt. Meetbouten zijn meetpunten in gevels waarvan de gemeente periodiek de hoogte meet. Per meetbout laat de kaart zien hoe snel die de laatste jaren is gezakt (de recente snelheid), of het zakken duidelijk is afgenomen (&ldquo;gestabiliseerd&rdquo;) en of er recente metingen zijn. De kaart van de gemeente kleurt op de gemiddelde snelheid sinds de eerste meting; daardoor kunnen de kleuren verschillen.</p>
    <h3>Heel Nederland: peilmerken van Rijkswaterstaat</h3>
    <p>Een peilmerk (officieel: NAP-peilmerk) is een metalen boutje in een huis, brug of ander bouwwerk, waarvan Rijkswaterstaat eens per 5 tot 15 jaar de hoogte meet. De kaart toont per peilmerk hoe snel het de laatste jaren is gezakt, met dezelfde kleuren als de meetbouten, maar als open rondje. Uitgezoomd zie je vakjes met de doorsnee van de peilmerken erin. Er zijn veel minder peilmerken dan meetbouten: meestal een paar per buurt. Een peilmerk zakt mee met de grond van de hele omgeving; het zegt dus iets over de plek, niet per se over dat ene gebouw. Peilmerken die meer dan 20 jaar niet zijn gemeten, of waarvan de metingen een onverklaarbare sprong maken, zijn grijs.</p>
    <h3>Heel Nederland: bodemdaling gemeten vanuit de ruimte</h3>
    <p>De laag &ldquo;Bodemdaling gemeten vanuit de ruimte&rdquo; komt van Copernicus, het aardobservatieprogramma van de Europese Unie (de dienst heet European Ground Motion Service, EGMS). Satellieten meten met radar elke paar dagen de hoogte van grond, straten en gebouwen. Daaruit is per blok van 100 bij 100 meter berekend hoeveel de grond per jaar zakte of steeg, over een periode van vijf jaar. Uitgezoomd zie je de doorsnee per blok van ongeveer 5 of 1 kilometer. Je ziet alleen blokken waar iets het radarsignaal goed terugkaatst, zoals daken, wegen en bruggen; in weilanden, akkers en op water is niets gemeten. Anders dan een meetbout of peilmerk meet een blok niet één gebouw, maar alles samen: een huis op palen kan stil staan terwijl de straat ervoor zakt, en andersom. Blokken met een onzekere meting zijn grijs. De laag staat standaard uit.</p>
    <h3>Wat de kleuren betekenen</h3>
    <p>De kleur hangt alleen af van hoe snel een meetpunt de laatste jaren is gezakt: groen bij minder dan 1 mm per jaar, geel bij 1 tot 2 mm per jaar, rood bij 2 mm per jaar of meer. Gestabiliseerd: eerst minstens 1,5 en nu minder dan 0,5 mm per jaar. Grijs: geen recente of te weinig metingen. De grenzen van 1 en 2 mm per jaar zijn ook klassegrenzen op de meetboutenkaart van de gemeente Amsterdam, die eveneens van groen via geel naar rood kleurt. De kleuren beschrijven alleen de gemeten snelheid; ze zeggen niet of iets goed of slecht is. Onder &ldquo;Grenzen aanpassen&rdquo; kun je zelf andere grenzen kiezen.</p>
    <h3>Disclaimer</h3>
    <p>Deze kaart geeft openbare meetgegevens van de gemeente Amsterdam en Rijkswaterstaat eenvoudiger weer en rekent ze om naar een snelheid. Daarbij kan ze afwijken van de weergave van de bron. De kaart geeft geen oordeel of advies over gebouwen, de grond of de waarde ervan, en kan onvolledig of verouderd zijn. De kaart wordt met zorg gemaakt, maar zonder garantie. Wie de kaart gebruikt, doet dat op eigen verantwoordelijkheid; aan de kaart kunnen geen rechten worden ontleend en de beheerder is niet aansprakelijk voor gevolgen van het gebruik, voor zover de wet dat toestaat.</p>
    <h3>Vragen over de gegevens</h3>
    <p>Deze site beheert de meetgegevens niet en kan ze niet aanpassen. Heb je een vraag over een meting, of denk je dat er iets niet klopt? Neem dan contact op met de bron:</p>
    <ul>
      <li><b>Meetbouten (Amsterdam):</b> gemeente Amsterdam, Basisinformatie, <a href="mailto:gebruik.basisinformatie@amsterdam.nl">gebruik.basisinformatie@amsterdam.nl</a> (contactadres bij de dataset <a href="https://data.overheid.nl/dataset/m6avdgklu2czqa" target="_blank" rel="noopener">Kernregistratie Meetbouten</a>).</li>
      <li><b>Peilmerken:</b> Rijkswaterstaat, <a href="https://www.rijkswaterstaat.nl/formulieren/contactformulier-servicedesk-data" target="_blank" rel="noopener">contactformulier Servicedesk Data</a>.</li>
    </ul>
    <p>Wat de bron aanpast of verwijdert, verdwijnt bij de volgende automatische update ook van deze kaart. De knop &ldquo;Feedback&rdquo; is voor opmerkingen over de website zelf.</p>
    <h3>Privacy</h3>
    <p>Deze site gebruikt geen cookies en volgt geen bezoekers. We tellen alleen anoniem hoe vaak de kaart per dag wordt geopend, vanaf mobiel of computer, via welke soort bron (bijvoorbeeld een zoekmachine of een andere website) en hoe vaak er wordt gezocht en geklikt. Er worden daarbij geen IP-adressen of andere kenmerken van jou opgeslagen. Je kaartinstellingen worden alleen in je eigen browser bewaard. De site wordt gehost door GitHub, de achtergrondkaart komt van PDOK, een gezocht adres buiten de kaartgegevens wordt opgezocht bij de PDOK Locatieserver van de overheid, het bouwjaar van het pand op een aangeklikte plek wordt opgevraagd bij de BAG-dienst van PDOK, en &ldquo;Alle meetgegevens bekijken&rdquo; haalt gegevens op bij de gemeente of bij Rijkswaterstaat; die partijen zien daarbij, zoals bij elke website, je IP-adres. Feedback komt als e-mail binnen bij de beheerder; er wordt geen naam of e-mailadres gevraagd. De adressen bij de meetpunten komen uit de openbare gegevens van de gemeente en Rijkswaterstaat; vragen daarover gaan naar de bron (zie &ldquo;Vragen over de gegevens&rdquo;).</p>
    <h3>Bronnen</h3>
    <p>Meetbouten: Gemeente Amsterdam (api.data.amsterdam.nl). Achtergrondkaart: PDOK / Kadaster (CC BY 4.0). Adressen zoeken: PDOK Locatieserver. Bouwjaar en adressen per pand: Basisregistratie Adressen en Gebouwen (BAG), Kadaster via PDOK (publiek domein). NAP-peilmerken: Rijkswaterstaat (CC0). Bodembeweging: Copernicus Land Monitoring Service, European Ground Motion Service (EGMS); bevat aangepaste Copernicus Sentinel-data. Gemeentegrenzen (voor de pagina's per gemeente): Kadaster, Bestuurlijke Gebieden (CC BY 4.0); buurtgrenzen Amsterdam: gemeente Amsterdam (CC0). Kaartbibliotheek: Leaflet (BSD-2). Broncode en rekenmethode: <a href="https://github.com/zakkingskaart/zakkingskaart.github.io" target="_blank" rel="noopener">GitHub</a>.</p>
  </div></div>
</div>
<div id="modal" hidden role="dialog" aria-modal="true" aria-labelledby="mTitle">
  <div class="mbox"><div class="mhead"><div><h2 id="mTitle"></h2><div class="sub" id="mSub"></div></div><button id="mClose" type="button" aria-label="Sluiten">&times;</button></div>
  <div id="mBody"></div></div>
</div>
<aside id="legend" class="panel">
  <div class="head"><h1>Zakkingskaart</h1><div class="btns"><button id="fbBtn" type="button" hidden>&#9733; Feedback</button><button class="toggle" id="toggle" aria-expanded="true" title="In- of uitklappen">–</button></div></div>
  <p class="indep">Onafhankelijk initiatief, los van gemeente en Rijkswaterstaat · <a href="#over" id="aboutLink">Over &amp; disclaimer</a>__GEBIEDLINK__</p>
  <div class="body">
    <div class="rows" id="rows"></div>
    <p class="note" id="viewnote">De getallen tellen de meetbouten die je nu op de kaart ziet.</p>
    <p class="note"><b>Amsterdam:</b> een stip is één meetbout in een gevel; een lijn is een stuk straat (de doorsnee van de meetbouten daar).</p>
    <p class="note"><b>Heel Nederland:</b> open rondjes zijn peilmerken van Rijkswaterstaat. Klik op iets voor details.</p>
    <p class="note"><b>In de buurt:</b> zoek een adres, of klik met rechts (op een telefoon: lang indrukken) ergens op de kaart.</p>
    <div class="naprow" id="boutrow">
      <label class="row"><input type="checkbox" id="cBout"><i class="sw" style="background:linear-gradient(90deg,#1a9850,#e6a700,#d7301f);border-radius:2px;height:4px"></i><span class="lbl">Meetbouten (Amsterdam)<small>Meetpunten van de gemeente in gevels van gebouwen</small></span><span class="cnt" id="cnt-bout">–</span></label>
    </div>
    <div class="naprow" id="naprow" hidden>
      <label class="row"><input type="checkbox" id="cNap"><i class="sw ring"></i><span class="lbl">Peilmerken (heel Nederland)<small>Meetpunten van Rijkswaterstaat in gebouwen en bruggen · eens per 5 tot 15 jaar gemeten</small></span><span class="cnt" id="cnt-nap">–</span></label>
      <p class="note" id="napnote"></p>
    </div>
    <div class="naprow" id="egrow" hidden>
      <label class="row"><input type="checkbox" id="cEgms"><i class="sw" style="background:linear-gradient(90deg,#1a9850,#e6a700,#d7301f);opacity:.6;border-radius:2px"></i><span class="lbl">Bodemdaling gemeten vanuit de ruimte<small>Satellietmetingen van de EU · per blok van 100 × 100 m · grond, straten en gebouwen samen</small></span></label>
      <p class="note" id="egnote"></p>
    </div>
    <details id="settings">
      <summary>Grenzen aanpassen</summary>
      <div class="set">
        <label for="sRed">Zakt snel vanaf (mm/jaar)</label><input id="sRed" type="number" step="0.1" min="0">
        <label for="sYellow">Zakt matig vanaf (mm/jaar)</label><input id="sYellow" type="number" step="0.1" min="0">
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
document.addEventListener("click",function(e){var a=e.target&&e.target.closest&&e.target.closest('a[href="#over"]');if(a&&a.id!=="aboutLink")open(e);});
document.getElementById("aClose").addEventListener("click",close);
m.addEventListener("click",function(e){if(e.target===m)close();});
document.addEventListener("keydown",function(e){if(e.key==="Escape"&&!m.hidden)close();});
if(location.hash==="#over")open();
window.addEventListener("hashchange",function(){if(location.hash==="#over"&&m.hidden)open();});
// Akkoord bij het openen: geldig voor 30 dagen in deze browser (localStorage, geen cookie).
// Verhoog iv als de tekst inhoudelijk verandert; dan ziet iedereen hem opnieuw.
var IK="zakkingskaart.intro",iv="2",GELDIG=30*864e5,it=document.getElementById("intro"),gezien=false;
try{var a=JSON.parse(localStorage.getItem(IK)||"null");gezien=!!a&&a.v===iv&&typeof a.t==="number"&&Date.now()-a.t>=0&&Date.now()-a.t<GELDIG;}catch(e){}
if(!gezien&&location.hash!=="#over"){it.hidden=false;setTimeout(function(){var b=document.getElementById("iOk");if(b)b.focus();},0);}
document.getElementById("iOk").addEventListener("click",function(){it.hidden=true;try{localStorage.setItem(IK,JSON.stringify({v:iv,t:Date.now()}));}catch(e){}});
})();
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
// Geen cookies en geen kenmerken per bezoeker. Alleen wie zelf ?niettellen=1 opent, krijgt een
// vlag in de eigen browser waardoor die browser niet meer meetelt (?niettellen=0 zet hem terug).
const TEL=(function(){
  const url=(M.statsUrl||"").trim(),geen={tel(){}};
  try{const q=new URLSearchParams(location.search).get("niettellen");
    if(q==="1")localStorage.setItem("zakkingskaart.niettellen","1");
    else if(q==="0")localStorage.removeItem("zakkingskaart.niettellen");
    if(q==="1"||q==="0")setTimeout(()=>{const m=document.getElementById("msg");if(!m)return;
      m.textContent=q==="1"?"Deze browser telt niet meer mee in de bezoekersstatistieken.":"Deze browser telt weer mee in de bezoekersstatistieken.";
      m.hidden=false;setTimeout(()=>{m.hidden=true;},6000);},400);
    if(localStorage.getItem("zakkingskaart.niettellen")==="1")return geen;}catch(e){}
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
  staleYears:M.defaults.staleYears,colorStale:false,showDots:true,show:{rood:true,geel:true,groen:true,stab:true,oud:true},base:"pdokgrijs",showNap:true,showEgms:false,showBout:true};
let S=JSON.parse(JSON.stringify(DEF));
try{const s=JSON.parse(localStorage.getItem("zakkingskaart.v1")||"null");if(s){S=Object.assign(S,s);S.show=Object.assign({},DEF.show,s.show||{});if(!s.napKeuze)S.showNap=DEF.showNap;}}catch(e){}
function save(){try{localStorage.setItem("zakkingskaart.v1",JSON.stringify(S));}catch(e){}}

const nf=(v,d=2)=>v==null?"–":Number(v).toLocaleString("nl-NL",{minimumFractionDigits:d,maximumFractionDigits:d});
const esc=s=>String(s==null?"":s).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const fmtDate=iso=>iso?iso.slice(8,10)+"-"+iso.slice(5,7)+"-"+iso.slice(0,4):"–";
function median(a){const s=a.slice().sort((x,y)=>x-y),m=s.length>>1;return s.length%2?s[m]:(s[m-1]+s[m])/2;}
function label(c){
  if(c==="rood")return["Zakt snel","vanaf "+nf(S.red,1)+" mm/jaar"];
  if(c==="geel")return["Zakt matig",nf(S.yellow,1)+" tot "+nf(S.red,1)+" mm/jaar"];
  if(c==="groen")return["Zakt nauwelijks","minder dan "+nf(S.yellow,1)+" mm/jaar"];
  if(c==="stab")return["Gestabiliseerd","zakte eerst flink, nu bijna niet meer"];
  return["Geen recente metingen","laatste meting ouder dan "+S.staleYears+" jaar, te weinig metingen of meetpunt vervallen"];
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
const map=L.map("map",{preferCanvas:true,zoomSnap:0.5,minZoom:7}).setView([52.36,4.9],13);
window.zkKaart={centrum:()=>map.getCenter(),zoom:()=>map.getZoom()};  // alleen-lezen, voor tests
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
map.attributionControl.addAttribution("Meetbouten: Gemeente Amsterdam"+(M.napUrl?" · Peilmerken: Rijkswaterstaat":"")+(M.egmsUrl?" · Satelliet: Copernicus (EU)":"")+" · Geen oordeel of advies");

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
if(M.napUrl)map.fitBounds([[50.75,3.36],[53.55,7.23]]);
else if(bounds.isValid())map.fitBounds(bounds,{padding:[20,20],maxZoom:15});

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
    const c=GC[pl._zkG][0],vis=S.showBout!==false&&!!S.show[c];
    pl.setStyle({color:COL[c],weight:sz.w,opacity:c==="oud"?.5:.9,dashArray:c==="stab"?(sz.w*0.3)+" "+(sz.w*1.6):null});
    if(vis&&!pl._zkVis){lineLayer.addLayer(pl);pl._zkVis=true;added=true;}
    else if(!vis&&pl._zkVis){lineLayer.removeLayer(pl);pl._zkVis=false;}
  }
  for(const m of dots){
    const i=m._zkI,c=C[i],vis=S.showBout!==false&&!!S.show[c]&&(S.showDots||!inLine[i]);
    const st=c==="stab";
    m.setStyle({radius:st?sz.r+0.5:sz.r,fillColor:COL[c],color:st?RING:"#ffffff",weight:st?Math.max(1.2,sz.sw+0.6):sz.sw,stroke:st||sz.sw>0,fillOpacity:c==="oud"?.75:1});
    if(vis&&!m._zkVis){dotLayer.addLayer(m);m._zkVis=true;}
    else if(!vis&&m._zkVis){dotLayer.removeLayer(m);m._zkVis=false;}
    else if(vis&&added&&m.bringToFront){m.bringToFront();}
  }
  counts();
  napRefresh(true);
  egRefresh(true);
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
  const tb=document.getElementById("cnt-bout");if(tb)tb.textContent=KL.reduce((a,k)=>a+n[k],0).toLocaleString("nl-NL");
}
const boutCb=document.getElementById("cBout");boutCb.checked=S.showBout!==false;
boutCb.addEventListener("change",()=>{S.showBout=boutCb.checked;save();styleAll();});
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
document.getElementById("reset").addEventListener("click",()=>{const base=S.base,nap=S.showNap,nk=S.napKeuze;S=JSON.parse(JSON.stringify(DEF));S.base=base;S.showNap=nap;boutCb.checked=true;if(nk)S.napKeuze=true;save();fillSettings();buildRows();styleAll();});
const legend=document.getElementById("legend"),tg=document.getElementById("toggle");
tg.addEventListener("click",()=>{const c=legend.classList.toggle("collapsed");tg.textContent=c?"+":"–";tg.setAttribute("aria-expanded",String(!c));});
if(window.innerWidth<640){legend.classList.add("collapsed");tg.textContent="+";tg.setAttribute("aria-expanded","false");}
L.DomEvent.disableClickPropagation(legend);L.DomEvent.disableScrollPropagation(legend);
L.DomEvent.disableClickPropagation(document.getElementById("search"));
document.getElementById("src").innerHTML="Gegevens van "+(M.napUrl?"Rijkswaterstaat (peilmerken), ":"")+"de gemeente Amsterdam (meetbouten, bijgewerkt "+fmtDate(M.generated)+")"+
  (M.egmsUrl?" en Copernicus, het aardobservatieprogramma van de EU (satelliet, gemeten "+esc(String(M.egmsRelease||"").replace("-","–"))+")":"")+". "+
  M.count.toLocaleString("nl-NL")+" meetbouten, waarvan "+M.withSeries.toLocaleString("nl-NL")+" met metingen over minstens "+nf(M.minSpan,0)+" jaar. "+
  "Hoe snel een punt zakt, is berekend uit de laatste meting en een meting van minstens "+nf(M.minSpan,0)+" jaar eerder. "+
  "Weergave van openbare meetgegevens, zonder oordeel of advies."+(M.proef?" <b>Proefrun: onvolledige data.</b>":"");

// ---------- popups ----------
function chart(b,c){return chartSeries(b[I.ser],b[I.rfrom],c);}
function chartSeries(s,rf,c){
  if(!s||s.length<4)return"";
  const W=290,H=130,l=38,r=10,t=10,bt=22;const ts=[],cs=[];
  for(let k=0;k<s.length;k+=2){ts.push(s[k]+1900);cs.push(s[k+1]);}
  const t0=Math.floor(ts[0]),t1=Math.ceil(ts[ts.length-1])||t0+1;
  const cmin=Math.min(0,...cs),cmax=Math.max(1,...cs);
  const X=v=>l+(v-t0)/((t1-t0)||1)*(W-l-r),Y=v=>t+(v-cmin)/((cmax-cmin)||1)*(H-t-bt);
  const all=ts.map((v,k)=>X(v).toFixed(1)+","+Y(cs[k]).toFixed(1)).join(" ");
  let rec="";
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
    '<div class="sub">Meetbout '+esc(b[I.id])+' van de gemeente'+(b[I.status]&&b[I.status]!=="Actueel"?" · "+esc(b[I.status]):"")+'</div>'+pill(c)+
    '<table class="kv">'+
    '<tr><th>Laatste jaren</th><td>'+(b[I.recent]==null?"te weinig metingen":nf(b[I.recent])+" mm/jaar <span class=m>"+y(b[I.rfrom])+"–"+y(b[I.lastY])+"</span>")+'</td></tr>'+
    '<tr><th>Daarvoor</th><td>'+(b[I.early]==null?"–":nf(b[I.early])+" mm/jaar <span class=m>"+y(b[I.efrom])+"–"+y(b[I.rfrom])+"</span>")+'</td></tr>'+
    '<tr><th>Sinds 1e meting</th><td>'+(b[I.total]==null?"–":nf(b[I.total])+" mm/jaar")+'</td></tr>'+
    '<tr><th>Totaal gezakt</th><td>'+(b[I.cum]==null?"–":nf(b[I.cum],1)+" mm")+'</td></tr>'+
    '<tr><th>Metingen</th><td>'+b[I.n]+(b[I.last]?" · laatste "+fmtDate(b[I.last])+(age!=null?" ("+age+" jaar geleden)":""):"")+'</td></tr>'+
    '</table>'+omgRegel(omgeving(b[I.lat],b[I.lon],ROND_M,i,-1))+(b[I.ser]?chart(b,c):serState==="laden"?'<p class="mnote">Grafiek wordt geladen …</p>':serState==="fout"&&b[I.n]>1?'<p class="mnote">Grafiek kon niet worden geladen.</p>':"")+
    '<a class="raw" href="#" data-meas="'+i+'">Alle meetgegevens bekijken</a> · '+RONDA+'<a class="raw" href="?punt='+encodeURIComponent(b[I.id])+'" data-link>Link naar dit punt</a>'+boutMeer(i)+'<p class="disc">Openbare meetgegevens, zonder oordeel of advies. <a href="#over">Meer</a></p>'+'</div>';
  const p=L.popup({maxWidth:340,autoPanPadding:[20,20]}).setLatLng([b[I.lat],b[I.lon]]).setContent(html).openOn(map);
  curPop={i:i,p:p};
  const el=p.getElement&&p.getElement();
  if(el){const a=el.querySelector("a[data-meas]");if(a)a.addEventListener("click",ev=>{ev.preventDefault();openMeas(i);});}
  bindLink(el);rondLink(el,b[I.lat],b[I.lon],b[I.addr]||("meetbout "+b[I.id]),6);
  egPopup(el,b[I.lat],b[I.lon]);bagPopup(el,b[I.lat],b[I.lon],6);
}
// "Link naar dit punt": kopieert de link; lukt dat niet, dan gewoon openen.
function bindLink(el){
  const a=el&&el.querySelector("a[data-link]");if(!a)return;
  a.addEventListener("click",ev=>{
    let url;try{url=new URL(a.getAttribute("href"),location.href).href;}catch(e){return;}
    if(navigator.clipboard&&navigator.clipboard.writeText&&window.isSecureContext){
      ev.preventDefault();
      navigator.clipboard.writeText(url).then(()=>say("Link naar dit punt gekopieerd."),()=>{location.href=url;});
    }
  });
}
function houseNo(a){const m=String(a||"").match(/\s(\d+)/);return m?parseInt(m[1],10):1e9;}
function groupTitle(gi){return G[gi][0]||"Zonder adres";}
function openList(title,idx,latlng){
  idx=idx.slice().sort((a,b)=>houseNo(B[a][I.addr])-houseNo(B[b][I.addr])||String(B[a][I.addr]).localeCompare(String(B[b][I.addr]),"nl"));
  const v=idx.filter(i=>C[i]!=="oud").map(i=>B[i][I.recent]);const med=v.length?median(v):null;
  const n={rood:0,geel:0,groen:0,stab:0,oud:0};idx.forEach(i=>n[C[i]]++);
  const mix=KL.filter(k=>n[k]).map(k=>'<span style="white-space:nowrap"><i class="sw" style="display:inline-block;width:9px;height:9px;border-radius:50%;background:'+COL[k]+'"></i> '+n[k]+"</span>").join(" &nbsp;");
  const rows=idx.map(i=>'<tr data-i="'+i+'"><td><i class="sw" style="background:'+COL[C[i]]+'"></i></td><td>'+esc(B[i][I.addr]||B[i][I.id])+'</td><td class="n">'+(B[i][I.recent]==null?"–":nf(B[i][I.recent]))+'</td><td class="n m">'+esc((B[i][I.last]||"").slice(0,4))+"</td></tr>").join("");
  const html='<div class="pop"><div class="ttl">'+esc(title)+'</div><div class="sub">'+idx.length+(idx.length>1?" meetbouten":" meetbout")+(med==null?"":" · "+snelZin(med,v.length))+'</div>'+
    '<div>'+mix+'</div><div class="list"><table><tbody>'+rows+'</tbody></table></div><div class="sub" style="margin-top:4px">Per meetbout: mm per jaar de laatste jaren · jaar van de laatste meting. Klik een regel voor details.</div>'+straatLinks(idx)+'<p class="disc">Openbare meetgegevens, zonder oordeel of advies. <a href="#over">Meer</a></p>'+'</div>';
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
const SRCQ={gem:'<p class="mnote">Vragen over deze meetgegevens, of klopt er iets niet? Neem contact op met de gemeente Amsterdam: <a href="mailto:gebruik.basisinformatie@amsterdam.nl">gebruik.basisinformatie@amsterdam.nl</a>.</p>',
  rws:'<p class="mnote">Vragen over deze meetgegevens, of klopt er iets niet? Neem contact op met Rijkswaterstaat via het <a href="https://www.rijkswaterstaat.nl/formulieren/contactformulier-servicedesk-data" target="_blank" rel="noopener">contactformulier Servicedesk Data</a>.</p>'};
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
      '<p class="mnote">'+items.length+' metingen. Geel gemarkeerd: de periode waarover de recente snelheid is berekend. Zakking is positief als het punt lager is komen te liggen.</p>'+SRCQ.gem+
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
function sayHtml(h,ms){msg.innerHTML=h;msg.hidden=false;clearTimeout(say._t);say._t=setTimeout(()=>msg.hidden=true,ms||9000);}
document.getElementById("search").addEventListener("submit",e=>{
  e.preventDefault();
  const q=document.getElementById("q").value.trim();if(!q)return;
  TEL.tel("zoek");
  if(M.napUrl&&/^\d{3}[a-z]\d{4}$/i.test(q)){napFind(q);return;}
  if(/^\d{5,}$/.test(q)){const i=B.findIndex(b=>String(b[I.id])===q);if(i>=0){map.setView([B[i][I.lat],B[i][I.lon]],18);setTimeout(()=>openBout(i),300);return;}}
  // 1. een Amsterdamse straat uit de meetboutgegevens (exacte naam, geen andere plaats genoemd)
  const qa=q.replace(/[,\s]+amsterdam\s*$/i,"").trim();
  const anders=/,|\b\d{4}\s?[a-z]{2}\b/i.test(qa);
  const base=qa.toLowerCase().replace(/\s+\d.*$/,"").trim();
  if(!anders&&byStreet.has(base)){showStreet(base,qa);return;}
  // 2. elk adres of elke plaats in Nederland via de PDOK Locatieserver (overheid, open data)
  zoekPdok(q).then(ok=>{if(!ok&&!fuzzyStreet(qa))say("Niets gevonden voor “"+q+"”.");})
    .catch(()=>{if(!fuzzyStreet(qa))say("Zoeken lukte niet. Probeer het later nog eens.");});
});
function showStreet(key,q){
  const idx=byStreet.get(key);
  const num=q.match(/\s(\d+)/);
  if(num){const n=parseInt(num[1],10);let best=null,bd=1e9;for(const i of idx){const h=houseNo(B[i][I.addr]);if(Math.abs(h-n)<bd){bd=Math.abs(h-n);best=i;}}
    if(best!=null&&bd<=4){map.setView([B[best][I.lat],B[best][I.lon]],17);setTimeout(()=>openBout(best),300);toonRond(B[best][I.lat],B[best][I.lon],B[best][I.addr]||q,6);return;}}
  const bb=L.latLngBounds(idx.map(i=>[B[i][I.lat],B[i][I.lon]]));
  map.fitBounds(bb,{maxZoom:17,padding:[60,60]});
  setTimeout(()=>openList(nameOf.get(key),idx,bb.getCenter()),400);
}
function fuzzyStreet(q){
  const base=q.toLowerCase().replace(/\s+\d.*$/,"").trim();if(base.length<3)return false;
  const key=streetKeys.find(k=>k.startsWith(base))||streetKeys.find(k=>k.includes(base));
  if(!key)return false;showStreet(key,q);return true;
}
const PDOK_LS="https://api.pdok.nl/bzk/locatieserver/search/v3_1/free";
function zoekPdok(q){
  const u=PDOK_LS+"?rows=1&fl=weergavenaam,type,centroide_ll,gemeentecode,gemeentenaam&fq="+encodeURIComponent("type:(adres OR weg OR postcode OR woonplaats OR gemeente)")+"&q="+encodeURIComponent(q);
  return fetch(u).then(r=>{if(!r.ok)throw new Error("HTTP "+r.status);return r.json();}).then(j=>{
    const d=j&&j.response&&j.response.docs&&j.response.docs[0];if(!d||!d.centroide_ll)return false;
    const m=String(d.centroide_ll).match(/POINT\(\s*([-\d.]+)\s+([-\d.]+)\s*\)/);if(!m)return false;
    const z=({adres:17,postcode:16,weg:16,woonplaats:13,gemeente:12})[d.type]||15;
    const la=parseFloat(m[2]),lo=parseFloat(m[1]);
    map.setView([la,lo],z);
    const gem={code:d.gemeentecode,naam:d.gemeentenaam};
    if(d.type==="adres"||d.type==="postcode"||d.type==="weg")toonRond(la,lo,d.weergavenaam||q,d.type==="adres"?3:null,gem);
    else{rondSluit();const x=gebCode(gem.code,gem.naam||(d.type==="gemeente"?String(d.weergavenaam||"").replace(/^Gemeente\s+/i,""):""));
      if(x)sayHtml(esc(d.weergavenaam||q)+' · <a href="'+esc(x.u)+'">Cijfers voor gemeente '+esc(x.n)+' →</a>');else say(d.weergavenaam||q);}
    return true;
  });
}

// ---------- links naar de pagina's per gemeente en buurt ----------
const GB=M.gebied||null;
function gebGem(i){return GB&&i>=0&&GB.g[i]?{u:"gemeente/"+GB.g[i][0]+"/",n:GB.g[i][1]}:null;}
function gebBuurt(i){return GB&&GB.a>=0&&i>=0&&GB.b&&GB.b[i]?{u:"gemeente/"+GB.g[GB.a][0]+"/"+GB.b[i][0]+"/",n:GB.b[i][1]}:null;}
function gebBout(i){if(!GB)return null;const bi=GB.bout?GB.bout[i]:-1;return gebBuurt(bi)||gebGem(GB.a);}
function gebNap(i){return GB&&NAP.g?gebGem(NAP.g[i]):null;}
function gebCode(code,naam){
  if(!GB)return null;const c=String(code||"").replace(/\D/g,"").padStart(4,"0"),nm=String(naam||"").toLowerCase();
  let i=code?GB.g.findIndex(g=>g[2]&&g[2].replace(/\D/g,"").padStart(4,"0")===c):-1;
  if(i<0&&nm)i=GB.g.findIndex(g=>g[1].toLowerCase()===nm);
  return gebGem(i);
}
function slugJs(t){return String(t||"").normalize("NFKD").replace(/[^\x00-\x7f]/g,"").toLowerCase().replace(/[^a-z0-9]+/g,"-").replace(/^-+|-+$/g,"")||"gebied";}
// Straat in de lijst-pop-up: link naar de buurtpagina('s), naar de regel van die straat
function gebStraat(si){return GB&&GB.a>=0&&GB.s&&si>=0&&GB.s[si]?{u:"gemeente/"+GB.g[GB.a][0]+"/straat/"+GB.s[si][0]+"/",n:GB.s[si][1]}:null;}
function gebWijk(wi){return GB&&GB.a>=0&&GB.w&&wi>=0&&GB.w[wi]?{u:"gemeente/"+GB.g[GB.a][0]+"/wijk/"+GB.w[wi][0]+"/",n:"wijk "+GB.w[wi][1]}:null;}
function straatVanBout(i){return GB&&GB.st?gebStraat(GB.st[i]):null;}
function straatLinks(idx){
  if(!GB||!GB.bout)return"";const tel=new Map();
  const sp=straatVanBout(idx[0]);const eerst=sp?'<a href="'+esc(sp.u)+'">Alle meetbouten in de '+esc(sp.n)+', per huisnummer →</a>':"";
  for(const i of idx){const bi=GB.bout[i];tel.set(bi,(tel.get(bi)||0)+1);}
  const st=B[idx[0]][I.street];const anker=st?"#straat-"+slugJs(st):"";
  const uit=[...tel.entries()].sort((a,b)=>b[1]-a[1]).slice(0,3).map(([bi])=>{const x=gebBuurt(bi);
    if(x)return '<a href="'+esc(x.u+anker)+'">Cijfers voor '+esc(st||"deze straat")+' in '+esc(x.n)+' →</a>';
    const g=gebGem(GB.a);return g?'<a href="'+esc(g.u)+'">Cijfers voor '+esc(g.n)+' →</a>':"";}).filter(Boolean);
  const alle=(eerst?[eerst]:[]).concat([...new Set(uit)]);
  return alle.length?'<p class="meer">'+alle.join("<br>")+'</p>':"";
}
function meerHtml(x){return x?'<p class="meer"><a href="'+esc(x.u)+'">Cijfers voor '+esc(x.n)+' →</a></p>':"";}
function boutMeer(i){const sp=straatVanBout(i),x=gebBout(i);
  const d=[];if(sp)d.push('<a href="'+esc(sp.u)+'">Alle meetbouten in de '+esc(sp.n)+' →</a>');
  if(x)d.push('<a href="'+esc(x.u)+'">Cijfers voor '+esc(x.n)+' →</a>');
  return d.length?'<p class="meer">'+d.join("<br>")+'</p>':"";}
// gemeente op een plek: via het dichtstbijzijnde adres bij de PDOK Locatieserver (alleen als het nodig is)
const PDOK_REV="https://api.pdok.nl/bzk/locatieserver/search/v3_1/reverse";
function gebBuurtPlek(lat,lon){
  // in Amsterdam: de buurt van de dichtstbijzijnde meetbout (binnen 120 m)
  if(!GB)return null;const kx=111320*Math.cos(lat*Math.PI/180);let bi=-1,bd=120;
  for(let i=0;i<B.length;i++){const dx=(B[i][I.lon]-lon)*kx,dy=(B[i][I.lat]-lat)*111320,d=Math.hypot(dx,dy);if(d<bd){bd=d;bi=i;}}
  return bi>=0?gebBout(bi):null;
}
function gebPlek(lat,lon){
  if(!GB)return Promise.resolve(null);
  // in Amsterdam: de buurt van de dichtstbijzijnde meetbout (binnen 120 m)
  const kx=111320*Math.cos(lat*Math.PI/180);let bi=-1,bd=120;
  for(let i=0;i<B.length;i++){const dx=(B[i][I.lon]-lon)*kx,dy=(B[i][I.lat]-lat)*111320,d=Math.hypot(dx,dy);if(d<bd){bd=d;bi=i;}}
  if(bi>=0){const x=gebBout(bi);if(x)return Promise.resolve(x);}
  return fetch(PDOK_REV+"?lat="+lat.toFixed(6)+"&lon="+lon.toFixed(6)+"&type=adres&rows=1&fl=gemeentecode,gemeentenaam")
    .then(r=>{if(!r.ok)throw new Error("HTTP "+r.status);return r.json();})
    .then(j=>{const d=j&&j.response&&j.response.docs&&j.response.docs[0];return d?gebCode(d.gemeentecode,d.gemeentenaam):null;})
    .catch(()=>null);
}

// ---------- omgeving: meetpunten binnen een straal ----------
const ROND_M=250;
function omgeving(lat,lon,r,skipB,skipN){
  const kx=111320*Math.cos(lat*Math.PI/180),ky=111320;
  const nieuw=()=>({n:0,k:{rood:0,geel:0,groen:0,stab:0,oud:0},v:[],laatste:0,dichtst:Infinity});
  const res={b:nieuw(),p:nieuw(),napKlaar:!!NAP.rows};
  for(let i=0;i<B.length;i++){
    if(i===skipB)continue;const b=B[i];
    const dx=(b[I.lon]-lon)*kx,dy=(b[I.lat]-lat)*ky,d=Math.sqrt(dx*dx+dy*dy);
    if(d<res.b.dichtst)res.b.dichtst=d;
    if(d>r)continue;const c=C[i]||"oud",t=res.b;t.n++;t.k[c]++;
    if(c!=="oud"&&b[I.recent]!=null)t.v.push(b[I.recent]);if(b[I.lastY])t.laatste=Math.max(t.laatste,Math.floor(b[I.lastY]));
  }
  if(NAP.rows)for(let i=0;i<NAP.rows.length;i++){
    if(i===skipN)continue;const q=NAP.rows[i];
    const dx=(q[NI.lon]-lon)*kx,dy=(q[NI.lat]-lat)*ky,d=Math.sqrt(dx*dx+dy*dy);
    if(d<res.p.dichtst)res.p.dichtst=d;
    if(d>r)continue;const c=napCls(q),t=res.p;t.n++;t.k[c]++;
    if(c!=="oud"&&q[NI.recent]!=null)t.v.push(q[NI.recent]);if(q[NI.lastY])t.laatste=Math.max(t.laatste,Math.floor(q[NI.lastY]));
  }
  return res;
}
// "zakt 1,2 mm per jaar" / "zakken doorgaans …" / "stijgt …"
function snelZin(v,n){const meer=n>1,st=v<0;return (st?(meer?"stijgen":"stijgt"):(meer?"zakken":"zakt"))+(meer?" doorgaans ":" ")+nf(Math.abs(v))+" mm per jaar";}
function afstand(m){return m<1000?Math.round(m/10)*10+" m":nf(m/1000,1)+" km";}
function omgRegel(o){
  const d=[];
  if(o.b.v.length)d.push(o.b.v.length+" andere meetbout"+(o.b.v.length>1?"en ":" ")+snelZin(median(o.b.v),o.b.v.length));
  if(o.p.v.length)d.push(o.p.v.length+" ander"+(o.p.v.length>1?"e peilmerken ":" peilmerk ")+snelZin(median(o.p.v),o.p.v.length));
  return d.length?'<div class="omg">In de buurt (binnen '+ROND_M+' m): '+d.join("; ")+'.</div>':"";
}
const ROND={el:document.getElementById("rond"),cirkel:null};
// De stippelcirkel staat in een eigen laag zonder klikken, zodat meetpunten eronder klikbaar blijven.
if(map.createPane){map.createPane("rond");const rp=map.getPane("rond");if(rp){rp.style.zIndex=450;rp.style.pointerEvents="none";}}
ROND.rend=L.svg({pane:"rond"});
function rondSluit(){ROND.el.hidden=true;ROND.terug=false;ROND.pandVorm=false;if(typeof bagVormWeg==="function")bagVormWeg();if(ROND.cirkel){map.removeLayer(ROND.cirkel);ROND.cirkel=null;}}
function toonRond(lat,lon,naam,pandM,gem){
  if(ROND.cirkel)map.removeLayer(ROND.cirkel);
  ROND.pandVorm=false;bagVormWeg();ROND.geb=null;
  ROND.cirkel=L.circle([lat,lon],{radius:ROND_M,color:RING,weight:2,dashArray:"5 5",fill:false,interactive:false,renderer:ROND.rend,pane:"rond"}).addTo(map);
  const blok=(t,soort,meer,enkel)=>{
    if(!t.n)return '<p class="rs"><b>'+soort+'</b>: geen binnen '+ROND_M+' m'+(isFinite(t.dichtst)?' (dichtstbijzijnde op '+afstand(t.dichtst)+')':'')+'.</p>';
    const bar=KL.filter(k=>t.k[k]).map(k=>'<i style="width:'+(100*t.k[k]/t.n).toFixed(1)+'%;background:'+COL[k]+'" title="'+esc(label(k)[0])+': '+t.k[k]+'"></i>').join("");
    return '<p class="rs"><b>'+t.n+' '+(t.n>1?meer:enkel)+'</b>'+(t.v.length?' · '+snelZin(median(t.v),t.v.length)+(t.v.length<t.n?' (snelheid van '+t.v.length+' bekend)':''):(t.n>1?' · van geen enkele is een recente snelheid bekend':' · geen recente snelheid bekend'))+
      '<span class="rbar">'+bar+'</span><span class="rk">'+KL.filter(k=>t.k[k]).map(k=>esc(label(k)[0])+' '+t.k[k]).join(" · ")+(t.laatste?' · laatste meting '+t.laatste:'')+'</span></p>';
  };
  const teken=()=>{
    const o=omgeving(lat,lon,ROND_M);
    ROND.el.innerHTML='<div class="rt"><div><h2>Rond '+esc(naam)+'</h2><div class="rk">Alle meetpunten binnen '+ROND_M+' meter (de stippellijn op de kaart)</div></div><button class="x" type="button" aria-label="Sluiten">×</button></div>'+
      (pandM!=null?'<div id="rondPand"></div>':"")+
      ((o.b.n||o.b.dichtst<3000)?blok(o.b,"Meetbouten","meetbouten","meetbout"):"")+(M.napUrl?(o.napKlaar?blok(o.p,"Peilmerken","peilmerken","peilmerk"):'<p class="rs rk">Peilmerken laden …</p>'):"")+
      (M.egmsUrl?'<div id="rondEg"><p class="rs rk">Satellietmetingen laden …</p></div>':"")+
      '<div id="rondGeb"></div><p class="rs rk">Gemeten snelheden, zonder oordeel of advies. Klik op een punt voor details.</p>';
    ROND.el.querySelector(".x").addEventListener("click",rondSluit);
    ROND.el.hidden=false;
    (ROND.geb?Promise.resolve(ROND.geb):(gebBuurtPlek(lat,lon)||(gem&&gebCode(gem.code,gem.naam)))?Promise.resolve(gebBuurtPlek(lat,lon)||gebCode(gem.code,gem.naam)):gebPlek(lat,lon)).then(x=>{
      ROND.geb=x;const box=document.getElementById("rondGeb");if(x&&box&&!ROND.el.hidden)box.innerHTML=meerHtml(x);});
    if(pandM!=null)bagPand(lat,lon,pandM).then(p=>{const box=document.getElementById("rondPand");
      if(!p||!box||ROND.el.hidden)return;box.innerHTML=bagTekst(p);ROND.pandVorm=true;bagVorm(p.geom);});
    if(M.egmsUrl)egRond(lat,lon,ROND_M).then(e=>{
      const box=document.getElementById("rondEg");if(!box)return;
      if(!e||!e.n){box.innerHTML='<p class="rs"><b>Satelliet</b>: hier is binnen '+ROND_M+' m niets gemeten.</p>';return;}
      const bar=KL.filter(k=>e.k[k]).map(k=>'<i style="width:'+(100*e.k[k]/e.n).toFixed(1)+'%;background:'+COL[k]+'"></i>').join("");
      box.innerHTML='<p class="rs"><b>Satelliet: '+e.n+' blok'+(e.n>1?'ken':'')+' van 100 m</b>'+(e.v.length?' · '+egZin(median(e.v)):'')+
        '<span class="rbar">'+bar+'</span><span class="rk">Grond, straten en gebouwen samen, gemeten vanuit de ruimte ('+esc(String(M.egmsRelease||"").replace("-","–"))+').</span></p>';
    });
    return o.napKlaar;
  };
  if(!teken()&&M.napUrl){if(!S.showNap){S.showNap=true;save();}napRefresh(true);
    const t=setInterval(()=>{if(NAP.rows){clearInterval(t);if(!ROND.el.hidden)teken();}else if(NAP.state==="fout")clearInterval(t);},200);}
}
L.DomEvent.disableClickPropagation(ROND.el);L.DomEvent.disableScrollPropagation(ROND.el);
map.on("contextmenu",e=>{if(e.originalEvent&&e.originalEvent.preventDefault)e.originalEvent.preventDefault();
  map.closePopup();toonRond(e.latlng.lat,e.latlng.lng,"dit punt",2);});
// link 'In de buurt' in een pop-up
function rondLink(el,lat,lon,naam,pandM){
  if(!el)return;const a=el.querySelector("a[data-rond]");if(!a)return;
  a.addEventListener("click",ev=>{ev.preventDefault();map.closePopup();toonRond(lat,lon,naam,pandM);});
}
const RONDA='<a class="raw" href="#" data-rond>Alles binnen '+ROND_M+' m</a> · ';
// Op mobiel is er geen ruimte voor paneel én pop-up: paneel even weg zolang een pop-up open is.
map.on("popupopen",()=>{if(!ROND.el.hidden&&matchMedia("(max-width:640px)").matches){ROND.el.hidden=true;ROND.terug=true;}});
// Op een telefoon schuift de inhoud van een pop-up van onderen in beeld (een 'sheet'); het punt blijft zichtbaar.
const SHEET={el:document.getElementById("sheet"),mob:()=>matchMedia("(max-width:640px)").matches};
if(SHEET.el){
  const origPan=L.Popup.prototype._adjustPan;
  L.Popup.prototype._adjustPan=function(){if(SHEET.mob())return;return origPan&&origPan.apply(this,arguments);};
  map.on("popupopen",e=>{
    if(!SHEET.mob())return;const c=e.popup.getElement();if(!c)return;SHEET.cur=e.popup;
    SHEET.el.querySelector(".body").appendChild(c);SHEET.el.hidden=false;document.body.classList.add("sheet");
    SHEET.el.querySelector(".body").scrollTop=0;
    // het punt boven de sheet houden
    requestAnimationFrame(()=>{const h=SHEET.el.offsetHeight,H=map.getSize().y,pt=map.latLngToContainerPoint(e.popup.getLatLng());
      const doel=Math.max(60,(H-h)/2);if(pt.y>H-h-30||pt.y<50)map.panBy([0,pt.y-doel],{animate:true});});
  });
  map.on("popupclose",e=>{if(e.popup!==SHEET.cur)return;SHEET.cur=null;
    const c=e.popup.getElement();if(c&&c.parentNode===SHEET.el.querySelector(".body"))c.remove();
    SHEET.el.hidden=true;document.body.classList.remove("sheet");});
  SHEET.el.querySelector(".dicht").addEventListener("click",()=>map.closePopup());
  L.DomEvent.disableClickPropagation(SHEET.el);L.DomEvent.disableScrollPropagation(SHEET.el);
}
map.on("popupclose",()=>{if(ROND.terug&&ROND.cirkel){ROND.el.hidden=false;}ROND.terug=false;});
document.addEventListener("keydown",e=>{if(e.key==="Escape"&&!ROND.el.hidden&&["modal","aboutModal","fbModal","intro"].every(id=>{const x=document.getElementById(id);return !x||x.hidden;}))rondSluit();});

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

// ---------- NAP-peilmerken (Rijkswaterstaat, heel Nederland) ----------
// Lui geladen: pas als de laag aan staat. Getekend binnen beeld vanaf zoom NAP.MINZ.
const NI={p:0,lat:1,lon:2,oms:3,status:4,n:5,first:6,last:7,lastY:8,recent:9,rfrom:10,total:11,sprong:12};
const NAP={rows:null,ser:null,serState:"",state:"",layer:L.layerGroup(),mk:new Map(),shown:new Set(),MINZ:11,STALE:20,cur:null,z:null,ver:0};
if(map.createPane){map.createPane("napgrid");const gp=map.getPane&&map.getPane("napgrid");if(gp)gp.style.zIndex=350;}
const GRID={layer:L.layerGroup(),rend:L.canvas({pane:"napgrid",padding:0.3}),cells:{}};
const napRow=document.getElementById("naprow"),napCb=document.getElementById("cNap"),napNoteEl=document.getElementById("napnote");
if(M.napUrl){napRow.hidden=false;napCb.checked=!!S.showNap;
  napCb.addEventListener("change",()=>{S.showNap=napCb.checked;S.napKeuze=true;save();napRefresh(true);});}
function napNote(t){napNoteEl.textContent=t||"";}
function napCls(r){
  if(r[NI.recent]==null||r[NI.sprong])return"oud";
  if(M.nowYear-r[NI.lastY]>NAP.STALE)return"oud";
  const v=r[NI.recent];return v>=S.red?"rood":v>=S.yellow?"geel":"groen";
}
function napLoad(){
  if(NAP.state==="laden")return;NAP.state="laden";napNote("Peilmerken laden …");
  fetch(M.napUrl).then(r=>{if(!r.ok)throw new Error("HTTP "+r.status);return r.json();})
  .then(d=>{NAP.rows=d.p||[];NAP.g=Array.isArray(d.g)&&d.g.length===NAP.rows.length?d.g:null;NAP.serUrl=d.reeksen;NAP.opgehaald=d.opgehaald||"";NAP.state="ok";napRefresh(true);})
  .catch(e=>{NAP.state="fout";napNote("De peilmerken konden niet worden geladen ("+(e&&e.message||e)+").");});
}
function napLoadSer(){
  if(NAP.ser||NAP.serState==="laden"||!NAP.serUrl)return;NAP.serState="laden";
  fetch(NAP.serUrl).then(r=>{if(!r.ok)throw new Error("HTTP "+r.status);return r.json();})
  .then(a=>{NAP.ser=a;NAP.serState="ok";if(NAP.cur&&map.hasLayer(NAP.cur.p))napOpen(NAP.cur.i);})
  .catch(()=>{NAP.serState="fout";if(NAP.cur&&map.hasLayer(NAP.cur.p))napOpen(NAP.cur.i);});
}
function napStyle(m,i){
  const c=napCls(NAP.rows[i]),z=map.getZoom();
  m.setStyle({radius:z>=16?7:z>=14?6:5,color:COL[c],weight:z>=14?3:2.5,opacity:c==="oud"?.75:1,fillColor:"#ffffff",fillOpacity:1});m._zkVer=NAP.ver;
}
function napRefresh(restyle){
  if(!M.napUrl)return;
  const on=!!S.showNap;napCb.checked=on;
  const cnt=document.getElementById("cnt-nap");
  if(!on){if(map.hasLayer(NAP.layer))map.removeLayer(NAP.layer);napGridClear();napNote("");cnt.textContent="–";return;}
  if(!NAP.rows){napLoad();return;}
  if(!map.hasLayer(NAP.layer))NAP.layer.addTo(map);
  const z=map.getZoom();if(z!==NAP.z){NAP.z=z;restyle=true;}
  if(restyle)NAP.ver++;
  if(z<NAP.MINZ){NAP.layer.clearLayers();NAP.shown=new Set();napGrid(z,cnt);return;}
  napGridClear();napNote("");
  const bb=map.getBounds(),wide=bb.pad?bb.pad(0.25):bb,want=new Set();let inView=0;
  for(let i=0;i<NAP.rows.length;i++){const r=NAP.rows[i];if(wide.contains([r[NI.lat],r[NI.lon]])){want.add(i);if(bb.contains([r[NI.lat],r[NI.lon]]))inView++;}}
  for(const i of NAP.shown)if(!want.has(i))NAP.layer.removeLayer(NAP.mk.get(i));
  for(const i of want){
    let m=NAP.mk.get(i);
    if(!m){const r=NAP.rows[i];m=L.circleMarker([r[NI.lat],r[NI.lon]],{renderer:rend});
      m.on("click",e=>{L.DomEvent.stopPropagation(e);napOpen(i);});NAP.mk.set(i,m);}
    if(m._zkVer!==NAP.ver)napStyle(m,i);
    if(!NAP.shown.has(i))NAP.layer.addLayer(m);
  }
  NAP.shown=want;cnt.textContent=inView.toLocaleString("nl-NL");
}
map.on("moveend",()=>napRefresh(false));
// Uitgezoomd: vakjes met de mediaan van de recente zakking van de peilmerken erin.
function napCells(zf){
  if(GRID.cells[zf])return GRID.cells[zf];
  const d=({7:0.16,8:0.08,9:0.04,10:0.02})[zf]||(zf<7?0.32:0.02),dl=d*1.6,cells=new Map();
  for(let i=0;i<NAP.rows.length;i++){const r=NAP.rows[i];
    const a=Math.floor(r[NI.lat]/d),o=Math.floor(r[NI.lon]/dl),k=a+":"+o;
    let c=cells.get(k);if(!c){c={s:a*d,w:o*dl,n:a*d+d,e:o*dl+dl,idx:[]};cells.set(k,c);}c.idx.push(i);}
  return GRID.cells[zf]=[...cells.values()];
}
function napGridClear(){GRID.layer.clearLayers();if(map.hasLayer(GRID.layer))map.removeLayer(GRID.layer);}
function napGrid(z,cnt){
  const zf=Math.floor(z),bb=map.getBounds(),wide=bb.pad?bb.pad(0.2):bb;let n=0;
  GRID.layer.clearLayers();if(!map.hasLayer(GRID.layer))GRID.layer.addTo(map);
  for(const c of napCells(zf)){
    if(!wide.contains([(c.s+c.n)/2,(c.w+c.e)/2]))continue;
    const vals=[];for(const i of c.idx){const r=NAP.rows[i];if(napCls(r)!=="oud")vals.push(r[NI.recent]);}
    n+=c.idx.length;
    const med=vals.length?median(vals):null,k=med==null?"oud":med>=S.red?"rood":med>=S.yellow?"geel":"groen";
    const rc=L.rectangle([[c.s,c.w],[c.n,c.e]],{renderer:GRID.rend,stroke:false,fillColor:COL[k],fillOpacity:k==="oud"?.22:.5});
    rc.on("click",e=>{L.DomEvent.stopPropagation(e);map.setView([(c.s+c.n)/2,(c.w+c.e)/2],NAP.MINZ);});
    rc._zkCell={n:c.idx.length,med:med,k:k};
    GRID.layer.addLayer(rc);
  }
  cnt.textContent=n.toLocaleString("nl-NL");
  napNote("Uitgezoomd: per vakje de doorsnee van de peilmerken erin. Klik op een vakje of zoom in voor losse peilmerken.");
}
function napOpen(i){
  const r=NAP.rows[i],c=napCls(r),y=v=>v==null?"":Math.floor(v);
  if(!NAP.cur||NAP.cur.i!==i)TEL.tel("punt");
  const age=r[NI.lastY]!=null?Math.max(0,Math.round(M.nowYear-r[NI.lastY])):null;
  const s=NAP.ser?NAP.ser[i]:null;
  const ch=s?chartSeries(s,r[NI.rfrom],c):NAP.serState==="fout"?'<p class="mnote">Grafiek kon niet worden geladen.</p>':r[NI.n]>1?'<p class="mnote">Grafiek wordt geladen …</p>':"";
  const lbl=c==="oud"?(r[NI.sprong]?"Onzekere metingen":"Geen recente metingen of te weinig metingen"):label(c)[0];
  const html='<div class="pop"><div class="ttl">Peilmerk '+esc(r[NI.p])+' <span style="font-weight:400;color:var(--muted)">(Rijkswaterstaat)</span></div>'+
    '<div class="sub">'+esc(r[NI.oms]||"")+'</div>'+
    '<div class="pill"><i style="background:#fff;box-shadow:inset 0 0 0 2.5px '+COL[c]+'"></i>'+esc(lbl)+'</div>'+
    '<table class="kv">'+
    '<tr><th>Laatste jaren</th><td>'+(r[NI.recent]==null?"te weinig metingen":nf(r[NI.recent])+" mm/jaar <span class=m>"+y(r[NI.rfrom])+"–"+y(r[NI.lastY])+"</span>")+'</td></tr>'+
    '<tr><th>Sinds 1e meting</th><td>'+(r[NI.total]==null?"–":nf(r[NI.total])+" mm/jaar <span class=m>"+esc((r[NI.first]||"").slice(0,4))+"–"+esc((r[NI.last]||"").slice(0,4))+"</span>")+'</td></tr>'+
    '<tr><th>Metingen</th><td>'+r[NI.n]+(r[NI.last]?" · laatste "+fmtDate(r[NI.last])+(age!=null?" ("+age+" jaar geleden)":""):"")+'</td></tr>'+
    '</table>'+omgRegel(omgeving(r[NI.lat],r[NI.lon],ROND_M,-1,i))+(r[NI.sprong]?'<p class="mnote">In de metingen zit een plotselinge sprong die niet te verklaren is (misschien een meetfout of een andere rekenmethode). Daarom is dit punt grijs.</p>':"")+ch+
    '<p class="mnote">Een peilmerk is een metalen boutje in een gebouw of brug. Rijkswaterstaat meet eens per 5 tot 15 jaar hoe hoog het ligt. Het zakt mee met de grond in de omgeving.</p>'+
    (M.napService?'<a class="raw" href="#" data-nmeas="'+i+'">Alle meetgegevens bekijken</a> · ':"")+RONDA+'<a class="raw" href="?punt='+encodeURIComponent(r[NI.p])+'" data-link>Link naar dit punt</a>'+
    '<div class="sub" style="margin-top:4px">Bron: Rijkswaterstaat</div>'+meerHtml(gebNap(i))+'<p class="disc">Openbare meetgegevens, zonder oordeel of advies. <a href="#over">Meer</a></p>'+'</div>';
  const p=L.popup({maxWidth:340,autoPanPadding:[20,20]}).setLatLng([r[NI.lat],r[NI.lon]]).setContent(html).openOn(map);
  NAP.cur={i:i,p:p};
  const el=p.getElement&&p.getElement();
  if(el){const a=el.querySelector("a[data-nmeas]");if(a)a.addEventListener("click",ev=>{ev.preventDefault();napMeas(i);});}
  bindLink(el);rondLink(el,r[NI.lat],r[NI.lon],"peilmerk "+r[NI.p],6);
  egPopup(el,r[NI.lat],r[NI.lon]);bagPopup(el,r[NI.lat],r[NI.lon],6);
  if(!NAP.ser)napLoadSer();
}
function napQueryUrl(p,fmt){
  return M.napService+"/4/query?"+new URLSearchParams({where:"puntnummer='"+String(p).replace(/'/g,"''")+"'",
    outFields:"puntnummer,project_id,project_titel,projectdatum,hoogte,orde",orderByFields:"projectdatum",returnGeometry:"false",f:fmt}).toString();
}
function napSummary(r){
  const c=napCls(r),y=v=>v==null?"":Math.floor(v);
  const lbl=c==="oud"?(r[NI.sprong]?"Onzekere metingen":"Geen recente metingen of te weinig metingen"):label(c)[0];
  return '<div class="msum">'+
    '<span>Laatste jaren: <b>'+(r[NI.recent]==null?"–":nf(r[NI.recent])+" mm/jaar")+'</b> '+(r[NI.rfrom]!=null?'<span class="m">'+y(r[NI.rfrom])+"–"+y(r[NI.lastY])+"</span>":"")+'</span>'+
    '<span>Sinds 1e meting: <b>'+(r[NI.total]==null?"–":nf(r[NI.total])+" mm/jaar")+'</b></span>'+
    '<span>Klasse: <b>'+esc(lbl)+'</b></span></div>';
}
function napRawSection(link,err){
  const a='<a href="'+esc(link)+'" target="_blank" rel="noopener">Open bij Rijkswaterstaat ↗</a>';
  if(err)return '<h3>Ruwe gegevens van Rijkswaterstaat</h3><p class="mnote">Kon de ruwe gegevens niet rechtstreeks ophalen ('+esc(err)+'). '+a+'</p>';
  return '<h3>Ruwe gegevens van Rijkswaterstaat</h3><p class="mnote" style="margin:0 0 6px">Precies zoals Rijkswaterstaat ze levert (JSON). '+a+'</p><pre class="rawjson" id="rawjson"></pre>';
}
function napMeas(i){
  TEL.tel("meet");
  const r=NAP.rows[i];lastFocus=document.activeElement;
  mTitle.textContent="Peilmerk "+r[NI.p]+" (Rijkswaterstaat)";
  mSub.textContent=[r[NI.oms],r[NI.status]].filter(Boolean).join(" · ");
  mBody.innerHTML='<p class="mnote">Meetgegevens ophalen bij Rijkswaterstaat …</p>';
  modal.hidden=false;document.getElementById("mClose").focus();
  const url=napQueryUrl(r[NI.p],"json"),link=napQueryUrl(r[NI.p],"html"),rf=r[NI.rfrom];
  fetch(url).then(x=>{if(!x.ok)throw new Error("HTTP "+x.status);return x.json();})
  .then(data=>{
    if(modal.hidden)return;
    if(data&&data.error)throw new Error(data.error.message||"fout bij Rijkswaterstaat");
    const items=((data&&data.features)||[]).map(f=>f.attributes||{}).filter(a=>a.projectdatum!=null)
      .sort((a,b)=>a.projectdatum-b.projectdatum||String(a.project_id).localeCompare(String(b.project_id)));
    const herz=items.some(a=>String(a.project_id||"").indexOf("=")>=0);
    let h0=null,prev=null,nUsed=0;
    const rows=items.map(a=>{
      const iso=new Date(a.projectdatum+432e5).toISOString().slice(0,10),t=yearFracOf(iso);
      const u=!herz||String(a.project_id||"").indexOf("=")>=0;
      let dz="–",tot="–";
      if(u&&a.hoogte!=null){nUsed++;if(h0==null)h0=a.hoogte;tot=nf((h0-a.hoogte)*1000,1);if(prev!=null)dz=nf((prev-a.hoogte)*1000,1);prev=a.hoogte;}
      const cls=!u?"oud":(rf!=null&&t!=null&&t>=rf-0.006?"rec":"");
      return '<tr'+(cls?' class="'+cls+'"':'')+'><td>'+fmtDate(iso)+'</td><td class="n">'+(a.hoogte==null?"–":nf(a.hoogte,4))+
        '</td><td class="n">'+dz+'</td><td class="n">'+tot+'</td><td>'+esc(a.project_id||"–")+'</td><td>'+esc(a.project_titel||"–")+
        '</td><td class="n">'+(a.orde==null?"–":esc(a.orde))+'</td><td>'+(u?"ja":"nee")+'</td></tr>';}).join("");
    mBody.innerHTML=napSummary(r)+
      '<div class="mt-wrap"><table class="mt"><thead><tr><th>Datum</th><th class="n">Hoogte t.o.v. NAP (m)</th><th class="n">Zakking sinds vorige (mm)</th><th class="n">Totaal gezakt (mm)</th><th>Project</th><th>Projecttitel</th><th class="n">Orde</th><th>Gebruikt</th></tr></thead><tbody>'+
      (rows||'<tr><td colspan="8">Geen metingen gevonden.</td></tr>')+'</tbody></table></div>'+
      '<p class="mnote">'+items.length+' metingen'+(items.length!==nUsed?', waarvan '+nUsed+' gebruikt in de berekening':'')+'. Geel gemarkeerd: de periode waarover de recente snelheid is berekend.'+
      (herz&&items.length!==nUsed?' Grijs: oude waarden uit vóór de herziening van het NAP-stelsel. Rijkswaterstaat heeft die metingen later herrekend; de kaart gebruikt alleen de herrekende waarden (project met &ldquo;=&rdquo;), anders ontstaan schijnbare sprongen van centimeters.':'')+
      ' Zakking is positief als het punt lager is komen te liggen. Metingen binnen één jaar tellen in de berekening als één meetcampagne.</p>'+SRCQ.rws+
      napRawSection(link,null);
    const pre=document.getElementById("rawjson");if(pre)pre.textContent=JSON.stringify(data,null,2);
  })
  .catch(err=>{
    if(modal.hidden)return;
    const s=NAP.ser?NAP.ser[i]:null,rows=[];
    for(let k=0;s&&k<s.length;k+=2){const t=s[k]+1900,y=Math.floor(t),d=new Date(Date.UTC(y,0,1)+(t-y)*365.25*864e5);
      rows.push('<tr><td>± '+d.toLocaleDateString("nl-NL",{month:"short",year:"numeric",timeZone:"UTC"})+'</td><td class="n">'+nf(s[k+1],1)+'</td></tr>');}
    mBody.innerHTML='<div class="warn">De actuele gegevens konden niet bij Rijkswaterstaat worden opgehaald. Hieronder de meetreeks zoals die bij het maken van deze kaart is opgeslagen (per meetcampagne, datum bij benadering).</div>'+
      napSummary(r)+
      '<div class="mt-wrap"><table class="mt"><thead><tr><th>Datum</th><th class="n">Totaal gezakt (mm)</th></tr></thead><tbody>'+(rows.join("")||'<tr><td colspan="2">Geen meetreeks opgeslagen.</td></tr>')+'</tbody></table></div>'+
      napRawSection(link,String(err&&err.message||err));
  });
}
function napFind(q){
  const go=()=>{const k=NAP.rows.findIndex(r=>String(r[NI.p]).toUpperCase()===q.toUpperCase());
    if(k<0){say("Peilmerk “"+q+"” niet gevonden.");return;}
    map.setView([NAP.rows[k][NI.lat],NAP.rows[k][NI.lon]],17);setTimeout(()=>{napRefresh(true);napOpen(k);},300);};
  if(!S.showNap){S.showNap=true;save();}
  if(NAP.rows)go();else{napRefresh(true);const t=setInterval(()=>{if(NAP.rows){clearInterval(t);go();}else if(NAP.state==="fout")clearInterval(t);},150);}
}

// ---------- EGMS: bodembeweging gemeten door satellieten (Copernicus) ----------
// Uit: niets geladen. Aan: uitgezoomd vakken van ~5 en ~1 km (mediaan), ingezoomd (zoom 14+) de vakken van 100 m.
const EG={ov:null,state:"",tegels:new Map(),wacht:new Map(),layer:L.layerGroup(),rend:null,DETAIL:14};
if(map.createPane){map.createPane("egms");const ep=map.getPane&&map.getPane("egms");if(ep){ep.style.zIndex=330;}}
EG.rend=L.canvas({pane:"egms",padding:0.3});
const egRow=document.getElementById("egrow"),egCb=document.getElementById("cEgms"),egNoteEl=document.getElementById("egnote");
if(M.egmsUrl){egRow.hidden=false;egCb.checked=!!S.showEgms;
  egCb.addEventListener("change",()=>{S.showEgms=egCb.checked;save();egRefresh(true);});}
function egNote(t){egNoteEl.textContent=t||"";}
function egCls(v,onz){if(onz)return"oud";const z=v/10;return z>=S.red?"rood":z>=S.yellow?"geel":"groen";}
function egTekst(v){const z=v/10;return Math.abs(z)<0.05?"beweegt niet meetbaar":z>0?"zakt "+nf(z,1)+" mm/jaar":"stijgt "+nf(-z,1)+" mm/jaar";}
function egLaad(){
  if(EG.state==="laden"||EG.state==="ok")return Promise.resolve();
  EG.state="laden";egNote("Satellietgegevens laden …");
  return fetch(M.egmsUrl).then(r=>{if(!r.ok)throw new Error("HTTP "+r.status);return r.json();})
    .then(d=>{EG.ov=d;EG.set=new Set(d.tegels||[]);EG.state="ok";egNote("");egRefresh(true);})
    .catch(()=>{EG.state="fout";egNote("De satellietgegevens konden niet worden geladen.");});
}
function egTegel(k){
  if(EG.tegels.has(k))return Promise.resolve(EG.tegels.get(k));
  if(EG.wacht.has(k))return EG.wacht.get(k);
  const pr=fetch("egms/"+k+".json").then(r=>{if(!r.ok)throw new Error("HTTP "+r.status);return r.json();})
    .then(a=>{EG.tegels.set(k,a);EG.wacht.delete(k);return a;}).catch(()=>{EG.wacht.delete(k);return [];});
  EG.wacht.set(k,pr);return pr;
}
function egSleutels(bb){
  const d=EG.ov.tegel,out=[];
  for(let x=Math.floor(bb.getWest()/d);x<=Math.floor(bb.getEast()/d);x++)
    for(let y=Math.floor(bb.getSouth()/d);y<=Math.floor(bb.getNorth()/d);y++){const k=x+"_"+y;if(EG.set.has(k))out.push(k);}
  return out;
}
function egVak(lat,lon,hy,hx,v,onz,info){
  const c=egCls(v,onz);
  const r=L.rectangle([[lat-hy,lon-hx],[lat+hy,lon+hx]],{renderer:EG.rend,stroke:false,fillColor:COL[c],fillOpacity:c==="oud"?.35:.5});
  r.on("click",e=>{L.DomEvent.stopPropagation(e);egOpen(lat,lon,v,onz,info);});
  return r;
}
function egOpen(lat,lon,v,onz,info){
  const html='<div class="pop"><div class="ttl">Bodemdaling gemeten vanuit de ruimte</div>'+
    '<div class="sub">'+(info?"Doorsnee van "+info+" blokken van 100 m":"Blok van 100 × 100 meter")+' · gemeten '+esc(String(M.egmsRelease||"").replace("-","–"))+'</div>'+
    '<p style="margin:6px 0 0;font-size:14px"><b>'+(onz?"Onzekere meting":egTekst(v))+'</b></p>'+
    '<p class="mnote">Satellieten meten met radar de hoogte van grond, straten en gebouwen samen; per blok is dat een gemiddelde. '+
    'Een huis op palen kan stil staan terwijl de straat ervoor zakt, en andersom. Meetbouten en peilmerken meten wél één punt aan één gebouw.</p>'+
    '<div class="sub" style="margin-top:4px">Bron: Copernicus, het aardobservatieprogramma van de EU ('+esc(M.egmsBron||"EGMS")+')</div>'+
    (info?'':'<a class="raw" href="#" data-rond>Alles binnen '+ROND_M+' m</a>')+
    '<p class="disc">Openbare meetgegevens, zonder oordeel of advies. <a href="#over">Meer</a></p></div>';
  const pp=L.popup({maxWidth:320,autoPanPadding:[20,20]}).setLatLng([lat,lon]).setContent(html).openOn(map);
  if(!info)rondLink(pp.getElement&&pp.getElement(),lat,lon,"dit vak");
}
function egRefresh(restyle){
  if(!M.egmsUrl)return;
  egCb.checked=!!S.showEgms;
  if(!S.showEgms){if(map.hasLayer(EG.layer))map.removeLayer(EG.layer);EG.layer.clearLayers();egNote("");return;}
  if(!map.hasLayer(EG.layer))EG.layer.addTo(map);
  if(!EG.ov){egLaad();return;}
  const z=map.getZoom(),bb=map.getBounds().pad(0.2);
  const tekenRaster=(arr,d,txt)=>{
    EG.layer.clearLayers();
    for(let i=0;i<arr.length;i+=4){const la=arr[i],lo=arr[i+1];
      if(la<bb.getSouth()-d[1]||la>bb.getNorth()+d[1]||lo<bb.getWest()-d[0]||lo>bb.getEast()+d[0])continue;
      EG.layer.addLayer(egVak(la,lo,d[1]/2,d[0]/2,arr[i+2],0,arr[i+3]));}
    egNote(txt);
  };
  if(z<11){tekenRaster(EG.ov.r5,EG.ov.d5,"Uitgezoomd: per blok van ongeveer 5 km de doorsnee. Zoom in voor meer detail.");return;}
  const sl=egSleutels(bb);
  if(z<EG.DETAIL||sl.length>16){tekenRaster(EG.ov.r1,EG.ov.d1,"Per blok van ongeveer 1 km de doorsnee. Zoom verder in voor blokken van 100 m.");return;}
  const zoomNu=z;
  Promise.all(sl.map(egTegel)).then(lijsten=>{
    if(map.getZoom()!==zoomNu||!S.showEgms)return;
    const b2=map.getBounds().pad(0.2);EG.layer.clearLayers();
    const hy=0.00045,hx=0.00045/Math.cos(map.getCenter().lat*Math.PI/180);
    for(const a of lijsten)for(let i=0;i<a.length;i+=4){
      if(!b2.contains([a[i],a[i+1]]))continue;
      EG.layer.addLayer(egVak(a[i],a[i+1],hy,hx,a[i+2],a[i+3],0));}
    egNote("Blokken van 100 m. Geen kleur betekent: daar is niets gemeten (zoals weilanden en water). Klik op een blok voor de snelheid.");
  });
}
map.on("moveend",()=>egRefresh(false));
// Leaflet geeft klikken alleen aan de bovenste tekenlaag. Een klik die daar niets raakt, komt als
// kaartklik binnen; zoek dan zelf het vakje eronder (eerst NAP-vakjes, dan satellietvakken).
map.on("click",e=>{
  for(const lg of [GRID.layer,EG.layer]){
    if(!map.hasLayer(lg))continue;
    const hit=lg.getLayers().find(r=>r.getBounds&&r.getBounds().contains(e.latlng));
    if(hit){hit.fire("click",{latlng:e.latlng});return;}
  }
});
// satellietmetingen binnen een straal (laadt de benodigde tegels)
function egRond(lat,lon,r){
  if(!M.egmsUrl)return Promise.resolve(null);
  const d=M.egmsTegel||0.1,dy=r/111320,dx=r/(111320*Math.cos(lat*Math.PI/180)),sl=[];
  for(let x=Math.floor((lon-dx)/d);x<=Math.floor((lon+dx)/d);x++)for(let y=Math.floor((lat-dy)/d);y<=Math.floor((lat+dy)/d);y++)
    if(!EG.set||EG.set.has(x+"_"+y))sl.push(x+"_"+y);
  {
    return Promise.all(sl.map(egTegel)).then(lijsten=>{
      const kx=111320*Math.cos(lat*Math.PI/180),o={n:0,v:[],k:{rood:0,geel:0,groen:0,stab:0,oud:0}};
      for(const a of lijsten)for(let i=0;i<a.length;i+=4){
        const ddx=(a[i+1]-lon)*kx,ddy=(a[i]-lat)*111320;if(ddx*ddx+ddy*ddy>r*r)continue;
        const c=egCls(a[i+2],a[i+3]);o.n++;o.k[c]++;if(c!=="oud")o.v.push(a[i+2]/10);}
      return o;
    });
  }
}
function egZin(m){return Math.abs(m)<0.05?"de grond beweegt doorgaans niet meetbaar":m>0?"de grond zakt doorgaans "+nf(m,1)+" mm per jaar":"de grond stijgt doorgaans "+nf(-m,1)+" mm per jaar";}
// In een pop-up: hoe beweegt de bodem eromheen volgens de satellieten?
function egPopup(el,lat,lon){
  if(!M.egmsUrl||!el)return;
  egRond(lat,lon,ROND_M).then(o=>{
    if(!o||!o.v.length||!el.isConnected)return;
    const t="Satelliet (binnen "+ROND_M+" m): "+egZin(median(o.v))+" ("+o.v.length+" blokken van 100 m).";
    const d=el.querySelector(".omg");
    if(d)d.insertAdjacentHTML("beforeend",'<br>'+esc(t));
    else{const a=el.querySelector("a.raw");if(a)a.insertAdjacentHTML("beforebegin",'<div class="omg">'+esc(t)+'</div>');}
  });
}

// ---------- pand op een plek: bouwjaar en aantal adressen (BAG via PDOK, publiek domein) ----------
const BAG_API="https://api.pdok.nl/kadaster/bag/ogc/v2/collections/pand/items";
const BAG={cache:new Map(),vorm:null};
function bagRingen(g){if(!g)return[];if(g.type==="Polygon")return[g.coordinates[0]];if(g.type==="MultiPolygon")return g.coordinates.map(p=>p[0]);return[];}
function bagAfstand(lat,lon,ringen){
  // 0 als het punt in het pand ligt, anders de afstand in meters tot de rand
  const kx=111320*Math.cos(lat*Math.PI/180),ky=111320;let best=Infinity;
  for(const r of ringen){let binnen=false;
    for(let i=0,j=r.length-1;i<r.length;j=i++){
      const xi=(r[i][0]-lon)*kx,yi=(r[i][1]-lat)*ky,xj=(r[j][0]-lon)*kx,yj=(r[j][1]-lat)*ky;
      if((yi>0)!==(yj>0)&&0<(xj-xi)*(0-yi)/((yj-yi)||1e-12)+xi)binnen=!binnen;
      const dx=xj-xi,dy=yj-yi,l2=dx*dx+dy*dy;let t=l2?-(xi*dx+yi*dy)/l2:0;t=Math.max(0,Math.min(1,t));
      const d=Math.hypot(xi+t*dx,yi+t*dy);if(d<best)best=d;}
    if(binnen)return 0;}
  return best;
}
function bagPand(lat,lon,maxM){
  const k=lat.toFixed(5)+","+lon.toFixed(5)+","+maxM;
  if(BAG.cache.has(k))return BAG.cache.get(k);
  const r=Math.max(maxM,2)+3,dy=r/111320,dx=r/(111320*Math.cos(lat*Math.PI/180));
  const u=BAG_API+"?f=json&limit=30&bbox="+[lon-dx,lat-dy,lon+dx,lat+dy].map(v=>v.toFixed(6)).join(",");
  const pr=fetch(u).then(x=>{if(!x.ok)throw new Error("HTTP "+x.status);return x.json();}).then(j=>{
    let best=null,bd=Infinity;
    for(const f of (j&&j.features)||[]){const pp=f.properties||{},st=String(pp.status||"");
      if(/gesloopt|niet gerealiseerd/i.test(st))continue;
      const d=bagAfstand(lat,lon,bagRingen(f.geometry));if(d<bd){bd=d;best=f;}}
    if(!best||bd>maxM)return null;
    const pp=best.properties||{};
    return{bouwjaar:pp.bouwjaar,adressen:pp.aantal_verblijfsobjecten,status:String(pp.status||""),geom:best.geometry};
  }).catch(()=>null);
  BAG.cache.set(k,pr);pr.then(v=>{if(v===null)setTimeout(()=>BAG.cache.delete(k),60000);});
  return pr;
}
function bagTekst(p){
  const d=[];const j=+p.bouwjaar;
  d.push(j>=1200&&j<=new Date().getFullYear()+5?"gebouwd in "+j:"bouwjaar onbekend");
  const n=+p.adressen;if(n>0)d.push(n+(n>1?" adressen":" adres"));
  if(p.status&&!/^pand in gebruik/i.test(p.status))d.push(p.status.toLowerCase());
  return '<p class="pand"><b>Pand:</b> '+esc(d.join(" · "))+' <span class="m">(BAG)</span></p>';
}
function bagVorm(geom){
  bagVormWeg();if(!geom)return;
  BAG.vorm=L.geoJSON(geom,{pane:"rond",renderer:ROND.rend,interactive:false,style:{color:"#24507a",weight:2,fill:true,fillOpacity:.08,dashArray:null}}).addTo(map);
}
function bagVormWeg(){if(BAG.vorm){map.removeLayer(BAG.vorm);BAG.vorm=null;}}
map.on("popupclose",()=>{if(!ROND.pandVorm)bagVormWeg();});
// In een pop-up: bouwjaar en aantal adressen van het pand waar het meetpunt in zit
function bagPopup(el,lat,lon,maxM){
  if(!el)return;
  bagPand(lat,lon,maxM).then(p=>{
    if(!p||!el.isConnected||el.querySelector(".pand"))return;
    const t=el.querySelector("table.kv");if(t)t.insertAdjacentHTML("afterend",bagTekst(p));
    ROND.pandVorm=false;bagVorm(p.geom);
  });
}

buildRows();fillSettings();styleAll();

// ---------- link naar een plek of meetpunt: ?kaart=lat,lon,zoom of ?punt=nummer ----------
(function(){
  let p;try{p=new URLSearchParams(location.search);}catch(e){return;}
  const k=p.get("kaart"),pt=(p.get("punt")||"").trim(),van=(p.get("van")||"").trim();
  if(GB&&/^gemeente\/[a-z0-9-]+\/(((wijk|straat)\/)?[a-z0-9-]+\/)?$/.test(van)){
    const d=van.split("/");const gi=GB.g.findIndex(g=>g[0]===d[1]);let x=null;
    if(gi>=0&&!d[2])x=gebGem(gi);
    else if(gi>=0&&gi===GB.a&&d[2]==="wijk"&&GB.w)x=gebWijk(GB.w.findIndex(w=>w[0]===d[3]));
    else if(gi>=0&&gi===GB.a&&d[2]==="straat"&&GB.s){const y=gebStraat(GB.s.findIndex(w=>w[0]===d[3]));x=y?{u:y.u,n:"de "+y.n}:null;}
    else if(gi>=0&&gi===GB.a&&GB.b){const bi=GB.b.findIndex(b=>b&&b[0]===d[2]);x=gebBuurt(bi);}
    if(x){const t=document.getElementById("terug");
      t.innerHTML='<a href="'+esc(x.u)+'">← Terug naar '+esc(x.n)+'</a><button type="button" aria-label="Verbergen">×</button>';
      t.hidden=false;document.body.classList.add("terug");
      t.querySelector("button").addEventListener("click",()=>{t.hidden=true;document.body.classList.remove("terug");});
      L.DomEvent.disableClickPropagation(t);}
  }
  if(k){const a=k.split(",").map(Number);
    if(a.length>=2&&a[0]>50&&a[0]<54&&a[1]>3&&a[1]<8)map.setView([a[0],a[1]],Math.max(7,Math.min(18,Math.round(a[2])||14)));}
  if(!pt)return;
  if(/^\d{3}[a-z]\d{4}$/i.test(pt)){if(M.napUrl)napFind(pt);return;}
  const i=B.findIndex(b=>String(b[I.id])===pt);
  if(i<0){say("Meetpunt “"+pt+"” niet gevonden.");return;}
  map.setView([B[i][I.lat],B[i][I.lon]],18);setTimeout(()=>openBout(i),0);
})();

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
