#!/usr/bin/env python3
# ============================================================================
# build_map.py  —  MyBuddy Vet Finder (V10.1)
#
# CHANGES in V10.1:
#   - Species parser now matches MULTIPLE species per value.
#     ("Dogs & Cats" -> ["cat","dog"], not just "dog").
#   - Added SPECIES_OVERRIDES dict for manual per-clinic fixes.
#   - Added [SPECIES] PROBE diagnostic for clinics containing "central".
#
# Still present from V10:
#   #1  🗺️ Get Directions button in every popup.
#   #2  Species filters (Dogs / Cats / Exotics) on desktop + mobile.
#   - Price-list fixes (strip IDs, keep $0, unmatched warnings)
#   - PRICE_EMAIL = info@mybuddy.sg for the ✍️ Review / Update Price button
# ============================================================================
import json, re, time, urllib.request, urllib.parse
from pathlib import Path
from openpyxl import load_workbook

WORKBOOK = "MyBuddy_Vet_Finder_MASTER_RECONCILED_2026-10-15_v9.xlsx"
OUTPUT   = "mybuddy-vet-map-v3.html"
CACHE_FILE = Path(__file__).parent / "geocache.json"
MANIFEST_OUTPUT = "manifest.json"
ICON_OUTPUT     = "icon.svg"

# ADD YOUR GOOGLE API KEY HERE
GOOGLE_API_KEY = "PASTE_YOUR_NEW_API_KEY_HERE"

# Auto-delete older maps to prevent caching confusion
for old in ["mybuddy-vet-map.html", "mybuddy-vet-map-v2.html"]:
    old_file = Path(__file__).parent / old
    if old_file.exists():
        old_file.unlink()

DAY_NAMES = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
DAY_ALIASES = {"mon":"Monday","monday":"Monday","tue":"Tuesday","tues":"Tuesday","tuesday":"Tuesday","wed":"Wednesday","weds":"Wednesday","wednesday":"Wednesday","thu":"Thursday","thur":"Thursday","thurs":"Thursday","thursday":"Thursday","fri":"Friday","friday":"Friday","sat":"Saturday","saturday":"Saturday","sun":"Sunday","sunday":"Sunday"}

POSTAL_FALLBACKS = {
    "563338": "560338" # Acacia Vet
}

EMERGENCY_OVERRIDES = {
    "MBV-0005",
    "MBV-0032"
}

# ─────────────────────────────────────────────────────────────────────────────
# Manual species overrides — use only when VET_SERVICES row is missing/wrong.
# Keys: MyBuddy Vet IDs. Values: list of "dog" / "cat" / "bird" / "exotic".
# Example:
#   SPECIES_OVERRIDES = {
#       "MBV-00XX": ["dog", "cat"],   # VET CENTRAL
#   }
# ─────────────────────────────────────────────────────────────────────────────
SPECIES_OVERRIDES = {
    # "MBV-00XX": ["dog", "cat"],
}

MANUAL_COORDS = {
    "MBV-0005": (1.3244, 103.9418),
    "MBV-0032": (1.3286, 103.8521)
}

def to_bool(v):
    if isinstance(v, bool): return v
    if isinstance(v, str):  return v.strip().lower() in ("true","yes","1","y")
    return False

def parse_time(s):
    s = str(s).strip().lower().replace(".", "").replace(" ", "")
    if not s: return None
    m = re.match(r"^(\d{1,2})(?::(\d{2}))?\s*(am|pm)?$", s)
    if not m: return None
    h = int(m.group(1)); mi = int(m.group(2) or 0); ap = m.group(3)
    if mi > 59: return None
    if ap == "am" and h == 12: h = 0
    elif ap == "pm" and h != 12: h += 12
    elif not ap and h > 23: return None
    return h * 60 + mi

def parse_intervals(text):
    if text is None: return None
    t = str(text).strip().lower()
    if not t: return None
    if re.search(r"\bby\s+appointment\b", t): return "appointment"
    if re.match(r"^closed\b", t): return []
    if re.search(r"\b(open\s+)?24\s*hours?\b", t) or "24/7" in t: return [(0, 1440)]
    ranges = re.findall(r"(\d{1,2}(?::\d{2})?\s*(?:am|pm)?)\s*(?:[\u2013\u2014\-]|to)\s*(\d{1,2}(?::\d{2})?\s*(?:am|pm)?)", t, re.I)
    if not ranges:
        return [] if re.search(r"\bclosed\b", t) else None
    intervals = []
    for a, b in ranges:
        sa = parse_time(a); sb = parse_time(b)
        if sa is None or sb is None: continue
        if sb <= sa: sb += 1440
        intervals.append((sa, sb))
    return intervals if intervals else None

def day_in_range(start, end, target):
    s = DAY_ALIASES.get(start.lower()); e = DAY_ALIASES.get(end.lower())
    if not s or not e: return False
    si = DAY_NAMES.index(s); ei = DAY_NAMES.index(e); cur = si
    while True:
        if DAY_NAMES[cur] == target: return True
        if cur == ei: break
        cur = (cur + 1) % 7
    return False

def extract_day_schedule(full_text, target_day):
    if not full_text: return ""
    chunks = re.split(r"[;\n]+", str(full_text))
    result = []; current_days = None
    for chunk in chunks:
        chunk = chunk.strip()
        if not chunk: continue
        dr = re.match(r"^(mon|monday|tue|tues|tuesday|wed|weds|wednesday|thu|thur|thurs|thursday|fri|friday|sat|saturday|sun|sunday)\s*[\u2013\u2014\-]\s*(mon|monday|tue|tues|tuesday|wed|weds|wednesday|thu|thur|thurs|thursday|fri|friday|sat|saturday|sun|sunday)\b", chunk, re.I)
        if dr:
            current_days = [d for d in DAY_NAMES if day_in_range(dr.group(1), dr.group(2), d)]
            if target_day in current_days:
                rest = chunk[dr.end():].strip()
                if rest: result.append(rest)
            continue
        ds = re.match(r"^(mon|monday|tue|tues|tuesday|wed|weds|wednesday|thu|thur|thurs|thursday|fri|friday|sat|saturday|sun|sunday)\b", chunk, re.I)
        if ds:
            d = DAY_ALIASES.get(ds.group(1).lower())
            current_days = [d]
            if d == target_day:
                rest = chunk[ds.end():].strip()
                if rest: result.append(rest)
            continue
        if current_days is None or target_day in current_days:
            result.append(chunk)
    return "; ".join(result)

def read_sheet(wb, name):
    ws = wb[name]
    rows = list(ws.iter_rows(values_only=True))
    if not rows: return []
    headers = [str(h).strip() if h is not None else f"col{i}" for i, h in enumerate(rows[0])]
    return [{headers[i]: (row[i] if i < len(row) else None) for i in range(len(headers))} for row in rows[1:] if not all(c is None for c in row)]

# ----------------------------------------------------------------------------
# SPECIES PARSER (V10.1)
# A value can match MULTIPLE species: "Dogs & Cats" -> {"dog","cat"}.
# Handles two VET_SERVICES layouts:
#   A) Long format:  a 'Species' column ("Dog", "Cat", "Dog & Cat", "Bird"...)
#   B) Wide format:  one row per clinic with boolean columns (Dogs / Cats / ...)
# ----------------------------------------------------------------------------
SPECIES_ALIASES = {
    "dog":    ["dog", "canine"],
    "cat":    ["cat", "feline"],
    "bird":   ["bird", "avian", "parrot"],
    "exotic": ["exotic", "rabbit", "small", "reptile", "pocket", "hamster", "guinea"],
}

def _species_from_text(s):
    """Return a SET of species matched in a string (can be multiple)."""
    s = str(s or "").strip().lower()
    if not s:
        return set()
    return {tag for tag, kws in SPECIES_ALIASES.items()
            if any(kw in s for kw in kws)}

def extract_species_tags(services_data):
    tags = {}
    if not services_data:
        return tags

    headers = list(services_data[0].keys()) if services_data else []
    has_species_col = any(str(h).strip().lower() == "species" for h in headers)

    wide_cols = []
    for h in headers:
        sp_set = _species_from_text(h)
        if sp_set:
            wide_cols.append((h, sp_set))

    for r in services_data:
        mbv = (r.get("MyBuddy Vet ID") or "").strip()
        if not mbv:
            continue

        if has_species_col:
            sp_set = _species_from_text(r.get("Species"))
            if sp_set:
                tags.setdefault(mbv, set()).update(sp_set)
        else:
            for col, sp_set in wide_cols:
                if to_bool(r.get(col)):
                    tags.setdefault(mbv, set()).update(sp_set)

    return {k: sorted(v) for k, v in tags.items()}

def geocode_postal(postal, cache):
    if postal in cache:
        return cache[postal]
    query = POSTAL_FALLBACKS.get(postal, postal)
    url = "https://www.onemap.gov.sg/api/common/elastic/search?" + urllib.parse.urlencode({
        "searchVal": query, "returnGeom": "Y", "getAddrDetails": "Y", "pageNum": 1})
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=8) as r:
            data = json.loads(r.read().decode("utf-8"))
        for p in (data.get("results") or []):
            try:
                lat = float(p["LATITUDE"]); lng = float(p["LONGITUDE"])
                if 1.1 <= lat <= 1.5 and 103.5 <= lng <= 104.2:
                    cache[postal] = (lat, lng)
                    return (lat, lng)
            except (KeyError, ValueError):
                continue
    except Exception:
        pass
    return None

def fetch_google_rating(name, address):
    query = urllib.parse.quote(f"{name} {address}")
    url = f"https://maps.googleapis.com/maps/api/place/textsearch/json?query={query}&key={GOOGLE_API_KEY}"
    try:
        req = urllib.request.Request(url)
        with urllib.request.urlopen(req, timeout=5) as r:
            data = json.loads(r.read().decode("utf-8"))
            if data.get("results"):
                first = data["results"][0]
                return first.get("rating", 0), first.get("user_ratings_total", 0)
    except Exception as e:
        print(f"Failed to fetch rating for {name}: {e}")
    return None, None

def write_pwa_files():
    manifest = {
        "name": "MyBuddy Vet Finder",
        "short_name": "Vet Finder",
        "start_url": "./mybuddy-vet-map-v3.html",
        "scope": "./",
        "display": "standalone",
        "orientation": "portrait",
        "background_color": "#ffffff",
        "theme_color": "#4CAF50",
        "icons": [
            {"src": "icon.svg", "sizes": "any", "type": "image/svg+xml", "purpose": "any"}
        ]
    }
    Path(MANIFEST_OUTPUT).write_text(json.dumps(manifest, indent=2), encoding="utf-8")

    icon_svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512">'
        '<rect width="512" height="512" rx="96" fill="#4CAF50"/>'
        '<circle cx="176" cy="180" r="44" fill="#ffffff"/>'
        '<circle cx="256" cy="140" r="44" fill="#ffffff"/>'
        '<circle cx="336" cy="180" r="44" fill="#ffffff"/>'
        '<circle cx="128" cy="272" r="40" fill="#ffffff"/>'
        '<circle cx="384" cy="272" r="40" fill="#ffffff"/>'
        '<path d="M256 240c-58 0-104 46-104 104 0 46 38 68 104 68s104-22 104-68c0-58-46-104-104-104z" fill="#ffffff"/>'
        '</svg>'
    )
    Path(ICON_OUTPUT).write_text(icon_svg, encoding="utf-8")

def build():
    wb = load_workbook(WORKBOOK, data_only=True)
    master   = read_sheet(wb, "MASTER_LIST")
    hours    = read_sheet(wb, "VET_CONTACT_HOURS")
    branches = read_sheet(wb, "BRANCH_RELATIONSHIPS")
    prices_data = read_sheet(wb, "VET_PRICES")

    try:
        services_data = read_sheet(wb, "VET_SERVICES")
    except KeyError:
        services_data = []
        print("[SPECIES] WARNING: no VET_SERVICES sheet — species filter will show everything.")

    hours_by_id = {r["MyBuddy Vet ID"]: r for r in hours if r.get("MyBuddy Vet ID")}

    # Process Prices
    prices_by_id = {}
    for p in prices_data:
        mbv = (p.get("MyBuddy Vet ID") or "").strip()
        if not mbv: continue
        service = p.get("Service")
        price = p.get("Price")
        if service and price not in (None, ""):
            prices_by_id.setdefault(mbv, []).append(f"{service}: ${price}")

    # Process Species
    species_by_id = extract_species_tags(services_data)

    # Apply manual overrides
    for _id, _sp in SPECIES_OVERRIDES.items():
        merged = set(species_by_id.get(_id, [])) | set(_sp)
        species_by_id[_id] = sorted(merged)
    if SPECIES_OVERRIDES:
        print(f"[SPECIES] Applied {len(SPECIES_OVERRIDES)} manual override(s).")

    groups = {}
    for r in branches:
        gid = r.get("Group ID")
        if not gid: continue
        groups.setdefault(gid, []).append({"id": r.get("MBV ID"), "name": r.get("Branch Name") or ""})
    siblings = {}
    for gid, members in groups.items():
        for m in members:
            siblings[m["id"]] = [x for x in members if x["id"] != m["id"]]

    clinics = []
    for r in master:
        if (r.get("Active / Closed / Relocated Status") or "").strip() != "Current":
            continue
        mbv = (r.get("MyBuddy Vet ID") or "").strip()
        h = hours_by_id.get(mbv, {})
        postal = r.get("Postal Code")
        postal = str(int(postal)) if isinstance(postal, (int, float)) else str(postal or "").strip()
        if postal.endswith(".0"): postal = postal[:-2]
        postal = postal.zfill(6)
        phone = (r.get("Phone / Main Contact") or h.get("Main Phone") or "").strip()
        wa    = (r.get("WhatsApp") or h.get("WhatsApp") or "").strip()
        booking = h.get("Booking URL") or ""
        if booking and not str(booking).lower().startswith("http"): booking = ""
        addr_parts = [r.get("Street Address"), r.get("Building Name")]
        addr = ", ".join(str(p).strip() for p in addr_parts if p)

        sched = {}
        for day in DAY_NAMES:
            raw = h.get(day)
            txt = extract_day_schedule(raw, day) if raw else ""
            sched[day] = parse_intervals(txt) if txt else None

        ecat = (h.get("Emergency Category") or "NOT_24_7_UNCONFIRMED").strip()
        if mbv in EMERGENCY_OVERRIDES:
            ecat = "E1_CONFIRMED_24_7"

        rating_val = r.get("Google Rating") or 0
        review_val = r.get("Review Count") or 0

        appt_req = str(h.get("Appointment Required") or "").strip()
        is_appt = "appointment" in appt_req.lower() and "no" not in appt_req.lower()

        clinics.append({
            "id": mbv,
            "name": r.get("Clinic / Provider Name") or "",
            "branch": r.get("Branch Name") or "",
            "address": addr,
            "postal": postal,
            "phone": phone,
            "whatsapp": wa,
            "booking": booking,
            "ecat": ecat,
            "schedule": sched,
            "house_call": to_bool(r.get("Offers_House_Calls")),
            "hospice":    to_bool(r.get("Hospice_Service")),
            "lat": None, "lng": None,
            "rating": rating_val,
            "review_count": review_val,
            "is_appt": is_appt,
            "prices":  prices_by_id.get(mbv, []),
            "species": species_by_id.get(mbv, [])
        })

    # --- Build-time diagnostics ---
    master_ids = {c["id"] for c in clinics}
    with_prices = sum(1 for c in clinics if c["prices"])
    print(f"[PRICE CHECK] Clinics with prices: {with_prices} / {len(clinics)}")
    unmatched = sorted(set(prices_by_id) - master_ids)
    if unmatched:
        print(f"[PRICE CHECK] WARNING: {len(unmatched)} price ID(s) in VET_PRICES "
              f"did not match any Current clinic in MASTER_LIST:")
        for u in unmatched:
            print(f"    - {u!r}")
    print(f"[PRICE CHECK] Clinics with NO price rows: {len(master_ids - set(prices_by_id))}")

    from collections import Counter
    with_species = sum(1 for c in clinics if c["species"])
    print(f"[SPECIES] Clinics with species tags: {with_species} / {len(clinics)}")
    if with_species:
        ctr = Counter(t for c in clinics for t in c["species"])
        print(f"[SPECIES] Tag counts: {dict(ctr)}")

    # --- PROBE: print species for any clinic whose name contains "central" ---
    for c in clinics:
        if "central" in c["name"].lower():
            print(f"[SPECIES] PROBE: {c['id']} {c['name']!r} -> {c['species']}")
    # -------------------------------------------------------------------------

    cache = {}
    if CACHE_FILE.exists():
        try:
            data = json.loads(CACHE_FILE.read_text())
            cache = {k: tuple(v) for k, v in data.items()}
        except Exception:
            pass

    for i, c in enumerate(clinics, 1):
        mbv = c["id"]
        postal = c["postal"]

        print(f"Fetching Google rating for {c['name']}...")
        live_rating, live_reviews = fetch_google_rating(c['name'], c['address'])

        if live_rating and live_reviews:
            c["rating"] = live_rating
            c["review_count"] = live_reviews

        time.sleep(0.1)

        if mbv in MANUAL_COORDS:
            coords = MANUAL_COORDS[mbv]
        else:
            coords = geocode_postal(postal, cache)

        if coords:
            c["lat"], c["lng"] = coords

        time.sleep(0.05)

    try:
        CACHE_FILE.write_text(json.dumps({k: v for k, v in cache.items()}, indent=2))
    except Exception:
        pass

    directory = {}
    for c in clinics:
        letter = c["name"][:1].upper()
        if not letter.isalpha(): letter = "#"
        directory.setdefault(letter, []).append({"id": c["id"], "name": c["name"]})
    for L in directory:
        directory[L].sort(key=lambda x: x["name"].lower())

    data_json = json.dumps({"clinics": clinics, "siblings": siblings, "directory": directory},
                           ensure_ascii=False, separators=(",", ":"))
    html = HTML_TEMPLATE.replace("/*DATA*/", data_json)
    Path(OUTPUT).write_text(html, encoding="utf-8")

    write_pwa_files()
    print(f"SUCCESS: Wrote {OUTPUT}, {MANIFEST_OUTPUT}, {ICON_OUTPUT}")

HTML_TEMPLATE = r'''<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover,maximum-scale=1">
<title>MyBuddy Vet Finder</title>
<link rel="manifest" href="manifest.json">
<meta name="theme-color" content="#4CAF50">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="default">
<meta name="apple-mobile-web-app-title" content="Vet Finder">
<link rel="icon" type="image/svg+xml" href="icon.svg">
<link rel="apple-touch-icon" href="icon.svg">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Baloo+2:wght@500;600;700;800&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css">
<style>
html,body,#map{height:100%;margin:0;font-family:'Baloo 2', cursive}
#map{background:#e8edf0}
#panel{position:absolute;z-index:1000;top:12px;left:12px;width:min(360px,calc(100vw - 40px));max-height:calc(100vh - 40px);overflow:visible;background:#fff;padding:16px;border-radius:20px;box-shadow:0 8px 25px rgba(0,0,0,.1);font-size:14px;display:flex;flex-direction:column;}
.panel-content { overflow-y: auto; flex: 1; padding-right: 4px; }
h1{font-size:22px;margin:0 0 4px;color:#388E3C;font-weight:800}
p.sub{font-size:12px;color:#90A4AE;margin:0 0 10px;font-weight:600}
.row{display:flex;align-items:center;gap:8px;margin:10px 0}
.row label{font-weight:700;color:#2C3E50;white-space:nowrap;font-size:13px}
#radius{flex:1;min-width:80px;accent-color:#4CAF50}
#radiusVal{min-width:52px;text-align:right;color:#4CAF50;font-weight:800}
.btn{background:#4CAF50;color:#fff;border:none;padding:10px 14px;border-radius:14px;font-size:13px;font-family:inherit;font-weight:700;cursor:pointer;flex:1;transition:transform 0.2s, background 0.2s;}
.btn:hover{background:#388E3C;transform:scale(1.02)}
.btn.ghost{background:#F1F8E9;color:#388E3C;border:1px solid #C8E6C9}
.btn.ghost:hover{background:#E8F5E9}
.input{width:100%;box-sizing:border-box;padding:10px 14px;border:2px solid #E0E0E0;border-radius:14px;font-size:13px;font-family:inherit;font-weight:600}
.input:focus{outline:none;border-color:#4CAF50}
.search-wrap{position:relative;margin:10px 0}
.search-wrap .input{padding-right:30px}
.search-wrap::after{content:"\1F50D";position:absolute;right:12px;top:50%;transform:translateY(-50%);font-size:14px;color:#90A4AE;pointer-events:none}
.directory{display:none;position:absolute;top:100%;left:0;right:0;background:#fff;border:2px solid #E0E0E0;border-radius:14px;margin-top:6px;max-height:300px;overflow-y:auto;box-shadow:0 12px 30px rgba(0,0,0,.15);z-index:1001}
.directory.open{display:block}
.dir-letter{border-bottom:1px solid #F0F0F0}
.dir-header{padding:10px 16px;font-weight:800;color:#388E3C;cursor:pointer;background:#F9FBF9;display:flex;justify-content:space-between}
.dir-header::after{content:"+";font-weight:600;color:#90A4AE;font-size:16px}
.dir-letter.open .dir-header::after{content:"−"}
.dir-items{display:none;padding:4px 0}
.dir-letter.open .dir-items{display:block}
.dir-item{padding:10px 16px 10px 24px;cursor:pointer;font-size:13px;font-weight:600;color:#2C3E50}
.dir-item:hover{background:#E8F5E9;color:#388E3C}
.filters{display:flex;flex-wrap:wrap;gap:6px;margin:10px 0}
.chip{font-size:12px;padding:8px 14px;border-radius:99px;border:2px solid #E0E0E0;background:#fff;cursor:pointer;font-weight:700;color:#90A4AE;transition:all 0.2s}
.chip.on{background:#4CAF50;color:#fff;border-color:#4CAF50}
.legend{font-size:12px;line-height:1.8;color:#90A4AE;border-top:2px dashed #E0E0E0;padding-top:12px;margin-top:12px;font-weight:600}
.dot{display:inline-block;width:12px;height:12px;border-radius:50%;margin-right:6px;vertical-align:middle;border:2px solid #fff;box-shadow:0 1px 3px rgba(0,0,0,.2)}
#status{font-size:12px;color:#388E3C;margin-top:10px;font-weight:700;background:#E8F5E9;padding:8px 12px;border-radius:10px;display:inline-block}

.disclaimer { font-size: 11px; color: #90A4AE; margin-top: 16px; padding-top: 12px; border-top: 2px dashed #E0E0E0; font-weight: 600; line-height: 1.4; }
.disclaimer b { color: #388E3C; }

.pin{width:18px;height:18px;border-radius:50%;border:3px solid #fff;box-shadow:0 2px 8px rgba(0,0,0,.3);display:flex;align-items:center;justify-content:center}
.pin.open{background:#4CAF50}
.pin.closed{background:#90A4AE}
.pin.emergency{background:#F44336;animation:pulse 1.5s infinite}
@keyframes pulse{0%{box-shadow:0 0 0 0 rgba(244,67,54,0.6)}70%{box-shadow:0 0 0 12px rgba(244,67,54,0)}100%{box-shadow:0 0 0 0 rgba(244,67,54,0)}}
.pin-hit{width:36px;height:36px;display:flex;align-items:center;justify-content:center;cursor:pointer}
.pin-wrap{background:transparent;border:none}
.pin.blink { animation: blink-anim 1s infinite !important; }
@keyframes blink-anim {
    0% { opacity: 1; transform: scale(1); box-shadow: 0 0 0 0 rgba(76, 175, 80, 0.7); }
    50% { opacity: 0.4; transform: scale(1.15); box-shadow: 0 0 0 10px rgba(76, 175, 80, 0); }
    100% { opacity: 1; transform: scale(1); box-shadow: 0 0 0 0 rgba(76, 175, 80, 0); }
}
.centre-handle-icon{background:transparent!important;border:none!important}
.centre-handle{width:46px;height:46px;border:3px solid rgba(33,150,243,.5);border-radius:50%;background:rgba(33,150,243,.1);display:flex;align-items:center;justify-content:center;color:#2196F3;font-size:30px;font-weight:bold;cursor:grab}
.centre-handle:active{cursor:grabbing;transform:scale(1.1)}

.popup{font-size:13px;line-height:1.5;min-width:260px;max-width:300px;font-family:'Baloo 2', cursive}
.popup h3{font-size:16px;margin:0 0 4px;color:#2C3E50;font-weight:800}
.popup p{margin:4px 0;font-weight:600}
.badge{display:inline-block;font-size:11px;padding:3px 8px;border-radius:99px;margin:0 4px 4px 0;color:#fff;font-weight:700}

.leaflet-popup-content .btn-link {
    display: inline-block;
    margin: 6px 6px 0 0;
    padding: 8px 12px;
    border-radius: 10px;
    text-decoration: none;
    color: #FFEB3B !important;
    text-shadow: 0 1px 2px rgba(0,0,0,0.4);
    background: #2196F3;
    font-size: 13px;
    font-weight: 800;
    text-align: center;
}
.leaflet-popup-content .btn-link.wa { background: #4CAF50; }

.price-list { max-height: 120px; overflow-y: auto; background: #F9FBFD; padding: 10px; border-radius: 10px; font-size: 12px; border: 1px solid #E1E8ED; margin-top: 8px; }
.price-item { padding: 4px 0; border-bottom: 1px dashed #E0E0E0; font-weight: 600; color: #555; }
.price-item:last-child { border-bottom: none; }

#mobile-top-bar, #mobile-species-bar, #mobile-radius-slider { display: none; }

@media (max-width: 768px) {
  #panel { display: none !important; }

  #mobile-top-bar {
    display: flex; position: fixed;
    top: calc(env(safe-area-inset-top, 0px) + 12px);
    left: 12px; right: 12px;
    z-index: 2500; gap: 8px;
  }

  .m-search-wrap { flex: 1; position: relative; }
  #mobile-search-input {
    width: 100%; height: 46px; padding: 0 16px;
    border: none; border-radius: 16px;
    box-shadow: 0 4px 15px rgba(0,0,0,.15);
    font-family: inherit; font-size: 15px; font-weight: 600;
    background: #fff; color: #2C3E50; box-sizing: border-box;
  }

  #mobile-house-btn {
    width: 46px; height: 46px;
    background: #fff; border: none; border-radius: 16px;
    box-shadow: 0 4px 15px rgba(0,0,0,.15);
    font-size: 20px; display: flex; align-items: center; justify-content: center;
    cursor: pointer; transition: background 0.2s;
  }
  #mobile-house-btn.on { background: #4CAF50; }

  #mobile-species-bar {
    display: flex; position: fixed;
    top: calc(env(safe-area-inset-top, 0px) + 68px);
    left: 12px; right: 12px; z-index: 2500; gap: 8px;
  }
  .m-sp-btn {
    flex: 1; height: 40px; border: none; border-radius: 12px;
    background: #fff; box-shadow: 0 4px 15px rgba(0,0,0,.12);
    font-size: 18px; cursor: pointer; transition: background 0.2s;
  }
  .m-sp-btn.on { background: #4CAF50; }

  #mobile-dir-dropdown {
    display: none; position: absolute; top: 52px; left: 0; right: 0;
    background: #fff; border-radius: 16px; box-shadow: 0 8px 25px rgba(0,0,0,.2);
    max-height: 65vh; overflow-y: auto; z-index: 2600;
  }
  #mobile-dir-dropdown.open { display: block; }
  .md-letter {
    padding: 12px 16px; font-weight: 800; font-size: 15px; color: #388E3C;
    background: #F9FBF9; border-bottom: 1px solid #EEE; cursor: pointer;
    display: flex; justify-content: space-between; align-items: center;
  }
  .md-letter::after { content: "+"; font-size: 18px; color: #90A4AE; }
  .md-letter.open::after { content: "−"; }
  .md-items { display: none; }
  .md-letter.open + .md-items { display: block; }
  .md-item {
    padding: 14px 16px 14px 24px; border-bottom: 1px solid #F1F5F9;
    font-weight: 600; font-size: 14px; color: #2C3E50; cursor: pointer;
  }
  .md-item:active { background: #E8F5E9; }

  #mobile-radius-slider {
    display: block; position: fixed;
    top: 50%; right: -85px; transform: translateY(-50%) rotate(-90deg);
    width: 220px; height: 44px; z-index: 2500;
    background: rgba(255,255,255,0.9); border-radius: 22px;
    box-shadow: 0 4px 15px rgba(0,0,0,.15);
    padding: 0 10px;
  }
  #mobile-radius-slider input { width: 100%; height: 100%; accent-color: #4CAF50; }

  .leaflet-control-zoom { display: none !important; }
  .leaflet-popup-content-wrapper { border-radius: 20px !important; }
  .leaflet-popup-content { margin: 14px 16px !important; }
  .popup { max-width: calc(100vw - 60px) !important; }
}

.leaflet-tile-pane {
    filter: saturate(1.5) contrast(0.95) hue-rotate(-5deg) brightness(1.05);
}
</style>
</head>
<body>
<div id="map"></div>

<div id="mobile-top-bar">
  <div class="m-search-wrap">
    <input id="mobile-search-input" placeholder="Search clinics..." autocomplete="off">
    <div id="mobile-dir-dropdown"></div>
  </div>
  <button id="mobile-house-btn" title="House Calls Only">🏠</button>
</div>

<div id="mobile-species-bar">
  <button class="m-sp-btn" data-sp="dog" title="Dogs">🐶 Dogs</button>
  <button class="m-sp-btn" data-sp="cat" title="Cats">🐱 Cats</button>
  <button class="m-sp-btn" data-sp="exotic" title="Exotics">🐰 Exotics</button>
</div>

<div id="mobile-radius-slider">
  <input id="mobile-radius-input" type="range" min="1" max="30" value="2" step="1">
</div>

<div id="fallback"><div><h2>Map failed to load</h2><p>Check your internet and open this file in Chrome or Edge.</p><button onclick="location.reload()">Retry</button></div></div>

<section id="panel">
  <div class="panel-content">
      <h1>🐾 MyBuddy Vet Finder</h1>
      <p class="sub">Singapore · 123 approved clinics</p>

      <div class="row">
        <label for="radius">Radius</label>
        <input id="radius" type="range" min="1" max="10" value="5" step="0.5">
        <span id="radiusVal">5.0 km</span>
      </div>

      <div class="row"><input id="centreInput" class="input" placeholder="Enter postal code..."></div>
      <div class="row" style="gap:8px">
        <button id="setCentre" class="btn">Set centre</button>
        <button id="resetCentre" class="btn ghost">Reset</button>
      </div>
      <div id="centreFeedback" class="feedback" style="color:#F44336; font-size:11px; font-weight:700; margin-top:4px;"></div>

      <div class="search-wrap">
        <input id="search" class="input" placeholder="Search clinic name...">
        <div id="directory" class="directory"></div>
      </div>

      <div class="filters" id="filters">
        <button class="chip sp-btn" data-sp="dog">🐶 Dogs</button>
        <button class="chip sp-btn" data-sp="cat">🐱 Cats</button>
        <button class="chip sp-btn" data-sp="exotic">🐰 Exotics</button>
      </div>

      <div class="legend">
        <span class="dot" style="background:#4CAF50"></span>Open now<br>
        <span class="dot" style="background:#90A4AE"></span>Closed / Unknown<br>
        <span class="dot" style="background:#F44336"></span>24/7 Emergency
      </div>

      <div id="status">Ready to sniff out a vet! 🐶</div>
  </div>
  <div class="disclaimer">
      <b>100% Unbiased:</b> MyBuddy takes <b>zero commission</b> from clinics, transport, or booking partners. All clinics are listed equally. No one can pay to rank higher.
  </div>
</section>

<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<script>
L.TileLayer.prototype.options.referrerPolicy = 'strict-origin-when-cross-origin';
const DATA = /*DATA*/;
const SINGAPORE = [1.3521, 103.8198];
const SGT_OFFSET_MIN = 8 * 60;

const PRICE_EMAIL = "info@mybuddy.sg";

function singaporeNow() {
  const now = new Date();
  const sgt = new Date(now.getTime() + SGT_OFFSET_MIN * 60000);
  return { day: ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"][sgt.getUTCDay()],
           mins: sgt.getUTCHours() * 60 + sgt.getUTCMinutes() };
}
function statusFor(c) {
  if (c.ecat === "E1_CONFIRMED_24_7") {
    return { key:"open", label:"OPEN NOW · 24/7 EMERGENCY" };
  }
  const { day, mins } = singaporeNow();
  const iv = c.schedule[day];
  if (iv === "appointment") return { key:"appt", label:"BY APPOINTMENT ONLY" };
  if (iv === null || iv === undefined) return { key:"unknown", label:"HOURS UNCERTAIN — CALL TO CONFIRM" };
  if (!iv.length) return { key:"closed", label:"CLOSED NOW" };
  return iv.some(([a,b]) => a <= mins && mins < b)
    ? { key:"open",  label:"OPEN NOW" }
    : { key:"closed", label:"CLOSED NOW" };
}

const map = L.map("map", { center: SINGAPORE, zoom: 11, zoomControl: false });
L.control.zoom({ position: 'bottomleft' }).addTo(map);

L.tileLayer("https://www.onemap.gov.sg/maps/tiles/Default/{z}/{x}/{y}.png", {
  minZoom: 11, maxZoom: 19,
  bounds: [[1.1304, 103.6020], [1.4784, 104.0945]],
  attribution: '&copy; <a href="https://www.onemap.gov.sg/">OneMap</a> · SLA'
}).addTo(map);

const centreCircle = L.circle(SINGAPORE, {
  radius: 5000, color:'#4CAF50', fillColor:'#4CAF50', fillOpacity:0.1, weight:3, dashArray: '5, 5', interactive:false
}).addTo(map);

const centreHandle = L.marker(SINGAPORE, {
  icon: L.divIcon({className:'centre-handle-icon',
    html:'<div class="centre-handle" title="Drag to move search centre">+</div>',
    iconSize:[46,46], iconAnchor:[23,23]}),
  draggable: true, zIndexOffset: 1000
}).addTo(map);

function moveCentre(ll) {
  centreCircle.setLatLng(ll); centreHandle.setLatLng(ll); updateVisibility();
}
centreHandle.on('drag', e => moveCentre(e.target.getLatLng()));
centreHandle.on('dragend', e => moveCentre(e.target.getLatLng()));

function pinIcon(c, st, hl, blink) {
  const colour = (c.ecat === "E1_CONFIRMED_24_7") ? "emergency"
              : (st.key === "open") ? "open"
              : "closed";
  const cls = "pin " + colour + (hl ? " highlight" : "") + (blink ? " blink" : "");
  return L.divIcon({
    className: "",
    html: '<div class="pin-hit"><div class="'+cls+'"></div></div>',
    iconSize: [36, 36], iconAnchor: [18, 18], popupAnchor: [0, -12]
  });
}
function digits(s){ return String(s||"").replace(/\D/g,""); }
function primaryPhone(p) {
    if (!p) return "";
    const parts = p.split('/');
    return digits(parts[0]);
}
function waUrl(v) {
  if (!v) return "";
  if (/^https?:\/\//i.test(v)) return v;
  const d = digits(v); if (!d) return "";
  return "https://wa.me/" + (d.startsWith("65") ? d : "65" + d);
}
function esc(s){ return String(s??"").replace(/[&<>"']/g, c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c])); }

window.reportPrice = function(id, name) {
    const subject = encodeURIComponent(`Price Update for ${name} (${id})`);
    const body = encodeURIComponent(`Hi MyBuddy,\n\nI would like to update/review the pricing for ${name}.\n\nCorrect Price: \n\nThank you!`);
    window.location.href = `mailto:${PRICE_EMAIL}?subject=${subject}&body=${body}`;
}

function popupHTML(c, st) {
  const color = st.key==="open" ? "#388E3C" : "#90A4AE";
  const sib = DATA.siblings[c.id] || [];
  const sibHTML = sib.length
    ? '<div style="border-top:2px dashed #EEE;margin-top:10px;padding-top:10px;font-size:12px;"><b style="color:#388E3C">Other branches</b><br>' + sib.map(s=>esc(s.name)).join("<br>") + '</div>'
    : "";
  const wa = waUrl(c.whatsapp);
  const badges = [];
  if (c.ecat==="E1_CONFIRMED_24_7") badges.push('<span class="badge" style="background:#F44336">24/7 Emergency</span>');
  if (c.ecat==="EXTENDED_HOURS_URGENT") badges.push('<span class="badge" style="background:#FF9800">Extended hours</span>');
  if (c.ecat==="REFERRAL_ONLY") badges.push('<span class="badge" style="background:#9C27B0">Referral only</span>');
  if (c.house_call) badges.push('<span class="badge" style="background:#00897B">House calls</span>');
  if (c.hospice) badges.push('<span class="badge" style="background:#5d4037">Hospice</span>');
  if (c.is_appt) badges.push('<span class="badge" style="background:#673AB7">📅 By Appointment</span>');

  let ratingHTML = "";
  if (c.rating > 0) {
    ratingHTML = `<div style="display:flex; align-items:center; gap:4px; margin:6px 0; font-size:13px; font-weight:800; color:#FF9800;">
      <span>⭐ ${parseFloat(c.rating).toFixed(1)}</span>
      <span style="color:#90A4AE; font-weight:600; font-size:12px;">(${c.review_count} reviews)</span>
    </div>`;
  }

  const directionsHTML = (c.lat && c.lng)
    ? '<a href="https://www.google.com/maps/dir/?api=1&destination=' + c.lat + ',' + c.lng + '"'
      + ' target="_blank" rel="noopener"'
      + ' style="display:flex;align-items:center;justify-content:center;gap:8px;margin:10px 0 8px;padding:12px;'
      + ' background:#FF9800;color:#fff !important;text-decoration:none;border-radius:12px;'
      + ' font-weight:800;font-size:14px;text-shadow:0 1px 2px rgba(0,0,0,.2);">'
      + '🗺️ Get Directions</a>'
    : '';

  let priceHTML = "";
  if (c.prices && c.prices.length > 0) {
      priceHTML = `<div style="font-weight:800; color:#388E3C; margin: 10px 0 4px; font-size:13px;">💰 Estimated Pricing</div>
      <div class="price-list">`;
      c.prices.forEach(p => {
          priceHTML += `<div class="price-item">${esc(p)}</div>`;
      });
      priceHTML += `</div>
      <div style="font-size:10px; color:#F44336; margin-top:4px; font-weight:700;">⚠️ Prices may be outdated. Please confirm with the clinic.</div>
      <button onclick="reportPrice('${c.id}', '${esc(c.name).replace(/'/g, "\\'")}')" style="width:100%; margin-top:6px; background:#fff; border:1px solid #388E3C; color:#388E3C; padding:8px; border-radius:8px; font-weight:bold; font-family:inherit; cursor:pointer; font-size:12px;">✍️ Review / Update Price</button>
      `;
  }

  let transportHtml = `
    <div style="margin-top:14px; border-top:2px dashed #EEE; padding-top:12px; margin-bottom:10px;">
      <div style="font-weight:800; color:#388E3C; margin-bottom:8px; font-size:14px;">🚗 Need transport to the vet?</div>
      <div style="display:flex; gap:6px; flex-wrap:wrap; margin-bottom:10px;">
        <button class="trans-btn" style="background-color:#FFF9C4; color:#F57F17; border:1px solid #FFD54F; font-weight:bold; padding:8px 12px; border-radius:10px; font-family:inherit; cursor:pointer;" onclick="L.DomEvent.stopPropagation(event); renderTransport('${c.id}', 'pet')">GrabPet</button>
        <button class="trans-btn" style="background-color:#FFEBEE; color:#C62828; border:1px solid #FFCDD2; font-weight:bold; padding:8px 12px; border-radius:10px; font-family:inherit; cursor:pointer;" onclick="L.DomEvent.stopPropagation(event); renderTransport('${c.id}', 'ambulance')">Petbulance</button>
        <button class="trans-btn" style="background-color:#F3E5F5; color:#6A1B9A; border:1px solid #E1BEE7; font-weight:bold; padding:8px 12px; border-radius:10px; font-family:inherit; cursor:pointer;" onclick="L.DomEvent.stopPropagation(event); renderTransport('${c.id}', 'mobile')">Mobile Vet</button>
      </div>
      <div id="transport-res-${c.id}" style="display:none; background:#F9FBFD; padding:10px; border-radius:10px; font-size:12px; border:1px solid #E1E8ED;"></div>
      <div style="font-size:10px; color:#90A4AE; margin-top:6px; font-weight:700;">💡 MyBuddy takes 0% commission. All transport partners are independent.</div>
    </div>
  `;

  return '<div class="popup"><h3>'+esc(c.name)+'</h3>'
    + (c.branch ? '<div style="color:#90A4AE;font-size:12px;font-weight:600;margin-bottom:6px;">'+esc(c.branch)+'</div>' : '')
    + ratingHTML
    + '<p style="color:'+color+'; font-weight:800; font-size:13px;">'+st.label+'</p>'
    + '<p style="font-size:13px; color:#555;">'+esc(c.address)+'<br>Singapore '+esc(c.postal)+'</p>'
    + directionsHTML
    + (badges.length ? '<div>'+badges.join("")+'</div>' : '')
    + priceHTML
    + transportHtml
    + '<div style="margin-top:10px; border-top:2px dashed #EEE; padding-top:10px;">'
    + (c.phone ? '<a class="btn-link" href="tel:'+primaryPhone(c.phone)+'">📞 Call '+esc(c.phone.split('/')[0].trim())+'</a>' : '')
    + (wa ? '<a class="btn-link wa" href="'+esc(wa)+'" target="_blank" rel="noopener">💬 WhatsApp</a>' : '')
    + (c.booking ? '<a class="btn-link" style="background:#1976D2;" href="'+esc(c.booking)+'" target="_blank" rel="noopener">📅 Book Online</a>' : '')
    + '</div>'
    + sibHTML
    + '</div>';
}

window.renderTransport = function(clinicId, type) {
  const resBox = document.getElementById('transport-res-' + clinicId);
  if (!resBox) return;

  if (resBox.style.display === 'block' && resBox.dataset.type === type) {
    resBox.style.display = 'none';
    return;
  }

  resBox.style.display = 'block';
  resBox.dataset.type = type;

  const data = TRANSPORT_DATA[type] || [];
  if (!data.length) {
    resBox.innerHTML = "<i>No providers listed.</i>";
    return;
  }

  let html = '<div style="padding:4px 0;">';
  data.forEach(p => {
    let linkHtml = "";
    if (p.url.startsWith('grab://')) {
        linkHtml = `<a class="main-action-btn" href="javascript:void(0)" onclick="window.location.href='${p.url}'; setTimeout(function(){ window.location.href='${p.fallback}'; }, 500);">${p.action}</a>`;
    } else {
        linkHtml = `<a class="main-action-btn" href="${p.url}" target="_blank">${p.action}</a>`;
    }

    html += `<div style="margin-bottom:12px; border-bottom:1px solid #EEE; padding-bottom:10px;">
        <div style="margin-bottom:6px;"><b>${p.provider}</b> <span style="color:#90A4AE; font-size:11px;">(${p.type})</span></div>
        ${linkHtml}
        ${p.note ? `<div style="color:#555; font-size:12px; font-weight:700; margin-top:6px; line-height:1.4; background:#fff; padding:8px; border-radius:8px; border:1px solid #EEE;">${p.note}</div>` : ''}
    </div>`;
  });
  html += '</div>';
  resBox.innerHTML = html;
};

const TRANSPORT_DATA = {
  pet: [
    {
      provider: "GrabPet",
      type: "Pet Transport",
      action: "Book a GrabPet Ride",
      url: "grab://open?screenType=TRANSPORT",
      fallback: "https://www.grab.com",
      note: "👇 After clicking: Enter your destination, tap \"View All\", and scroll down to select GrabPet."
    }
  ],
  ambulance: [
    { provider: "Royal Animal Rescue", type: "Petbulance", action: "Call +65 9489 1781", url: "tel:+6594891781" },
    { provider: "SPCA Community Emergencies", type: "Community Rescue", action: "Call +65 6287 5355 (Ext. 9)", url: "tel:+6562875355" }
  ],
  mobile: [
    { provider: "Fur Folks Home / VetMobile", type: "Mobile Vet", action: "Call +65 9755 2015", url: "tel:+6597552015" }
  ]
};

async function geocodeCentre(query) {
  const isPostal = /^\d{6}$/.test(query);
  const url = "https://www.onemap.gov.sg/api/common/elastic/search?searchVal="
    + encodeURIComponent(isPostal ? query : query + ", Singapore")
    + "&returnGeom=Y&getAddrDetails=Y&pageNum=1";
  const r = await fetch(url); const j = await r.json();
  const first = (j.results || [])[0];
  if (!first) return null;
  const lat = parseFloat(first.LATITUDE), lng = parseFloat(first.LONGITUDE);
  if (isNaN(lat) || isNaN(lng)) return null;
  return { lat, lng, isPostal };
}

const layer = L.layerGroup().addTo(map);
const markers = [];
const markersById = {};
const clinicsById = {};
let currentFilter = "all";
let currentSpecies = "all";
let searchText = "";
let selectedId = null;

function updateVisibility() {
  const centre = centreCircle.getLatLng();
  const radius = centreCircle.getRadius();
  const q = searchText.trim().toLowerCase();
  const hasSearch = q.length > 0;
  let visible = 0;

  if (selectedId) {
    markers.forEach(m => {
      const on = m.clinic.id === selectedId;
      if (on) { if (!layer.hasLayer(m.marker)) layer.addLayer(m.marker); visible++; }
      else if (layer.hasLayer(m.marker)) layer.removeLayer(m.marker);
    });
    const c = clinicsById[selectedId];
    document.getElementById("status").textContent = c ? "Selected: " + c.name : "";
    return;
  }

  markers.forEach(m => {
    const c = m.clinic, st = m.status;
    const isEmergency = c.ecat === "E1_CONFIRMED_24_7";
    let show;
    if (isEmergency) show = true;
    else if (currentFilter === "house") show = c.house_call;
    else if (hasSearch) show = c.name.toLowerCase().includes(q) || c.id.toLowerCase().includes(q);
    else show = map.distance(centre, m.marker.getLatLng()) <= radius;

    if (show && !isEmergency && currentFilter === "open" && st.key !== "open") show = false;
    if (show && currentFilter === "emergency" && !isEmergency) show = false;

    // Species filter: hide only when we KNOW it doesn't treat that species.
    if (show && currentSpecies !== "all") {
      const sp = c.species || [];
      if (sp.length > 0 && !sp.includes(currentSpecies)) show = false;
    }

    const shouldBlink = (currentFilter === "house" && c.house_call);
    const desiredIcon = pinIcon(c, st, false, shouldBlink);
    if (m.marker.options.icon.options.html !== desiredIcon.options.html) m.marker.setIcon(desiredIcon);

    if (show) { if (!layer.hasLayer(m.marker)) layer.addLayer(m.marker); visible++; }
    else if (layer.hasLayer(m.marker)) layer.removeLayer(m.marker);
  });
  document.getElementById("status").textContent = hasSearch
    ? "Search: " + visible + " matching clinic(s)"
    : "Showing " + visible + " of " + markers.length + " clinics inside radius";
}

document.querySelectorAll('.sp-btn, .m-sp-btn').forEach(b => {
  b.addEventListener('click', () => {
    const sp = b.dataset.sp;
    currentSpecies = (currentSpecies === sp) ? "all" : sp;
    document.querySelectorAll('.sp-btn, .m-sp-btn').forEach(x => {
      x.classList.toggle('on', x.dataset.sp === currentSpecies);
    });
    selectedId = null;
    updateVisibility();
  });
});

document.getElementById("radius").addEventListener("input", e => {
  const km = parseFloat(e.target.value);
  centreCircle.setRadius(km * 1000);
  document.getElementById("radiusVal").textContent = km.toFixed(1) + " km";
  updateVisibility();
});

async function setCentreFromInput() {
  const q = document.getElementById("centreInput").value.trim();
  if (!q) return;
  const fb = document.getElementById("centreFeedback");
  fb.textContent = "Finding location…";
  try {
    const res = await geocodeCentre(q);
    if (!res) throw new Error();
    const ll = L.latLng(res.lat, res.lng);
    centreCircle.setLatLng(ll); centreHandle.setLatLng(ll);
    map.setView(ll, res.isPostal ? 15 : 13);
    fb.textContent = ""; updateVisibility();
  } catch(e) { fb.textContent = "Could not find: " + q; }
}
document.getElementById("setCentre").addEventListener("click", setCentreFromInput);
document.getElementById("centreInput").addEventListener("keydown", e => { if (e.key === "Enter") setCentreFromInput(); });
document.getElementById("resetCentre").addEventListener("click", () => {
  centreCircle.setLatLng(SINGAPORE); centreHandle.setLatLng(SINGAPORE);
  map.setView(SINGAPORE, 11);
  document.getElementById("centreFeedback").textContent = "";
  clearSearch();
});

const searchInput = document.getElementById("search");
const directory   = document.getElementById("directory");
const statusEl    = document.getElementById("status");

function renderDirectory(filter) {
  const q = (filter||"").trim().toLowerCase();
  directory.innerHTML = "";
  const letters = Object.keys(DATA.directory).sort();
  letters.forEach(L => {
    const clinics = DATA.directory[L];
    const matches = q ? clinics.filter(c => c.name.toLowerCase().includes(q)) : clinics;
    if (!matches.length) return;
    const section = document.createElement("div");
    section.className = "dir-letter" + (q ? " open" : "");
    const header = document.createElement("div");
    header.className = "dir-header";
    header.textContent = L + " (" + matches.length + ")";
    header.onclick = () => section.classList.toggle("open");
    section.appendChild(header);
    const list = document.createElement("div");
    list.className = "dir-items";
    matches.forEach(c => {
      const item = document.createElement("div");
      item.className = "dir-item";
      item.textContent = c.name;
      item.onclick = () => selectClinic(c.id);
      list.appendChild(item);
    });
    section.appendChild(list);
    directory.appendChild(section);
  });
}
function openDirectory() { renderDirectory(searchInput.value); directory.classList.add("open"); }
function closeDirectory() { directory.classList.remove("open"); }

searchInput.addEventListener("focus", openDirectory);
searchInput.addEventListener("click", openDirectory);
searchInput.addEventListener("input", () => {
  searchText = searchInput.value;
  selectedId = null;
  const q = searchText.trim().toLowerCase();
  if (q) {
    const matches = Object.values(markersById).filter(m =>
      m.clinic.name.toLowerCase().includes(q) || m.clinic.id.toLowerCase().includes(q));
    if (matches.length === 1) selectClinic(matches[0].clinic.id, { keepInput: true });
    else updateVisibility();
  } else {
    updateVisibility();
  }
  openDirectory();
});
searchInput.addEventListener("keydown", e => { if (e.key === "Escape") { clearSearch(); closeDirectory(); } });
document.addEventListener("click", e => { if (!e.target.closest(".search-wrap")) closeDirectory(); });

function clearSearch() {
  searchInput.value = "";
  searchText = "";
  selectedId = null;
  updateVisibility();
}

function selectClinic(id, opts) {
  const m = markersById[id];
  if (!m) return;
  selectedId = id;
  if (!(opts && opts.keepInput)) {
    searchInput.value = m.clinic.name;
    searchText = m.clinic.name.toLowerCase();
  }
  closeDirectory();

  if (window.innerWidth <= 768) {
    document.getElementById('mobile-dir-dropdown').classList.remove('open');
  }

  updateVisibility();
  const ll = m.marker.getLatLng();
  map.flyTo(ll, 16, { duration: 0.8 });
  m.marker.openPopup();
}

let unresolved = 0;
DATA.clinics.forEach(c => {
  clinicsById[c.id] = c;
  if (c.lat == null || c.lng == null) { unresolved++; return; }
  const st = statusFor(c);
  const marker = L.marker([c.lat, c.lng], { icon: pinIcon(c, st, false) });
  marker.bindPopup(() => popupHTML(c, statusFor(c)));
  marker.clinic = c; marker.status = st;
  markers.push({ marker, clinic: c, status: st });
  markersById[c.id] = { marker, clinic: c, status: st };
  layer.addLayer(marker);
});

updateVisibility();
if (unresolved) statusEl.textContent += " · " + unresolved + " without coordinates";
setInterval(updateVisibility, 60000);

(function mobileInit() {
  const mSearchInput = document.getElementById('mobile-search-input');
  const mDropdown = document.getElementById('mobile-dir-dropdown');
  const mRadiusSlider = document.getElementById('mobile-radius-input');
  const mHouseBtn = document.getElementById('mobile-house-btn');

  function buildMobileDropdown(filter = "") {
    mDropdown.innerHTML = "";
    const q = filter.toLowerCase();
    const letters = Object.keys(DATA.directory).sort();

    letters.forEach(letter => {
      const clinics = DATA.directory[letter].filter(c => c.name.toLowerCase().includes(q));
      if (!clinics.length) return;

      const header = document.createElement('div');
      header.className = 'md-letter';
      header.textContent = letter + ' (' + clinics.length + ')';

      const list = document.createElement('div');
      list.className = 'md-items';

      header.onclick = () => {
          header.classList.toggle('open');
      };

      mDropdown.appendChild(header);

      clinics.forEach(c => {
        const item = document.createElement('div');
        item.className = 'md-item';
        item.textContent = c.name;
        item.onclick = () => {
          mSearchInput.value = c.name;
          mDropdown.classList.remove('open');
          selectClinic(c.id);
        };
        list.appendChild(item);
      });

      mDropdown.appendChild(list);
    });
  }

  mSearchInput.addEventListener('focus', () => {
    buildMobileDropdown(mSearchInput.value);
    mDropdown.classList.add('open');
  });
  mSearchInput.addEventListener('input', () => {
    buildMobileDropdown(mSearchInput.value);
  });
  document.addEventListener('click', (e) => {
    if (!e.target.closest('.m-search-wrap')) mDropdown.classList.remove('open');
  });

  mHouseBtn.addEventListener('click', () => {
      if (currentFilter === "house") {
          currentFilter = "all";
          mHouseBtn.classList.remove('on');
          updateVisibility();
      } else {
          currentFilter = "house";
          mHouseBtn.classList.add('on');
          updateVisibility();
          map.fitBounds([[1.1304, 103.6020], [1.4784, 104.0945]]);
      }
  });

  mRadiusSlider.addEventListener('input', (e) => {
    const km = parseFloat(e.target.value);
    centreCircle.setRadius(km * 1000);
    if (km >= 30) map.fitBounds([[1.1304, 103.6020], [1.4784, 104.0945]]);
    else updateVisibility();
  });

  if (window.innerWidth <= 768) {
    mRadiusSlider.value = 2;
    centreCircle.setRadius(2000);

    if (navigator.geolocation) {
      navigator.geolocation.getCurrentPosition((pos) => {
        const ll = L.latLng(pos.coords.latitude, pos.coords.longitude);
        centreCircle.setLatLng(ll);
        centreHandle.setLatLng(ll);
        map.setView(ll, 14);
        updateVisibility();
      });
    }
  }
})();
</script>
</body>
</html>
'''

if __name__ == "__main__":
    build()