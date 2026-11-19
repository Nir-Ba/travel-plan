#!/usr/bin/env python3
"""Turn a Google Takeout export (Google Maps saved lists + starred places) into one
spreadsheet (CSV) to review in Excel (save back as CSV UTF-8), then import into Travel Plan.

Usage (Mac):  python3 tools/takeout_to_sheet.py ~/Downloads/Takeout
Output:       <Takeout folder>/Travel Plan import.csv

Each row gets a City, a Category and a "Check" note that says how sure the city is:
  OK          - the place has a map point, or it was found near its list's city
  CHECK       - found, but far from the list's city, or the city is a guess
  NOT FOUND   - no match on the map; the city (if any) comes from the list name
Map search: Photon (OpenStreetMap, free, no key). Results are cached in
<Takeout folder>/.travelplan-cache.json, so a second run is fast.
"""
import csv, json, math, os, re, sys, time, urllib.parse, urllib.request

VERSION = "1.0.0"
UA = "TravelPlan-import/" + VERSION

# Hebrew / other list names -> what the list is.
# ("city", name) = all places in that city. ("country", CC, centre query) = places across a country.
LISTS = {
    "מדריד": ("city", "Madrid"), "פראג": ("city", "Prague"), "ניו יורק": ("city", "New York"),
    "בנגקוק": ("city", "Bangkok"), "לוס אנגלס": ("city", "Los Angeles"), "סן פרנסיסקו": ("city", "San Francisco"),
    "ברלין": ("city", "Berlin"), "מלבורן - מלונות": ("city", "Melbourne"), "טוקיו": ("city", "Tokyo"),
    "אתונה": ("city", "Athens"), "בריסל": ("city", "Brussels"), "אמסטרדם": ("city", "Amsterdam"),
    "פריז": ("city", "Paris"), "לונדון": ("city", "London"), "ברטיסלבה": ("city", "Bratislava"),
    "מרסיי": ("city", "Marseille"), "מילאנו": ("city", "Milan"), "דבלין": ("city", "Dublin"),
    "מיאמי": ("city", "Miami"), "ברצלונה": ("city", "Barcelona"), "פוקט": ("city", "Phuket"),
    "רומא": ("city", "Rome"), "בוסטון": ("city", "Boston"), "ג׳נבה": ("city", "Geneva"),
    "פרנקפורט": ("city", "Frankfurt"),
    "פורטוגל": ("country", "PT", "Portugal"), "קפריסין": ("country", "CY", "Cyprus"),
    "אוסטרליה": ("country", "AU", "Australia"), "מרוקו": ("country", "MA", "Morocco"),
    "סרדיניה": ("country", "IT", "Sardinia"), "דרום איטליה": ("country", "IT", "Naples"),
    "יפן טיול": ("country", "JP", "Japan"), "פיליפינים": ("country", "PH", "Philippines"),
    "יוון": ("country", "GR", "Greece"), "New Zealand": ("country", "NZ", "New Zealand"),
    "Faroe Islands": ("country", "FO", "Faroe Islands"),
}
SKIP_LISTS = {"Images"}  # not places (image links)
NEAR_KM = 40  # "in this city" if the place is this close to the city centre

CATS = {  # OpenStreetMap tag -> app category
    ("amenity", "restaurant"): "Restaurant", ("amenity", "fast_food"): "Restaurant", ("amenity", "food_court"): "Restaurant",
    ("amenity", "cafe"): "Café", ("amenity", "ice_cream"): "Café", ("shop", "bakery"): "Café",
    ("amenity", "bar"): "Bar", ("amenity", "pub"): "Bar", ("amenity", "biergarten"): "Bar", ("amenity", "nightclub"): "Bar",
    ("amenity", "place_of_worship"): "Sight", ("amenity", "theatre"): "Sight", ("place", "square"): "Sight",
    ("leisure", "park"): "Park", ("leisure", "garden"): "Park", ("leisure", "nature_reserve"): "Park",
    ("boundary", "national_park"): "Park", ("tourism", "zoo"): "Park",
    ("tourism", "museum"): "Museum", ("tourism", "gallery"): "Museum",
    ("tourism", "hotel"): "Hotel", ("tourism", "hostel"): "Hotel", ("tourism", "guest_house"): "Hotel",
    ("tourism", "apartment"): "Hotel", ("tourism", "motel"): "Hotel", ("tourism", "camp_site"): "Hotel",
    ("tourism", "caravan_site"): "Hotel",
}

def category(key, val):
    if (key, val) in CATS: return CATS[(key, val)]
    if key == "shop": return "Shop"
    if key in ("tourism", "historic", "man_made"): return "Sight"
    return "Other"

def km(a, b):
    r = math.pi / 180
    h = math.sin((b[0] - a[0]) * r / 2) ** 2 + math.cos(a[0] * r) * math.cos(b[0] * r) * math.sin((b[1] - a[1]) * r / 2) ** 2
    return 2 * 6371 * math.asin(math.sqrt(h))

# ---- Photon with a disk cache ----
cache, cache_path, calls = {}, "", 0
def photon(path, params):
    global calls
    url = "https://photon.komoot.io/" + path + "?" + urllib.parse.urlencode(params, doseq=True)
    if url in cache: return cache[url]
    for attempt in range(3):
        try:
            time.sleep(0.25)
            with urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": UA}), timeout=20) as r:
                feats = json.load(r).get("features", [])
            calls += 1
            cache[url] = feats
            if calls % 25 == 0: save_cache()
            return feats
        except Exception as e:
            print(f"  photon retry {attempt + 1}: {e}", file=sys.stderr); time.sleep(2 * (attempt + 1))
    return []

def save_cache():
    with open(cache_path, "w") as f: json.dump(cache, f)

def city_of(p):
    """City name from Photon properties. Tokyo's wards count as Tokyo."""
    if (p.get("countrycode") or "").upper() == "JP" and (p.get("state") or "").lower().startswith("tokyo"): return "Tokyo"
    return p.get("city") or p.get("town") or p.get("village") or p.get("county") or p.get("state") or ""

centres = {}
def centre(q):
    if q not in centres:
        f = photon("api", {"q": q, "limit": 1, "lang": "en", "osm_tag": "place"}) or photon("api", {"q": q, "limit": 1, "lang": "en"})
        centres[q] = (f[0]["geometry"]["coordinates"][1], f[0]["geometry"]["coordinates"][0]) if f else None
        if not centres[q]: print(f"  no centre found for {q!r}", file=sys.stderr)
    return centres[q]

def find(name, near=None, cc=None):
    """Best venue match for a name. Skips towns/streets; prefers the sight over a station named after it."""
    params = {"q": name, "limit": 8, "lang": "en", "osm_tag": ["!place", "!boundary", "!highway"]}
    if near: params.update(lat=near[0], lon=near[1])
    hits = []
    for f in photon("api", params):
        p = f["properties"]; ll = (f["geometry"]["coordinates"][1], f["geometry"]["coordinates"][0])
        if cc and (p.get("countrycode") or "").upper() != cc: continue
        hits.append((ll, p))
    venue = [h for h in hits if h[1].get("osm_key") not in ("railway", "public_transport")]
    return (venue or hits or [None])[0]

def reverse(ll):
    f = photon("reverse", {"lat": ll[0], "lon": ll[1], "lang": "en"})
    return f[0]["properties"] if f else {}

def ll_from_url(u):
    m = re.search(r"!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)", u) or re.search(r"/search/(-?\d+\.\d+),\s*(-?\d+\.\d+)", u) \
        or re.search(r"[?&]q=(-?\d+\.\d+),(-?\d+\.\d+)", u) or re.search(r"@(-?\d+\.\d+),(-?\d+\.\d+)", u)
    return (float(m.group(1)), float(m.group(2))) if m else None

def address(p):
    parts = [p.get("name"), p.get("street"), p.get("district"), city_of(p), p.get("country")]
    out = []
    for x in parts:
        if x and x not in out: out.append(x)
    return ", ".join(out)

def read_list(path):
    """Rows of a Saved/<list>.csv. Some lists have a description line before the header."""
    with open(path, encoding="utf-8-sig", newline="") as f: rows = list(csv.reader(f))
    for i, r in enumerate(rows):
        if r[:1] == ["Title"]:
            head = r; body = rows[i + 1:]; break
    else:
        print(f"  skipped {os.path.basename(path)}: no Title header", file=sys.stderr); return []
    col = {h: k for k, h in enumerate(head)}
    out = []
    for r in body:
        g = lambda h: (r[col[h]] if h in col and col[h] < len(r) else "").strip()
        if not g("Title") and not g("URL"): continue
        if "google.com/maps" not in g("URL"):
            print(f"  skipped (not a Google Maps link): {g('Title')[:50]!r}", file=sys.stderr); continue
        out.append({"name": " ".join(g("Title").split()) or "Dropped pin", "note": g("Note") or g("Comment"), "url": g("URL")})
    return out

def main():
    global cache, cache_path
    root = os.path.expanduser(sys.argv[1] if len(sys.argv) > 1 else "~/Downloads/Takeout")
    print(f"takeout_to_sheet {VERSION}: reading {root}")
    cache_path = os.path.join(root, ".travelplan-cache.json")
    if os.path.exists(cache_path):
        cache = json.load(open(cache_path)); print(f"  cache: {len(cache)} earlier answers")
    rows = []

    def add(name, note, src, url, city, cat, check, found="", ll=None):
        rows.append({"Include": "yes", "City": city, "Category": cat, "Name": name, "Note": note, "Check": check,
                     "Found on map": found, "Source list": src,
                     "Lat": f"{ll[0]:.6f}" if ll else "", "Lon": f"{ll[1]:.6f}" if ll else "", "Google link": url})

    # 1. Saved lists
    sdir = os.path.join(root, "Saved")
    files = sorted(f for f in os.listdir(sdir) if f.endswith(".csv")) if os.path.isdir(sdir) else []
    if not files: print("  no Saved/ folder or no lists in it", file=sys.stderr)
    for fn in files:
        lst = fn[:-4]
        if lst in SKIP_LISTS: print(f"- {lst}: skipped (not places)"); continue
        kind = LISTS.get(lst, ("theme",))
        items = read_list(os.path.join(sdir, fn))
        print(f"- {lst}: {len(items)} place(s), list type {kind[0]}")
        for it in items:
            ll = ll_from_url(it["url"])
            if kind[0] == "city":
                c = centre(kind[1])
                if ll:  # a dropped pin: we know the point
                    p = reverse(ll); far = c and km(ll, c) > NEAR_KM
                    add(it["name"], it["note"], lst, it["url"], kind[1] if not far else city_of(p), category(p.get("osm_key"), p.get("osm_value")),
                        "CHECK: pin is far from " + kind[1] if far else "OK", address(p), ll); continue
                h = find(it["name"], c)
                if not h:
                    add(it["name"], it["note"], lst, it["url"], kind[1], "Other", "NOT FOUND: city from list name"); continue
                d = km(h[0], c) if c else 0
                if d <= NEAR_KM:
                    add(it["name"], it["note"], lst, it["url"], kind[1], category(h[1].get("osm_key"), h[1].get("osm_value")), "OK", address(h[1]), h[0])
                else:  # best match is far away: likely a different branch. Keep the list's city, no point.
                    add(it["name"], it["note"], lst, it["url"], kind[1], category(h[1].get("osm_key"), h[1].get("osm_value")),
                        f"CHECK: best match is {round(d)} km away", address(h[1]))
            else:
                cc = kind[1] if kind[0] == "country" else None
                near = centre(kind[2]) if kind[0] == "country" else None
                if ll:
                    p = reverse(ll)
                    add(it["name"], it["note"], lst, it["url"], city_of(p), category(p.get("osm_key"), p.get("osm_value")), "OK", address(p), ll); continue
                h = find(it["name"], near, cc)
                if not h:
                    add(it["name"], it["note"], lst, it["url"], "", "Other", "NOT FOUND: add the city"); continue
                add(it["name"], it["note"], lst, it["url"], city_of(h[1]), category(h[1].get("osm_key"), h[1].get("osm_value")),
                    "CHECK: city is a guess from the name", address(h[1]), h[0])

    # 2. Starred places (have a point + address)
    sp = os.path.join(root, "Maps (your places)", "Saved Places.json")
    if os.path.exists(sp):
        feats = json.load(open(sp, encoding="utf-8")).get("features", [])
        print(f"- Starred places: {len(feats)} place(s)")
        for f in feats:
            pr = f.get("properties", {}); loc = pr.get("location", {}); c = f.get("geometry", {}).get("coordinates") or [0, 0]
            url = pr.get("google_maps_url", "")
            ll = (c[1], c[0]) if c != [0, 0] else ll_from_url(url)
            name = loc.get("name") or ("Dropped pin" if ll else "")
            if not ll: print(f"  skipped starred place without a point: {name!r}", file=sys.stderr); continue
            p = reverse(ll)
            h = find(name, ll) if loc.get("name") else None
            cat = category(h[1].get("osm_key"), h[1].get("osm_value")) if h and km(h[0], ll) < 0.4 else category(p.get("osm_key"), p.get("osm_value"))
            add(name, pr.get("Comment", "") if loc else "", "Starred", url, city_of(p), cat, "OK", loc.get("address") or address(p), ll)
    else:
        print("  no Maps (your places)/Saved Places.json - skipped starred places", file=sys.stderr)

    save_cache()
    # No city yet, but the name or note says one of the known cities ("Yumoto Spa, ..., Bangkok"): use it.
    known = sorted({r["City"] for r in rows if r["City"]}, key=len, reverse=True)
    fixed = 0
    for r in rows:
        if r["City"]: continue
        text = f'{r["Name"]} {r["Note"]}'.lower()
        hit = next((c for c in known if re.search(r"(^|[\s,(])" + re.escape(c.lower()) + r"($|[\s,)])", text)), None)
        if hit: r["City"] = hit; r["Check"] = "CHECK: city taken from the name"; fixed += 1
    print(f"  city taken from the name for {fixed} row(s)")
    # Problems first, then by city.
    order = {"NOT": 0, "CHE": 1, "OK": 2}
    rows.sort(key=lambda r: (order.get(r["Check"][:3], 3), r["City"].lower(), r["Source list"], r["Name"].lower()))
    out = os.path.join(root, "Travel Plan import.csv")
    with open(out, "w", encoding="utf-8-sig", newline="") as f:  # BOM: Excel shows Hebrew right
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()) if rows else ["Include", "City", "Name"])
        w.writeheader(); w.writerows(rows)
    n = lambda k: sum(1 for r in rows if r["Check"].startswith(k))
    print(f"\nWrote {len(rows)} rows to {out}")
    write_xlsx(rows, os.path.join(root, "Travel Plan import.xlsx"))
    print(f"  OK: {n('OK')}   CHECK: {n('CHECK')}   NOT FOUND: {n('NOT')}   (Photon calls this run: {calls})")

APP_CATS = ["Restaurant", "Café", "Bar", "Park", "Museum", "Sight", "Shop", "Hotel", "Other"]

def write_xlsx(rows, path):
    """Same rows as an Excel file to review: dropdowns, colours, filters. Save it back as CSV UTF-8 for the app."""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.worksheet.datavalidation import DataValidation
    except ImportError:
        print("  skipped Excel file: openpyxl is not installed (pip3 install openpyxl)", file=sys.stderr); return
    if not rows: print("  skipped Excel file: no rows", file=sys.stderr); return
    wb = Workbook(); ws = wb.active; ws.title = "Places"
    cols = list(rows[0].keys()); ws.append(cols)
    for r in rows: ws.append([r[c] for c in cols])
    n = len(rows) + 1
    ws.freeze_panes = "A2"; ws.auto_filter.ref = f"A1:{chr(64 + len(cols))}{n}"
    for c in ws[1]: c.font = Font(bold=True)
    widths = {"Include": 9, "City": 18, "Category": 13, "Name": 34, "Note": 30, "Check": 34, "Found on map": 40,
              "Source list": 18, "Lat": 11, "Lon": 11, "Google link": 14}
    for i, c in enumerate(cols, 1): ws.column_dimensions[chr(64 + i)].width = widths.get(c, 14)
    fills = {"NOT": PatternFill("solid", fgColor="F8D7D3"), "CHE": PatternFill("solid", fgColor="FFF1C2")}
    gl = cols.index("Google link") + 1
    for i, r in enumerate(rows, 2):
        f = fills.get(r["Check"][:3])
        if f:
            for c in ws[i]: c.fill = f
        if not r["City"]: ws.cell(i, cols.index("City") + 1).fill = PatternFill("solid", fgColor="F0A39A")
        link = ws.cell(i, gl)
        if r["Google link"]: link.hyperlink = r["Google link"]; link.font  # keep the URL as the value: the app reads it from the CSV = Font(color="0A7C6E", underline="single")
        for c in ws[i]: c.alignment = Alignment(vertical="top")
    # Dropdowns. City allows new names too (no error), Include and Category do not.
    cities = sorted({r["City"] for r in rows if r["City"]}, key=str.lower)
    cs = wb.create_sheet("Cities")
    for c in cities: cs.append([c])
    col = lambda name: chr(64 + cols.index(name) + 1)
    dv_inc = DataValidation(type="list", formula1='"yes,no"', allow_blank=True)
    dv_cat = DataValidation(type="list", formula1='"' + ",".join(APP_CATS) + '"', allow_blank=True)
    dv_city = DataValidation(type="list", formula1=f"Cities!$A$1:$A${max(1, len(cities))}", allow_blank=True, showErrorMessage=False)
    for dv, name in ((dv_inc, "Include"), (dv_cat, "Category"), (dv_city, "City")):
        ws.add_data_validation(dv); dv.add(f"{col(name)}2:{col(name)}{n}")
    wb.active = 0
    wb.save(path)
    print(f"Wrote Excel file {path}")

if __name__ == "__main__":
    main()
