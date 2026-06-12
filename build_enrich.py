#!/usr/bin/env python3
"""
Enrich PS2-Explorer/stations.json (PSMS extension export, 554 stations) with two
signals derived from the three official/crowdsourced spreadsheets:

  1. projectDomains + clean bizDomain   <- "List of stations with project titles..."
                                           (merged by Station Id, exact)
  2. visited / visitedCompanies         <- "PS Tracker" interaction-company list
                                           (conservative company-name match)

Run:  python3 build_enrich.py            # dry run, prints a report only
      python3 build_enrich.py --write    # also rewrites stations.json in place

Source .xlsx live in ~/Downloads (override with env vars below).
"""
import json, re, sys, os, datetime
from pathlib import Path
import openpyxl

HERE = Path(__file__).resolve().parent
DL = Path(os.path.expanduser("~/Downloads"))
F_PROJ = Path(os.environ.get("PROJ_XLSX",
    DL / "List of stations with project titles and description - June 12th, 2026.xlsx"))
F_TRACK = Path(os.environ.get("TRACK_XLSX",
    DL / "PS Tracker - Sem 1 2026-27.xlsx"))
STATIONS_JSON = HERE / "stations.json"
WRITE = "--write" in sys.argv


def cells(fn, sheet=None):
    wb = openpyxl.load_workbook(fn, read_only=True, data_only=True)
    ws = wb[sheet] if sheet else wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    hdr = [str(c).strip() if c is not None else "" for c in rows[0]]
    out = []
    for r in rows[1:]:
        out.append({hdr[i]: r[i] for i in range(min(len(hdr), len(r)))})
    return out


# ---------------------------------------------------------------- project file
def clean(s):
    return ("" if s is None else str(s)).replace("\xa0", " ").strip()


proj_rows = cells(F_PROJ)
proj_by_id = {}   # int id -> {bizDomain, city, domains:set}
for r in proj_rows:
    sid = r.get("Station Id")
    if sid is None:
        continue
    try:
        sid = int(sid)
    except (ValueError, TypeError):
        continue
    e = proj_by_id.setdefault(sid, {"bizDomain": "", "city": "", "domains": set()})
    bd = clean(r.get("Business Domain"))
    if bd and bd != "-":
        e["bizDomain"] = bd
    city = clean(r.get("City"))
    if city and not e["city"]:
        e["city"] = city
    pd = clean(r.get("Project Domain"))
    if pd and pd not in ("-", "Yet to be finalized", "Details awaited", "Others"):
        for tok in pd.split(","):
            tok = tok.strip()
            if tok and tok not in ("-",):
                e["domains"].add(tok)


# --------------------------------------------------------------- tracker file
track_rows = cells(F_TRACK, "PS Tracker")
companies = []
last = None
for r in track_rows:
    c = clean(r.get("Company"))
    if c and c.lower() != "company":
        last = c
        if c not in companies:
            companies.append(c)
# `companies` = unique interaction-company names in announce order


def normalize(name):
    n = name.lower()
    n = re.sub(r"[^a-z0-9 ]", " ", n)
    n = re.sub(r"\s+", " ", n).strip()
    for w in (" limited", " ltd", " pvt", " private", " inc", " llp", " co ",
              " india", " technologies", " technology", " services", " solutions",
              " systems", " group", " labs", " lab", " corp", " company"):
        n = n.replace(w, " ")
    return re.sub(r"\s+", " ", n).strip()


def nospace(s):
    return s.replace(" ", "")


MANUAL_MATCHES = {
    "ABSTC Mumbai": ["aditya birla science technology company abstc mumbai"],
    "Kisetsu Saison Finance (India) Private Limited (Credit Saison India)":
        ["credit saison engineering"],
    "PwC US Advisory, Bangalore": ["pwc us advisory oscs"],
    "Dixon/PADGET": ["padget electronics dixon"],
    "Oliver Wyman": ["oliver wayman"],
    "Swiss Re GBS": ["swiss re global business"],
    "ENDPOINT eCLINICAL": ["endpoint clinical india"],
    "McKinsey Client Capability network": [
        "mckinsey research insights intern with arr",
        "mckinsey shape", "mckinsey tech ecosystem"],
}
MANUAL_NORM = {normalize(k): [v for v in vs] for k, vs in MANUAL_MATCHES.items()}


def matches(company, station):
    """Replicates the conservative matcher from the Claude.ai pipeline."""
    cn, sn = normalize(company), normalize(station)
    if not cn or not sn:
        return False
    cl = company.lower()
    # ---- special cases ----
    if cl.startswith("csir-4pi") or cl == "csir-4pi":
        return station.strip().lower() == "csir-4pi"
    if "siemens eda" in cl:
        s = station.lower()
        return "siemens" in s and "eda" in s
    if "siemens energy" in cl:          # SEIL — distinct entity from Siemens EDA
        return "siemens energy" in station.lower()
    if cl.startswith("itc"):           # match all ITC stations
        return station.lower().startswith("itc")
    if cl.startswith("v-gaurd") or cl.startswith("v-guard"):   # tracker typo
        return station.lower().startswith("v-guard")
    if "qapita" in cl:                  # QFIPL(...) obfuscated station name
        return "qapita" in station.lower()
    # ---- manual ----
    if cn in MANUAL_NORM and sn in MANUAL_NORM[cn]:
        return True
    # ---- generic strict prefix ----
    strict = (cn == sn) or sn.startswith(cn + " ") or cn.startswith(sn + " ")
    if not strict and len(nospace(cn)) >= 6 and len(nospace(sn)) >= 6:
        strict = nospace(sn).startswith(nospace(cn)) or nospace(cn).startswith(nospace(sn))
    return strict


# ------------------------------------------------------------------- merge
doc = json.load(open(STATIONS_JSON))
stations = doc["stations"]

n_dom = n_biz = n_vis = 0
matched_companies = set()
for s in stations:
    try:
        sid = int(s["stationId"])
    except (ValueError, TypeError, KeyError):
        sid = None
    info = proj_by_id.get(sid)
    if info:
        if info["domains"]:
            s["projectDomains"] = sorted(info["domains"]); n_dom += 1
        else:
            s["projectDomains"] = []
        if info["bizDomain"]:
            s["bizDomain"] = info["bizDomain"]; n_biz += 1
        if info["city"] and not s.get("city"):
            s["city"] = info["city"]
    else:
        s.setdefault("projectDomains", [])

    name = s.get("stationName", "")
    hits = [c for c in companies if matches(c, name)]
    s["visited"] = bool(hits)
    s["visitedCompanies"] = hits
    if hits:
        n_vis += 1
        matched_companies.update(hits)

doc["enriched_at"] = datetime.datetime.now().isoformat(timespec="seconds")

print(f"stations            : {len(stations)}")
print(f"projectDomains added: {n_dom}")
print(f"clean bizDomain     : {n_biz}")
print(f"interaction companies in tracker: {len(companies)}")
print(f"visited stations    : {n_vis}")
print(f"distinct companies matched: {len(matched_companies)} / {len(companies)}")
unmatched = [c for c in companies if c not in matched_companies]
print(f"unmatched companies ({len(unmatched)}):")
for c in unmatched:
    print("   -", c)

if WRITE:
    json.dump(doc, open(STATIONS_JSON, "w"), ensure_ascii=False, indent=0)
    print(f"\nWROTE {STATIONS_JSON}")
else:
    print("\n(dry run — pass --write to update stations.json)")
