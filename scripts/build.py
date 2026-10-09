"""Build data/programs.js for the dashboard.

Steps:
  1. Read the shortlist workbook (data/source/*.xlsx, sheet "All Programs").
  2. Re-check every program against the live Études en France catalogue API
     (cached in data/cache/live.json; pass --refresh to fetch again).
  3. Optionally check that program websites still load (--check-web).
  4. Recompute eligibility, years, fees and attach warnings for anything suspicious.

Usage:
  python scripts/build.py                 # use cached live data
  python scripts/build.py --refresh       # re-download all programs from EEF (~5 min)
  python scripts/build.py --check-web     # also re-test program websites
"""
import concurrent.futures as cf
import datetime as dt
import html
import json
import re
import ssl
import sys
import time
import urllib.error
import urllib.request
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
XLSX = ROOT / "data/source/France_Masters_Shortlist_2027.xlsx"
LIVE = ROOT / "data/cache/live.json"
WEB = ROOT / "data/cache/webcheck.json"
OUT = ROOT / "data/programs.js"
API = "https://etudesenfrance.diplomatie.gouv.fr/parcours-api/public/catalogue/formations/{}/site/{}"

# Verified on official pages, 9 Oct 2026.
NON_EU_MASTER = 3950          # service-public.gouv.fr A18927 (2026-27)
EU_MASTER = 255               # service-public.gouv.fr A18927 (2026-27)
EU_ENGINEER = 630             # INSA Toulouse 2026-27 fee page
OLD_MASTER_FEE = 3770         # 2025-26 differentiated master fee

POLICY_OVERRIDES = {
    "Université Paris Cité": (
        "UPDATED: 2026-27 non-EU master students paid the EU rate (€255), but for 2027-28 "
        "Paris Cité says all new non-EU students may be charged the differentiated fee "
        "(€3,950). Only students enrolled since 2025-26/2026-27 keep the old rate.",
        "https://u-paris.fr/droits-dinscription-tarifs-exonerations/",
    ),
    "Université de Tours": (
        "2027-28: partial exemptions capped at 25% of non-EU students, on request "
        "(continuing students, double degrees, personal circumstances). Engineering "
        "diplomas are not subject to differentiated fees at Tours.",
        "https://www.univ-tours.fr/international/etudes-stages-a-tours/droits-dinscription-differencies-pour-les-etudiants-extra-communautaires-2026-2027/droits-dinscription-differencies-pour-les-etudiants-extra-communautaires-2027-2028-2",
    ),
    "Université Grenoble Alpes": (
        "2026-27 was a transition year with broad exemptions (new entrants could apply). "
        "From 2027-28 (your intake) requests are assessed on criteria, priority to "
        "financial/social need, within the 25% cap.",
        "https://www.univ-grenoble-alpes.fr/actualites/a-la-une/actualites-universite/etudiants-internationaux-ce-que-change-le-nouveau-decret-sur-les-droits-d-inscription-a-l-uga-1775870.kjsp",
    ),
}
UNOFFICIAL_SOURCES = ("observalgerie.com", "letudiant.fr", "mastersportal.com")

FRENCH_REQ = re.compile(
    r"DELF|DALF|TCF|TEF\b|"
    r"(fran[cç]ais|french)[^.;]{0,50}\b(B1|B2|C1)\b|\b(B1|B2|C1)\b[^.;]{0,50}(fran[cç]ais|french)",
    re.I,
)
OLD_YEAR = re.compile(r"20(23|24|25)\s*[-/–]\s*20(24|25|26)")
OFF_TOPIC = re.compile(
    r"biom[ée]dicament|pharma|chimie|chemistry|biotechnolog|agro|[ée]cologie|management|droit|"
    r"economi|finance|marketing|gestion",
    re.I,
)


def clean(s):
    if not s:
        return ""
    s = re.sub(r"<br\s*/?>|</p>|</li>", " ", s)
    s = re.sub(r"<[^>]+>", "", s)
    return " ".join(html.unescape(s).split())


def lvl_num(v):
    m = re.search(r"\+(\d)", v or "")
    return int(m.group(1)) if m else None


def read_workbook():
    wb = openpyxl.load_workbook(XLSX)
    ws = wb["All Programs"]
    hdr = [c.value for c in ws[1]]
    rows = []
    for row in ws.iter_rows(min_row=2):
        if not any(c.value for c in row):
            continue
        d = {}
        for h, c in zip(hdr, row):
            d[h] = c.value
            if c.hyperlink:
                d[h + "__link"] = c.hyperlink.target
        rows.append(d)
    return rows


def fetch_one(key):
    url = API.format(*key.split("/"))
    err = "?"
    for attempt in range(6):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0", "Accept": "application/json"})
            with urllib.request.urlopen(req, timeout=30) as r:
                body = r.read().decode("utf-8")
            if body.lstrip().startswith("{"):
                return key, json.loads(body)
            err = "maintenance page"
        except urllib.error.HTTPError as e:
            if e.code in (404, 410):
                return key, {"__error": e.code}
            err = str(e.code)
        except Exception as e:  # network hiccup, retry
            err = type(e).__name__
        time.sleep(2 + attempt * 2)
    return key, {"__error": err}


def fetch_live(keys):
    out = {}
    with cf.ThreadPoolExecutor(6) as ex:
        for i, (k, v) in enumerate(ex.map(fetch_one, keys)):
            out[k] = v
            if i % 100 == 0:
                print(f"  fetched {i}/{len(keys)}", flush=True)
    out["__fetched"] = dt.date.today().isoformat()
    return out


def check_web(urls):
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE

    def chk(u):
        full = u if u.startswith("http") else "https://" + u
        for attempt in range(2):
            try:
                req = urllib.request.Request(full, headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0) Chrome/130"})
                return u, urllib.request.urlopen(req, timeout=25, context=ctx).status
            except urllib.error.HTTPError as e:
                if attempt:
                    return u, e.code
            except Exception as e:
                if attempt:
                    return u, "ERR:" + type(e).__name__
        return u, "?"

    with cf.ThreadPoolExecutor(20) as ex:
        res = dict(ex.map(chk, sorted(urls)))
    res["__checked"] = dt.date.today().isoformat()
    return res


def first_url(s):
    m = re.search(r"https?://\S+", s or "") or re.match(r"www\.\S+", s or "")
    return m.group(0).rstrip(".,;)") if m else None


def city_of(address):
    """Town from the line holding the 5-digit postcode, e.g. '25211 MONTBELIARD CEDEX' -> 'Montbeliard'."""
    for line in reversed((address or "").strip().splitlines()):
        m = re.search(r"\b\d{5}\s+([^\d,]+)", line)
        if m:
            town = re.sub(r"\bcedex\b.*", "", m.group(1), flags=re.I).strip(" ,-")
            return town.title()
    return ""


def city_from_site(site):
    """'Site de Montbéliard' / 'Campus de Lyon' -> town, if the site name looks like one."""
    m = re.match(r"(?:site|campus|p[oô]le)\s+(?:universitaire\s+)?(?:de |d'|du |des )\s*(.+)", site or "", re.I)
    if m and not re.search(r"\d|rue|avenue|bâtiment|batiment", m.group(1), re.I):
        return m.group(1).strip()
    return ""


def route_for(degree, open_levels, any_open):
    """Best way in for a 4-year BS holder (Bac+4 completed). EEF levels are the
    level of the study year you join: M1 = Bac+4, M2 = Bac+5."""
    if not any_open:
        return "closed", "Closed in EEF for 2027", None, None
    if degree == "PhD Doctorat":
        return "no", "PhD – needs a Master first", None, 3
    if degree == "Mastère spécialisé":
        return "possible", "Mastère Spécialisé: Bac+4 + 3 yrs work experience accepted", 6, 1
    if 4 in open_levels:
        return "eligible", "Eligible – join M1 / 4th year (Bac+4 level)", 4, 2
    if 5 in open_levels:
        return "m2", "Only final year (M2, Bac+5) open – possible, needs a strong match", 5, 1
    if 3 in open_levels:
        return "eligible", "Eligible – joins at 3rd-year (Bac+3) level, longer study", 3, 3
    return "no", "Undergraduate entry only – not a fit", min(open_levels), None


def build():
    rows = read_workbook()
    keys = []
    for d in rows:
        m = re.search(r"formation/(\d+)/site/(\d+)", d.get("Catalogue link__link") or "")
        d["_key"] = f"{m.group(1)}/{m.group(2)}"
        keys.append(d["_key"])

    if "--refresh" in sys.argv or not LIVE.exists():
        print("Fetching live data from Études en France …")
        live = fetch_live(sorted(set(keys)))
        LIVE.write_text(json.dumps(live, ensure_ascii=False), encoding="utf-8")
    else:
        live = json.loads(LIVE.read_text(encoding="utf-8"))
    fetched = live.get("__fetched", "2026-10-09")

    urls = set()
    for v in live.values():
        if isinstance(v, dict) and "formationInformationGenerale" in v:
            u = first_url(v["formationInformationGenerale"].get("UrlInfoSpecifique"))
            if u:
                urls.add(u)
    if "--check-web" in sys.argv:
        print(f"Checking {len(urls)} websites …")
        WEB.write_text(json.dumps(check_web(urls), indent=0), encoding="utf-8")
    web = json.loads(WEB.read_text(encoding="utf-8")) if WEB.exists() else {}

    dup = Counter((" ".join(d["Program title (as in Études en France)"].split()).lower(), d["University / group"] or d["Institution / faculty"]) for d in rows)

    programs, changes = [], Counter()
    for d in rows:
        L = live.get(d["_key"], {"__error": "missing"})
        warn = []

        def w(level, code, msg):
            warn.append({"l": level, "c": code, "m": msg})

        ok = "__error" not in L
        g = L.get("formationInformationGenerale", {}) if ok else {}
        E = L.get("formationsInformationsEntree", []) if ok else []
        degree = d["Degree type"]
        inst_type = d["Institution type"]
        public = inst_type in ("Public university", "Public school / institute")

        if not ok:
            w("high", "unverified", f"Could not be re-checked on the EEF portal (server error {L['__error']}). Open the EEF link to confirm it still exists.")

        entries = [
            {
                "y": e["annee"],
                "lvl": e["valeurComNiveau"],
                "open": not e["ferme"],
                "months": e["dureeMois"],
                "fee": int(e["coutFormation"]) if e["coutFormation"] else None,
                "from": e["dateDebut"],
                "to": e["dateFin"],
                "extra": e["procedureParallele"],
                "polytech": e["processusCandidature"] == "Consortium Polytech",
            }
            for e in E
        ]
        open_levels = {lvl_num(e["lvl"]) for e in entries if e["open"] and lvl_num(e["lvl"])}
        if not ok:  # fall back to the workbook
            open_levels = {lvl_num(x) for x in (d["Entry level required"] or "").split("/")} if d["EEF status"] == "Open" else set()

        any_open = any(e["open"] for e in entries) if ok else d["EEF status"] == "Open"
        status, route_txt, entry_lvl, years = route_for(degree, open_levels, any_open)
        if not ok and status != "closed":
            route_txt += " (from spreadsheet, not re-verified)"

        old_elig = d["Your eligibility (4-yr BS, 3 yrs exp.)"] or ""
        if old_elig.startswith("Requires Bac+5") and status == "m2":
            changes["Bac+5 programs re-classed from 'needs a Master' to 'M2 entry possible'"] += 1
        if old_elig.startswith("Eligible") and status in ("m2", "closed", "no"):
            changes["programs marked eligible but whose M1 / Bac+4 year is closed or absent"] += 1
            if status == "m2":
                w("med", "m1closed", "The spreadsheet said you could enter, but on the live portal only the final year (M2) is open for 2027 – the M1 year is closed.")

        closed_years = [e for e in entries if not e["open"]]
        if closed_years and status != "closed":
            w("info", "partclosed", "Some entry years are closed on EEF: " + ", ".join(f"year {e['y']} ({e['lvl']})" for e in closed_years) + ".")
        if status == "closed":
            w("high", "closed", "Not accepting applications through Études en France for the 2027 intake.")
        if status == "m2":
            w("med", "m2only", "Only M2 (final year) entry is open. A 4-year BS formally gives Bac+4, but French M2 jury usually expects an M1 in the same field – apply, but keep M1 options as backup.")
        if status == "possible" and lvl_num(d["Entry level required"]) == 6:
            w("info", "ms", "Mastère Spécialisé normally requires Bac+5; Bac+4 is accepted only with 3+ years of relevant experience (and usually within a quota). Confirm with the school.")

        # ---- fees ----
        cat_fees = sorted({e["fee"] for e in entries if e["fee"]}) if ok else []
        per_year = d["Est. tuition / year non-EU (€)"]
        low = d["Lowest possible with exemption (€)"]
        basis = d["Fee basis & source"] or ""
        conf = d["Fee confidence"] or "Low"
        apprenticeship = inst_type == "Apprenticeship (CFA/ITII)"
        engineer = degree == "Diplôme d'ingénieur"

        if apprenticeship:
            w("high", "apprentice", "Apprenticeship track: free only with a French employer contract and residence permit – not realistic to apply from abroad.")
        elif public and cat_fees and max(cat_fees) < 1000:
            low = max(cat_fees)
            per_year = NON_EU_MASTER
            basis = f"Catalogue shows €{max(cat_fees):,}, which is the EU/French rate. Most public schools charge non-EU students €3,950 (2026-27), so the estimate is corrected to €3,950. Some schools waive it – ask."
            conf = "Medium"
            changes["fee estimates raised from the EU rate to the non-EU rate (€3,950)"] += 1
            w("high", "eurate", f"Fee was understated: the catalogue's €{max(cat_fees):,} is the EU rate. Budget €3,950/yr unless the school confirms non-EU students pay the same.")
        elif public and OLD_MASTER_FEE in cat_fees:
            per_year = max(per_year or 0, NON_EU_MASTER)
            w("med", "oldfee", "The school entered €3,770 – that is the old 2025-26 rate. The 2026-27 non-EU master fee is €3,950 (and may rise again for 2027-28).")
        elif public and engineer and not cat_fees:
            per_year, low = NON_EU_MASTER, EU_ENGINEER
            basis = "Not in catalogue. Public engineering schools: non-EU €3,950/yr at most schools (e.g. INSA Toulouse, ENIM 2026-27); some (e.g. Tours) charge everyone the €630 national rate."
            conf = "Low"
            w("med", "engfee", "Engineering-diploma fee for non-EU students varies by school (€630 or €3,950). Check the school's fee page.")
        if public and not apprenticeship and low in (254, 255):
            low = EU_MASTER

        if conf == "Low" and not any(x["c"] == "engfee" for x in warn):
            w("med", "lowconf", "Fee not published by the school – this is an estimate (" + basis[:140] + ").")
        if any(s in basis.lower() for s in UNOFFICIAL_SOURCES):
            w("med", "unofficialfee", "Fee figure comes from a third-party site, not the school. Confirm on the official page.")
        if inst_type == "Private / EESPIG" and per_year and per_year < 2000:
            w("med", "cheapprivate", f"Unusually low fee (€{per_year:,}) for a private school – may be a deposit or per-semester figure.")
        if per_year and per_year > 10000:
            w("info", "expensive", f"High tuition (€{per_year:,}/yr). Look for school scholarships before applying.")

        # ---- policy ----
        uni = d["University / group"] or d["Institution / faculty"] or ""
        policy, policy_src = d["University exemption policy 2026-27"], d["Policy source"]
        for name, (txt, src) in POLICY_OVERRIDES.items():
            if name.lower() in uni.lower() and public:
                policy, policy_src = txt, src
                if name == "Université Paris Cité":
                    w("high", "paris", "Paris Cité exempted non-EU master students in 2026-27, but says new non-EU students in 2027-28 may pay €3,950. Do not count on the €255 rate.")
                    low = EU_MASTER
        if policy_src and any(s in policy_src for s in UNOFFICIAL_SOURCES):
            w("info", "newsource", "Exemption policy comes from a news article, not the university. Check the university's own fee page.")

        # ---- language & content ----
        lang_map = {"Français": "French", "Anglais": "English", "Français et anglais": "French & English"}
        lang = lang_map.get(g.get("langueEnseignement"), d["Teaching language"]) if ok else d["Teaching language"]
        req = clean((g.get("preRequis") or {}).get("value")) if ok else (d["Requirements (excerpt)"] or "")
        if lang == "English" and FRENCH_REQ.search(req):
            w("med", "langconflict", "Listed as English-taught, but the requirements mention a French-language certificate/level. Ask which language the courses are in.")
        if lang in (None, "n/a", ""):
            w("med", "nolang", "Teaching language not stated in the catalogue.")
        proc_text = " ".join(clean((e.get("procedureInscription") or {}).get("value")) for e in E)
        if OLD_YEAR.search(proc_text):
            w("info", "stale", "The school's catalogue text still mentions past academic years (e.g. 2025-26) – some details may be out of date.")

        title = " ".join((g.get("libelleFormation") if ok else None or d["Program title (as in Études en France)"]).split())
        if d["Relevance"].count("★") >= 4 and OFF_TOPIC.search(title):
            w("med", "relevance", "Rated as a strong fit, but the title looks like pharma/biotech/management – check that it really matches embedded/biomedical engineering.")
        if dup[(" ".join(d["Program title (as in Études en France)"].split()).lower(), uni or d["Institution / faculty"])] > 1:
            w("info", "dup", "Listed more than once in EEF (different sites or partners). Make sure you pick the right campus.")

        website = (g.get("UrlInfoSpecifique") if ok else d["Program website"]) or ""
        wurl = first_url(website)
        ws = web.get(wurl) if wurl else None
        if not wurl:
            w("info", "noweb", "No program website in the catalogue.")
        elif ws in (404, 410, 500, 502, 503) or str(ws) == "ERR:URLError":
            w("med", "deadweb", f"Program website did not load when checked ({ws}). The page may have moved.")

        extra = any(e["extra"] for e in entries) if ok else d["Extra school application required?"] == "Yes"
        if extra:
            w("info", "extra", "You must ALSO apply on the school's own website (separate application), in addition to EEF.")
        if any(e["polytech"] for e in entries):
            w("info", "polytech", "Application goes through the Polytech network (Concours Polytech) – check its own deadline.")

        if years is None:
            years = d["Years to finish (est.)"]
        total = per_year * years if (per_year is not None and years) else None
        deadline = max((e["to"] for e in entries if e["open"] and e["to"]), default=d["Apply until (EEF)"])
        site = g.get("libelleSite") if ok else d["Campus / site"]
        address = (L.get("siteInformationContact") or {}).get("adresse") if ok else ""
        contact = (L.get("siteInformationContact") or {}).get("emailContactSite") if ok else ""

        sev = {"high": 0, "med": 1, "info": 2}
        warn.sort(key=lambda x: sev[x["l"]])
        programs.append({
            "id": d["#"],
            "rel": d["Relevance"].count("★"),
            "field": d["Field (category)"],
            "title": title,
            "en": re.sub(r"^(major|track):\s*", "", d["Approx. English title (auto-translated)"] or ""),
            "degree": degree,
            "inst": d["Institution / faculty"],
            "uni": uni,
            "site": site,
            "city": city_of(address) or city_from_site(site),
            "type": inst_type,
            "lang": lang,
            "status": status,
            "route": route_txt,
            "entryLvl": entry_lvl,
            "years": years,
            "entries": entries,
            "fee": per_year,
            "low": low,
            "total": total,
            "basis": basis,
            "conf": conf,
            "catFees": cat_fees,
            "policy": policy,
            "policySrc": policy_src,
            "extra": extra,
            "deadline": deadline,
            "web": wurl,
            "contact": contact,
            "req": req[:700],
            "eef": d["Catalogue link__link"],
            "verified": ok,
            "warn": warn,
        })

    # Fill missing towns from the same university's other listings.
    uni_city = defaultdict(Counter)
    for p in programs:
        if p["city"]:
            uni_city[p["uni"]][p["city"]] += 1
    for p in programs:
        if not p["city"] and uni_city[p["uni"]]:
            p["city"] = uni_city[p["uni"]].most_common(1)[0][0]

    meta = {
        "built": dt.date.today().isoformat(),
        "liveChecked": fetched,
        "webChecked": web.get("__checked", "2026-10-09"),
        "count": len(programs),
        "verified": sum(p["verified"] for p in programs),
        "changes": [{"what": k, "n": v} for k, v in changes.most_common()],
    }
    OUT.write_text(
        "// Generated by scripts/build.py – do not edit by hand.\n"
        "window.META = " + json.dumps(meta, ensure_ascii=False) + ";\n"
        "window.PROGRAMS = " + json.dumps(programs, ensure_ascii=False, separators=(",", ":")) + ";\n",
        encoding="utf-8",
    )
    print(json.dumps(meta, indent=1, ensure_ascii=False))
    print("status:", Counter(p["status"] for p in programs))
    print("warnings:", Counter(x["c"] for p in programs for x in p["warn"]))


if __name__ == "__main__":
    build()
