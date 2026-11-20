# =====================================================================
# INDIA AGRICULTURE WEATHER INTELLIGENCE PLATFORM - v6 (Colab cell OR GitHub Actions script)
# GitHub build: notebook logic unchanged; additions are marked [GH] (name standard, tehsil-code clean-up, IMERG diagnostics, RUN_STATES env for test runs).
# Fixed Google Sheet. Sheet tabs: 4 requested tabs + 3 history tabs. Everything else -> Excel/Drive.
# Agri Intelligence workbook is never written.
# =====================================================================
import subprocess, sys
def _pip(*p): subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", *p])
for _m, _p in {"ee": "earthengine-api", "gspread": "gspread", "openpyxl": "openpyxl", "rapidfuzz": "rapidfuzz", "bs4": "beautifulsoup4", "pypdf": "pypdf"}.items():
    try: __import__(_m)
    except ImportError: _pip(_p)
import ee, gspread, json, os, re, glob, time, pickle, unicodedata, io, random, shutil, datetime as dt, warnings
import requests, pandas as pd, numpy as np
from rapidfuzz import fuzz, process
from google.oauth2 import service_account
warnings.filterwarnings("ignore")
pd.set_option("display.width", 280); pd.set_option("display.max_columns", 80); pd.set_option("display.max_colwidth", 40)
try:
    import google.colab; IN_COLAB = True
except ImportError: IN_COLAB = False

# ============================ CONFIG (fixed identifiers - do not change) ============================
AGRI_SHEET_ID = "1XorY5NC-ZcqM3_vhjanncwTNgbL1MWVHIHZBLDWKAoA"      # NEVER written
TRACTOR_SHEET_ID = "1HxtmPL4GTMYlsKkN8s_1hNkvmZpLvrIDYBHgzt7q7Mc"    # read-only search for TEHSIL_MASTER
FIXED_WEATHER_SHEET_ID = "1TneGWHkc8mhF3EC7drV3tTPfBhxLwgavCP5PcP4apxc"
def secret(name):
    if IN_COLAB:
        try:
            from google.colab import userdata
            v = userdata.get(name)
            if v and v.strip(): return v.strip()
        except Exception: pass
    v = os.environ.get(name); return v.strip() if v and v.strip() else None
WEATHER_SHEET_ID = secret("WEATHER_SHEET_ID") or FIXED_WEATHER_SHEET_ID
WEATHER_SHEET_URL = f"https://docs.google.com/spreadsheets/d/{WEATHER_SHEET_ID}/edit"
RUN_STATES = None                       # None = all India; test: ["Tamil Nadu", "Puducherry"]
if os.environ.get("RUN_STATES", "").strip(): RUN_STATES = [x.strip() for x in os.environ["RUN_STATES"].split(",") if x.strip()]   # [GH] test runs from the Actions "Run workflow" box; empty = unchanged (all India)
RUN_TEHSIL_FORECASTS = True
BUILD_DAILY_HISTORY = (os.environ.get("BUILD_DAILY_HISTORY", "1" if IN_COLAB else "0") == "1")   # Open-Meteo archive (rate-limited); CHIRPS monthly history is always built
ARCHIVE_MAX_POINTS_PER_RUN, ARCHIVE_BATCH, OBS_MAX_KM = 60, 5, 75
RAIN_DAY_MM, FORECAST_DAYS, PAST_DAYS, BATCH, BATCH_SLEEP = 2.5, 7, 2, 50, 4
IMD = "https://api.imd.gov.in/api/v1"; OM = "https://api.open-meteo.com/v1/forecast"
CKAN_HOSTS = ["https://ckandev.indiadataportal.com", "https://indiadataportal.com"]; CKAN_RES = "42cd589e-2dc4-4f9c-a245-f201cee357e7"
UA = {"User-Agent": "Mozilla/5.0 (AgriWeatherIntelligence; research)"}
IST = dt.timezone(dt.timedelta(hours=5, minutes=30)); UTC = dt.timezone.utc
now = dt.datetime.now(IST); today = now.date(); RUN_TS = now.strftime("%Y-%m-%d %H:%M IST"); STAMP = now.strftime("%Y%m%d_%H%M")
SHEET_TABS = ["WEATHER MANAGEMENT SUMMARY", "INDIA DISTRICT INTELLIGENCE", "INDIA TEHSIL INTELLIGENCE", "APPROACHING WEATHER ALERTS", "HISTORY MOM DISTRICT", "HISTORY MOM TEHSIL", "HISTORY MOM STATE"]
LOG = []
def log(step, status, detail=""): LOG.append([dt.datetime.now(IST).strftime("%H:%M:%S"), step, status, str(detail)[:300]]); print(f"  [{status:5}] {step}: {str(detail)[:150]}")
print("Run time:", RUN_TS, "| environment:", "Colab" if IN_COLAB else "GitHub/CLI")

# ============================ KEY, STORAGE, SHEETS CLIENTS ============================
def _parse(raw):
    try:
        d = json.loads(raw); return d if isinstance(d, dict) else None
    except Exception: return None
def load_key():
    d = _parse(secret("GEE_SERVICE_ACCOUNT_JSON") or "")
    if d and d.get("type") == "service_account": return d
    for p in glob.glob("/content/*.json") + glob.glob("/content/drive/MyDrive/*.json"):
        try: d = _parse(open(p, encoding="utf-8").read())
        except Exception: d = None
        if d and d.get("type") == "service_account" and d.get("private_key"): return d
    if IN_COLAB:
        from google.colab import files
        for _, c in files.upload().items():
            d = _parse(c.decode("utf-8", errors="ignore"))
            if d and d.get("type") == "service_account": return d
    raise SystemExit("STOP: service-account JSON not found (set secret/env GEE_SERVICE_ACCOUNT_JSON).")
info = load_key(); SA_JSON = json.dumps(info); SA_EMAIL = info["client_email"]; print("Service account:", SA_EMAIL)
HISTORY_WRITE = True
if IN_COLAB:
    HISTORY_WRITE = False; PROJ = "/content/India_Weather_Intelligence_TEMP"; OUT_DIR = "/content"
    try:
        from google.colab import drive
        drive.mount("/content/drive")
        if os.path.isdir("/content/drive/MyDrive"): PROJ = "/content/drive/MyDrive/India_Weather_Intelligence"; HISTORY_WRITE = True
    except Exception as e: print("  Drive mount failed:", type(e).__name__, str(e)[:100])
else: PROJ = os.environ.get("WEATHER_DATA_DIR", "data"); OUT_DIR = "output"
DIRS = {k: f"{PROJ}/{k}" for k in ["history", "forecast_vintages", "masters", "audit", "logs", "checkpoints", "failed", "imd_public", "backups", "alerts"]}
for p in list(DIRS.values()) + [OUT_DIR]: os.makedirs(p, exist_ok=True)
print(("Persistent folder: " if HISTORY_WRITE else "WARNING: Drive not mounted - nothing historical is saved: ") + PROJ)
CLIENTS = []
if IN_COLAB:
    try:
        from google.colab import auth as _auth; _auth.authenticate_user()
        import google.auth; _cu, _ = google.auth.default(); CLIENTS.append(("Colab user account", gspread.authorize(_cu)))
    except Exception as e: log("Colab user login", "CHECK", f"{type(e).__name__}: {str(e)[:90]}")
try:
    _sa = service_account.Credentials.from_service_account_info(info, scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]); CLIENTS.append(("Service account", gspread.authorize(_sa)))
except Exception as e: log("Service-account Sheets client", "CHECK", str(e)[:100])
ee.Initialize(ee.ServiceAccountCredentials(SA_EMAIL, key_data=SA_JSON), project=secret("GEE_PROJECT") or info.get("project_id")); log("Earth Engine", "OK", "initialised")

# ============================ HELPERS ============================
def http_get(url, params=None, tries=3, timeout=40):
    last = "no attempt"
    for i in range(tries):
        try:
            r = requests.get(url, params=params, headers=UA, timeout=timeout)
            if r.status_code == 200: return r, "OK (HTTP 200)"
            last = f"HTTP {r.status_code}"
            if r.status_code == 429: time.sleep(65 + random.uniform(0, 15)); continue
            if r.status_code in (500, 502, 503, 504): time.sleep(2 ** (i + 1) + random.uniform(0, 2)); continue
            return None, last
        except Exception as e: last = f"{type(e).__name__}: {str(e)[:80]}"; time.sleep(2 ** (i + 1))
    return None, "FAILED: " + last
def f(x):
    try:
        v = float(x); return np.nan if np.isnan(v) else v
    except Exception: return np.nan
def clean(v):
    if v is None: return ""
    if isinstance(v, np.integer): return int(v)
    if isinstance(v, (np.floating, float)): return "" if np.isnan(v) else float(v)
    if isinstance(v, (bool, np.bool_)): return str(bool(v))
    return v if isinstance(v, (int, str)) else str(v)
def vals(df, drop=("geom", "sgeom", "cell")):
    d = df.drop(columns=[c for c in drop if c in df.columns]); d = d.loc[:, ~d.columns.duplicated()]
    return [list(map(str, d.columns))] + [[clean(v) for v in row] for row in d.values.tolist()]
def shift_year(d, n):
    try: return d.replace(year=d.year - n)
    except ValueError: return d.replace(year=d.year - n, day=28)
def find_rows(js):
    if isinstance(js, list): return js
    if isinstance(js, dict):
        for k in ("data", "rows", "results", "items", "features"):
            if isinstance(js.get(k), list): return js[k]
        if isinstance(js.get("data"), dict): return [js["data"]]
    return []
def pick_key(keys, pats, avoid=""):
    for p in pats:
        for k in keys:
            kl = str(k).lower()
            if re.search(p, kl) and not (avoid and re.search(avoid, kl)): return k
    return None
def kget(row, *names):
    low = {str(k).lower().replace(" ", "").replace("_", ""): v for k, v in row.items()}
    for n in names:
        v = low.get(n.lower().replace(" ", "").replace("_", ""))
        if v not in (None, ""): return v
    return None
def hav_km(la1, lo1, la2, lo2):
    p = np.pi / 180; a = np.sin((la2 - la1) * p / 2) ** 2 + np.cos(la1 * p) * np.cos(la2 * p) * np.sin((lo2 - lo1) * p / 2) ** 2
    return 12742 * np.arcsin(np.sqrt(a))
def pid_of(la, lo): return f"{float(la):.3f},{float(lo):.3f}"

# ============================ IMD (disabled until credentials verified) ============================
IMD_KEY = secret("IMD_API_KEY"); IMD_SCHEME = (secret("IMD_AUTH_SCHEME") or "").lower()
def imd_request(path, params=None, auth=True):
    h = dict(UA); p = dict(params or {})
    if auth and IMD_KEY and IMD_SCHEME:
        if IMD_SCHEME == "bearer": h["Authorization"] = "Bearer " + IMD_KEY
        elif IMD_SCHEME == "x-api-key": h["x-api-key"] = IMD_KEY
        elif IMD_SCHEME.startswith("query:"): p[IMD_SCHEME.split(":", 1)[1]] = IMD_KEY
        elif IMD_SCHEME.startswith("header:"): h[IMD_SCHEME.split(":", 1)[1]] = IMD_KEY
    try: return requests.get(f"{IMD}/{path}", params=p, headers=h, timeout=60), None
    except Exception as e: return None, f"{type(e).__name__}: {str(e)[:80]}"
r0, e0 = imd_request("districtrainfall", {"id": 164}, auth=False); IMD_AUTH, imd_diag = "UNKNOWN", ""
if r0 is None: IMD_AUTH, imd_diag = "NETWORK ERROR", e0
else:
    imd_diag = f"HTTP {r0.status_code}; body={r0.text[:100]!r}"
    if r0.status_code == 200: IMD_AUTH = "NOT REQUIRED"
    elif IMD_KEY and IMD_SCHEME:
        r1, _ = imd_request("districtrainfall", {"id": 164}); IMD_AUTH = "AUTHENTICATED" if (r1 is not None and r1.status_code == 200) else "CREDENTIALS REJECTED"
    else: IMD_AUTH = "WAITING FOR CREDENTIALS"
IMD_ACTIVE = IMD_AUTH in ("NOT REQUIRED", "AUTHENTICATED"); log("IMD authentication", "OK" if IMD_ACTIVE else "FAIL", f"{IMD_AUTH} | {imd_diag}")
aws = []
if IMD_ACTIVE:
    r, _ = imd_request("districtrainfall")
    if r is not None and r.status_code == 200 and find_rows(r.json()):
        for sid in range(1, 37):
            rr, _ = imd_request("aws_data", {"sid": sid})
            try: aws += find_rows(rr.json()) if (rr is not None and rr.status_code == 200) else []
            except Exception: pass
            time.sleep(0.3)
imd_by = {}

# ============================ NAME ENGINE ============================
def norm_txt(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower().replace("_", " ").replace("&", " and ")
    s = re.sub(r"\(.*?\)", " ", s); s = re.sub(r"\b(district|dist|taluk|taluka|tehsil|tahsil|mandal|block|ut|u t)\b", " ", s)
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9 ]", " ", s)).strip()
STATE_ALIAS = {"orissa": "odisha", "pondicherry": "puducherry", "nct of delhi": "delhi", "national capital territory of delhi": "delhi", "andaman and nicobar": "andaman and nicobar islands", "jammu kashmir": "jammu and kashmir",
    "uttaranchal": "uttarakhand", "telengana": "telangana", "tamilnadu": "tamil nadu", "chattisgarh": "chhattisgarh", "dadra and nagar haveli": "dadra and nagar haveli and daman and diu", "daman and diu": "dadra and nagar haveli and daman and diu"}
_GROUPS = [("baleshwar", "balasore"), ("visakhapatnam", "vishakhapatnam", "vizag"), ("gadchiroli", "garhchiroli"), ("bengaluru urban", "bangalore urban"), ("bengaluru rural", "bangalore rural"), ("mysuru", "mysore"), ("belagavi", "belgaum"),
    ("kalaburagi", "gulbarga"), ("vijayapura", "bijapur"), ("shivamogga", "shimoga"), ("tumakuru", "tumkur"), ("ballari", "bellary"), ("chikkamagaluru", "chikmagalur"), ("thiruvananthapuram", "trivandrum"), ("kozhikode", "calicut"),
    ("thoothukudi", "tuticorin", "thoothukkudi"), ("tiruchirappalli", "trichy", "tiruchirapalli", "tiruchchirappalli"), ("kancheepuram", "kanchipuram"), ("viluppuram", "villupuram"), ("thiruvallur", "tiruvallur"), ("kanniyakumari", "kanyakumari"),
    ("the nilgiris", "nilgiris"), ("sivagangai", "sivaganga"), ("tiruvarur", "thiruvarur"), ("tirunelveli", "tirunelveli kattabo"), ("ahilyanagar", "ahmednagar", "ahmadnagar"), ("gurugram", "gurgaon"), ("prayagraj", "allahabad"), ("ayodhya", "faizabad"),
    ("purba medinipur", "east midnapore"), ("paschim medinipur", "west midnapore"), ("puducherry", "pondicherry"), ("sri potti sriramulu nellore", "nellore", "spsr nellore"), ("mahbubnagar", "mahabubnagar"), ("rangareddy", "ranga reddy")]
ALIAS_D = {}
for _g in _GROUPS:
    for _n in _g: ALIAS_D[norm_txt(_n)] = norm_txt(_g[0])
def mk_key(code, *parts):
    return str(code).strip() if (code is not None and pd.notna(code) and str(code).strip() not in ("", "nan", "None", "0")) else "N:" + "|".join(norm_txt(p) for p in parts)
def reconcile_districts(src, canon, source_label):
    cs = {norm_txt(s): s for s in canon["State"].unique()}; cs_list = list(cs)
    canon = canon.assign(_n=canon["District"].map(norm_txt)); canon["_na"] = canon["_n"].map(lambda n: ALIAS_D.get(n, n)); by_state = {s: g for s, g in canon.groupby("State")}; audit, mp = [], {}
    for r in src.itertuples():
        os_, od = str(r.State), str(r.District); ns = STATE_ALIAS.get(norm_txt(os_), norm_txt(os_)); ms, sm_ = None, "none"
        if ns in cs: ms, sm_ = cs[ns], "exact"
        else:
            m = process.extractOne(ns, cs_list, scorer=fuzz.token_sort_ratio, score_cutoff=90)
            if m: ms, sm_ = cs[m[0]], f"fuzzy {m[1]:.0f}"
        cand = by_state[ms] if ms in by_state else canon; nd0 = norm_txt(od); nd = ALIAS_D.get(nd0, nd0)
        hit, method, score, conf, status, reason = None, "Unmatched", 0.0, "Unmatched", "UNMATCHED", "no candidate above 75"
        ex = cand[cand["_n"] == nd0]; al = cand[cand["_na"] == nd]
        if len(ex) == 1: hit, method, score, conf = ex.iloc[0], "Exact normalized name", 100.0, "Exact name match"
        elif len(ex) > 1: method, conf, status, reason = "Duplicate names", "Ambiguous match", "MANUAL REVIEW", "duplicate canonical names in parent"
        elif len(al) == 1: hit, method, score, conf = al.iloc[0], "Alias/historical name", 100.0, "Exact name match (alias)"
        elif len(al) > 1: method, conf, status, reason = "Alias duplicates", "Ambiguous match", "MANUAL REVIEW", "alias maps to several units"
        else:
            top_ = process.extract(nd, list(cand["_na"]), scorer=fuzz.token_sort_ratio, limit=2)
            if top_:
                s1 = top_[0][1]; s2 = top_[1][1] if len(top_) > 1 else 0; row = cand.iloc[top_[0][2]]; method, score = "Fuzzy within parent state", float(s1)
                if s1 >= 98 and (s1 - s2) >= 3 and ms: hit, conf = row, "High-confidence fuzzy match"
                elif s1 >= 75: conf, status, reason = ("Medium-confidence fuzzy match" if s1 >= 90 else "Low-confidence match"), "MANUAL REVIEW", f"score {s1:.0f}: no geographic corroboration, NOT merged"
                if (s1 - s2) < 3 and s1 >= 75 and hit is None: conf, status, reason = "Ambiguous match", "MANUAL REVIEW", "two close candidates"
        if hit is not None: mp[r.SKEY] = hit.DKEY; status = "AUTO-ACCEPTED"; reason = "name evidence agrees with parent state"
        audit.append({"Original source name": od, "Canonical LGD name": hit.District if hit is not None else "", "Original state name": os_, "Matched state": ms or "", "State match": sm_, "LGD state code": hit.state_code if hit is not None else "",
            "LGD district code": hit.district_code if hit is not None else "", "LGD sub-district code": "", "Matching method": method, "Matching score": round(score, 1), "Matching confidence": conf, "Verification status": status, "Source": source_label, "Timestamp": RUN_TS, "Reason for correction": reason})
    return audit, mp
def reconcile_tehsils(tu, of):
    of = of.copy(); of["_n"] = of["Tehsil"].map(norm_txt); by_code = dict(zip(of["tehsil_code"].astype(str), of.index)); groups = {k: g for k, g in of.groupby("DKEY")}; audit, codes, matched = [], [], set()
    for r in tu.itertuples():
        code = str(r.tehsil_code).strip() if (r.tehsil_code is not None and pd.notna(r.tehsil_code)) else ""; hit = None; meth, conf, status, reason = "Unmatched", "Unmatched", "UNMATCHED", "not found in official extract"
        if code and code in by_code:
            o = of.loc[by_code[code]]
            if o["DKEY"] == r.DKEY: hit, meth, conf, status, reason = o, "Exact LGD code match", "Exact code match", "AUTO-ACCEPTED", ""
            else: meth, conf, status, reason = "Code under different district", "Ambiguous match", "MANUAL REVIEW", "code exists under a different official district"
        else:
            g = groups.get(r.DKEY)
            if g is not None:
                nt = norm_txt(r.Tehsil); ex = g[g["_n"] == nt]
                if len(ex) == 1: hit, meth, conf, status, reason = ex.iloc[0], "Exact normalized name in parent district", "Exact name match", "AUTO-ACCEPTED", "code filled from official extract"
                elif len(ex) > 1: meth, conf, status, reason = "Duplicate names in parent", "Ambiguous match", "MANUAL REVIEW", "identical names within one district"
                else:
                    top_ = process.extract(nt, list(g["_n"]), scorer=fuzz.token_sort_ratio, limit=2)
                    if top_:
                        s1 = top_[0][1]; s2 = top_[1][1] if len(top_) > 1 else 0
                        if s1 >= 98 and s1 - s2 >= 3: hit, meth, conf, status, reason = g.iloc[top_[0][2]], "Fuzzy within parent district", "High-confidence fuzzy match", "AUTO-ACCEPTED", f"score {s1:.0f}"
                        elif s1 >= 75: meth, conf, status, reason = "Fuzzy within parent district", "Medium/Low-confidence match", "MANUAL REVIEW", f"score {s1:.0f}: NOT merged"
        if hit is not None: matched.add(str(hit["tehsil_code"])); codes.append(hit["tehsil_code"])
        else: codes.append(r.tehsil_code)
        if meth != "Exact LGD code match": audit.append({"Original source name": r.Tehsil, "Canonical LGD name": hit["Tehsil"] if hit is not None else "", "Original state name": r.State, "Matched state": r.State, "LGD state code": hit["state_code"] if hit is not None else "",
            "LGD district code": hit["district_code"] if hit is not None else "", "LGD sub-district code": hit["tehsil_code"] if hit is not None else "", "Matching method": meth, "Matching confidence": conf, "Verification status": status, "Source": "TEHSIL_MASTER vs official LGD extract", "Timestamp": RUN_TS, "Reason for correction": reason})
    return codes, audit, matched

# ============================ ADMIN MASTERS (with persistent caches for GitHub runs) ============================
TM_CACHE = DIRS["masters"] + "/tehsil_master_cached.csv"; OFF_CACHE = DIRS["masters"] + "/official_lgd_subdistricts.csv"
def uniq_headers(h):
    seen, out = {}, []
    for c in h:
        c = str(c).strip() or "col"; seen[c] = seen.get(c, 0) + 1; out.append(c if seen[c] == 1 else f"{c}_{seen[c]}")
    return out
def std_master(df, label):
    cols = list(df.columns); K = lambda pats, av="": pick_key(cols, pats, av)
    kT = K([r"^tehsil", r"sub.?district.*name", r"^sub.?district$", r"taluk", r"^name$"], "code|id$"); kS = K([r"^state", r"state.?name"], "code|id$"); kD = K([r"^district", r"district.?name"], "code|sub|id$")
    kSC = K([r"state.*code", r"^st.?code", r"lgd.*state"]); kDC = K([r"district.*code", r"^dt.?code", r"lgd.*district"], "sub"); kTC = K([r"(tehsil|sub.?district|sdt).*code", r"lgd.*(tehsil|sub)"])
    kLa = K([r"^lat", r"centroid.*lat", r"latitude"]); kLo = K([r"^lon", r"^lng", r"centroid.*lon", r"longitude"]); log("Master columns", "INFO", f"state={kS},dist={kD},tehsil={kT},codes={kSC}/{kDC}/{kTC},lat/lon={kLa}/{kLo}")
    if not (kS and kD and kT): return None
    out = pd.DataFrame({"State": df[kS].astype(str).str.strip(), "state_code": df[kSC] if kSC else None, "District": df[kD].astype(str).str.strip(), "district_code": df[kDC] if kDC else None, "Tehsil": df[kT].astype(str).str.strip(),
        "tehsil_code": df[kTC] if kTC else None, "lat": pd.to_numeric(df[kLa], errors="coerce") if kLa else np.nan, "lon": pd.to_numeric(df[kLo], errors="coerce") if kLo else np.nan})
    used = {kT, kS, kD, kSC, kDC, kTC, kLa, kLo}
    for c in [c for c in cols if c not in used and re.search(r"agro|sown|cropped|cultivable|intensity|elevation|\bdem\b|area|zone|orchard|cropland", c.lower())][:14]: out["TM: " + c.replace("TM: ", "")] = df[c].values
    out = out[(out["Tehsil"].str.len() > 0) & (~out["Tehsil"].str.lower().isin(["nan", "none"]))]; out["Source"] = label; return out.reset_index(drop=True)
def master_from_values(v, label):
    if not v or len(v) < 3: return None
    hi = next((i for i, r in enumerate(v[:15]) if sum(1 for c in r if str(c).strip()) >= 4 and any(re.search(r"tehsil|sub.?district|taluk", str(c).lower()) for c in r)), 0)
    hdr = uniq_headers(v[hi]); rows = [(r + [""] * len(hdr))[:len(hdr)] for r in v[hi + 1:] if any(str(c).strip() for c in r)]
    return std_master(pd.DataFrame(rows, columns=hdr), label)
def find_tehsil_master():
    tried = []
    for p in [DIRS["masters"] + "/TEHSIL_MASTER.csv", "/content/TEHSIL_MASTER.csv", "/content/drive/MyDrive/TEHSIL_MASTER.csv"]:
        if os.path.exists(p):
            try:
                t = std_master(pd.read_csv(p, dtype=str, on_bad_lines="skip"), f"TEHSIL_MASTER csv ({p})")
                if t is not None: return t, tried + [f"csv {p}: {len(t)} rows"]
            except Exception as e: tried.append(f"csv {p}: {type(e).__name__}")
    seen = set()
    def scan(cl, g, sid, lb):
        if (cl, sid) in seen or not sid: return None
        seen.add((cl, sid))
        try: sh = g.open_by_key(sid); ws = sh.worksheets()
        except Exception as e: tried.append(f"{cl}/{lb}: cannot open ({type(e).__name__})"); return None
        tried.append(f"{cl}/{lb}: '{sh.title}' {len(ws)} tabs")
        for w in ws:
            if w.title.strip().upper().replace(" ", "_") == "TEHSIL_MASTER":
                try: t = master_from_values(w.get_all_values(), f"TEHSIL_MASTER tab in '{sh.title}' (read-only)")
                except Exception as e: tried.append(f"read failed: {type(e).__name__}"); t = None
                if t is not None: return t
        return None
    for cl, g in CLIENTS:
        for lb, sid in [("secret", secret("TEHSIL_MASTER_SHEET_ID")), ("Weather sheet", WEATHER_SHEET_ID), ("Tractor sheet", TRACTOR_SHEET_ID), ("Agri sheet (read-only)", AGRI_SHEET_ID)]:
            t = scan(cl, g, sid, lb)
            if t is not None: return t, tried
        try:
            files_ = g.list_spreadsheet_files(); pri = [x for x in files_ if re.search(r"tehsil|master|agri|weather|intelligence|lgd|crop", x.get("name", ""), re.I)]
            for x in (pri + [x for x in files_ if x not in pri])[:70]:
                t = scan(cl, g, x["id"], "Drive:" + x.get("name", "")[:30])
                if t is not None: return t, tried
        except Exception as e: tried.append(f"{cl}/Drive listing: {type(e).__name__}")
    if IN_COLAB:
        try:
            m = re.search(r"/d/([a-zA-Z0-9-_]+)", input("TEHSIL_MASTER not found. Paste the link of the workbook containing the TEHSIL_MASTER tab (Enter to skip): ").strip())
            if m:
                for cl, g in CLIENTS:
                    t = scan(cl, g, m.group(1), "pasted link")
                    if t is not None: return t, tried
        except Exception: pass
    if os.path.exists(TM_CACHE):      # persistent cached copy from an earlier successful run
        try:
            c = pd.read_csv(TM_CACHE, dtype=str); c["lat"] = pd.to_numeric(c["lat"], errors="coerce"); c["lon"] = pd.to_numeric(c["lon"], errors="coerce"); c["Source"] = "TEHSIL_MASTER (cached copy: " + str(c["Source"].iloc[0])[:60] + ")"
            return c, tried + [f"cached copy {TM_CACHE}: {len(c)} rows"]
        except Exception as e: tried.append(f"cache read: {type(e).__name__}")
    return None, tried
def fetch_official():
    for host in CKAN_HOSTS:
        df, recs, off, total, ok = None, [], 0, None, False
        for _ in range(40):
            r, st = http_get(f"{host}/api/3/action/datastore_search", {"resource_id": CKAN_RES, "limit": 5000, "offset": off}, tries=3, timeout=90)
            if r is None: log("CKAN LGD sub-districts", "FAIL", f"{host}: {st}"); break
            try: res = r.json().get("result", {})
            except Exception: break
            batch = res.get("records", []); total = res.get("total", total)
            if not batch: break
            recs += batch; off += len(batch)
            if total is not None and off >= int(total): ok = True; break
        if ok and recs: df = pd.DataFrame(recs)
        elif recs: log("CKAN LGD sub-districts", "FAIL", f"partial {len(recs)} of {total}: rejected")
        if df is None:
            r, st = http_get(f"{host}/api/3/action/resource_show", {"id": CKAN_RES}, tries=2, timeout=60)
            try:
                if r is not None:
                    rr, _ = http_get(r.json()["result"]["url"], tries=2, timeout=120)
                    if rr is not None: df = pd.read_csv(io.StringIO(rr.text), dtype=str)
            except Exception as e: log("CKAN resource_show", "FAIL", type(e).__name__)
        if df is not None:
            cl_ = {c.lower(): c for c in df.columns}; need = ["sub_district_name", "sub_district_code", "district_code", "district_name", "state_name", "state_code"]
            if not all(n in cl_ for n in need): log("CKAN LGD sub-districts", "FAIL", f"unexpected columns {list(df.columns)[:10]}"); continue
            o = pd.DataFrame({"State": df[cl_["state_name"]].astype(str).str.strip(), "state_code": df[cl_["state_code"]].astype(str).str.strip(), "District": df[cl_["district_name"]].astype(str).str.strip(),
                              "district_code": df[cl_["district_code"]].astype(str).str.strip(), "Tehsil": df[cl_["sub_district_name"]].astype(str).str.strip(), "tehsil_code": df[cl_["sub_district_code"]].astype(str).str.strip()})
            o = o[(o["Tehsil"] != "") & (o["tehsil_code"] != "")].drop_duplicates("tehsil_code").reset_index(drop=True); o["DKEY"] = o["district_code"]
            log("CKAN LGD sub-districts", "OK", f"{len(o)} sub-districts, {o['district_code'].nunique()} districts, {o['state_code'].nunique()} states from {host}")
            try: o.to_csv(OFF_CACHE, index=False)
            except Exception: pass
            return o, "live"
    if os.path.exists(OFF_CACHE):
        try: o = pd.read_csv(OFF_CACHE, dtype=str); log("Official LGD extract", "CHECK", f"CKAN unreachable; using cached copy {len(o)} rows"); return o, "cached"
        except Exception: pass
    return None, "none"
# [GH] ---------- NAME STANDARD (shared rule for Agri / Weather / Crop / Groundwater) ----------
# Rows whose district matched an official LGD district get the OFFICIAL LGD State + District names; rows whose tehsil code
# matched an official sub-district (same district) get the OFFICIAL LGD tehsil name. The original TEHSIL_MASTER spelling is kept
# in "Source ... name" columns. Unmatched rows keep their own names (state name still mapped to the official spelling).
# Forecast / history / scoring logic is untouched. Set env STANDARDIZE_NAMES=0 to reproduce the notebook output exactly.
STANDARDIZE_NAMES = os.environ.get("STANDARDIZE_NAMES", "1") == "1"
def standardize_names(tm, od):
    tm = tm.copy()
    for a, b in (("State", "Source state name"), ("District", "Source district name"), ("Tehsil", "Source tehsil name")):
        if b not in tm.columns: tm[b] = tm[a]
    lk = od.drop_duplicates("DKEY").set_index("DKEY"); hit = tm["DKEY"].isin(lk.index)
    tm.loc[hit, "State"] = tm.loc[hit, "DKEY"].map(lk["State"]); tm.loc[hit, "District"] = tm.loc[hit, "DKEY"].map(lk["District"])
    st = {STATE_ALIAS.get(norm_txt(x), norm_txt(x)): x for x in od["State"].unique()}
    tm.loc[~hit, "State"] = [st.get(STATE_ALIAS.get(norm_txt(x), norm_txt(x)), x) for x in tm.loc[~hit, "State"]]
    if "official" in globals() and official is not None:
        ot = official.drop_duplicates("tehsil_code").set_index("tehsil_code")
        code = tm["tehsil_code"].astype(str).str.strip()
        ok = code.isin(ot.index) & (code.map(ot["DKEY"]) == tm["DKEY"].astype(str))
        tm.loc[ok, "Tehsil"] = code[ok].map(ot["Tehsil"])
    n_ch = int((tm["State"] != tm["Source state name"]).sum() + (tm["District"] != tm["Source district name"]).sum() + (tm["Tehsil"] != tm["Source tehsil name"]).sum())
    log("Name standard (official LGD)", "OK", f"{n_ch} name fields standardised; {int((~hit).sum())} rows in districts not matched to LGD keep source names")
    return tm

# [GH] ---------- TEHSIL CODE CLEAN-UP (duplicates from pre-split districts) ----------
# TEHSIL_MASTER has some rows with a zero-padded code ("05796") under an OLD parent district, while the same tehsil also
# exists with the correct code under its current district. Original logic compares codes as exact text, so both rows survive.
# With FIX_TEHSIL_CODES on: (1) a code that differs from an official LGD code only by leading zeros is unpadded;
# (2) a row whose code is an official code under a DIFFERENT official district is dropped if a correct-district copy exists,
#     otherwise it is moved to the official district (it keeps its coordinates). Every action goes to LGD NAME MATCHING AUDIT.
# Forecast / history / scoring logic is untouched. Set env FIX_TEHSIL_CODES=0 to keep the notebook behaviour.
FIX_TEHSIL_CODES = os.environ.get("FIX_TEHSIL_CODES", "1") == "1"
def fix_code_padding(tm, official):
    tm = tm.copy(); off = set(official["tehsil_code"].astype(str).str.strip()); n = 0; new = []
    for c in tm["tehsil_code"]:
        cs = "" if (c is None or (isinstance(c, float) and np.isnan(c))) else str(c).strip()
        if cs and cs not in off and cs.isdigit() and cs.lstrip("0") in off: new.append(cs.lstrip("0")); n += 1
        else: new.append(c)
    tm["tehsil_code"] = new; log("Tehsil code padding", "OK", f"{n} zero-padded TEHSIL_MASTER codes aligned to official LGD codes"); return tm
def dedupe_old_districts(tm, official, matched, audit):
    tm = tm.reset_index(drop=True); ot = official.drop_duplicates("tehsil_code").set_index("tehsil_code"); code = tm["tehsil_code"].astype(str).str.strip()
    known = code.isin(ot.index); right = known & (code.map(ot["DKEY"]) == tm["DKEY"].astype(str)); good = set(code[right]); drop, moved, seen = [], 0, set()
    for i in tm.index[known & ~right]:
        c = code[i]; o = ot.loc[c]; rec = {"Original source name": tm.at[i, "Tehsil"], "Canonical LGD name": o["Tehsil"], "Original state name": tm.at[i, "State"], "Matched state": o["State"], "LGD state code": o["state_code"],
            "LGD district code": o["district_code"], "LGD sub-district code": c, "Source": "TEHSIL_MASTER vs official LGD extract", "Timestamp": RUN_TS}
        if c in good or c in seen:
            drop.append(i); rec.update({"Matching method": "Duplicate under old parent district", "Matching confidence": "Exact code match", "Verification status": "AUTO-ACCEPTED",
                "Reason for correction": f"row under '{tm.at[i, 'District']}' dropped: same LGD code already present under official district '{o['District']}'"})
        else:
            tm.at[i, "DKEY"] = o["DKEY"]; seen.add(c); matched.add(c); moved += 1
            rec.update({"Matching method": "Moved to official parent district", "Matching confidence": "Exact code match", "Verification status": "AUTO-ACCEPTED",
                "Reason for correction": f"was under '{tm.at[i, 'District']}'; LGD code belongs to official district '{o['District']}'"})
        for j in range(len(audit) - 1, -1, -1):      # replace the MANUAL REVIEW entry reconcile_tehsils wrote for this row
            a = audit[j]
            if a.get("Matching method") == "Code under different district" and a.get("Original source name") == tm.at[i, "Tehsil"] and a.get("Original state name") == tm.at[i, "State"]: del audit[j]; break
        audit.append(rec)
    tm = tm.drop(index=drop).reset_index(drop=True); log("Old-district duplicates", "OK", f"{len(drop)} duplicate tehsil rows removed, {moved} rows moved to their official district"); return tm, matched

print("\nLoading administrative masters...")
official, off_mode = fetch_official(); tm_user, tried = find_tehsil_master()
for t_ in tried: log("Master search", "INFO", t_)
log("TEHSIL_MASTER", "OK" if tm_user is not None else "FAIL", f"{len(tm_user)} tehsils: {tm_user['Source'].iloc[0]}" if tm_user is not None else "not found (share the workbook with your login/" + SA_EMAIL + ", set TEHSIL_MASTER_SHEET_ID, or place TEHSIL_MASTER.csv in " + DIRS["masters"] + ")")
if tm_user is not None and HISTORY_WRITE and "cached copy" not in str(tm_user["Source"].iloc[0]):
    try: tm_user.to_csv(TM_CACHE, index=False)
    except Exception: pass
tehsil_audit, dist_audit = [], []; OFFICIAL_NOTE = f"official LGD sub-district extract ({off_mode}; India Data Portal CKAN; LGD data dated 2023-03-01 - newer changes NOT included)"
if official is not None:
    od = official.drop_duplicates("district_code")[["State", "state_code", "District", "district_code", "DKEY"]].reset_index(drop=True)
    if tm_user is not None:
        src = tm_user.groupby(["State", "District"], dropna=False).agg(lat=("lat", "mean"), lon=("lon", "mean")).reset_index(); src["SKEY"] = src["State"] + "|" + src["District"]
        dist_audit, dmap_ = reconcile_districts(src[["SKEY", "State", "District", "lat", "lon"]], od, "TEHSIL_MASTER districts vs official LGD")
        tm_user["DKEY"] = [dmap_.get(s + "|" + d, "N:" + norm_txt(s) + "|" + norm_txt(d)) for s, d in zip(tm_user["State"], tm_user["District"])]
        if FIX_TEHSIL_CODES: tm_user = fix_code_padding(tm_user, official)        # [GH] "05796" -> "5796" when only the leading zero differs
        codes, tehsil_audit, matched = reconcile_tehsils(tm_user, official); tm_user["tehsil_code"] = codes
        if FIX_TEHSIL_CODES: tm_user, matched = dedupe_old_districts(tm_user, official, matched, tehsil_audit)   # [GH] drop old-district duplicate rows
        tm = tm_user.copy(); lk = od.set_index("DKEY")
        tm["district_code"] = [lk.loc[k, "district_code"] if k in lk.index else None for k in tm["DKEY"]]; tm["state_code"] = [lk.loc[k, "state_code"] if k in lk.index else None for k in tm["DKEY"]]
        extra = official[~official["tehsil_code"].isin(matched)].copy(); extra["lat"] = np.nan; extra["lon"] = np.nan; extra["Source"] = "official LGD extract only (not in TEHSIL_MASTER; no coordinates)"
        tm = pd.concat([tm, extra], ignore_index=True); master_src = f"TEHSIL_MASTER (coordinates) reconciled to {OFFICIAL_NOTE}; {len(extra)} official tehsils absent from TEHSIL_MASTER appended"
    else:
        tm = official.copy(); tm["lat"] = np.nan; tm["lon"] = np.nan; tm["Source"] = OFFICIAL_NOTE + "; no coordinates"; master_src = OFFICIAL_NOTE + " (NO coordinates: tehsils inherit district forecasts)"
    if STANDARDIZE_NAMES: tm = standardize_names(tm, od)    # [GH] one name standard (official LGD) for State/District/Tehsil
    dm = od.copy()
    for k in set(tm["DKEY"]) - set(dm["DKEY"]):
        g = tm[tm["DKEY"] == k].iloc[0]; dm.loc[len(dm)] = {"State": g["State"], "state_code": g["state_code"], "District": g["District"], "district_code": g["district_code"], "DKEY": k}
    lgd_ok = True
elif tm_user is not None:
    tm = tm_user.copy(); tm["DKEY"] = [mk_key(c, s, d) for c, s, d in zip(tm["district_code"], tm["State"], tm["District"])]
    dm = tm.groupby(["DKEY"], dropna=False).agg(State=("State", "first"), state_code=("state_code", "first"), District=("District", "first"), district_code=("district_code", "first")).reset_index(); master_src = "TEHSIL_MASTER only (NOT validated against official LGD codes)"; lgd_ok = True
else: tm, dm, lgd_ok, master_src = None, None, False, "none"
def load_gaul():
    gaul = ee.FeatureCollection("FAO/GAUL/2015/level2").filter(ee.Filter.eq("ADM0_NAME", "India"))
    def _c(ft): c = ft.geometry().centroid(1000).coordinates(); return ft.set({"lon": c.get(0), "lat": c.get(1)})
    gi = gaul.map(_c).select(["ADM1_NAME", "ADM2_NAME", "ADM2_CODE", "lon", "lat"], None, False).getInfo()
    return pd.DataFrame([x["properties"] for x in gi["features"]]).rename(columns={"ADM1_NAME": "State", "ADM2_NAME": "District", "ADM2_CODE": "gaul_code"})
try: gaul_df = load_gaul(); log("GAUL 2015 (temporary coordinates only)", "OK", f"{len(gaul_df)} districts")
except Exception as e: gaul_df = pd.DataFrame(); log("GAUL 2015", "FAIL", str(e)[:100])
if not lgd_ok:
    log("LGD", "FAIL", "No official/TEHSIL_MASTER source: GAUL 2015 districts used as TEMPORARY labelled fallback (zero tehsils)")
    dm = gaul_df.assign(state_code=None, district_code=gaul_df["gaul_code"], DKEY=gaul_df["gaul_code"].astype(str)); master_src = "FAO GAUL 2015 (NOT LGD; temporary)"
dm["gaul_code"] = None; dm["gaul_lat"] = np.nan; dm["gaul_lon"] = np.nan
if lgd_ok and len(gaul_df):
    gs = gaul_df.rename(columns={"gaul_code": "SKEY"})[["SKEY", "State", "District", "lat", "lon"]].copy(); gs["SKEY"] = gs["SKEY"].astype(str)
    ga, gmap = reconcile_districts(gs, dm[["State", "state_code", "District", "district_code", "DKEY"]], "FAO GAUL 2015 vs LGD districts"); dist_audit += ga; gl = gaul_df.assign(k=gaul_df["gaul_code"].astype(str)).set_index("k")
    for sk, dk in gmap.items():
        i = dm.index[dm["DKEY"] == dk]
        if len(i): dm.loc[i[0], ["gaul_code", "gaul_lat", "gaul_lon"]] = [sk, gl.loc[sk, "lat"], gl.loc[sk, "lon"]]
    log("GAUL to LGD districts", "OK", f"{len(gmap)}/{len(gs)} auto-matched")
if "lat" not in dm.columns: dm["lat"] = np.nan; dm["lon"] = np.nan
if tm is not None and tm["lat"].notna().any():
    for i, r in dm.iterrows():
        g = tm[(tm["DKEY"] == r["DKEY"]) & tm["lat"].notna()]
        if len(g): j = ((g["lat"] - g["lat"].mean()) ** 2 + (g["lon"] - g["lon"].mean()) ** 2).idxmin(); dm.loc[i, ["lat", "lon"]] = [g.loc[j, "lat"], g.loc[j, "lon"]]
miss = dm["lat"].isna() & dm["gaul_lat"].notna(); dm.loc[miss, "lat"] = dm.loc[miss, "gaul_lat"]; dm.loc[miss, "lon"] = dm.loc[miss, "gaul_lon"]
dm["arch_lat"] = np.where(dm["gaul_lat"].notna(), dm["gaul_lat"], dm["lat"]); dm["arch_lon"] = np.where(dm["gaul_lon"].notna(), dm["gaul_lon"], dm["lon"]); dm["Source"] = master_src
if tm is not None: tm["TKEY"] = [mk_key(c, s, d, t) for c, s, d, t in zip(tm["tehsil_code"], tm["State"], tm["District"], tm["Tehsil"])]; tm = tm.drop_duplicates("TKEY").reset_index(drop=True)
if RUN_STATES:
    keep = {norm_txt(s) for s in RUN_STATES}; dm = dm[dm["State"].map(norm_txt).isin(keep)]
    if tm is not None: tm = tm[tm["State"].map(norm_txt).isin(keep)]
dm = dm.reset_index(drop=True); tm = tm.reset_index(drop=True) if tm is not None else None; nT = 0 if tm is None else len(tm); sm = dm.groupby(["State", "state_code"], dropna=False).size().reset_index(name="Districts")
print(f"Administrative master: {len(sm)} states/UTs | {len(dm)} districts | {nT} tehsils | {master_src}")
audit_rows = dist_audit + tehsil_audit

# ============================ LIVE OBSERVATIONS (METAR, IMERG, IMD bulletins) ============================
PUB_MASTER = DIRS["imd_public"] + "/imd_public_records.csv"
def retain_prior_bulletins():
    rows = []
    for p in glob.glob(DIRS["backups"] + "/*.xlsx") + glob.glob(OUT_DIR + "/*.xlsx"):
        try:
            x = pd.read_excel(p, sheet_name="IMD PUBLIC BULLETINS", header=None, dtype=str); df = x.iloc[1:].copy(); df.columns = [str(c) for c in x.iloc[0].tolist()]
            if "URL" in df.columns: df["Retained from"] = os.path.basename(p); rows.append(df)
        except Exception: continue
    old = pd.read_csv(PUB_MASTER, dtype=str) if os.path.exists(PUB_MASTER) else pd.DataFrame(); allr = pd.concat([old] + rows, ignore_index=True) if (len(old) or rows) else pd.DataFrame()
    if len(allr) and "URL" in allr.columns:
        allr = allr.drop_duplicates(subset=[c for c in ["URL", "Retrieved"] if c in allr.columns], keep="first")
        if HISTORY_WRITE: allr.to_csv(PUB_MASTER, index=False)
    return allr
try: PUB_PRIOR = retain_prior_bulletins(); log("IMD bulletin records retained", "OK" if len(PUB_PRIOR) else "CHECK", f"{len(PUB_PRIOR)} prior records")
except Exception as e: PUB_PRIOR = pd.DataFrame(); log("IMD bulletin retention", "ERROR", str(e)[:100])
def imd_public_bulletins():
    from bs4 import BeautifulSoup
    import pypdf, urllib.robotparser as rpm, urllib.parse as up
    base = "https://mausam.imd.gov.in/"; rob = rpm.RobotFileParser(); rob.set_url(base + "robots.txt")
    try: rob.read()
    except Exception: pass
    ok = lambda u: rob.can_fetch(UA["User-Agent"], u); rows = []; r, st = http_get(base, tries=2, timeout=40)
    if r is None or not ok(base): return pd.DataFrame([{"status": "IMD homepage not fetched: " + st, "Retrieved": RUN_TS}])
    links = []
    for a in BeautifulSoup(r.text, "html.parser").find_all("a", href=True):
        t = a.get_text(" ", strip=True); h = up.urljoin(base, a["href"])
        if re.search(r"press|release|bulletin|summary|warning|forecast|cyclone|nowcast", t + h, re.I) and "pdf" in h.lower() and h not in [l[1] for l in links]: links.append((t, h))
    for t, h in links[:8]:
        rec = {"Link text": t[:80], "URL": h, "Retrieved": RUN_TS, "Class": "OFFICIAL TEXT BULLETIN (unstructured; not station data)"}
        if not ok(h): rec["status"] = "robots.txt disallows"; rows.append(rec); continue
        rr, st = http_get(h, tries=2, timeout=120)
        if rr is None or len(rr.content) > 80e6: rec["status"] = "not downloaded: " + (st if rr is None else "larger than 80 MB"); rows.append(rec); continue
        try:
            txt = " ".join((p.extract_text() or "") for p in pypdf.PdfReader(io.BytesIO(rr.content)).pages[:3]); rec["status"] = "OK" if txt.strip() else "OK (no extractable text)"; m = re.search(r"(\d{1,2})[./-](\d{1,2})[./-](20\d{2})", txt)
            rec["Date in text"] = m.group(0) if m else "not found (freshness UNKNOWN)"; rec["Key lines"] = " | ".join([l.strip() for l in re.split(r"[\n\.]", txt) if re.search(r"heavy rain|thunder|heat ?wave|cold ?wave|cyclone|depression|monsoon|fog|lightning", l, re.I)][:5])[:600]
        except Exception as e: rec["status"] = f"PDF parse failed: {type(e).__name__}"
        rows.append(rec)
    return pd.DataFrame(rows) if rows else pd.DataFrame([{"status": "no bulletin PDF links found (structure unverified)", "Retrieved": RUN_TS}])
try:
    PUBLIC = imd_public_bulletins(); log("IMD public bulletins", "OK" if "URL" in PUBLIC.columns else "CHECK", f"{len(PUBLIC)} rows")
    if HISTORY_WRITE and "URL" in PUBLIC.columns: pd.concat([pd.read_csv(PUB_MASTER, dtype=str) if os.path.exists(PUB_MASTER) else pd.DataFrame(), PUBLIC.astype(str)], ignore_index=True).drop_duplicates(subset=["URL", "Retrieved"], keep="first").to_csv(PUB_MASTER, index=False)
except Exception as e: PUBLIC = pd.DataFrame([{"status": "error: " + str(e)[:100], "Retrieved": RUN_TS}]); log("IMD public bulletins", "ERROR", str(e)[:100])
PUBLIC_ALL = pd.concat([PUB_PRIOR, PUBLIC.astype(str)], ignore_index=True).drop_duplicates(subset=[c for c in ["URL", "Retrieved"] if c in set(PUB_PRIOR.columns) | set(PUBLIC.columns)], keep="first") if (len(PUB_PRIOR) or len(PUBLIC)) else pd.DataFrame()
def fetch_metar():
    try:
        r = requests.get("https://aviationweather.gov/api/data/metar", params={"bbox": "6,68,38,98", "format": "json", "hours": 4}, headers=UA, timeout=60)
        if r.status_code == 204: return pd.DataFrame(), "EMPTY (HTTP 204)"
        if r.status_code != 200: return pd.DataFrame(), f"HTTP {r.status_code}"
        rows = []
        for x in r.json():
            ic = str(x.get("icaoId", ""))
            if ic[:2] not in ("VA", "VE", "VI", "VO"): continue
            ot = x.get("obsTime"); t = pd.to_datetime(ot, unit="s", utc=True) if isinstance(ot, (int, float)) else pd.to_datetime(x.get("reportTime"), utc=True, errors="coerce")
            rows.append({"id": ic, "name": x.get("name", ""), "lat": f(x.get("lat")), "lon": f(x.get("lon")), "time": t, "temp": f(x.get("temp")), "wind": round(f(x.get("wspd")) * 1.852, 1), "vis": f(str(x.get("visib", "")).replace("+", "")), "wx": x.get("wxString") or "", "src": "NOAA aviationweather METAR"})
        df = pd.DataFrame(rows)
        if len(df): df = df.sort_values("time").groupby("id").tail(1).reset_index(drop=True); df["age_h"] = ((pd.Timestamp.now(tz="UTC") - df["time"]).dt.total_seconds() / 3600).round(1)
        return df, f"OK {len(df)} Indian airport stations"
    except Exception as e: return pd.DataFrame(), f"{type(e).__name__}: {str(e)[:80]}"
METAR, metar_note = fetch_metar(); log("NOAA METAR (live airport observations)", "OK" if len(METAR) else "FAIL", metar_note)
OBS_ST = METAR.copy() if len(METAR) else pd.DataFrame(columns=["id", "lat", "lon", "time", "temp", "wind", "vis", "wx", "src", "age_h"])
IMERG_NOTE, IMERG_FRESH, IMG_IMERG = "not attempted", False, None
def imerg_setup():
    global IMERG_NOTE, IMERG_FRESH
    notes = []      # [GH] keep the reason for EVERY dataset tried (V07 error was being overwritten by the V06 note)
    for cid, band in (("NASA/GPM_L3/IMERG_V07", "precipitation"), ("NASA/GPM_L3/IMERG_V06", "precipitationCal")):
        try:
            # [GH] look only at the last 10 days: same "latest image" result, but avoids scanning the full 1998-now collection (can time out)
            col = ee.ImageCollection(cid).filterDate(dt.datetime.now(UTC) - dt.timedelta(days=10), dt.datetime.now(UTC) + dt.timedelta(days=1)); lt = col.aggregate_max("system:time_start").getInfo()
            if lt is None: notes.append(f"{cid}: no images in the last 10 days STALE, not used"); continue
            latest = dt.datetime.fromtimestamp(lt / 1000, UTC); age = (dt.datetime.now(UTC) - latest).total_seconds() / 3600
            if age > 48: notes.append(f"{cid}: latest {latest:%Y-%m-%d} UTC ({age/24:.0f} days behind) STALE, not used"); continue
            IMERG_FRESH = True; IMERG_NOTE = " | ".join(notes + [f"{cid}: latest {latest:%Y-%m-%d %H:%M} UTC ({age:.1f} h old)"])
            return col.filterDate(latest - dt.timedelta(hours=24), latest + dt.timedelta(minutes=1)).select(band).sum().multiply(0.5).rename("imerg24")
        except Exception as e: notes.append(f"{cid}: {type(e).__name__}: {str(e)[:120]}")
    IMERG_NOTE = " | ".join(notes) or IMERG_NOTE
    return None
try: IMG_IMERG = imerg_setup()
except Exception as e: IMERG_NOTE = f"error {str(e)[:70]}"
log("IMERG satellite rainfall", "OK" if IMERG_FRESH else "CHECK", IMERG_NOTE)

# ============================ FORECASTS (Open-Meteo; vintages never overwritten) ============================
DAILY = "precipitation_sum,precipitation_probability_max,temperature_2m_max,temperature_2m_min,wind_speed_10m_max"; CURRENT = "temperature_2m,relative_humidity_2m,precipitation,cloud_cover,wind_speed_10m,wind_direction_10m"
FC_CACHE = DIRS["checkpoints"] + "/fc_cache_v2.pkl"
def get_forecasts(cells):
    cache = {}
    if os.path.exists(FC_CACHE):
        try:
            c = pickle.load(open(FC_CACHE, "rb"))
            if c.get("v") == 2 and c["date"] == str(today) and time.time() - c["ts"] < 6 * 3600: cache = c["data"]
        except Exception: pass
    def run(todo, bs):
        for i in range(0, len(todo), bs):
            sub = todo[i:i + bs]; p = {"latitude": ",".join(f"{a:.2f}" for a, _ in sub), "longitude": ",".join(f"{b:.2f}" for _, b in sub), "daily": DAILY, "current": CURRENT, "timezone": "Asia/Kolkata", "forecast_days": FORECAST_DAYS, "past_days": PAST_DAYS}
            r, st = http_get(OM, p, tries=4, timeout=90)
            if r is None and "400" in st: p.pop("current"); r, st = http_get(OM, p, tries=4, timeout=90)
            if r is None: log("Forecast batch", "FAIL", f"{i//bs+1}: {st}"); time.sleep(BATCH_SLEEP + random.uniform(0, 3)); continue
            js = r.json(); js = js if isinstance(js, list) else [js]
            for pt, it in zip(sub, js):
                if isinstance(it, dict) and it.get("daily"): cache[pt] = {"daily": it["daily"], "current": it.get("current"), "grid_lat": it.get("latitude"), "grid_lon": it.get("longitude"), "elev": it.get("elevation")}
            if (i // bs) % 10 == 0: print(f"    forecasts {min(i+bs, len(todo))}/{len(todo)}"); pickle.dump({"v": 2, "date": str(today), "ts": time.time(), "data": cache}, open(FC_CACHE, "wb"))
            time.sleep(BATCH_SLEEP + random.uniform(0, 2))
    todo = [c for c in cells if c not in cache]; log("Forecast cache", "INFO", f"{len(cache)} cached, {len(todo)} to fetch"); run(todo, BATCH)
    again = [c for c in cells if c not in cache]
    if again: log("Forecast retry", "INFO", f"{len(again)} failed cells retried"); time.sleep(30 + random.uniform(0, 10)); run(again, 10)
    pickle.dump({"v": 2, "date": str(today), "ts": time.time(), "data": cache}, open(FC_CACHE, "wb")); return cache
def cell_of(la, lo): return (round(float(la), 1), round(float(lo), 1)) if pd.notna(la) and pd.notna(lo) else None
dm["cell"] = [cell_of(a, b) for a, b in zip(dm["lat"], dm["lon"])]
if tm is not None: tm["cell"] = [cell_of(a, b) for a, b in zip(tm["lat"], tm["lon"])]
cells = set(c for c in dm["cell"] if c)
if tm is not None and RUN_TEHSIL_FORECASTS: cells |= set(c for c in tm["cell"] if c)
print(f"\nForecast cells needed: {len(cells)}"); FC = get_forecasts(sorted(cells)); failed_cells = sorted(c for c in cells if c not in FC)
log("Forecasts", "OK" if FC and not failed_cells else "CHECK" if FC else "FAIL", f"{len(FC)}/{len(cells)} cells; failed {len(failed_cells)}")
if failed_cells and HISTORY_WRITE: pd.DataFrame(failed_cells, columns=["lat", "lon"]).to_csv(f"{DIRS['failed']}/failed_forecast_cells_{STAMP}.csv", index=False)
cell_users = {}
for df_ in (dm, tm):
    if df_ is not None:
        for c in df_["cell"]:
            if c: cell_users[c] = cell_users.get(c, 0) + 1
def split_daily(d):
    t = d["time"]; i0 = next((i for i, x in enumerate(t) if x >= str(today)), 0)
    return {k: (v[i0:] if isinstance(v, list) else v) for k, v in d.items()}, {k: (v[:i0] if isinstance(v, list) else v) for k, v in d.items()}
VINTAGE_FILE = ""
if HISTORY_WRITE and FC:
    try:
        rows = []
        for (la, lo), v in FC.items():
            d = v["daily"]
            for k, t in enumerate(d["time"]): rows.append((RUN_TS, "Open-Meteo forecast API (best-match mix)", la, lo, v.get("grid_lat"), v.get("grid_lon"), v.get("elev"), t, t < str(today), d["precipitation_sum"][k], (d.get("precipitation_probability_max") or [None] * len(d["time"]))[k], d["temperature_2m_max"][k], d["temperature_2m_min"][k], d["wind_speed_10m_max"][k]))
        VINTAGE_FILE = f"{DIRS['forecast_vintages']}/forecast_vintage_{STAMP}.csv.gz"
        while os.path.exists(VINTAGE_FILE): VINTAGE_FILE = VINTAGE_FILE.replace(".csv.gz", "_b.csv.gz")
        pd.DataFrame(rows, columns=["retrieved", "source", "cell_lat", "cell_lon", "grid_lat", "grid_lon", "elevation_m", "date", "is_past_model_analysis", "precip_mm", "precip_prob_max_pct", "tmax_c", "tmin_c", "wind_max_kmh"]).to_csv(VINTAGE_FILE, index=False, compression="gzip"); log("Forecast vintage", "OK", VINTAGE_FILE)
    except Exception as e: VINTAGE_FILE = ""; log("Forecast vintage", "ERROR", str(e)[:120])

# ============================ EVENT ENGINE (daily rain + source probability columns) ============================
def rain_class(mm):
    if mm is None or np.isnan(mm): return "No data"
    for lim, nm in [(204.5, "Extremely heavy"), (115.6, "Very heavy"), (64.5, "Heavy"), (15.6, "Moderate"), (2.5, "Light"), (0.1, "Very light")]:
        if mm >= lim: return nm
    return "None"
LEVEL_ORDER = {"RED": 0, "AMBER": 1, "BLUE": 2, "GREEN": 3, "NO DATA": 4}
def window_label(i): return ["Today (<=24h)", "24-48h", "48-72h", "3-5 days", "3-5 days", "5-7 days", "5-7 days"][i] if 0 <= i < 7 else "Beyond 7 days"
def runs(mask):
    out, i, n = [], 0, len(mask)
    while i < n:
        if mask[i]:
            j = i
            while j + 1 < n and mask[j + 1]: j += 1
            out.append((i, j)); i = j + 1
        else: i += 1
    return out
RAIN_D = [f"Rain D{k} mm" for k in range(7)]; PROB_D = [f"Rain prob D{k} %" for k in range(7)]
def analyse(daily, cur, elev, past):
    dates = daily["time"]; n = len(dates); pr = [f(x) for x in daily["precipitation_sum"]]; pp = daily.get("precipitation_probability_max") or [None] * n
    tx = [f(x) for x in daily["temperature_2m_max"]]; tn = [f(x) for x in daily["temperature_2m_min"]]; wd = [f(x) for x in daily["wind_speed_10m_max"]]; ev = []; hi = (f(elev) >= 1500)
    for a, b in runs([(not np.isnan(v)) and v >= RAIN_DAY_MM for v in pr]):
        tot = float(np.nansum(pr[a:b + 1])); peak = float(np.nanmax(pr[a:b + 1]))
        ev.append({"Type": "Rain", "grp": "rain", "a": a, "Start": dates[a], "End": dates[b], "Days": b - a + 1, "Total mm": round(tot, 1), "Peak value": round(peak, 1), "Intensity": rain_class(peak), "Level": "RED" if peak >= 64.5 else "AMBER" if (peak >= 15.6 or tot >= 50) else "BLUE", "Prob %": pp[a] if a < len(pp) else None, "Lead": window_label(a)})
    for a, b in runs([(not np.isnan(v)) and v < RAIN_DAY_MM for v in pr]):
        if b - a + 1 >= 5: ev.append({"Type": "Dry spell", "grp": "dry", "a": a, "Start": dates[a], "End": dates[b], "Days": b - a + 1, "Total mm": round(float(np.nansum(pr[a:b + 1])), 1), "Peak value": np.nan, "Intensity": "n/a", "Level": "BLUE", "Prob %": None, "Lead": window_label(a)})
    for a, b in runs([(not np.isnan(v)) and v >= 40 for v in tx]):
        m = float(np.nanmax(tx[a:b + 1])); ev.append({"Type": "High temperature (indicative)", "grp": "heat", "a": a, "Start": dates[a], "End": dates[b], "Days": b - a + 1, "Total mm": np.nan, "Peak value": m, "Intensity": f"Tmax {m:.1f}C", "Level": "RED" if m >= 45 else "AMBER", "Prob %": None, "Lead": window_label(a)})
    for a, b in runs([(not np.isnan(v)) and v <= 5 for v in tn]):
        m = float(np.nanmin(tn[a:b + 1])); ev.append({"Type": "Low temperature (high elevation: usual for altitude)" if hi else ("Frost risk (indicative)" if m <= 0 else "Low temperature (indicative)"), "grp": "cold", "a": a, "Start": dates[a], "End": dates[b], "Days": b - a + 1, "Total mm": np.nan, "Peak value": m, "Intensity": f"Tmin {m:.1f}C", "Level": "BLUE" if hi else ("RED" if m <= 0 else "AMBER"), "Prob %": None, "Lead": window_label(a)})
    for a, b in runs([(not np.isnan(v)) and v >= 50 for v in wd]):
        m = float(np.nanmax(wd[a:b + 1])); ev.append({"Type": "Strong wind (indicative, model 10 m)", "grp": "wind", "a": a, "Start": dates[a], "End": dates[b], "Days": b - a + 1, "Total mm": np.nan, "Peak value": m, "Intensity": f"{m:.0f} km/h", "Level": "RED" if m >= 62 else "AMBER", "Prob %": None, "Lead": window_label(a)})
    ev.sort(key=lambda e: (LEVEL_ORDER[e["Level"]], e["a"])); rain_ev = sorted([e for e in ev if e["grp"] == "rain"], key=lambda e: e["a"]); top_ = ev[0] if ev else None
    wet_sum = float(np.nansum([v for v in pr if not np.isnan(v) and v >= RAIN_DAY_MM])); ev_sum = float(sum(e["Total mm"] for e in rain_ev)); pv = [f(x) for x in (past or {}).get("precipitation_sum", [])]
    probs = [f(p) for p in pp[:7]]; pvalid = [p for p in probs if not np.isnan(p)]
    s = {"Model alert level": top_["Level"] if top_ else "GREEN", "Hazards": "; ".join(sorted({e["Type"] for e in ev})) or "None", "Top event": f'{top_["Type"]} ({top_["Intensity"]})' if top_ else "None", "Top event start": top_["Start"] if top_ else "", "Top event lead": top_["Lead"] if top_ else "",
         "_lead": top_["a"] if top_ else None, "_grp": top_["grp"] if top_ else None, "Rain events (7d)": len(rain_ev), "1st rain event start": rain_ev[0]["Start"] if rain_ev else "", "1st rain event end": rain_ev[0]["End"] if rain_ev else "",
         "1st rain event duration d": rain_ev[0]["Days"] if rain_ev else 0, "1st rain event total mm": rain_ev[0]["Total mm"] if rain_ev else 0.0, "1st rain event peak daily mm": rain_ev[0]["Peak value"] if rain_ev else 0.0, "1st rain prob %": rain_ev[0]["Prob %"] if rain_ev else None,
         "1st rain lead": rain_ev[0]["Lead"] if rain_ev else "", "Cumulative 5-day mm": round(float(np.nansum(pr[:5])), 1), "Cumulative 7-day mm": round(float(np.nansum(pr[:7])), 1), "Rainy days (7d)": int(sum(1 for v in pr if not np.isnan(v) and v >= RAIN_DAY_MM)),
         "Max daily mm": round(float(np.nanmax(pr)), 1) if not all(np.isnan(pr)) else np.nan, "Tmax 7d C": float(np.nanmax(tx)) if not all(np.isnan(tx)) else np.nan, "Tmin 7d C": float(np.nanmin(tn)) if not all(np.isnan(tn)) else np.nan,
         "Max wind km/h": float(np.nanmax(wd)) if not all(np.isnan(wd)) else np.nan, "Elevation m": elev, "Event total check vs daily": "OK" if abs(ev_sum - wet_sum) < 0.15 else f"MISMATCH ({ev_sum:.1f} vs {wet_sum:.1f})",
         "Model-analysis rain D-1 mm": pv[-1] if len(pv) >= 1 else np.nan, "Model-analysis rain D-2 mm": pv[-2] if len(pv) >= 2 else np.nan, "_dates": dates, "Forecast base date (D0)": dates[0],
         "Rain probability score % (source-supplied)": max(pvalid) if pvalid else None, "Probability basis": "Open-Meteo source-supplied daily max precipitation probability; blank = source did not supply (not estimated)"}
    c = cur or {}; s.update({"Model-current time": c.get("time", ""), "Model-current temp C": c.get("temperature_2m", np.nan), "Model-current RH %": c.get("relative_humidity_2m", np.nan), "Model-current wind km/h": c.get("wind_speed_10m", np.nan), "Model-current cloud %": c.get("cloud_cover", np.nan)})
    for k in range(7):
        s[RAIN_D[k]] = pr[k] if k < n else np.nan; s[PROB_D[k]] = probs[k] if k < len(probs) and not np.isnan(probs[k]) else None
    return s, ev

# ============================ HISTORY: CHIRPS (historical only) + MONTH-ON-MONTH YEAR-WISE ============================
chirps = ee.ImageCollection("UCSB-CHG/CHIRPS/DAILY"); chirps_latest = dt.datetime.fromtimestamp(chirps.aggregate_max("system:time_start").getInfo() / 1000, UTC).date(); lag_days = (today - chirps_latest).days
log("CHIRPS", "CHECK" if lag_days > 7 else "OK", f"latest {chirps_latest} ({lag_days} days behind): historical estimate only")
def csum(s, e): return chirps.filterDate(str(s), str(e)).select("precipitation").sum()
def crain(s, e): return chirps.filterDate(str(s), str(e)).select("precipitation").map(lambda im: im.gte(RAIN_DAY_MM)).sum()
month_start = today.replace(day=1); HB = ["y1", "y2", "y3", "r1", "r2", "r3", "mtd0", "mtd1", "mtd2", "mtd3", "recent7", "recent30"]; img = None
for n in (1, 2, 3):
    s = shift_year(today, n); e = s + dt.timedelta(days=7)
    for nm, im in ((f"y{n}", csum(s, e)), (f"r{n}", crain(s, e))): img = im.rename(nm) if img is None else img.addBands(im.rename(nm))
for n in (0, 1, 2, 3):
    if n == 0: end_ = min(chirps_latest, today); im = csum(month_start, end_ + dt.timedelta(days=1)) if end_ >= month_start else ee.Image.constant(0).rename("precipitation")
    else: im = csum(shift_year(month_start, n), shift_year(today, n) + dt.timedelta(days=1))
    img = img.addBands(im.rename(f"mtd{n}"))
img = img.addBands(csum(chirps_latest - dt.timedelta(days=6), chirps_latest + dt.timedelta(days=1)).rename("recent7")).addBands(csum(chirps_latest - dt.timedelta(days=29), chirps_latest + dt.timedelta(days=1)).rename("recent30"))
M_YEARS = list(range(today.year - 3, today.year + 1)); IM_M, M_BANDS = None, []
for y in M_YEARS:
    for m in range(1, 13):
        ms, me = dt.date(y, m, 1), dt.date(y + (m == 12), (m % 12) + 1, 1)
        if me - dt.timedelta(days=1) > chirps_latest: continue           # only COMPLETE months
        nm = f"m{y}{m:02d}"; b = csum(ms, me).rename(nm); IM_M = b if IM_M is None else IM_M.addBands(b); M_BANDS.append(nm)
def ee_points(df, key, image, bands, scale=5566, chunk=800):
    out = {}; items = [(str(k), la, lo) for k, la, lo in df[[key, "lat", "lon"]].itertuples(index=False, name=None) if pd.notna(la)]
    for i in range(0, len(items), chunk):
        feats = [ee.Feature(ee.Geometry.Point([float(lo), float(la)]), {"KEY": k}) for k, la, lo in items[i:i + chunk]]
        try:
            rr = image.reduceRegions(collection=ee.FeatureCollection(feats), reducer=ee.Reducer.first(), scale=scale).select(["KEY"] + bands, None, False).getInfo()
            for ft in rr["features"]: out[ft["properties"]["KEY"]] = ft["properties"]
        except Exception as e: log("EE point chunk", "ERROR", str(e)[:120])
    return out
print("\nComputing CHIRPS history (same-window + month-by-year)...")
H_D = ee_points(dm, "DKEY", img, HB); has_t = tm is not None and tm["lat"].notna().any(); H_T = ee_points(tm, "TKEY", img, HB) if has_t else {}
H_MD = ee_points(dm, "DKEY", IM_M, M_BANDS) if IM_M is not None else {}; H_MT = ee_points(tm, "TKEY", IM_M, M_BANDS) if (has_t and IM_M is not None) else {}
log("CHIRPS history", "OK" if H_D else "FAIL", f"districts {len(H_D)}/{len(dm)}, tehsils {len(H_T)}/{nT}; monthly bands {len(M_BANDS)}")
IM_D, IM_T = {}, {}
if IMERG_FRESH and IMG_IMERG is not None:
    try: IM_D = ee_points(dm, "DKEY", IMG_IMERG, ["imerg24"], 11132); IM_T = ee_points(tm, "TKEY", IMG_IMERG, ["imerg24"], 11132) if has_t else {}
    except Exception as e: log("IMERG points", "ERROR", str(e)[:100])
def hist_cols(h, fc7):
    if not h: return {"Hist method": "no data"}
    ys = [f(h.get(f"y{n}")) for n in (1, 2, 3)]; rs = [f(h.get(f"r{n}")) for n in (1, 2, 3)]; ms = [f(h.get(f"mtd{n}")) for n in (1, 2, 3)]
    o = {"Hist method": "CHIRPS point sample (satellite-gauge estimate; HISTORICAL ONLY)", "Same-window y-1 mm": ys[0], "y-2 mm": ys[1], "y-3 mm": ys[2], "Hist avg mm": round(float(np.nanmean(ys)), 1), "Hist median mm": round(float(np.nanmedian(ys)), 1), "Hist min mm": round(float(np.nanmin(ys)), 1),
         "Hist max mm": round(float(np.nanmax(ys)), 1), "Hist rainy days avg": round(float(np.nanmean(rs)), 1), "Month-to-date prior yrs avg mm": round(float(np.nanmean(ms)), 1)}
    if not np.isnan(fc7):
        a = o["Hist avg mm"]; o["Forecast 7d vs hist avg mm"] = round(fc7 - a, 1); o["Departure %"] = round(100 * (fc7 - a) / a, 0) if a > 0.5 else np.nan; o["Exceeds N of 3 prior yrs"] = int(sum(fc7 > y for y in ys if not np.isnan(y)))
        o["Interpretation"] = f"Forecast {fc7:.1f} mm vs 3-yr same-window avg {a:.1f} mm; exceeds {o['Exceeds N of 3 prior yrs']} of 3 prior years. Recent pattern only, not a climate normal."
    return o
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
def mom_table(df, keycol, H, level):
    rows = []
    if df is None: return pd.DataFrame()
    for r in df.itertuples():
        d = r._asdict(); h = H.get(str(getattr(r, keycol))); basis = "CHIRPS point sample at representative point (complete months only)"
        if not h and level == "tehsil": h = H_MD.get(str(d["DKEY"])); basis = "DISTRICT-INHERITED (no tehsil coordinates)" if h else basis
        for y in M_YEARS:
            v = [round(f(h.get(f"m{y}{m:02d}")), 1) if (h and h.get(f"m{y}{m:02d}") is not None) else np.nan for m in range(1, 13)]
            rec = {"State": d["State"], "District": d["District"], "LGD district code": d.get("district_code")}
            if level == "tehsil": rec.update({"Tehsil": d["Tehsil"], "LGD sub-district code": d.get("tehsil_code")})
            rec["Year"] = y
            for i, m in enumerate(MONTHS): rec[m + " mm"] = v[i]
            rec["Year total mm (complete months)"] = round(float(np.nansum(v)), 1) if not all(np.isnan(v)) else np.nan; rec["Months available"] = int(sum(1 for x in v if not np.isnan(x))); rec["Source/basis"] = basis; rec["Data through"] = str(chirps_latest); rows.append(rec)
    return pd.DataFrame(rows)
MOM_D = mom_table(dm, "DKEY", H_MD, "district"); MOM_T = mom_table(tm, "TKEY", H_MT, "tehsil")
if len(MOM_D):
    mc = [m + " mm" for m in MONTHS]; st_ = MOM_D.groupby(["State", "Year"])[mc].mean().round(1).reset_index(); al_ = MOM_D.groupby("Year")[mc].mean().round(1).reset_index(); al_.insert(0, "State", "ALL INDIA (mean of districts)")
    MOM_S = pd.concat([al_, st_], ignore_index=True); MOM_S["Basis"] = "mean of district point samples (CHIRPS; complete months only)"
else: MOM_S = pd.DataFrame()

# ============================ DAILY ARCHIVE (optional; rate-limit aware; Colab default) ============================
HIST_V2 = DIRS["history"] + "/district_daily_history_v2.csv"; ARCH_STATE = DIRS["checkpoints"] + "/archive_state.json"; ARCH = {"done_points": 0, "total_points": 0, "rate_limited": 0, "unavailable": 0, "empty": 0, "stopped": "", "range": "n/a", "new_rows": 0, "pending": 0}
HCOLS = ["point_id", "lat", "lon", "date", "precip_mm", "tmax_c", "tmin_c", "source", "spatial_resolution", "units", "retrieved"]
def load_state():
    try: return json.load(open(ARCH_STATE))
    except Exception: return {"points": {}}
def archive_request(sub, start, end):
    p = {"latitude": ",".join(str(la) for _, la, _ in sub), "longitude": ",".join(str(lo) for _, _, lo in sub), "start_date": str(start), "end_date": str(end), "daily": "precipitation_sum,temperature_2m_max,temperature_2m_min", "timezone": "Asia/Kolkata"}
    for attempt in range(4):
        try: r = requests.get("https://archive-api.open-meteo.com/v1/archive", params=p, headers=UA, timeout=240)
        except Exception as e: return "unavailable", type(e).__name__
        if r.status_code == 200:
            try: return "ok", r.json()
            except Exception: return "unavailable", "bad JSON"
        if r.status_code == 429:
            reason = ""
            try: reason = str(r.json().get("reason", ""))
            except Exception: pass
            rl = reason.lower()
            if "daily" in rl: return "rate_limited_daily", reason
            if "hourly" in rl: return "rate_limited_hourly", reason
            if attempt == 3: return "rate_limited_minutely", reason or "HTTP 429"
            ra = r.headers.get("Retry-After"); time.sleep(min(float(ra) if ra and re.match(r"^\d+(\.\d+)?$", ra) else 60 * (2 ** attempt), 300) + random.uniform(0, 15)); continue
        if r.status_code >= 500: time.sleep(15 * (2 ** attempt) + random.uniform(0, 8)); continue
        return "unavailable", f"HTTP {r.status_code}"
    return "unavailable", "retries exhausted"
def run_archive():
    pts = {}
    for r in dm.itertuples():
        if pd.notna(r.arch_lat): pts.setdefault(pid_of(r.arch_lat, r.arch_lon), (float(r.arch_lat), float(r.arch_lon)))
    st = load_state(); done = set(pd.read_csv(HIST_V2, usecols=["point_id"], dtype=str)["point_id"]) if os.path.exists(HIST_V2) else set(); ARCH["total_points"] = len(pts); nowts = time.time()
    todo = [p for p in pts if p not in done and st["points"].get(p, {}).get("status") != "empty" and st["points"].get(p, {}).get("next_retry", 0) <= nowts][:ARCHIVE_MAX_POINTS_PER_RUN]; start, end = shift_year(today, 3), today - dt.timedelta(days=6); consec = 0
    for i in range(0, len(todo), ARCHIVE_BATCH):
        sub = [(p, *pts[p]) for p in todo[i:i + ARCHIVE_BATCH]]; res, data = archive_request([(a, round(b, 3), round(c, 3)) for a, b, c in sub], start, end)
        if res == "ok":
            consec = 0; js = data if isinstance(data, list) else [data]; rows = []
            for (p, la, lo), it in zip(sub, js):
                d = it.get("daily") if isinstance(it, dict) else None
                if not d or all(v is None for v in d.get("precipitation_sum", [])): st["points"][p] = {"status": "empty", "next_retry": nowts + 30 * 86400}; ARCH["empty"] += 1; continue
                rows += [(p, la, lo, t, a, b, c, "Open-Meteo archive (reanalysis; not station data)", "~0.1-0.25 deg model grid", "mm, degC", RUN_TS) for t, a, b, c in zip(d["time"], d["precipitation_sum"], d["temperature_2m_max"], d["temperature_2m_min"])]; st["points"][p] = {"status": "done"}
            if rows: pd.DataFrame(rows, columns=HCOLS).to_csv(HIST_V2, mode="a", header=not os.path.exists(HIST_V2), index=False); ARCH["new_rows"] += len(rows)
        else:
            for p, _, _ in sub: st["points"][p] = {"status": "rate_limited" if res.startswith("rate_limited") else "unavailable", "attempts": st["points"].get(p, {}).get("attempts", 0) + 1, "last": str(data)[:100], "next_retry": nowts + (86400 if res == "rate_limited_daily" else 7200 if res.startswith("rate_limited") else 21600)}
            ARCH["rate_limited" if res.startswith("rate_limited") else "unavailable"] += len(sub); consec += 1; log("Archive batch", "FAIL", f"{res}: {str(data)[:90]}")
            if res in ("rate_limited_daily", "rate_limited_hourly") or consec >= 2: ARCH["stopped"] = f"stopped early ({res}); queued for retry"; json.dump(st, open(ARCH_STATE, "w")); break
        json.dump(st, open(ARCH_STATE, "w")); time.sleep(random.uniform(10, 20))
    done = set(pd.read_csv(HIST_V2, usecols=["point_id"], dtype=str)["point_id"]) if os.path.exists(HIST_V2) else set(); ARCH["done_points"] = len(done & set(pts)); ARCH["pending"] = sum(1 for v in st["points"].values() if v.get("status") in ("rate_limited", "unavailable"))
    if os.path.exists(HIST_V2): d_ = pd.read_csv(HIST_V2, usecols=["date"], dtype=str)["date"]; ARCH["range"] = f"{d_.min()} to {d_.max()}"
if BUILD_DAILY_HISTORY and HISTORY_WRITE:
    try: run_archive(); log("Historical daily archive", "OK" if ARCH["done_points"] else "CHECK", f"{ARCH['done_points']}/{ARCH['total_points']} points; {ARCH['range']}; rate-limited {ARCH['rate_limited']}; {ARCH['stopped']}")
    except Exception as e: log("Historical daily archive", "ERROR", f"{type(e).__name__}: {str(e)[:120]}")
else: log("Historical daily archive", "INFO", "skipped (disabled for this environment or Drive not mounted); month-by-year CHIRPS history is built regardless")

# ============================ AGRONOMY RULES (generic, UNVALIDATED) ============================
CAL = {"Paddy": {1: "Flowering/Grain filling (rabi/southern)", 2: "Harvesting (rabi/southern)", 3: "Harvesting (rabi)", 5: "Land preparation/Nursery", 6: "Nursery/Sowing", 7: "Transplanting", 8: "Vegetative", 9: "Vegetative/Flowering", 10: "Flowering/Grain filling", 11: "Harvesting", 12: "Vegetative (rabi/southern)"},
 "Wheat": {10: "Land preparation", 11: "Sowing", 12: "Germination/Vegetative", 1: "Vegetative/Flowering", 2: "Flowering/Grain filling", 3: "Grain filling/Maturity", 4: "Harvesting"}, "Maize": {6: "Sowing", 7: "Vegetative", 8: "Vegetative/Flowering", 9: "Grain filling", 10: "Harvesting", 11: "Sowing (rabi)", 12: "Vegetative", 1: "Vegetative", 2: "Flowering", 3: "Grain filling"},
 "Cotton": {5: "Sowing", 6: "Sowing/Germination", 7: "Vegetative", 8: "Flowering", 9: "Boll formation", 10: "Boll formation/Harvesting", 11: "Harvesting", 12: "Harvesting"}, "Soybean": {6: "Sowing", 7: "Vegetative", 8: "Flowering", 9: "Pod development", 10: "Harvesting"},
 "Groundnut": {6: "Sowing", 7: "Vegetative", 8: "Flowering", 9: "Pod development", 10: "Harvesting", 11: "Sowing (rabi)", 12: "Vegetative", 1: "Flowering", 2: "Pod development", 3: "Harvesting"}, "Sugarcane": {2: "Sowing/Planting", 3: "Sowing/Planting", 10: "Sowing/Planting", 11: "Sowing/Planting", 6: "Vegetative", 7: "Vegetative", 8: "Vegetative", 9: "Vegetative", 12: "Harvesting", 1: "Harvesting"},
 "Pulses": {6: "Sowing", 7: "Vegetative", 8: "Flowering", 9: "Pod development", 10: "Sowing (rabi)", 11: "Vegetative", 12: "Flowering", 1: "Pod development", 2: "Harvesting", 3: "Harvesting"}, "Millets": {6: "Sowing", 7: "Vegetative", 8: "Flowering", 9: "Grain filling", 10: "Harvesting"}, "Mustard": {10: "Sowing", 11: "Vegetative", 12: "Flowering", 1: "Pod development", 2: "Maturity", 3: "Harvesting"}}
VK = {"rain": {"harvest": 1.0, "maturity": .9, "flower": .9, "grain": .8, "boll": .8, "pod": .8, "germination": .7, "sowing": .6, "vegetative": .5, "transplant": .5, "land": .3, "nursery": .5}, "dry": {"flower": 1.0, "grain": .9, "boll": .9, "pod": .9, "germination": .9, "transplant": .9, "sowing": .8, "vegetative": .7, "nursery": .8, "maturity": .3, "harvest": .1, "land": .2},
      "heat": {"flower": 1.0, "grain": 1.0, "pod": .9, "boll": .8, "sowing": .5, "vegetative": .5, "harvest": .3, "maturity": .4, "land": .2, "nursery": .5, "germination": .6, "transplant": .6}, "cold": {"flower": .9, "pod": .8, "grain": .7, "germination": .6, "vegetative": .4, "sowing": .4, "boll": .6, "nursery": .6, "harvest": .2, "maturity": .3, "land": .1, "transplant": .4},
      "wind": {"maturity": .8, "grain": .8, "flower": .7, "boll": .7, "harvest": .6, "pod": .6, "vegetative": .4, "sowing": .2, "land": .1, "nursery": .3, "germination": .2, "transplant": .3}}
def vuln(stage, grp): sl = stage.lower(); return max([v for k, v in VK.get(grp, {}).items() if k in sl] or [0.3])
ACT = {"rain": {"harvest": "Advance harvest where mature; keep harvested produce covered/dry; arrange drainage and drying space", "flower": "Avoid spraying/irrigation before rain; ensure field drainage", "sowing": "Postpone sowing/tillage if heavy rain is forecast; prepare drainage channels", "vegetative": "Drain excess water; postpone fertiliser and spray", "default": "Review field drainage; postpone spraying and irrigation until the event passes"},
 "dry": {"default": "Review irrigation schedule; prioritise critical stages (flowering/grain/pod); consider mulching"}, "heat": {"default": "Irrigate in early morning/evening; avoid spraying in peak heat; monitor flowering/grain stages"}, "cold": {"default": "Prepare frost/cold protection where relevant; validate with state Agromet advisory"}, "wind": {"default": "Postpone spraying; support tall crops; secure harvested produce"}}
def action_for(grp, stage):
    d = ACT.get(grp, {}); sl = stage.lower()
    for k, v in d.items():
        if k != "default" and k in sl: return v
    return d.get("default", "Monitor and follow IMD/State Agromet advisory")
TRACT = {"rain": ("Field operations postponed during/after event; retail conversion may slip; demand bunches up once fields are accessible", "Post-rain inspection (clutch, hydraulics, electricals), waterlogging service, implement readiness", "Pre-position technicians/spares; schedule post-rain doorstep service camps"),
 "dry": ("Tillage/dryland operations may stall; irrigation-equipment use may rise", "Irrigation-pump/PTO service demand may rise", "Plan irrigation-related service and preventive maintenance"), "heat": ("Working hours shift; higher engine cooling load", "Cooling-system and filter inspections", "Run preventive-maintenance doorstep campaigns"),
 "cold": ("Cold starts affected in cold areas", "Battery/starter/fuel-system checks", "Winter-readiness checks where relevant"), "wind": ("Spraying and open-field operations disrupted", "Sprayer/implement checks", "Hold spraying-related promotions during high wind")}
def conf_class(prob):
    if prob is not None and not (isinstance(prob, float) and np.isnan(prob)): return "High" if prob >= 70 else "Medium" if prob >= 40 else "Low"
    return "Not supplied"
LV_SC = {"RED": 25, "AMBER": 17, "BLUE": 8, "GREEN": 0, "NO DATA": 0}; LEAD_SC = [10, 8, 6, 4, 4, 2, 2]; CONF_SC = {"High": 10, "Medium": 6, "Low": 3, "Not supplied": 2}; month = today.month

# ============================ BUILD LOCATION TABLES ============================
def build_rows(df, keycol, level):
    rows, events = [], []; extras = [c for c in df.columns if str(c).startswith("TM: ")]
    for r in df.itertuples():
        d = r._asdict(); key = getattr(r, keycol); cell = d.get("cell")
        base = {"State": d["State"], "District": d["District"], "LGD state code": d.get("state_code"), "LGD district code": d.get("district_code")}
        if level == "tehsil":
            base.update({"Tehsil": d["Tehsil"], "LGD sub-district code": d.get("tehsil_code"), "Master source": d.get("Source", "")})
            for c in ("Source state name", "Source district name", "Source tehsil name"): base[c] = df.loc[r.Index, c] if c in df.columns else ""   # [GH] carry original spellings to the sheet
            for c in extras: base[c] = df.loc[r.Index, c]
        base.update({"Rep lat": round(d["lat"], 4) if pd.notna(d["lat"]) else np.nan, "Rep lon": round(d["lon"], 4) if pd.notna(d["lon"]) else np.nan, "_key": key})
        fc = FC.get(cell) if cell else None
        if not fc: base.update({"Model alert level": "NO DATA", "Alert level (combined)": "NO DATA", "Coverage status": "NO FORECAST: " + ("no representative coordinates" if not cell else "forecast fetch failed for this cell")}); rows.append(base); continue
        fut, past = split_daily(fc["daily"]); s, ev = analyse(fut, fc.get("current"), fc.get("elev"), past); base.update(s); base["Coverage status"] = "forecast OK"; dk = d["DKEY"]; lead = s["_lead"]
        base["Confidence"] = conf_class(s["1st rain prob %"]) if s["Rain events (7d)"] else "Not supplied"; base["Confidence basis"] = f"source probability {s['1st rain prob %']}% (blank = not supplied); IMD corroboration unavailable; lead {s['1st rain lead'] or lead}"
        il = imd_by.get(dk, {}).get("IMD warning level", "NOT AVAILABLE"); il2 = il if il in LEVEL_ORDER else "GREEN"; comb = s["Model alert level"] if LEVEL_ORDER[s["Model alert level"]] <= LEVEL_ORDER[il2] else il2
        base["IMD warning level (district-level)"] = il if IMD_ACTIVE else "NOT AVAILABLE (IMD credentials pending)"; base["Alert level (combined)"] = comb; base["Alert source"] = "Model" if comb != "GREEN" else "n/a"
        base["Priority label"] = comb + (" (low confidence - verify)" if comb in ("RED", "AMBER") and base["Confidence"] in ("Low", "Not supplied") else "")
        base["Forecast source"] = "Open-Meteo global model (best-match mix; NOT IMD)"; base["Forecast retrieved"] = RUN_TS; base["Forecast issue/model run"] = "Not supplied by source"; base["Forecast valid"] = f'{s["_dates"][0]} to {s["_dates"][-1]}'
        base["Forecast grid lat/lon (actual)"] = f'{fc.get("grid_lat")},{fc.get("grid_lon")}'; base["Locations sharing this 0.1-deg cell"] = cell_users.get(cell, 1)
        base["Resolution"] = "MODELLED POINT FORECAST at tehsil representative point (~0.1 deg grid); NOT a tehsil observation" if level == "tehsil" else "MODELLED POINT FORECAST at district representative point (~0.1 deg grid)"
        base.update(hist_cols((H_T if level == "tehsil" else H_D).get(str(key)), s["Cumulative 7-day mm"]))
        im = (IM_T if level == "tehsil" else IM_D).get(str(key)); base["IMERG 24h mm"] = round(f(im["imerg24"]), 1) if im and im.get("imerg24") is not None else np.nan
        grp = s["_grp"]; act = [(c, CAL[c][month]) for c in CAL if month in CAL[c]]; base["_cv"] = max([vuln(st, grp) for _, st in act] or [0]) if grp else 0; base["_ncrop"] = len(act); base["_lead"] = lead; base["_grp"] = grp
        base["Active crop groups (calendar-estimated)"] = ", ".join(c for c, _ in act); rows.append(base)
        for e in ev: events.append({"State": d["State"], "District": d["District"], "Tehsil": d.get("Tehsil", ""), "Level": level, "Event": e["Type"], "Start": e["Start"], "End": e["End"], "Duration d": e["Days"], "Total mm": e["Total mm"], "Peak": e["Peak value"], "Intensity": e["Intensity"], "Severity": e["Level"], "Lead": e["Lead"], "Source prob % (blank = not supplied)": e["Prob %"], "Forecast retrieved": RUN_TS})
    return pd.DataFrame(rows), pd.DataFrame(events)
print("\nBuilding district intelligence..."); RD, ED = build_rows(dm, "DKEY", "district")
print("Building tehsil intelligence..."); RT, ET = (build_rows(tm, "TKEY", "tehsil") if tm is not None and len(tm) else (pd.DataFrame(), pd.DataFrame()))
NEED = {"Alert level (combined)": "NO DATA", "Model alert level": "NO DATA", "Confidence": "Not supplied", "Coverage status": "", "Top event": "", "IMD warning level (district-level)": "NOT AVAILABLE", "Hazards": "", "Event total check vs daily": ""}
for T in (RD, RT):
    for c, v in NEED.items():
        if len(T) and c not in T.columns: T[c] = v
    if len(T): T["Alert level (combined)"] = T["Alert level (combined)"].fillna("NO DATA")
INH_COLS = ["Model alert level", "Hazards", "Top event", "Top event start", "Top event lead", "Cumulative 5-day mm", "Cumulative 7-day mm", "1st rain event total mm", "Alert level (combined)", "Priority label", "Confidence", "Hist avg mm", "Interpretation", "Model-analysis rain D-1 mm", "Model-analysis rain D-2 mm",
            "Model-current temp C", "IMERG 24h mm", "_cv", "_ncrop", "_lead", "_grp", "Forecast source", "Forecast retrieved", "Forecast valid", "Forecast base date (D0)", "Rain probability score % (source-supplied)", "Probability basis"] + RAIN_D + PROB_D
if len(RT):
    inh = RD.assign(_k=RD["LGD district code"].astype(str)).drop_duplicates("_k").set_index("_k")
    for i in RT[RT["Coverage status"] != "forecast OK"].index:
        k = str(RT.loc[i, "LGD district code"])
        if k in inh.index and inh.loc[k, "Model alert level"] != "NO DATA":
            for c in INH_COLS:
                if c in inh.columns: RT.loc[i, c] = inh.loc[k, c]
            RT.loc[i, "Resolution"] = "DISTRICT-LEVEL INHERITED DATA (no tehsil coordinates). NOT a tehsil observation or forecast."; RT.loc[i, "Coverage status"] = "inherited from district"
def attach_obs(T, src_df, keyname):
    if not len(T): return T
    for c in ["Obs source", "Obs station", "Obs time UTC", "Obs weather"]: T[c] = ""
    for c in ["Obs distance km", "Obs age h", "Obs temp C", "Obs wind km/h", "Obs visibility km"]: T[c] = np.nan
    T["Obs status"] = "NOT AVAILABLE (no live station source)"; S = OBS_ST.dropna(subset=["lat", "lon"]).reset_index(drop=True)
    if len(S):
        la = pd.to_numeric(src_df["lat"], errors="coerce").values; lo = pd.to_numeric(src_df["lon"], errors="coerce").values; pos = {k: i for i, k in enumerate(src_df[keyname])}
        for i in T.index:
            j0 = pos.get(T.loc[i, "_key"]) if "_key" in T.columns else None
            if j0 is None or np.isnan(la[j0]): continue
            dd = hav_km(la[j0], lo[j0], S["lat"].values, S["lon"].values); j = int(np.argmin(dd))
            if dd[j] <= OBS_MAX_KM:
                s_ = S.loc[j]; fresh = pd.notna(s_.get("age_h")) and s_["age_h"] <= 6
                T.loc[i, ["Obs source", "Obs station", "Obs distance km", "Obs time UTC", "Obs age h", "Obs temp C", "Obs wind km/h", "Obs visibility km", "Obs weather"]] = [s_["src"], str(s_["id"]), round(float(dd[j]), 1), str(s_["time"])[:16], s_.get("age_h", np.nan), s_["temp"], s_["wind"], s_["vis"], s_["wx"] or "no precipitation/phenomena reported"]
                T.loc[i, "Obs status"] = "NEARBY STATION (not an in-tehsil observation)" if fresh else "NEARBY STATION - STALE (>6 h)"
            else: T.loc[i, "Obs status"] = f"NO STATION within {OBS_MAX_KM} km"
    return T
RD = attach_obs(RD, dm, "DKEY"); RT = attach_obs(RT, tm, "TKEY") if tm is not None and len(RT) else RT
def live_evidence(r):
    d1, d2 = r.get("Model-analysis rain D-1 mm", np.nan), r.get("Model-analysis rain D-2 mm", np.nan)
    if str(r.get("Obs status", "")).startswith("NEARBY STATION (") and pd.notna(r.get("Obs age h")) and r.get("Obs age h") <= 6: return pd.Series(["2 STATION OBSERVATION (live, nearby)", f'{r["Obs weather"]} at {r["Obs station"]} ({r["Obs distance km"]} km)', f'{r["Obs age h"]} h old'])
    if pd.notna(r.get("IMERG 24h mm", np.nan)): return pd.Series(["3 SATELLITE ESTIMATE (IMERG)", f'{r["IMERG 24h mm"]} mm in last 24 h', IMERG_NOTE[:60]])
    if pd.notna(d1) or pd.notna(d2): return pd.Series(["4 MODEL ANALYSIS (NOT observed)", f"{d1} mm yesterday; {d2} mm day before", f"retrieved {RUN_TS}"])
    return pd.Series(["5 NOT AVAILABLE", "", ""])
for T in (RD, RT):
    if len(T): T[["Live evidence class", "Live rain evidence", "Evidence freshness"]] = T.apply(live_evidence, axis=1)
share = {}
if len(RT): share = RT[RT["Coverage status"] == "forecast OK"].groupby(RT["LGD district code"].astype(str))["Alert level (combined)"].apply(lambda s: float(s.isin(["RED", "AMBER"]).mean())).to_dict()
def score(T, level):
    if not len(T): return T
    cols = {c: [] for c in ["sev", "ex", "cv", "ec", "cf", "ur", "tot"]}
    for i in T.index:
        lv = T.loc[i, "Alert level (combined)"]; sv = share.get(str(T.loc[i, "LGD district code"]), 0.0); s_ = LV_SC.get(lv, 0); e_ = 20 * sv if (level == "district" or lv in ("RED", "AMBER")) else 0.0
        cvv = T.loc[i, "_cv"] if "_cv" in T.columns else np.nan; c_ = 20 * float(cvv) if pd.notna(cvv) else 0.0; nc = T.loc[i, "_ncrop"] if "_ncrop" in T.columns else np.nan
        n_ = 15 * min(1.0, float(nc) / 4) if pd.notna(nc) and lv not in ("GREEN", "NO DATA") else 0.0; q_ = CONF_SC.get(T.loc[i, "Confidence"], 2); ld = T.loc[i, "_lead"] if "_lead" in T.columns else np.nan
        u_ = LEAD_SC[int(ld)] if pd.notna(ld) and int(ld) < 7 and lv not in ("GREEN", "NO DATA") else 0
        for kx, vx in zip(cols, (s_, round(e_, 1), round(c_, 1), round(n_, 1), q_, u_, round(s_ + e_ + c_ + n_ + q_ + u_, 1))): cols[kx].append(vx)
    T["Score: weather severity /25"] = cols["sev"]; T["Score: geographic exposure /20"] = cols["ex"]; T["Score: crop-stage vulnerability /20 (UNVALIDATED)"] = cols["cv"]; T["Score: economic impact /15 (PROXY)"] = cols["ec"]
    T["Score: forecast confidence /10"] = cols["cf"]; T["Score: urgency /10"] = cols["ur"]; T["PRIORITY SCORE /100"] = cols["tot"]; return T
RD = score(RD, "district"); RT = score(RT, "tehsil")
mism = int(RD["Event total check vs daily"].astype(str).str.startswith("MISMATCH").sum()) if len(RD) else 0
crop_rows, adv_rows, tr_rows = [], [], []
sel = RD[(RD["Alert level (combined)"].isin(["RED", "AMBER"])) | RD["State"].map(norm_txt).isin(["tamil nadu", "puducherry"])] if len(RD) else RD
for i in sel.index:
    T_ = RD.loc[i]; grp = T_.get("_grp")
    if not grp or pd.isna(grp): continue
    lead = T_.get("_lead"); bucket = "TODAY" if lead == 0 else "NEXT 24-48H" if lead == 1 else "48-72H" if lead == 2 else "3-5 DAYS" if lead in (3, 4) else "5-7 DAYS"
    for c, st in [(c, CAL[c][month]) for c in CAL if month in CAL[c]]:
        v = vuln(st, grp); crop_rows.append({"State": T_["State"], "District": T_["District"], "Crop": c, "Stage (CALENDAR-ESTIMATED, not verified)": st, "Crop presence in district": "NOT VERIFIED", "Hazard": T_["Top event"], "Hazard lead": T_["Top event lead"], "Stage vulnerability (0-1, generic heuristic)": v, "Alert level": T_["Alert level (combined)"], "Evidence status": "UNVALIDATED - confirm with ICAR/KVK/IMD Agromet"})
        if v >= .5: adv_rows.append({"State": T_["State"], "District": T_["District"], "Crop": c, "Crop stage (calendar-estimated)": st, "Weather condition": T_["Top event"], "Risk category": T_["Alert level (combined)"], "Action required": f"{bucket}: " + action_for(grp, st), "Advisory validity": T_["Forecast valid"], "Source": "Rules + model forecast; NOT an official advisory", "Issue date": str(today), "Confidence": T_["Confidence"]})
for i in (RD[RD["Alert level (combined)"].isin(["RED", "AMBER", "BLUE"])].index if len(RD) else []):
    T_ = RD.loc[i]; grp = T_.get("_grp")
    if grp and not pd.isna(grp) and grp in TRACT: imp, svc, act = TRACT[grp]; tr_rows.append({"State": T_["State"], "District": T_["District"], "Weather condition": T_["Top event"], "Timing": T_["Top event lead"], "Tractor industry impact (HYPOTHESIS)": imp, "Service opportunity (HYPOTHESIS)": svc, "Recommended action (HYPOTHESIS)": act, "Alert level": T_["Alert level (combined)"], "Priority score": T_["PRIORITY SCORE /100"], "Measured data attached": "NO (hypothesis only)"})
CROPX, ADV, TRC = pd.DataFrame(crop_rows), pd.DataFrame(adv_rows), pd.DataFrame(tr_rows)

# ============================ COVERAGE, STATE SUMMARY, REPORT ============================
def pct(a, b): return round(100 * a / b, 1) if b else 0.0
tfc = int((RT["Coverage status"] == "forecast OK").sum()) if len(RT) else 0; tinh = int(RT["Coverage status"].astype(str).str.startswith("inherited").sum()) if len(RT) else 0
tobs = int(RT["Obs status"].astype(str).str.startswith("NEARBY STATION (").sum()) if len(RT) else 0; dobs = int(RD["Obs status"].astype(str).str.startswith("NEARBY STATION (").sum()) if len(RD) else 0; dist_fc = int((RD["Coverage status"] == "forecast OK").sum()) if len(RD) else 0
tcoord = int(tm["lat"].notna().sum()) if tm is not None else 0; thist = sum(1 for k in (tm["TKEY"] if tm is not None else []) if str(k) in H_T); manual = int(sum(1 for a in audit_rows if a.get("Verification status") != "AUTO-ACCEPTED")); live_ev = RT["Live evidence class"].value_counts().to_dict() if len(RT) else {}
COUNTS = [("1. States/UTs in master", len(sm)), ("2. Districts in master", len(dm)), ("3. Tehsils in master", nT), ("4. Districts with forecast", dist_fc), ("5. Tehsils with representative coordinates", tcoord), ("6. Tehsils with OWN point forecast", tfc), ("7. Tehsils with INHERITED district forecast", tinh),
          ("8. Tehsils with CHIRPS history (own point)", thist), ("9. Tehsils with a fresh nearby station observation", tobs), ("10. Tehsils with no forecast at all", nT - tfc - tinh), ("11. Records needing manual verification", manual)]
alert_order = ["RED", "AMBER", "BLUE", "GREEN"]; st_rows = []
for s_, g in RD.groupby("State"):
    tg = RT[RT["State"] == s_] if len(RT) else pd.DataFrame()
    st_rows.append({"State/UT": s_, "Districts": len(g), "Tehsils": len(tg), **{f"Districts {k}": int((g["Alert level (combined)"] == k).sum()) for k in alert_order}, **{f"Tehsils {k}": int((tg["Alert level (combined)"] == k).sum()) if len(tg) else 0 for k in alert_order[:3]},
        "District forecast coverage %": pct(int((g["Coverage status"] == "forecast OK").sum()), len(g)), "Tehsil forecast coverage % (own+inherited)": pct(int(tg["Coverage status"].astype(str).str.contains("forecast OK|inherited").sum()), len(tg)) if len(tg) else 0.0,
        "STATE RISK SCORE": round(0.5 * float(g["PRIORITY SCORE /100"].max()) + 0.5 * float(g["PRIORITY SCORE /100"].mean()), 1)})
STS = pd.DataFrame(st_rows).sort_values("STATE RISK SCORE", ascending=False).reset_index(drop=True)
for s_ in sm["State"]:
    if s_ not in set(STS["State/UT"]): STS.loc[len(STS)] = {"State/UT": s_, "Districts": 0}
tn_d, tn_t = (RD[RD["State"].map(norm_txt) == "tamil nadu"] if len(RD) else RD), (RT[RT["State"].map(norm_txt) == "tamil nadu"] if len(RT) else RT)
CL = {1: "1 SUCCESSFULLY TESTED", 2: "2 IMPLEMENTED, NOT TESTED", 3: "3 WAITING FOR CREDENTIALS/PERMISSION", 4: "4 SOURCE UNAVAILABLE", 5: "5 DATA STALE", 6: "6 DATA NOT AVAILABLE"}; REPORT = []
def rp(sct, item, cls, verdict, ev): REPORT.append([sct, item, CL[cls], verdict, ev])
rp("IMD", "Authentication", 1 if IMD_ACTIVE else 3, "PASS" if IMD_ACTIVE else "FAIL", f"{IMD_AUTH}; {imd_diag[:120]}; connector disabled until verified credentials (RTI: free for non-commercial use only)")
rp("IMD", "Public bulletin records retained", 1 if len(PUBLIC_ALL) else 6, "PASS" if len(PUBLIC_ALL) >= 8 else "CHECK", f"{len(PUBLIC_ALL)} records")
rp("LGD", "Official automated extraction", 4, "FAIL", "lgdirectory.gov.in CAPTCHA (not bypassed); data.gov.in has no API for sub-districts")
rp("LGD", "Official sub-district extract (CKAN, 2023)", 1 if official is not None else 4, "PASS" if official is not None else "FAIL", f"{0 if official is None else len(official)} sub-districts ({off_mode})")
rp("LGD", "TEHSIL_MASTER", 1 if tm_user is not None else 4, "PASS" if tm_user is not None else "FAIL", f"{0 if tm_user is None else len(tm_user)} rows; {' | '.join(tried)[:200]}")
rp("LGD", "All-India tehsils", 1 if nT > 5000 else 4, "PASS" if nT >= 6400 else "CHECK" if nT else "FAIL", f"{nT} tehsils; states/UTs {len(sm)}; districts {len(dm)}; manual-review {manual}")
rp("Tamil Nadu", "Coverage", 1 if len(tn_d) else 6, "PASS" if len(tn_d) >= 38 else "CHECK", f"{len(tn_d)} districts (ref 38), {len(tn_t)} tehsils")
rp("Forecast", "Open-Meteo", 1 if FC and not failed_cells else 2, "PASS" if FC and not failed_cells else "CHECK", f"{len(FC)}/{len(cells)} cells; failed {len(failed_cells)}; mismatches {mism}; vintage {VINTAGE_FILE or 'not written'}")
rp("Forecast", "Tehsil coverage", 1 if tfc else 6, "PASS" if nT and tfc == nT else "CHECK", f"own {tfc}, inherited {tinh}")
rp("Live", "NOAA METAR", 1 if len(METAR) else 4, "PASS" if len(METAR) else "FAIL", metar_note); rp("Live", "IMERG", 1 if IMERG_FRESH else 5, "PASS" if IMERG_FRESH else "CHECK", IMERG_NOTE); rp("Live", "CHIRPS", 5, "CHECK", f"latest {chirps_latest} ({lag_days} d behind): historical only")
rp("History", "Month-by-year CHIRPS history", 1 if len(MOM_D) else 6, "PASS" if len(MOM_D) else "FAIL", f"{len(M_BANDS)} complete months; district rows {len(MOM_D)}, tehsil rows {len(MOM_T)}; through {chirps_latest}")
rp("History", "Daily archive", 1 if ARCH["done_points"] else 6, "CHECK", f"{ARCH['done_points']}/{ARCH['total_points']} points; {ARCH['range']}; rate-limited {ARCH['rate_limited']}; pending {ARCH['pending']}")
rp("Persistence", "Storage", 1 if HISTORY_WRITE else 4, "PASS" if HISTORY_WRITE else "FAIL", PROJ)

# ============================ TABS ============================
def top(df, n, cols, sort="PRIORITY SCORE /100"): return df.sort_values(sort, ascending=False).head(n)[[c for c in cols if c in df.columns]] if len(df) else pd.DataFrame()
CORE = ["Alert level (combined)", "Priority label", "Confidence", "Rain probability score % (source-supplied)", "Hazards", "Top event", "Top event start", "Top event lead", "Forecast base date (D0)"] + RAIN_D + PROB_D + \
       ["1st rain event start", "1st rain event end", "1st rain event duration d", "1st rain event total mm", "1st rain event peak daily mm", "Cumulative 5-day mm", "Cumulative 7-day mm", "Rainy days (7d)", "Max daily mm", "Tmax 7d C", "Tmin 7d C", "Max wind km/h", "Elevation m",
        "Live evidence class", "Live rain evidence", "Evidence freshness", "Obs station", "Obs distance km", "Obs temp C", "Obs weather", "Hist avg mm", "Departure %", "Interpretation", "IMD warning level (district-level)", "Active crop groups (calendar-estimated)",
        "Score: weather severity /25", "Score: geographic exposure /20", "Score: crop-stage vulnerability /20 (UNVALIDATED)", "Score: economic impact /15 (PROXY)", "Score: forecast confidence /10", "Score: urgency /10", "PRIORITY SCORE /100",
        "Probability basis", "Forecast source", "Forecast retrieved", "Forecast valid", "Forecast grid lat/lon (actual)", "Locations sharing this 0.1-deg cell", "Resolution", "Coverage status"]
DIST_COLS = ["State", "District", "LGD state code", "LGD district code"] + CORE; TEH_COLS = ["State", "District", "Tehsil", "LGD state code", "LGD district code", "LGD sub-district code"] + CORE + ["Rep lat", "Rep lon", "Master source"] + [c for c in (RT.columns if len(RT) else []) if str(c).startswith("TM: ")] + ["Source state name", "Source district name", "Source tehsil name"]   # [GH] original spellings (blank if STANDARDIZE_NAMES=0)
def fixed(df, cols): return df.reindex(columns=cols) if len(df) else pd.DataFrame(columns=cols)
alert_parts = []
if len(ED): alert_parts.append(ED[ED["Severity"].isin(["RED", "AMBER", "BLUE"])])
if len(ET): alert_parts.append(ET[ET["Severity"].isin(["RED", "AMBER"])])
AL_COLS = ["State", "District", "Tehsil", "Level", "Event", "Severity", "Start", "End", "Duration d", "Total mm", "Peak", "Intensity", "Lead", "Source prob % (blank = not supplied)", "Forecast retrieved"]
ALERTS = pd.concat(alert_parts, ignore_index=True) if alert_parts else pd.DataFrame(columns=AL_COLS)
if len(ALERTS): ALERTS = ALERTS.assign(_o=ALERTS["Severity"].map(LEVEL_ORDER)).sort_values(["_o", "Start", "State", "District"]).drop(columns="_o").head(40000)
SUMMARY = [["WEATHER MANAGEMENT SUMMARY", RUN_TS], ["Fixed Google Sheet", WEATHER_SHEET_URL], ["Master", master_src], ["Forecast", "Open-Meteo model (NOT IMD); probabilities are source-supplied only"], ["Live evidence", "airport METAR (<=6 h) > IMERG (if fresh) > model analysis (NOT observed); CHIRPS historical only"], ["IMD", IMD_AUTH + " (connector disabled until credentials verified)"],
           ["History", f"month-by-year CHIRPS through {chirps_latest} ({lag_days} d behind); tabs HISTORY MOM DISTRICT / TEHSIL / STATE"], [], ["COVERAGE COUNTS", ""]] + [[k, v] for k, v in COUNTS] + [[], ["TOP 25 STATES BY RISK"]] + vals(STS.head(25)[[c for c in ["State/UT", "Districts", "Tehsils", "Districts RED", "Districts AMBER", "Districts BLUE", "Districts GREEN", "STATE RISK SCORE"] if c in STS.columns]])
TABS = {"WEATHER MANAGEMENT SUMMARY": [[clean(x) for x in r] for r in SUMMARY], "INDIA DISTRICT INTELLIGENCE": vals(fixed(RD, DIST_COLS)), "INDIA TEHSIL INTELLIGENCE": vals(fixed(RT, TEH_COLS)) if len(RT) else [["Tehsil master unavailable"]],
        "APPROACHING WEATHER ALERTS": vals(fixed(ALERTS, AL_COLS)), "HISTORY MOM DISTRICT": vals(MOM_D) if len(MOM_D) else [["no history"]], "HISTORY MOM TEHSIL": vals(MOM_T) if len(MOM_T) else [["no history"]], "HISTORY MOM STATE": vals(MOM_S) if len(MOM_S) else [["no history"]],
        # Excel-only (full detail kept, not written to the Google Sheet):
        "STATE SUMMARY": vals(STS), "TAMIL NADU DISTRICTS": vals(fixed(tn_d, DIST_COLS)) if len(tn_d) else [[""]], "TAMIL NADU TEHSILS": vals(fixed(tn_t, TEH_COLS)) if len(tn_t) else [[""]], "CROP WEATHER EXPOSURE": vals(CROPX) if len(CROPX) else [["none"]], "FARMER ADVISORIES": vals(ADV) if len(ADV) else [["none"]],
        "TRACTOR & SERVICE IMPACT": vals(TRC) if len(TRC) else [["none"]], "LIVE WEATHER OBSERVATIONS": vals(OBS_ST.assign(Retrieved=RUN_TS)) if len(OBS_ST) else [["NOT AVAILABLE"]], "IMD PUBLIC BULLETINS": vals(PUBLIC_ALL) if len(PUBLIC_ALL) else [["none"]],
        "LGD NAME MATCHING AUDIT": vals(pd.DataFrame(audit_rows)) if audit_rows else [["none"]], "DATA QUALITY & VALIDATION": [["Section", "Item", "Status class", "Verdict", "Evidence"]] + REPORT, "EXECUTION LOG": [["Time", "Step", "Status", "Detail"]] + LOG}

# ============================ OUTPUTS: Excel (fixed name) + FIXED Google Sheet ============================
XLSX = f"{OUT_DIR}/India_Weather_Intelligence_LATEST.xlsx"
def write_excel():
    names = {}
    with pd.ExcelWriter(XLSX, engine="openpyxl") as xw:
        for t, v in TABS.items():
            nm = t[:31]; nm = nm if nm not in names else nm[:28] + str(len(names)); names[nm] = 1; w = max(len(r) for r in v) if v else 1
            pd.DataFrame([r + [""] * (w - len(r)) for r in v]).to_excel(xw, sheet_name=nm, header=False, index=False)
excel_ok = False
try: write_excel(); excel_ok = True; shutil.copy(XLSX, f"{DIRS['backups']}/India_Weather_Intelligence_{STAMP}.xlsx")
except Exception as e: print("Excel write error:", type(e).__name__, str(e)[:120])
def retry(fn, *a, **k):
    for i in range(4):
        try: return fn(*a, **k)
        except gspread.exceptions.APIError as e:
            if "429" in str(e) or "quota" in str(e).lower(): time.sleep(65 + random.uniform(0, 10)); continue
            raise
    return fn(*a, **k)
def put(sh, title, v):
    if sh.id == AGRI_SHEET_ID: raise RuntimeError("refusing to write to the Agri Intelligence workbook")
    w = max(len(r) for r in v); v = [r + [""] * (w - len(r)) for r in v]
    try: ws = retry(sh.worksheet, title)
    except gspread.WorksheetNotFound: ws = retry(sh.add_worksheet, title, rows=max(len(v) + 5, 100), cols=max(w + 1, 10))
    if ws.row_count < len(v) + 1 or ws.col_count < w: retry(ws.resize, rows=max(ws.row_count, len(v) + 1), cols=max(ws.col_count, w))
    for i in range(0, len(v), 3000): retry(ws.update, range_name=f"A{i+1}", values=v[i:i + 3000], value_input_option="RAW"); time.sleep(1.5)
    if ws.row_count > len(v) + 1 or ws.col_count > w + 1: retry(ws.resize, rows=len(v) + 1, cols=w + 1)   # trims leftovers only AFTER the new data is written
sheet_ok, sheet_client, SH, tabs_written = False, "none", None, 0
for lab, g in CLIENTS:
    try: SH = retry(g.open_by_key, WEATHER_SHEET_ID); sheet_client = lab; break
    except Exception as e:
        resp = getattr(getattr(e, "__cause__", None), "response", None); log("Open Weather sheet via " + lab, "FAIL", (resp.text[:160].replace("\n", " ") if resp is not None else f"{type(e).__name__}: {str(e)[:120]}"))
if SH is not None:
    for t in SHEET_TABS:
        try: put(SH, t, TABS[t]); tabs_written += 1; print("  Sheet tab written:", t)
        except Exception as e: log("Sheet tab " + t, "FAIL", f"{type(e).__name__}: {str(e)[:100]}")
    sheet_ok = tabs_written == len(SHEET_TABS)
else: print("\nGOOGLE SHEET NOT WRITTEN. Colab: re-run and approve the Google login popup (account must have edit access). GitHub: enable Sheets API (https://console.developers.google.com/apis/api/sheets.googleapis.com/overview?project=251142614090) and share the sheet with", SA_EMAIL)
if HISTORY_WRITE:
    try:      # persistent daily alert history (for trend graphs outside the sheet)
        ah = RD[RD["Alert level (combined)"].isin(["RED", "AMBER"])][["State", "District", "Alert level (combined)", "Hazards", "Top event start", "1st rain event total mm", "Confidence", "PRIORITY SCORE /100"]].assign(run_date=str(today), retrieved=RUN_TS); AH = DIRS["alerts"] + "/alert_history.csv"
        pd.concat([pd.read_csv(AH, dtype=str) if os.path.exists(AH) else pd.DataFrame(), ah.astype(str)], ignore_index=True).drop_duplicates(["run_date", "State", "District", "Top event start"], keep="first").to_csv(AH, index=False)
        pd.DataFrame(REPORT, columns=["Section", "Item", "Class", "Verdict", "Evidence"]).to_csv(f"{DIRS['audit']}/final_report_{STAMP}.csv", index=False); pd.DataFrame(LOG, columns=["Time", "Step", "Status", "Detail"]).to_csv(f"{DIRS['logs']}/execution_log_{STAMP}.csv", index=False)
        if audit_rows: pd.DataFrame(audit_rows).to_csv(f"{DIRS['audit']}/name_matching_audit_{STAMP}.csv", index=False)
    except Exception as e: print("history save error:", str(e)[:100])
if IN_COLAB and excel_ok:
    try: from google.colab import files as _f; _f.download(XLSX)
    except Exception: pass

# ============================ PRINT ONLY THE FOUR REQUESTED TABS ============================
def show(title, df, n=30):
    print("\n" + "=" * 8 + " " + title + " " + "=" * 8); print(df.head(n).to_string(index=False) if len(df) else "(none)")
print("\n" + "=" * 8 + " WEATHER MANAGEMENT SUMMARY " + "=" * 8)
for row in TABS["WEATHER MANAGEMENT SUMMARY"][:36]: print("  " + " | ".join(str(x) for x in row if str(x) != ""))
dcols_show = ["State", "District", "Alert level (combined)", "Rain probability score % (source-supplied)", "Top event lead"] + RAIN_D + ["Cumulative 7-day mm", "Confidence", "PRIORITY SCORE /100"]
show("INDIA DISTRICT INTELLIGENCE (top 30 by priority score)", top(RD, 30, dcols_show))
show("INDIA TEHSIL INTELLIGENCE (top 30 by priority score)", top(RT, 30, ["State", "District", "Tehsil", "Alert level (combined)", "Rain probability score % (source-supplied)"] + RAIN_D + ["Coverage status"]) if len(RT) else pd.DataFrame())
show("APPROACHING WEATHER ALERTS", ALERTS[["State", "District", "Tehsil", "Level", "Event", "Severity", "Start", "Total mm", "Lead"]] if len(ALERTS) else pd.DataFrame())
n_fail = sum(1 for r in REPORT if r[3] == "FAIL"); n_chk = sum(1 for r in REPORT if r[3] == "CHECK")
print(f"\nSTATUS: sheet tabs written {tabs_written}/{len(SHEET_TABS)} (via {sheet_client}) | Excel {'OK' if excel_ok else 'FAILED'}: {XLSX} | report items FAIL {n_fail}, CHECK {n_chk} (see Excel tab 'DATA QUALITY & VALIDATION' and {DIRS['audit']})")
print("\nWEATHER INTELLIGENCE GOOGLE SHEET:", WEATHER_SHEET_URL, "-", "UPDATED" if sheet_ok else "NOT FULLY UPDATED (see message above)")