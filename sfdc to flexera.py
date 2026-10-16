"""
sfdc_to_flexera.py  -  Salesforce users -> CSV -> Flexera SaaS user custom fields, in one run.

  1. EXTRACT  Salesforce OAuth (refresh token) -> SOQL on User (all pages) -> CSV in the output folder
              Columns: Id, Name, Username, LastLoginDate, ProfileName, UserLicenseName, Email,
                       IsActive, UserPermissionsSupportUser
  2. UPDATE   For each user, set Flexera SaaS custom fields SFDC_<column>:
                  IsActive                    -> SFDC_IsActive
                  UserPermissionsSupportUser  -> SFDC_UserPermissionsSupportUser
              Matched by email (default), username or Flexera user ID.

OTHER CUSTOM FIELDS ARE NEVER REMOVED
  Flexera's write endpoint replaces a user's whole custom-field list. For every user the script
  first reads ALL current custom fields (GET /users/{id}) - with or without values, from any
  source - changes only the SFDC_ fields and sends everything back. If a user's current fields
  cannot be read, that user is skipped, never written blind.

.env (same folder as this script)
  # Salesforce
  SFDC_CLIENT_ID=...
  SFDC_CLIENT_SECRET=...
  SFDC_REFRESH_TOKEN=...
  SFDC_TOKEN_URL=https://login.salesforce.com/services/oauth2/token
  SFDC_INSTANCE_URL=https://ontime.my.salesforce.com
  SFDC_API_VERSION=v66.0
  SFDC_SOQL_WHERE=                       # optional, e.g.  IsActive = true

  # Flexera (same keys as the other Flexera scripts)
  FLEXERA_TOKEN_URL=...  FLEXERA_CLIENT_ID=...  FLEXERA_CLIENT_SECRET=...
  FLEXERA_API_BASE=https://snowatlas-westeurope.flexera.eu

  # What to do
  RUN_STEPS=extract,update               # extract | update | extract,update
  SFDC_INPUT_FILE=                       # for RUN_STEPS=update only: CSV/XLSX to use (default: latest extract)
  SFDC_MATCH_BY=email                    # email | username | flexera_id
  SFDC_COLUMNS=IsActive,UserPermissionsSupportUser
  SFDC_FIELD_PREFIX=SFDC_
  SFDC_DUP_RULE=skip                     # same email on several SFDC users with different values:
                                         #   skip | first | last | any_true
  SFDC_LIMIT=                            # optional: only the first N users (testing)
  SFDC_DRY_RUN=yes                       # yes = write nothing to Flexera. No = write.

  # Output + log
  OUTPUT_DIR=output                      # default: output (relative to this script, or an absolute path)
  LOG_FILE=                              # default: <OUTPUT_DIR>/sfdc_to_flexera.log
  LOG_MAX_MB=5                           # when the log passes this size it is renamed
                                         #   sfdc_to_flexera_<timestamp>.log and a fresh log is started

Before the first update: create in Flexera  SaaS settings > Custom fields > Add (Category: User, Type: Text)
  SFDC_IsActive   and   SFDC_UserPermissionsSupportUser

Output files (in OUTPUT_DIR)
  salesforce_users_<timestamp>.csv          full Salesforce extract of this run
  salesforce_users_latest.csv               copy of the newest extract
  sfdc_custom_fields_results_<timestamp>.csv one row per user: status, values, other fields kept

Requires:  python -m pip install requests python-dotenv     (+ openpyxl only to read an .xlsx input)
"""
import csv
import json
import os
import shutil
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

try:
    import requests
except ImportError:
    sys.exit("requests not installed ->  python -m pip install requests")
try:
    from dotenv import load_dotenv
except ImportError:
    sys.exit("python-dotenv not installed ->  python -m pip install python-dotenv")

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

# Older Flexera .env names are accepted too
for _new, _old in (("FLEXERA_TOKEN_URL", "TOKEN_URL"), ("FLEXERA_CLIENT_ID", "CLIENT_ID"),
                   ("FLEXERA_CLIENT_SECRET", "CLIENT_SECRET"), ("FLEXERA_API_BASE", "API_BASE_URL")):
    if not (os.getenv(_new) or "").strip() and (os.getenv(_old) or "").strip():
        os.environ[_new] = os.environ[_old]


def env(key, default=""):
    return (os.getenv(key) or default).strip()


def required(key):
    val = env(key)
    if not val or val.startswith("PASTE_") or "<" in val or set(val) == {"*"}:
        sys.exit(f"[config] {key} is missing in .env")
    return val


# ── Settings ────────────────────────────────────────────────────────
STEPS = {s.strip().lower() for s in env("RUN_STEPS", "extract,update").split(",") if s.strip()}
if not STEPS or not STEPS <= {"extract", "update"}:
    sys.exit("[config] RUN_STEPS must be extract, update, or extract,update")

SFDC_TOKEN_URL = env("SFDC_TOKEN_URL", "https://login.salesforce.com/services/oauth2/token")
SFDC_INSTANCE_URL = env("SFDC_INSTANCE_URL", "https://ontime.my.salesforce.com").rstrip("/")
SFDC_API_VERSION = env("SFDC_API_VERSION", "v66.0")
SFDC_SOQL_WHERE = env("SFDC_SOQL_WHERE")
SOQL_FIELDS = ["Id", "Name", "Username", "LastLoginDate", "Profile.Name", "Profile.UserLicense.Name",
               "Email", "IsActive", "UserPermissionsSupportUser"]
# CSV header for each SOQL field
CSV_COLUMNS = {"Id": "Id", "Name": "Name", "Username": "Username", "LastLoginDate": "LastLoginDate",
               "Profile.Name": "ProfileName", "Profile.UserLicense.Name": "UserLicenseName",
               "Email": "Email", "IsActive": "IsActive",
               "UserPermissionsSupportUser": "UserPermissionsSupportUser"}

FIELDS_PATH = env("CF_FIELDS_PATH", "/api/saas/consolidated-view/v1/custom-fields")
VALUES_PATH = env("CF_VALUES_PATH", FIELDS_PATH.rstrip("/") + "/values")
USERS_PATH = env("CF_USERS_PATH", "/api/saas/consolidated-view/v1/users")

MATCH_BY = env("SFDC_MATCH_BY", "email").lower()
if MATCH_BY in ("id", "flexeraid", "flexera-id"):
    MATCH_BY = "flexera_id"
if MATCH_BY not in ("email", "username", "flexera_id"):
    sys.exit(f"[config] SFDC_MATCH_BY must be email, username or flexera_id (got '{MATCH_BY}')")
KEY_COLUMN = env("SFDC_KEY_COLUMN", {"email": "Email", "username": "Username", "flexera_id": "FlexeraId"}[MATCH_BY])
SOURCE_COLUMNS = [c.strip() for c in env("SFDC_COLUMNS", "IsActive,UserPermissionsSupportUser").split(",") if c.strip()]
PREFIX = env("SFDC_FIELD_PREFIX", "SFDC_")
DUP_RULE = env("SFDC_DUP_RULE", "skip").lower()
if DUP_RULE not in ("skip", "first", "last", "any_true"):
    sys.exit(f"[config] SFDC_DUP_RULE must be skip, first, last or any_true (got '{DUP_RULE}')")
LIMIT = int(env("SFDC_LIMIT") or 0)
DRY_RUN = env("SFDC_DRY_RUN", "yes").lower() != "no"

_out = Path(env("OUTPUT_DIR", "output"))
OUTPUT_DIR = _out if _out.is_absolute() else BASE_DIR / _out
_log = Path(env("LOG_FILE") or OUTPUT_DIR / "sfdc_to_flexera.log")
LOG_FILE = _log if _log.is_absolute() else BASE_DIR / _log
LOG_MAX_BYTES = int(float(env("LOG_MAX_MB", "5")) * 1024 * 1024)

USERS_PAGE_SIZE = 1000
FIELDS_PAGE_SIZE = 100
SLEEP_BETWEEN_WRITES = 0.2
MAX_RETRIES = 5
TIMEOUT = 60
TRUE_WORDS = {"true", "yes", "y", "1"}
FALSE_WORDS = {"false", "no", "n", "0"}
STAMP = f"{datetime.now():%Y%m%d_%H%M%S}"


# ── Log: console + file, rotated when it passes LOG_MAX_MB ──────────
class RotatingLog:
    def __init__(self, path, max_bytes):
        self.path, self.max_bytes = path, max_bytes
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.f = None
        self._open()

    def _open(self):
        if self.path.exists() and self.path.stat().st_size >= self.max_bytes:
            self._rotate_file()
        self.f = open(self.path, "a", encoding="utf-8")

    def _rotate_file(self):
        rotated = self.path.with_name(f"{self.path.stem}_{datetime.now():%Y%m%d_%H%M%S}{self.path.suffix}")
        n = 2
        while rotated.exists():
            rotated = self.path.with_name(f"{self.path.stem}_{datetime.now():%Y%m%d_%H%M%S}_{n}{self.path.suffix}")
            n += 1
        os.replace(self.path, rotated)

    def write_line(self, line):
        self.f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S}  {line}\n")
        self.f.flush()
        if self.f.tell() >= self.max_bytes:          # rotate mid-run too
            self.f.close()
            self._rotate_file()
            self.f = open(self.path, "a", encoding="utf-8")
            self.f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S}  (log continued after rotation)\n")
            self.f.flush()


class _Tee:
    """Console + log. Lines ending in \\r (progress counters) go to the console only."""
    def __init__(self, stream, log):
        self.stream, self.log, self.buf = stream, log, ""

    def write(self, text):
        self.stream.write(text)
        self.buf += text
        while True:
            cuts = [i for i in (self.buf.find("\n"), self.buf.find("\r")) if i >= 0]
            if not cuts:
                break
            cut = min(cuts)
            line, ending, self.buf = self.buf[:cut], self.buf[cut], self.buf[cut + 1:]
            if ending == "\n" and line.strip():
                self.log.write_line(line)

    def flush(self):
        self.stream.flush()


def start_log():
    log = RotatingLog(LOG_FILE, LOG_MAX_BYTES)
    log.write_line("=" * 70)
    log.write_line("RUN START")
    sys.stdout = _Tee(sys.__stdout__, log)
    sys.stderr = _Tee(sys.__stderr__, log)


def norm_value(v):
    """Salesforce/Excel TRUE/FALSE (bool or text) -> 'TRUE'/'FALSE'; blank -> None; else text."""
    if v is None:
        return None
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    s = str(v).strip()
    if not s:
        return None
    if s.lower() in TRUE_WORDS:
        return "TRUE"
    if s.lower() in FALSE_WORDS:
        return "FALSE"
    return s


def unique_path(path):
    n, p = 2, path
    while p.exists():
        p = path.with_name(f"{path.stem}_{n}{path.suffix}")
        n += 1
    return p


# ══════════════════════════ STEP 1: SALESFORCE EXTRACT ══════════════
def sfdc_token():
    resp = requests.post(SFDC_TOKEN_URL, timeout=30, data={
        "grant_type": "refresh_token",
        "client_id": required("SFDC_CLIENT_ID"),
        "client_secret": required("SFDC_CLIENT_SECRET"),
        "refresh_token": required("SFDC_REFRESH_TOKEN"),
    })
    if not resp.ok:
        # Salesforce errors are safe to show (no token in them); never print a successful response
        sys.exit(f"[sfdc] Authentication failed: HTTP {resp.status_code} {resp.text[:300]}")
    data = resp.json()
    print("[sfdc] Authenticated")
    return data["access_token"]


def flatten(record):
    """{'Profile': {'Name': 'X', 'UserLicense': {'Name': 'Y'}}} -> {'Profile.Name': 'X', 'Profile.UserLicense.Name': 'Y'}"""
    out = {}

    def walk(prefix, obj):
        for k, v in obj.items():
            if k == "attributes":
                continue
            key = f"{prefix}.{k}" if prefix else k
            if isinstance(v, dict):
                walk(key, v)
            else:
                out[key] = v
    walk("", record)
    return out


def extract_salesforce():
    soql = f"SELECT {', '.join(SOQL_FIELDS)} FROM User"
    if SFDC_SOQL_WHERE:
        soql += f" WHERE {SFDC_SOQL_WHERE}"
    print(f"[sfdc] Instance: {SFDC_INSTANCE_URL} | API {SFDC_API_VERSION}")
    print(f"[sfdc] SOQL: {soql}")

    token = sfdc_token()
    url = f"{SFDC_INSTANCE_URL}/services/data/{SFDC_API_VERSION}/query"
    params = {"q": soql}
    records, refreshed = [], False
    while True:
        resp = None
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                resp = requests.get(url, headers={"Authorization": f"Bearer {token}"},
                                    params=params, timeout=TIMEOUT)
            except requests.RequestException as e:
                if attempt == MAX_RETRIES:
                    raise
                print(f"[sfdc] network error ({e.__class__.__name__}) - retrying")
                time.sleep(2 ** attempt)
                continue
            if resp.status_code in (500, 502, 503, 504) and attempt < MAX_RETRIES:
                time.sleep(2 ** attempt)
                continue
            break
        if resp.status_code == 401 and not refreshed:
            print("[sfdc] session expired - re-authenticating")
            token, refreshed = sfdc_token(), True
            continue
        if not resp.ok:
            sys.exit(f"[sfdc] Query failed: HTTP {resp.status_code} {resp.text[:300]}")
        data = resp.json()
        batch = data.get("records", [])
        records.extend(flatten(r) for r in batch)
        print(f"[sfdc] Retrieved {len(records):,} / {data.get('totalSize', '?')} users", end="\r")
        if data.get("done", True):
            break
        url, params = SFDC_INSTANCE_URL + data["nextRecordsUrl"], None
    print()
    if not records:
        sys.exit("[sfdc] Salesforce returned 0 users - stopping, nothing written to Flexera.")

    rows = []
    for r in records:
        row = {}
        for f in SOQL_FIELDS:
            v = r.get(f)
            row[CSV_COLUMNS[f]] = ("TRUE" if v else "FALSE") if isinstance(v, bool) else ("" if v is None else v)
        rows.append(row)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    out = unique_path(OUTPUT_DIR / f"salesforce_users_{STAMP}.csv")
    with open(out, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(CSV_COLUMNS.values()))
        w.writeheader()
        w.writerows(rows)
    latest = OUTPUT_DIR / "salesforce_users_latest.csv"
    try:
        shutil.copyfile(out, latest)
    except PermissionError:
        print(f"[sfdc] Could not refresh {latest.name} (open in Excel?) - the timestamped file is complete")
    active = sum(1 for r in rows if r.get("IsActive") == "TRUE")
    print(f"[sfdc] {len(rows):,} users ({active:,} active) -> {out}")
    return rows


# ══════════════════════════ STEP 2: FLEXERA UPDATE ══════════════════
class Flexera:
    def __init__(self):
        explicit = env("FLEXERA_API_BASE").rstrip("/")
        src = explicit if explicit and "<" not in explicit else required("FLEXERA_UPLOAD_URL")
        u = urlparse(src)
        if not u.scheme or not u.netloc:
            sys.exit("[config] Set FLEXERA_API_BASE=https://snowatlas-westeurope.flexera.eu in .env")
        self.base = f"{u.scheme}://{u.netloc}"
        self.s = requests.Session()
        self.s.headers["Accept"] = "application/json"
        self.content_type = None
        self.refresh_token()
        print(f"[flexera] Authenticated ({self.base})")

    def refresh_token(self):
        resp = requests.post(required("FLEXERA_TOKEN_URL"), timeout=30,
                             headers={"Content-Type": "application/x-www-form-urlencoded"},
                             data={"grant_type": "client_credentials",
                                   "client_id": required("FLEXERA_CLIENT_ID"),
                                   "client_secret": required("FLEXERA_CLIENT_SECRET")})
        if resp.status_code in (400, 401):
            sys.exit("[flexera] Token refused. Check FLEXERA_CLIENT_ID / FLEXERA_CLIENT_SECRET.")
        if not resp.ok:
            sys.exit(f"[flexera] Token request failed: HTTP {resp.status_code} {resp.text[:300]}")
        self.s.headers["Authorization"] = f"Bearer {resp.json()['access_token']}"

    def call(self, method, path, **kw):
        refreshed, resp = False, None
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                resp = self.s.request(method, self.base + path, timeout=TIMEOUT, **kw)
            except requests.RequestException as e:
                if attempt == MAX_RETRIES:
                    raise
                print(f"          network error ({e.__class__.__name__}) - retrying")
                time.sleep(2 ** attempt)
                continue
            if resp.status_code == 401 and not refreshed:
                print("          Flexera token expired - getting a new one")
                self.refresh_token()
                refreshed = True
                continue
            if resp.status_code in (429, 500, 502, 503, 504) and attempt < MAX_RETRIES:
                time.sleep(int(resp.headers.get("Retry-After") or 2 ** attempt))
                continue
            return resp
        return resp

    def paged(self, path, page_size, label):
        items, page = [], 1
        while True:
            resp = self.call("GET", path, params={"page_number": page, "page_size": page_size})
            if not resp.ok:
                sys.exit(f"[flexera] Get {label} failed: HTTP {resp.status_code} {resp.text[:300]}")
            body = resp.json()
            items.extend(body.get("items", []))
            total = (body.get("pagination") or {}).get("total_pages") or 1
            print(f"[flexera] {label}: page {page}/{total}", end="\r")
            if page >= total:
                break
            page += 1
        print()
        return items

    def user_custom_fields(self, user_id):
        """All current custom fields of one user; 'NOT_FOUND'; or None if unreadable."""
        resp = self.call("GET", f"{USERS_PATH.rstrip('/')}/{user_id}")
        if resp.status_code == 404:
            return "NOT_FOUND"
        if not resp.ok:
            return None
        try:
            cf = resp.json().get("customFields")
        except ValueError:
            return None
        return cf if isinstance(cf, list) else None

    def write_custom_fields(self, user_id, fields):
        body = json.dumps({"customFields": fields, "id": user_id, "saasCategory": "User"}).encode("utf-8")
        resp = None
        for ct in ([self.content_type] if self.content_type else ["application/json", "application/merge-patch+json"]):
            resp = self.call("PATCH", VALUES_PATH, data=body, headers={"Content-Type": ct})
            if resp.status_code != 400:
                if resp.ok:
                    self.content_type = ct
                return resp, body
        return resp, body


def read_input_file():
    """RUN_STEPS=update only: rows from SFDC_INPUT_FILE (csv/xlsx) or the latest extract."""
    name = env("SFDC_INPUT_FILE")
    path = Path(name) if name else OUTPUT_DIR / "salesforce_users_latest.csv"
    if not path.is_absolute():
        path = BASE_DIR / path if name else path
    if not path.exists():
        sys.exit(f"[input] {path} not found. Run with RUN_STEPS=extract,update or set SFDC_INPUT_FILE.")
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        try:
            import openpyxl
        except ImportError:
            sys.exit("[input] Reading .xlsx needs openpyxl ->  python -m pip install openpyxl")
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb[env("SFDC_SHEET")] if env("SFDC_SHEET") else wb.worksheets[0]
        it = ws.iter_rows(values_only=True)
        header = [str(h).strip() if h is not None else "" for h in next(it, [])]
        rows = [dict(zip(header, r)) for r in it if any(c not in (None, "") for c in r)]
        wb.close()
    else:
        with open(path, newline="", encoding="utf-8-sig") as f:
            rows = [{(k or "").strip(): v for k, v in r.items()} for r in csv.DictReader(f)]
    print(f"[input] {len(rows):,} rows from {path}")
    return rows


def find_col(header, wanted):
    return next((h for h in header if h and h.lower() == wanted.lower()), None)


def prepare_updates(rows):
    """rows -> {key: {SFDC_field: value}}, conflicts, list of SFDC field names."""
    header = list(rows[0].keys()) if rows else []
    key_col = find_col(header, KEY_COLUMN)
    if not key_col:
        sys.exit(f"[update] Column '{KEY_COLUMN}' not found for SFDC_MATCH_BY={MATCH_BY}. "
                 f"Columns: {', '.join(h for h in header if h)}")
    src = {}
    for c in SOURCE_COLUMNS:
        h = find_col(header, c)
        if not h:
            sys.exit(f"[update] Column '{c}' (SFDC_COLUMNS) not found. Columns: {', '.join(h for h in header if h)}")
        src[h] = PREFIX + c
    print(f"[update] Match on '{key_col}' | " + ", ".join(f"{h} -> {f}" for h, f in src.items()))

    grouped, blank = {}, 0
    for r in rows:
        raw = r.get(key_col)
        key = "" if raw is None else str(raw).strip()
        if MATCH_BY != "flexera_id":
            key = key.lower()
        if not key or (MATCH_BY == "email" and "@" not in key):
            blank += 1
            continue
        grouped.setdefault(key, []).append({f: norm_value(r.get(h)) for h, f in src.items()})

    users, conflicts, combined = {}, {}, 0
    for key, vals in grouped.items():
        if len(vals) == 1:
            users[key] = vals[0]
            continue
        merged, conflict = {}, False
        for f in src.values():
            seen = [v[f] for v in vals if v[f] is not None]
            distinct = list(dict.fromkeys(seen))
            if len(distinct) <= 1:
                merged[f] = distinct[0] if distinct else None
            elif DUP_RULE == "first":
                merged[f] = seen[0]
            elif DUP_RULE == "last":
                merged[f] = seen[-1]
            elif DUP_RULE == "any_true":
                merged[f] = "TRUE" if "TRUE" in distinct else seen[-1]
            else:
                conflict = True
        if conflict:
            conflicts[key] = vals
        else:
            users[key] = merged
            combined += 1
    if blank:
        print(f"[update] Skipped {blank:,} Salesforce user(s) with a blank/invalid {key_col}")
    if combined:
        print(f"[update] {combined:,} {key_col}(s) used by several Salesforce users - combined (SFDC_DUP_RULE={DUP_RULE})")
    if conflicts:
        print(f"[update] {len(conflicts):,} {key_col}(s) used by several Salesforce users with DIFFERENT values "
              f"- not updated (set SFDC_DUP_RULE=first|last|any_true to choose)")
    if LIMIT:
        users = dict(list(users.items())[:LIMIT])
        print(f"[update] SFDC_LIMIT={LIMIT}: processing the first {len(users)} user(s) only")
    return users, conflicts, list(src.values())


def merge(existing, updates):
    """Keep every existing custom field; only set the fields in `updates` (None = leave as is)."""
    upd = {k.lower(): (k, v) for k, v in updates.items() if v is not None}
    out, done = [], set()
    for f in existing:
        if not isinstance(f, dict) or not f.get("name"):
            continue
        k = f["name"].strip().lower()
        if k in upd:
            out.append({"name": upd[k][0], "value": upd[k][1]})
            done.add(k)
        else:
            out.append({"name": f["name"], "value": f.get("value")})
    for k, (name, v) in upd.items():
        if k not in done:
            out.append({"name": name, "value": v})
    return out


def update_flexera(rows):
    print(f"[update] Match by {MATCH_BY} | {'DRY RUN (nothing written)' if DRY_RUN else 'LIVE - values will be written'}")
    users, conflicts, sfdc_fields = prepare_updates(rows)
    if not users and not conflicts:
        sys.exit("[update] No users to process")

    fx = Flexera()
    defs = fx.paged(FIELDS_PATH, FIELDS_PAGE_SIZE, "custom fields")
    user_defs = {(d.get("fieldName") or "").strip().lower(): d["fieldName"] for d in defs
                 if d.get("fieldName") and "user" in str(d.get("saasCategory") or "").lower()}
    missing = [f for f in sfdc_fields if f.lower() not in user_defs]
    if missing:
        sys.exit(f"[flexera] Create these SaaS User custom fields first (SaaS settings > Custom fields, "
                 f"Category: User, Type: Text): {', '.join(missing)}")
    exact = {f: user_defs[f.lower()] for f in sfdc_fields}
    print(f"[flexera] Custom fields found: {', '.join(exact.values())}")

    if MATCH_BY == "flexera_id":
        lookup = None
    else:
        all_users = fx.paged(USERS_PATH, USERS_PAGE_SIZE, "SaaS users")
        lookup = {}
        for u in all_users:
            k = (u.get(MATCH_BY) or "").strip().lower()
            if k and u.get("id"):
                lookup.setdefault(k, []).append(u["id"])
        found = sum(1 for k in users if k in lookup)
        print(f"[flexera] {len(all_users):,} SaaS users loaded")
        print(f"[match] {found:,} of {len(users):,} Salesforce {MATCH_BY}s found in Flexera "
              f"({(100 * found / len(users)) if users else 0:.1f}%)")

    out_path = unique_path(OUTPUT_DIR / f"sfdc_custom_fields_results_{STAMP}.csv")
    out_f = open(out_path, "w", newline="", encoding="utf-8-sig")
    w = csv.writer(out_f)
    w.writerow([MATCH_BY, "flexera_user_id", "status"] + list(exact.values()) + ["other_fields_kept", "detail"])
    summary = {}

    def record(key, uid, status, vals, kept="", detail=""):
        s = status.split()[0]
        summary[s] = summary.get(s, 0) + 1
        w.writerow([key, uid, status] + [vals.get(f, "") or "" for f in sfdc_fields] + [kept, detail])
        out_f.flush()

    for key, vals_list in conflicts.items():
        record(key, "", "CONFLICT", {}, "", "different values: " + " | ".join(json.dumps(v) for v in vals_list))

    shown, total, first_err = 0, len(users), False
    try:
        for i, (key, vals) in enumerate(users.items(), 1):
            updates = {exact[f]: v for f, v in vals.items()}
            if all(v is None for v in updates.values()):
                record(key, "", "SKIPPED", vals, "", "all SFDC values blank")
                continue
            ids = [key] if lookup is None else lookup.get(key)
            if not ids:
                record(key, "", "NOT_FOUND", vals, "", f"{MATCH_BY} not in Flexera SaaS users")
                continue
            for uid in ids:
                existing = fx.user_custom_fields(uid)
                if existing == "NOT_FOUND":
                    record(key, uid, "NOT_FOUND", vals, "", "Flexera user ID not found")
                    continue
                if existing is None:
                    record(key, uid, "SKIPPED", vals, "", "could not read current custom fields - not written")
                    continue
                cur = {(f.get("name") or "").strip().lower(): f.get("value") for f in existing if isinstance(f, dict)}
                tkeys = {k.lower() for k in updates}
                kept = "; ".join(f"{f.get('name')}={f.get('value')}" for f in existing if isinstance(f, dict)
                                 and (f.get("name") or "").strip().lower() not in tkeys)
                was = ", ".join(f"{n} was {cur.get(n.lower())!r}" for n in updates)
                if all(v is None or str(cur.get(n.lower())) == v for n, v in updates.items()):
                    record(key, uid, "UNCHANGED", vals, kept, "already up to date")
                    continue
                payload = merge(existing, updates)
                if DRY_RUN:
                    record(key, uid, "DRY_RUN", vals, kept, was)
                    if shown < 5:
                        print(f"          {key}: would send {json.dumps(payload)}")
                        shown += 1
                    continue
                resp, sent = fx.write_custom_fields(uid, payload)
                if resp.ok:
                    record(key, uid, "OK", vals, kept, was)
                else:
                    record(key, uid, f"FAILED {resp.status_code}", vals, kept, resp.text[:300])
                    if not first_err:
                        print(f"[flexera] First failure for {key}: HTTP {resp.status_code} {resp.text[:300]}")
                        print(f"[flexera] Body sent: {sent.decode('utf-8')}")
                        first_err = True
                    if resp.status_code == 403:
                        sys.exit("[flexera] 403 - app registration not allowed to write custom fields. Stopping.")
                time.sleep(SLEEP_BETWEEN_WRITES)
            if i % 25 == 0 or i == total:
                print(f"[progress] {i:,}/{total:,} users", end="\r")
        print()
    finally:
        out_f.close()
        print(f"[update] {summary} -> {out_path}")


# ══════════════════════════ MAIN ══════════════════════════════════
def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    print(f"[config] Steps: {', '.join(s for s in ('extract', 'update') if s in STEPS)} | "
          f"output: {OUTPUT_DIR} | log: {LOG_FILE} (rotates at {LOG_MAX_BYTES / 1024 / 1024:g} MB)")
    rows = extract_salesforce() if "extract" in STEPS else read_input_file()
    if "update" in STEPS:
        update_flexera(rows)


if __name__ == "__main__":
    start_log()
    try:
        main()
        print("[run] Finished OK")
    except SystemExit as e:
        if e.code not in (None, 0):
            if not isinstance(e.code, int):
                print(e.code, file=sys.stderr)
            print("[run] Stopped", file=sys.stderr)
            sys.exit(1)
        print("[run] Finished OK")
    except KeyboardInterrupt:
        print("\n[run] Interrupted - partial results are in the output folder; re-running is safe "
              "(users already updated show UNCHANGED)", file=sys.stderr)
        sys.exit(1)
    except Exception:
        traceback.print_exc()
        print("[run] Crashed", file=sys.stderr)
        sys.exit(1)
