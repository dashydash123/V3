"""
flexera_custom_field_test.py  -  set a SaaS user custom field in Flexera (Snow Atlas) SaaS, by email.

Test version: sets CF_FIELD_NAME (default "abcde") = CF_FIELD_VALUE (default "yes")
for every email in CF_EMAILS_CSV (one column: email).

Reuses the SAME .env as ado_users_to_flexera.py:
  FLEXERA_TOKEN_URL, FLEXERA_CLIENT_ID, FLEXERA_CLIENT_SECRET   -> token (client credentials)
  FLEXERA_UPLOAD_URL                                            -> API host is taken from this
  (or set FLEXERA_API_BASE=https://<host> to override)

Extra .env keys for this script (all optional):
  CF_FIELD_NAME=abcde
  CF_FIELD_VALUE=yes
  CF_FIELDS=abcde=yes;Zendesk Role=yes   # several fields at once (overrides the two above)
  CF_EMAILS_CSV=test_emails.csv
  CF_DRY_RUN=yes        # yes = show what would change, write nothing. Set to no to write.

Steps
  1. GET   /api/saas/consolidated-view/v2/custom-fields         -> confirm the field exists
  2. GET   /api/saas/consolidated-view/v1/users                 -> email -> SaaS user ID map
  3. PATCH /api/saas/consolidated-view/v2/custom-fields/values  -> one call per user
  -> output/custom_field_results_<timestamp>.csv
  -> output/flexera_custom_field_test.log (appended every run; secrets never printed)

Permission: the app registration needs SaaS consolidation write (saas.consolidation.view.crud).
The ADO registration may only have the data-import permission -> expect HTTP 403 if so.

Requires:  python -m pip install python-dotenv requests
"""
import csv
import json
import os
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

# Routes can be overridden in .env if your host uses a different version (see flexera_route_probe.py)
# v1 confirmed on snowatlas-westeurope.flexera.eu (v2 in the public spec returns 404 there)
CUSTOM_FIELDS_PATH = (os.getenv("CF_FIELDS_PATH") or "").strip() or "/api/saas/consolidated-view/v1/custom-fields"
CF_VALUES_PATH = (os.getenv("CF_VALUES_PATH") or "").strip() or CUSTOM_FIELDS_PATH.rstrip("/") + "/values"
USERS_PATH = (os.getenv("CF_USERS_PATH") or "").strip() or "/api/saas/consolidated-view/v1/users"

# Accept the older .env names too (API_BASE_URL, TOKEN_URL, CLIENT_ID, CLIENT_SECRET)
for new, old in (("FLEXERA_TOKEN_URL", "TOKEN_URL"), ("FLEXERA_CLIENT_ID", "CLIENT_ID"),
                 ("FLEXERA_CLIENT_SECRET", "CLIENT_SECRET"), ("FLEXERA_API_BASE", "API_BASE_URL")):
    if not (os.getenv(new) or "").strip() and (os.getenv(old) or "").strip():
        os.environ[new] = os.environ[old]
USERS_PAGE_SIZE = 1000   # API max
CF_PAGE_SIZE = 100
SLEEP_BETWEEN_PATCHES = 0.2
MAX_RETRIES = 3
TIMEOUT = 60


# ── Log: everything printed also goes to the log file ──
class _Tee:
    def __init__(self, stream, log):
        self.stream, self.log = stream, log

    def write(self, text):
        self.stream.write(text)
        for line in text.splitlines():
            if line.strip():
                self.log.write(f"{datetime.now():%Y-%m-%d %H:%M:%S}  {line}\n")
        self.log.flush()

    def flush(self):
        self.stream.flush()


def out_dir():
    d = BASE_DIR / (os.getenv("OUTPUT_DIR") or "output")
    d.mkdir(parents=True, exist_ok=True)
    return d


def start_log():
    log = open(out_dir() / "flexera_custom_field_test.log", "a", encoding="utf-8")
    log.write("\n" + "=" * 70 + f"\n{datetime.now():%Y-%m-%d %H:%M:%S}  RUN START\n")
    log.flush()
    sys.stdout = _Tee(sys.__stdout__, log)
    sys.stderr = _Tee(sys.__stderr__, log)


def cfg(key, default=None):
    val = (os.getenv(key) or default or "").strip()
    if not val or val.startswith("PASTE_"):
        sys.exit(f"[config] {key} is missing in .env")
    return val


def api_base():
    explicit = (os.getenv("FLEXERA_API_BASE") or "").strip().rstrip("/")
    if explicit and "<" not in explicit:
        u = urlparse(explicit)
        return f"{u.scheme}://{u.netloc}" if u.scheme and u.netloc else explicit
    u = urlparse(cfg("FLEXERA_UPLOAD_URL"))
    if not u.scheme or not u.netloc:
        sys.exit("[config] Could not read a host from FLEXERA_UPLOAD_URL; set FLEXERA_API_BASE in .env")
    return f"{u.scheme}://{u.netloc}"


FIELD_NAME = (os.getenv("CF_FIELD_NAME") or "abcde").strip()
FIELD_VALUE = (os.getenv("CF_FIELD_VALUE") or "yes").strip()


def _parse_targets():
    """CF_FIELDS=abcde=yes;Zendesk Role=yes  -> {'abcde': 'yes', 'Zendesk Role': 'yes'}
    Falls back to CF_FIELD_NAME / CF_FIELD_VALUE when CF_FIELDS is not set."""
    raw = (os.getenv("CF_FIELDS") or "").strip().strip('"').strip("'")
    if not raw:
        return {FIELD_NAME: FIELD_VALUE}
    targets = {}
    for part in raw.split(";"):
        if not part.strip():
            continue
        if "=" not in part:
            sys.exit(f"[config] CF_FIELDS entry '{part}' needs the form name=value")
        name, value = part.split("=", 1)
        targets[name.strip()] = value.strip()
    return targets


TARGETS = _parse_targets()   # field name -> value to set (names fixed to exact case later)
EMAILS_CSV = os.getenv("CF_EMAILS_CSV") or "test_emails.csv"
DRY_RUN = (os.getenv("CF_DRY_RUN") or "yes").strip().lower() != "no"
# Single-user test: patch only this SaaS user ID (skips the email lookup)
TEST_USER_ID = (os.getenv("CF_TEST_USER_ID") or "").strip()
# Content types to try for the PATCH body, in order (first non-400 wins)
_ct = (os.getenv("CF_CONTENT_TYPE") or "").strip()
CONTENT_TYPES = [_ct] if _ct else ["application/json", "application/merge-patch+json"]


# ── Flexera API ─────────────────────────────────────────────────────
def get_token():
    resp = requests.post(
        cfg("FLEXERA_TOKEN_URL"),
        data={"grant_type": "client_credentials",
              "client_id": cfg("FLEXERA_CLIENT_ID"),
              "client_secret": cfg("FLEXERA_CLIENT_SECRET")},
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        timeout=30,
    )
    if resp.status_code in (400, 401):
        sys.exit("[flexera] Token refused. Check FLEXERA_CLIENT_ID / FLEXERA_CLIENT_SECRET.")
    if not resp.ok:
        sys.exit(f"[flexera] Token request failed: HTTP {resp.status_code} {resp.text[:500]}")
    print("[flexera] Access token obtained")
    return resp.json()["access_token"]


def call(session, method, url, **kwargs):
    """Retry on throttling / transient server errors."""
    for attempt in range(1, MAX_RETRIES + 1):
        resp = session.request(method, url, timeout=TIMEOUT, **kwargs)
        if resp.status_code in (429, 500, 502, 503, 504) and attempt < MAX_RETRIES:
            wait = int(resp.headers.get("Retry-After") or 2 ** attempt)
            print(f"          HTTP {resp.status_code} - retrying in {wait}s")
            time.sleep(wait)
            continue
        return resp
    return resp


def explain(resp, step):
    hint = ""
    if resp.status_code == 403:
        hint = ("\n          -> 403: this app registration lacks SaaS consolidation permission "
                "(saas.consolidation.view.crud). Add it in Application registrations.")
    elif resp.status_code == 404:
        hint = ("\n          -> 404: this route is not on this host. Run flexera_route_probe.py and set "
                "CF_FIELDS_PATH / CF_USERS_PATH in .env to the routes that return 200.")
    sys.exit(f"[flexera] {step} failed: HTTP {resp.status_code} {resp.text[:500]}{hint}")


def paged_get(session, url, page_size, label):
    items, page = [], 1
    while True:
        resp = call(session, "GET", url, params={"page_number": page, "page_size": page_size})
        if not resp.ok:
            explain(resp, f"Get {label}")
        body = resp.json()
        items.extend(body.get("items", []))
        total = (body.get("pagination") or {}).get("total_pages") or 1
        print(f"[flexera] {label}: page {page}/{total}", end="\r")
        if page >= total:
            break
        page += 1
    print()
    return items


def find_field(session, base):
    fields = paged_get(session, base + CUSTOM_FIELDS_PATH, CF_PAGE_SIZE, "custom fields")
    cats = {}
    for f in fields:
        c = str(f.get("saasCategory"))
        cats[c] = cats.get(c, 0) + 1
    print(f"[flexera] {len(fields):,} custom field(s) visible, by category: {cats}")
    user_defs = {(f.get("fieldName") or "").strip().lower(): f for f in fields
                 if "user" in str(f.get("saasCategory") or "").lower()}
    resolved, missing = {}, []
    for name, value in TARGETS.items():
        f = user_defs.get(name.lower())
        if not f:
            missing.append(name)
            continue
        print(f"          - {f.get('fieldName')} | category={f.get('saasCategory')} | type={f.get('dataType')}")
        resolved[f["fieldName"]] = value   # exact name/case as stored
    if missing:
        names = sorted(str(f.get("fieldName")) for f in user_defs.values())
        print(f"[flexera] SaaS User custom fields that exist: {', '.join(names) or '(none)'}")
        sys.exit(f"[flexera] User custom field(s) not found: {', '.join(missing)}. "
                 f"Create them in SaaS settings > Custom fields (category: User) - not SAM settings - "
                 f"or fix the names in CF_FIELDS.")
    print(f"[flexera] Will set: " + ", ".join(f"{n}={v}" for n, v in resolved.items()))
    return resolved


_existing_cache = {}   # user ID -> list of current custom fields (when the users list includes them)


def email_map(session, base):
    users = paged_get(session, base + USERS_PATH, USERS_PAGE_SIZE, "SaaS users")
    m = {}
    for u in users:
        e = (u.get("email") or "").strip().lower()
        if e and u.get("id"):
            m.setdefault(e, []).append(u["id"])
            if isinstance(u.get("customFields"), list):
                _existing_cache[u["id"]] = u["customFields"]
    print(f"[flexera] {len(users):,} SaaS users, {len(m):,} distinct emails")
    return m


def read_existing(session, base, user_id):
    """Current custom fields for one user, as the UI reads them: GET .../users/{id} -> customFields.
    Returns a list (possibly empty), or None if they could not be read (then we must NOT write)."""
    if user_id in _existing_cache:
        return _existing_cache[user_id]
    resp = call(session, "GET", f"{base}{USERS_PATH.rstrip('/')}/{user_id}")
    if not resp.ok:
        print(f"          read failed for {user_id}: HTTP {resp.status_code} {resp.text[:150]}")
        return None
    try:
        body = resp.json()
    except ValueError:
        return None
    cf = body.get("customFields")
    if cf is None:
        # Some responses omit the key when a user has no values; treat explicit null/absent as unknown
        print(f"          no 'customFields' key in user response for {user_id}; keys: {', '.join(list(body)[:12])}")
        return None
    return cf if isinstance(cf, list) else None


def merge_fields(existing, targets):
    """Keep every existing field; set/replace each field in targets. Same shape the UI sends."""
    lower_targets = {n.lower(): (n, v) for n, v in targets.items()}
    merged, done = [], set()
    for f in existing:
        name = f.get("name") if isinstance(f, dict) else None
        if not name:
            continue
        key = name.strip().lower()
        if key in lower_targets:
            n, v = lower_targets[key]
            merged.append({"name": n, "value": v})
            done.add(key)
        else:
            merged.append({"name": name, "value": f.get("value")})
    for key, (n, v) in lower_targets.items():
        if key not in done:
            merged.append({"name": n, "value": v})
    return merged


def read_emails():
    path = Path(EMAILS_CSV)
    if not path.is_absolute():
        path = BASE_DIR / path
    if not path.exists():
        sys.exit(f"[input] {path} not found")
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        col = next((c for c in reader.fieldnames or [] if c.strip().lower() == "email"), None)
        if not col:
            sys.exit(f"[input] {path.name} needs a column named 'email'")
        emails = [(r[col] or "").strip().lower() for r in reader]
    emails = [e for e in dict.fromkeys(emails) if "@" in e]
    print(f"[input] {len(emails)} email(s) from {path.name}")
    return emails


_working_ct = None   # remembered once a content type is accepted


def set_value(session, base, user_id, custom_fields):
    """Same body shape the Flexera UI sends: {customFields:[ALL fields], id, saasCategory}.
    The endpoint REPLACES the user's whole list, so custom_fields must include the existing ones."""
    global _working_ct
    payload = {"customFields": custom_fields, "id": user_id, "saasCategory": "User"}
    body = json.dumps(payload).encode("utf-8")
    resp = None
    for ct in ([_working_ct] if _working_ct else CONTENT_TYPES):
        resp = call(session, "PATCH", base + CF_VALUES_PATH, data=body,
                    headers={"Content-Type": ct})
        if resp.status_code != 400:
            if resp.ok and _working_ct is None:
                _working_ct = ct
                print(f"[flexera] PATCH accepted with Content-Type: {ct}")
            return resp, body
        print(f"          400 with Content-Type {ct}: {resp.text[:200]}")
    return resp, body


# ══════════════════════════ MAIN ══════════════════════════════════
def main():
    base = api_base()
    print(f"[config] API host: {base}")
    print(f"[config] Routes: fields={CUSTOM_FIELDS_PATH} | values={CF_VALUES_PATH} | users={USERS_PATH}")
    print(f"[config] Fields: " + ", ".join(f"'{n}'='{v}'" for n, v in TARGETS.items()) + " | "
          f"{'DRY RUN (nothing written)' if DRY_RUN else 'LIVE - values will be written'}")

    session = requests.Session()
    session.headers.update({"Authorization": f"Bearer {get_token()}", "Accept": "application/json"})
    targets = find_field(session, base)          # exact field name -> value
    tkeys = {n.lower() for n in targets}
    field_name = ", ".join(targets)              # for messages

    if TEST_USER_ID:
        # Single-user test: no email lookup, patch just this ID
        print(f"[test] Single-user mode: SaaS user ID {TEST_USER_ID}")
        emails = ["(CF_TEST_USER_ID)"]
        by_email = {"(CF_TEST_USER_ID)": [TEST_USER_ID]}
    else:
        emails = read_emails()
        if not emails:
            sys.exit("[input] No valid emails to process")
        by_email = email_map(session, base)

    results, first_error_shown = [], False
    for email in emails:
        ids = by_email.get(email)
        if not ids:
            results.append([email, "", "SKIPPED", "email not found in SaaS users"])
            continue
        for uid in ids:   # one email can map to more than one SaaS user record
            existing = read_existing(session, base, uid)
            if existing is None:
                results.append([email, uid, "SKIPPED",
                                "could not read existing custom fields - not written (protects other values)"])
                continue
            current = {(f.get("name") or "").strip().lower(): f.get("value")
                       for f in existing if isinstance(f, dict)}
            was = "; ".join(f"{n} was {current.get(n.lower())!r}" for n in targets)
            change = ", ".join(f"{n}={v}" for n, v in targets.items())
            others = [f.get("name") for f in existing if isinstance(f, dict)
                      and (f.get("name") or "").strip().lower() not in tkeys]
            keep_note = f"keeps {len(others)} other field(s): {', '.join(others) or '-'}"
            if all(n.lower() in current and str(current[n.lower()]) == v
                   for n, v in targets.items()):
                results.append([email, uid, "UNCHANGED", f"already {change}"])
                continue
            merged = merge_fields(existing, targets)
            if DRY_RUN:
                results.append([email, uid, "DRY_RUN", f"would set {change} ({was}); {keep_note}"])
                print(f"          {email}: would send {json.dumps(merged)}")
                continue
            resp, sent = set_value(session, base, uid, merged)
            if resp.ok:
                results.append([email, uid, "OK", f"{change} ({was}); {keep_note}"])
            else:
                results.append([email, uid, f"FAILED {resp.status_code}", resp.text[:300]])
                if not first_error_shown:
                    print(f"[flexera] First failure: HTTP {resp.status_code} {resp.text[:300]}")
                    print(f"[flexera] Body sent: {sent.decode('utf-8')}")
                    first_error_shown = True
                if resp.status_code in (401, 403):
                    sys.exit("[flexera] Permission problem - stopping. App registration needs "
                             "saas.consolidation.view.crud.")
            time.sleep(SLEEP_BETWEEN_PATCHES)

    out = out_dir() / f"custom_field_results_{datetime.now():%Y%m%d_%H%M%S}.csv"
    with open(out, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["email", "saas_user_id", "status", "detail"])
        w.writerows(results)

    summary = {}
    for r in results:
        k = r[2].split()[0]
        summary[k] = summary.get(k, 0) + 1
    print(f"[done] {summary} -> {out}")
    if not DRY_RUN and summary.get("OK"):
        print(f"[done] Check a user in Flexera > SaaS > Users > Custom information for '{field_name}'")


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
    except Exception:
        traceback.print_exc()
        print("[run] Crashed", file=sys.stderr)
        sys.exit(1)
