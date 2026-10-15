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

CUSTOM_FIELDS_PATH = "/api/saas/consolidated-view/v2/custom-fields"
CF_VALUES_PATH = "/api/saas/consolidated-view/v2/custom-fields/values"
USERS_PATH = "/api/saas/consolidated-view/v1/users"
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
    if explicit:
        return explicit
    u = urlparse(cfg("FLEXERA_UPLOAD_URL"))
    if not u.scheme or not u.netloc:
        sys.exit("[config] Could not read a host from FLEXERA_UPLOAD_URL; set FLEXERA_API_BASE in .env")
    return f"{u.scheme}://{u.netloc}"


FIELD_NAME = (os.getenv("CF_FIELD_NAME") or "abcde").strip()
FIELD_VALUE = (os.getenv("CF_FIELD_VALUE") or "yes").strip()
EMAILS_CSV = os.getenv("CF_EMAILS_CSV") or "test_emails.csv"
DRY_RUN = (os.getenv("CF_DRY_RUN") or "yes").strip().lower() != "no"


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
        hint = "\n          -> 404: endpoint not found on this host. Check FLEXERA_API_BASE."
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
    print(f"[flexera] {len(fields)} custom field(s) visible:")
    for f in fields:
        print(f"          - {f.get('fieldName')} | category={f.get('saasCategory')} | type={f.get('dataType')}")
    match = [f for f in fields
             if (f.get("fieldName") or "").strip().lower() == FIELD_NAME.lower()
             and "user" in str(f.get("saasCategory") or "").lower()]
    if not match:
        sys.exit(f"[flexera] User custom field '{FIELD_NAME}' not found. "
                 f"Create it in SaaS settings > Custom fields (category: User) first.")
    name = match[0]["fieldName"]   # exact name/case as stored
    print(f"[flexera] Using field '{name}'")
    return name


def email_map(session, base):
    users = paged_get(session, base + USERS_PATH, USERS_PAGE_SIZE, "SaaS users")
    m = {}
    for u in users:
        e = (u.get("email") or "").strip().lower()
        if e and u.get("id"):
            m.setdefault(e, []).append(u["id"])
    print(f"[flexera] {len(users):,} SaaS users, {len(m):,} distinct emails")
    return m


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


def set_value(session, base, user_id, field_name):
    payload = {"id": user_id, "saasCategory": "User",
               "customFields": [{"name": field_name, "value": FIELD_VALUE}]}
    return call(session, "PATCH", base + CF_VALUES_PATH, json=payload,
                headers={"Content-Type": "application/merge-patch+json"})


# ══════════════════════════ MAIN ══════════════════════════════════
def main():
    base = api_base()
    print(f"[config] API host: {base}")
    print(f"[config] Field '{FIELD_NAME}' = '{FIELD_VALUE}' | "
          f"{'DRY RUN (nothing written)' if DRY_RUN else 'LIVE - values will be written'}")

    emails = read_emails()
    if not emails:
        sys.exit("[input] No valid emails to process")

    session = requests.Session()
    session.headers.update({"Authorization": f"Bearer {get_token()}", "Accept": "application/json"})

    field_name = find_field(session, base)
    by_email = email_map(session, base)

    results, first_error_shown = [], False
    for email in emails:
        ids = by_email.get(email)
        if not ids:
            results.append([email, "", "SKIPPED", "email not found in SaaS users"])
            continue
        for uid in ids:   # one email can map to more than one SaaS user record
            if DRY_RUN:
                results.append([email, uid, "DRY_RUN", f"would set {field_name}={FIELD_VALUE}"])
                continue
            resp = set_value(session, base, uid, field_name)
            if resp.ok:
                results.append([email, uid, "OK", ""])
            else:
                results.append([email, uid, f"FAILED {resp.status_code}", resp.text[:300]])
                if not first_error_shown:
                    print(f"[flexera] First failure: HTTP {resp.status_code} {resp.text[:300]}")
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
