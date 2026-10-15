"""
Test: set a SaaS user custom field in Flexera One SaaS Management via API.

Sets custom field FIELD_NAME (default "abcde") = FIELD_VALUE (default "yes")
for every email listed in EMAILS_CSV, matching SaaS users by email.

APIs used (SaaS: Users + SaaS: Custom fields):
  GET   /api/saas/consolidated-view/v2/custom-fields          -> confirm field exists
  GET   /api/saas/consolidated-view/v1/users                  -> build email -> user ID map
  PATCH /api/saas/consolidated-view/v2/custom-fields/values   -> set value per user

Run:
  pip install requests python-dotenv
  python flexera_custom_field_test.py            # dry run by default (DRY_RUN=true)
  set DRY_RUN=false in .env to actually write
"""

import csv
import os
import sys
import time
from datetime import datetime

import requests
from dotenv import load_dotenv

load_dotenv()

# ---------- Config (.env) ----------
API_BASE_URL = os.getenv("API_BASE_URL", "").rstrip("/")   # same host your ADO script uses
TOKEN_URL = os.getenv("TOKEN_URL", "")                     # same token endpoint your ADO script uses
CLIENT_ID = os.getenv("CLIENT_ID", "")
CLIENT_SECRET = os.getenv("CLIENT_SECRET", "")

FIELD_NAME = os.getenv("FIELD_NAME", "abcde")
FIELD_VALUE = os.getenv("FIELD_VALUE", "yes")
EMAILS_CSV = os.getenv("EMAILS_CSV", "test_emails.csv")    # one column: email
DRY_RUN = os.getenv("DRY_RUN", "true").strip().lower() != "false"

PAGE_SIZE = 1000          # API max
SLEEP_BETWEEN_CALLS = 0.2 # gentle throttle for per-user PATCH
MAX_RETRIES = 3
TIMEOUT = 60


def fail(msg):
    print(f"ERROR: {msg}")
    sys.exit(1)


def get_token():
    resp = requests.post(
        TOKEN_URL,
        data={
            "grant_type": "client_credentials",
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET,
        },
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        timeout=TIMEOUT,
    )
    if resp.status_code != 200:
        fail(f"Token request failed ({resp.status_code}): {resp.text[:500]}")
    return resp.json()["access_token"]


def request_with_retry(session, method, url, **kwargs):
    for attempt in range(1, MAX_RETRIES + 1):
        resp = session.request(method, url, timeout=TIMEOUT, **kwargs)
        if resp.status_code in (429, 500, 502, 503, 504) and attempt < MAX_RETRIES:
            wait = int(resp.headers.get("Retry-After", 2 ** attempt))
            print(f"  {resp.status_code} on {method} {url} - retrying in {wait}s")
            time.sleep(wait)
            continue
        return resp
    return resp


def check_custom_field(session):
    url = f"{API_BASE_URL}/api/saas/consolidated-view/v2/custom-fields"
    found = []
    page = 1
    while True:
        resp = request_with_retry(session, "GET", url, params={"page_number": page, "page_size": 100})
        if resp.status_code != 200:
            fail(f"Get custom fields failed ({resp.status_code}): {resp.text[:500]}")
        body = resp.json()
        found.extend(body.get("items", []))
        if page >= (body.get("pagination", {}).get("total_pages") or 1):
            break
        page += 1

    print(f"Custom fields visible to this app registration: {len(found)}")
    for f in found:
        print(f"  - {f.get('fieldName')} | category={f.get('saasCategory')} | type={f.get('dataType')}")

    match = [f for f in found if (f.get("fieldName") or "").strip().lower() == FIELD_NAME.lower()]
    if not match:
        fail(f"Custom field '{FIELD_NAME}' not found. Create it in the UI (category: User) first.")
    field = match[0]
    print(f"Using field '{field.get('fieldName')}' (category={field.get('saasCategory')})")
    return field.get("fieldName")  # use exact name/case as stored


def build_email_map(session):
    url = f"{API_BASE_URL}/api/saas/consolidated-view/v1/users"
    email_map = {}
    page = 1
    while True:
        resp = request_with_retry(session, "GET", url, params={"page_number": page, "page_size": PAGE_SIZE})
        if resp.status_code != 200:
            fail(f"Get users failed ({resp.status_code}): {resp.text[:500]}")
        body = resp.json()
        for u in body.get("items", []):
            email = (u.get("email") or "").strip().lower()
            if email:
                email_map.setdefault(email, []).append(u["id"])
        total_pages = body.get("pagination", {}).get("total_pages") or 1
        print(f"  users page {page}/{total_pages}")
        if page >= total_pages:
            break
        page += 1
    print(f"SaaS users with email: {len(email_map)}")
    return email_map


def read_emails():
    if not os.path.exists(EMAILS_CSV):
        fail(f"{EMAILS_CSV} not found")
    with open(EMAILS_CSV, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        col = next((c for c in reader.fieldnames or [] if c.strip().lower() == "email"), None)
        if not col:
            fail(f"{EMAILS_CSV} needs a column named 'email'")
        emails = [(row[col] or "").strip().lower() for row in reader]
    return [e for e in dict.fromkeys(emails) if "@" in e]


def set_field(session, user_id, field_name):
    url = f"{API_BASE_URL}/api/saas/consolidated-view/v2/custom-fields/values"
    payload = {
        "id": user_id,
        "saasCategory": "User",
        "customFields": [{"name": field_name, "value": FIELD_VALUE}],
    }
    return request_with_retry(
        session, "PATCH", url, json=payload,
        headers={"Content-Type": "application/merge-patch+json"},
    )


def main():
    for k, v in {"API_BASE_URL": API_BASE_URL, "TOKEN_URL": TOKEN_URL,
                 "CLIENT_ID": CLIENT_ID, "CLIENT_SECRET": CLIENT_SECRET}.items():
        if not v:
            fail(f"{k} missing in .env")

    print(f"Mode: {'DRY RUN (no writes)' if DRY_RUN else 'LIVE'}")
    session = requests.Session()
    session.headers.update({"Authorization": f"Bearer {get_token()}", "Accept": "application/json"})

    field_name = check_custom_field(session)
    email_map = build_email_map(session)
    emails = read_emails()
    print(f"Emails to update: {len(emails)}")

    results = []
    for email in emails:
        ids = email_map.get(email)
        if not ids:
            results.append((email, "", "SKIPPED", "email not found in SaaS users"))
            continue
        for uid in ids:  # same email can map to >1 SaaS user record
            if DRY_RUN:
                results.append((email, uid, "DRY_RUN", f"would set {field_name}={FIELD_VALUE}"))
                continue
            resp = set_field(session, uid, field_name)
            ok = resp.status_code == 200
            results.append((email, uid, "OK" if ok else f"FAILED {resp.status_code}",
                            "" if ok else resp.text[:300]))
            time.sleep(SLEEP_BETWEEN_CALLS)

    out = f"custom_field_results_{datetime.now():%Y%m%d_%H%M%S}.csv"
    with open(out, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["email", "saas_user_id", "status", "detail"])
        w.writerows(results)

    summary = {}
    for r in results:
        key = r[2].split()[0]
        summary[key] = summary.get(key, 0) + 1
    print(f"Done: {summary}  ->  {out}")


if __name__ == "__main__":
    main()
