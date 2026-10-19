"""
Pull every user (with all fields) for one SaaS application from Flexera / Snow Atlas.
Mirrors the Power Query M logic: token -> find application by name -> page through users.
No transformation: the raw API records are saved as-is to JSON.

Settings come from a .env file next to this script:
    BASE_URL, CLIENT_ID, CLIENT_SECRET, APP_NAME, PAGE_SIZE

Standard library only (no pip installs needed).
"""

import json
import os
import sys
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime

MAX_PAGES = 500  # same safety cap as the M query


# ---------- .env loader ----------
def load_env(path):
    if not os.path.exists(path):
        sys.exit(f"ERROR: .env file not found at {path}")
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
load_env(os.path.join(SCRIPT_DIR, ".env"))

BASE_URL = os.environ.get("BASE_URL", "https://westeurope.snowsoftware.io").rstrip("/")
CLIENT_ID = os.environ.get("CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET", "")
APP_NAME = os.environ.get("APP_NAME", "")
PAGE_SIZE = int(os.environ.get("PAGE_SIZE", "100") or 100)

missing = [k for k, v in {"CLIENT_ID": CLIENT_ID, "CLIENT_SECRET": CLIENT_SECRET, "APP_NAME": APP_NAME}.items() if not v]
if missing:
    sys.exit(f"ERROR: missing values in .env: {', '.join(missing)}")


# ---------- HTTP helpers ----------
def http_request(url, method="GET", headers=None, data=None):
    req = urllib.request.Request(url, data=data, headers=headers or {}, method=method)
    try:
        with urllib.request.urlopen(req, timeout=120) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {e.code} on {method} {url}\n{body[:1000]}") from e


def get_token():
    body = urllib.parse.urlencode({
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
    }).encode("utf-8")
    resp = http_request(
        f"{BASE_URL}/idp/api/connect/token",
        method="POST",
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        data=body,
    )
    return resp["access_token"]


TOKEN = {"value": None}


def api_get(path, params=None):
    """GET with bearer token; refreshes the token once if it has expired (401)."""
    url = f"{BASE_URL}/{path}"
    if params:
        url += "?" + urllib.parse.urlencode(params)
    for attempt in range(2):
        if TOKEN["value"] is None:
            TOKEN["value"] = get_token()
        headers = {"Authorization": f"Bearer {TOKEN['value']}", "Accept": "application/json"}
        try:
            return http_request(url, headers=headers)
        except RuntimeError as e:
            if "HTTP 401" in str(e) and attempt == 0:
                TOKEN["value"] = None  # token expired -> get a new one and retry
                continue
            raise


# ---------- Main ----------
def main():
    print(f"Getting token from {BASE_URL} ...")
    TOKEN["value"] = get_token()

    # Find the application
    print(f"Looking up application '{APP_NAME}' ...")
    app_resp = api_get(
        "api/saas/consolidated-view/v1/applications",
        {"filter": f"name -eq '{APP_NAME}'"},
    )
    apps = app_resp.get("items") or []
    if not apps:
        sys.exit(f"ERROR: no application found with name '{APP_NAME}'")
    if len(apps) > 1:
        print(f"WARNING: {len(apps)} applications matched; using the first one.")
    application = apps[0]
    app_id = application["id"]
    print(f"Application id: {app_id}")

    # Page through users (same stop rules as the M query)
    users = []
    page = 1
    while page <= MAX_PAGES:
        resp = api_get(
            f"api/saas/consolidated-view/v1/applications/{app_id}/users",
            {"page_number": page, "page_size": PAGE_SIZE},
        )
        items = resp.get("items") or []
        users.extend(items)
        print(f"  page {page}: {len(items)} users (total {len(users)})")
        if len(items) < PAGE_SIZE:
            break
        page += 1

    # Save raw output
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_name = "".join(c if c.isalnum() else "_" for c in APP_NAME)
    out_path = os.path.join(SCRIPT_DIR, f"{safe_name}_users_{stamp}.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(
            {
                "extracted_at": datetime.now().isoformat(timespec="seconds"),
                "application": application,
                "user_count": len(users),
                "users": users,
            },
            f,
            indent=2,
            ensure_ascii=False,
        )
    print(f"\nDone. {len(users)} users saved to {out_path}")


if __name__ == "__main__":
    main()
