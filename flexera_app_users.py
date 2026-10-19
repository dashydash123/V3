"""
Pull every user (with all fields) for one SaaS application from Flexera / Snow Atlas.
Mirrors the Power Query M logic: token -> find application by name -> page through users.
No transformation: every field the API returns is written to Excel as-is.

Output: <APP_NAME>_users.xlsx next to this script (overwritten each run)
  - Sheet "Users"       : one row per user, one column per field returned by the API.
                          Nested fields (customFields, discoveryTypes, extraFields, ...)
                          are kept whole as JSON text in their cell.
  - Sheet "Application" : the application record (field / value).

Settings come from a .env file next to this script:
    BASE_URL, CLIENT_ID, CLIENT_SECRET, APP_NAME, PAGE_SIZE, PAGE_DELAY (optional)

If Flexera returns 429 (rate limit), the script waits and retries automatically.

Standard library only (no pip installs needed).
"""

import json
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from xml.sax.saxutils import escape

MAX_PAGES = 500           # same safety cap as the M query
EXCEL_CELL_LIMIT = 32767  # max characters Excel allows in one cell


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
PAGE_DELAY = float(os.environ.get("PAGE_DELAY", "2") or 2)  # seconds between page calls


# ---------- HTTP helpers ----------
RETRY_STATUSES = {429, 500, 502, 503, 504}
MAX_RETRIES = 8


def _wait_seconds(e, attempt):
    """Use the server's Retry-After / reset header if present, else back off 5,10,20,40,60,60..."""
    for h in ("Retry-After", "X-RateLimit-Reset", "RateLimit-Reset"):
        v = e.headers.get(h) if e.headers else None
        if v:
            try:
                secs = float(v)
                if secs > 1e9:  # epoch timestamp, not seconds
                    secs = secs - time.time()
                if 0 < secs <= 300:
                    return secs + 1
            except ValueError:
                pass
    return min(5 * (2 ** attempt), 60)


def http_request(url, method="GET", headers=None, data=None):
    for attempt in range(MAX_RETRIES + 1):
        req = urllib.request.Request(url, data=data, headers=headers or {}, method=method)
        try:
            with urllib.request.urlopen(req, timeout=120) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", errors="replace")
            if e.code in RETRY_STATUSES and attempt < MAX_RETRIES:
                wait = _wait_seconds(e, attempt)
                print(f"    HTTP {e.code} (rate limit / server busy) - waiting {wait:.0f}s, "
                      f"retry {attempt + 1}/{MAX_RETRIES} ...")
                time.sleep(wait)
                continue
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


# ---------- Minimal .xlsx writer (standard library only) ----------
_ILLEGAL_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _col_letter(n):
    """1 -> A, 27 -> AA"""
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _cell_xml(ref, value, style=0):
    s = f' s="{style}"' if style else ""
    if value is None:
        return ""
    if isinstance(value, bool):
        return f'<c r="{ref}" t="b"{s}><v>{int(value)}</v></c>'
    if isinstance(value, (int, float)) and abs(value) < 1e15:
        return f'<c r="{ref}"{s}><v>{value}</v></c>'
    if isinstance(value, (dict, list)):
        value = json.dumps(value, ensure_ascii=False)
    text = _ILLEGAL_XML.sub("", str(value))
    if len(text) > EXCEL_CELL_LIMIT:
        text = text[: EXCEL_CELL_LIMIT - 15] + " ...[TRUNCATED]"
    return f'<c r="{ref}" t="inlineStr"{s}><is><t xml:space="preserve">{escape(text)}</t></is></c>'


def _sheet_xml(headers, rows):
    ncols = max(len(headers), 1)
    nrows = len(rows) + 1
    last = f"{_col_letter(ncols)}{nrows}"

    # Column widths based on content (capped)
    widths = []
    for i, h in enumerate(headers):
        longest = len(str(h))
        for r in rows[:500]:
            v = r[i] if i < len(r) else None
            if v is not None:
                longest = max(longest, len(json.dumps(v, ensure_ascii=False) if isinstance(v, (dict, list)) else str(v)))
        widths.append(min(max(longest + 2, 10), 60))

    out = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        f'<dimension ref="A1:{last}"/>',
        '<sheetViews><sheetView workbookViewId="0">'
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        '</sheetView></sheetViews>',
        "<cols>" + "".join(
            f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>' for i, w in enumerate(widths, 1)
        ) + "</cols>",
        "<sheetData>",
        '<row r="1">' + "".join(
            _cell_xml(f"{_col_letter(c)}1", h, style=1) for c, h in enumerate(headers, 1)
        ) + "</row>",
    ]
    for r_idx, row in enumerate(rows, 2):
        cells = "".join(_cell_xml(f"{_col_letter(c)}{r_idx}", v) for c, v in enumerate(row, 1))
        out.append(f'<row r="{r_idx}">{cells}</row>')
    out.append("</sheetData>")
    if headers:
        out.append(f'<autoFilter ref="A1:{last}"/>')
    out.append("</worksheet>")
    return "".join(out)


def write_xlsx(path, sheets):
    """sheets: list of (sheet_name, headers, rows)"""
    n = len(sheets)
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        + "".join(
            f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            for i in range(1, n + 1)
        )
        + "</Types>"
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>"
    )
    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
        + "".join(
            f'<sheet name="{escape(name[:31])}" sheetId="{i}" r:id="rId{i}"/>'
            for i, (name, _, _) in enumerate(sheets, 1)
        )
        + "</sheets>"
        + "<definedNames>"
        + "".join(
            f'<definedName name="_xlnm._FilterDatabase" localSheetId="{i}" hidden="1">'
            f"'{escape(name[:31])}'!$A$1:${_col_letter(max(len(h), 1))}${len(r) + 1}</definedName>"
            for i, (name, h, r) in enumerate(sheets)
        )
        + "</definedNames></workbook>"
    )
    workbook_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(
            f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
            for i in range(1, n + 1)
        )
        + f'<Relationship Id="rId{n + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        "</Relationships>"
    )
    styles = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>'
        '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts>'
        '<fills count="3"><fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/><bgColor indexed="64"/></patternFill></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        "</styleSheet>"
    )

    tmp_path = path + ".tmp"
    with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        z.writestr("xl/styles.xml", styles)
        for i, (_, headers, rows) in enumerate(sheets, 1):
            z.writestr(f"xl/worksheets/sheet{i}.xml", _sheet_xml(headers, rows))
    try:
        os.replace(tmp_path, path)
    except PermissionError:
        os.remove(tmp_path)
        sys.exit(f"ERROR: could not overwrite {path} - close it in Excel and run again.")


# ---------- Main ----------
def main():
    missing = [k for k, v in {"CLIENT_ID": CLIENT_ID, "CLIENT_SECRET": CLIENT_SECRET, "APP_NAME": APP_NAME}.items() if not v]
    if missing:
        sys.exit(f"ERROR: missing values in .env: {', '.join(missing)}")

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
        time.sleep(PAGE_DELAY)  # stay under the tenant rate limit

    # Every field returned by the API becomes a column (order of first appearance)
    columns = []
    for u in users:
        for k in u:
            if k not in columns:
                columns.append(k)
    user_rows = [[u.get(c) for c in columns] for u in users]

    app_rows = [[k, v] for k, v in application.items()]

    safe_name = "".join(c if c.isalnum() else "_" for c in APP_NAME)
    out_path = os.path.join(SCRIPT_DIR, f"{safe_name}_users.xlsx")
    write_xlsx(out_path, [
        ("Users", columns, user_rows),
        ("Application", ["field", "value"], app_rows),
    ])
    print(f"\nDone. {len(users)} users x {len(columns)} fields saved to {out_path}")


if __name__ == "__main__":
    main()
