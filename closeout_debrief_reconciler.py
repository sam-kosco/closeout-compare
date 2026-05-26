"""
crosswinds_generate_data.py
============================
Called by crosswinds_refresh.yml with the PAYLOAD env var set to a JSON
array of inspections from Power Automate (last 7 days from SafetyCulture).

What this script does:
  1. Downloads Crosswind_Debriefs.xlsx from SharePoint (READ ONLY)
  2. Reads all historical records from the Crosswind_Debriefs table
  3. Merges incoming payload records with the historical data in memory
     (Power Automate handles all Excel writes — this script never modifies the file)
  4. Calculates rolling 7-day compliance for all 41 tails
  5. Writes crosswinds_data.json for the GitHub Pages dashboard

Environment variables (GitHub Secrets):
  TENANT_ID, CLIENT_ID, CLIENT_SECRET
"""

import os, json, sys, io, requests
from datetime import datetime, timezone, date, timedelta
from openpyxl import load_workbook

# ── CONSTANTS ────────────────────────────────────────────────────────────────
TENANT_ID     = os.environ["TENANT_ID"]
CLIENT_ID     = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]

DRIVE_ID  = "b!_bzXaIx86kOufgJN3ih-BaDIDthKYuxJkJtLi1Bm5irGjCEnK-VHSpBRRm3_SDKU"
FILE_PATH = "Power Flows/Debriefs/Crosswinds Debriefs.xlsx"
TABLE_NAME = "Crosswind_Debriefs"

TAILS = [
    "N351DC","N383CA","N478DC","N2322Y","N536DC","N723AG","N830BS",
    "N390JA","N727MZ","N705Q","N238E","N154DS","N70KK","N633DS",
    "N582DC","N572CW","N82ZZ","N618DC","N595GL","N267DC","N150PT",
    "N557DS","N942RM","N599DC","N11ZM","N625DC","N627DC","N419JS",
    "N37EE","N90FF","N60VU","N518MA","N123TV","N321JE","N84VV",
    "N568DC","N58HH","N141DK","N29FT","N650KN","N714KJ"
]

WINDOW_DAYS  = 7   # rolling compliance window
SHOW_HISTORY = 5   # recent inspections shown in detail panel


# ── GRAPH API HELPERS ─────────────────────────────────────────────────────────
def get_token():
    print("Acquiring Graph API token...")
    r = requests.post(
        f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token",
        data={
            "grant_type":    "client_credentials",
            "client_id":     CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "scope":         "https://graph.microsoft.com/.default",
        }
    )
    r.raise_for_status()
    print("  Token acquired.")
    return r.json()["access_token"]


def download_excel(token):
    encoded = FILE_PATH.replace(" ", "%20")
    url = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/root:/{encoded}:/content"
    print(f"Downloading: {FILE_PATH}")
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"})
    r.raise_for_status()
    print(f"  Downloaded {len(r.content):,} bytes")
    return io.BytesIO(r.content)


# ── EXCEL READING ─────────────────────────────────────────────────────────────
def read_table(buffer):
    """
    Read all rows from the Crosswind_Debriefs table.
    Returns a list of dicts with keys matching column headers.
    """
    wb = load_workbook(buffer, data_only=True)

    # Find the sheet containing the table
    ws = None
    for sheet in wb.worksheets:
        for tbl in sheet.tables.values():
            if tbl.name == TABLE_NAME:
                ws = sheet
                break
        if ws:
            break

    if ws is None:
        # Fallback: just read the first sheet
        print(f"  Warning: table '{TABLE_NAME}' not found — reading first sheet")
        ws = wb.worksheets[0]

    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    headers = [str(h).strip() if h is not None else "" for h in rows[0]]
    records = []
    for row in rows[1:]:
        if not any(v is not None for v in row):
            continue
        rec = dict(zip(headers, row))
        records.append(rec)

    print(f"  Read {len(records)} records from SharePoint")
    return records


# ── COMPLIANCE CALCULATION ────────────────────────────────────────────────────
def parse_record_date(val):
    """Parse a date value from the Excel cell — handles date objects, datetimes, and strings."""
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    try:
        return datetime.fromisoformat(str(val).replace("Z", "+00:00")).date()
    except Exception:
        return None


def normalize_records(excel_records, incoming):
    """
    Merge incoming payload into the Excel records in memory.
    Incoming records take precedence (they are the freshest data).
    Returns a unified list of normalized dicts.
    """
    # Build dict from Excel keyed by Audit ID
    by_audit = {}
    for rec in excel_records:
        aid = str(rec.get("Audit ID") or "").strip()
        if aid:
            by_audit[aid] = {
                "audit_id":    aid,
                "date":        parse_record_date(rec.get("Date")),
                "tech":        str(rec.get("Name") or "").strip(),
                "location":    str(rec.get("Location") or "").strip(),
                "tail":        str(rec.get("Tail") or "").strip().upper(),
                "template_id": str(rec.get("Template ID") or "").strip(),
                "report_url":  str(rec.get("Report Link") or "").strip(),
            }

    # Overlay with incoming payload (newer/corrected data)
    for insp in incoming:
        aid = (insp.get("audit_id") or "").strip()
        if not aid:
            continue
        date_val = insp.get("date", "")
        try:
            d = datetime.fromisoformat(date_val.replace("Z", "+00:00")).date()
        except Exception:
            d = None
        by_audit[aid] = {
            "audit_id":    aid,
            "date":        d,
            "tech":        insp.get("tech", ""),
            "location":    insp.get("location", ""),
            "tail":        (insp.get("tail") or "").strip().upper(),
            "template_id": insp.get("template_id", ""),
            "report_url":  insp.get("report_url", ""),
        }

    return list(by_audit.values())


def build_compliance(all_records):
    """
    Calculate rolling 7-day compliance for every tail in the master list.
    """
    today        = date.today()
    window_start = today - timedelta(days=WINDOW_DAYS - 1)

    # Group by tail, sorted newest-first
    by_tail = {}
    for rec in all_records:
        tail = rec.get("tail", "")
        if tail not in TAILS:
            continue
        if rec.get("date") is None:
            continue
        if tail not in by_tail:
            by_tail[tail] = []
        by_tail[tail].append(rec)

    for tail in by_tail:
        by_tail[tail].sort(key=lambda x: x["date"], reverse=True)

    planes = []
    for tail in TAILS:
        records   = by_tail.get(tail, [])
        recent_7d = [r for r in records if r["date"] >= window_start]
        count_7d  = len(recent_7d)

        last_clean    = records[0]["date"].isoformat() if records else None
        last_location = records[0]["location"]         if records else None
        last_tech     = records[0]["tech"]             if records else None
        days_since    = (today - date.fromisoformat(last_clean)).days if last_clean else None

        if count_7d >= 2:
            status = "compliant"
        elif count_7d == 1:
            status = "soon"
        else:
            status = "noncompliant"

        recent_inspections = [
            {
                "audit_id":   r["audit_id"],
                "date":       r["date"].isoformat(),
                "location":   r["location"],
                "tech":       r["tech"],
                "report_url": r["report_url"],
            }
            for r in records[:SHOW_HISTORY]
        ]

        planes.append({
            "tail":              tail,
            "status":            status,
            "count7d":           count_7d,
            "daysSinceLast":     days_since,
            "lastClean":         last_clean,
            "lastLocation":      last_location,
            "lastTech":          last_tech,
            "recentInspections": recent_inspections,
        })

    return planes


# ── MAIN ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("=== Crosswinds Compliance Data Relay ===")
    print(f"Run time: {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')}\n")

    # Parse incoming payload from Power Automate
    try:
        incoming = json.loads(os.environ["PAYLOAD"])
    except Exception as e:
        print(f"ERROR: Failed to parse PAYLOAD: {e}")
        sys.exit(1)

    print(f"Incoming inspections from Power Automate: {len(incoming)}")
    incoming = [i for i in incoming if (i.get("tail") or "").strip().upper() in TAILS]
    print(f"  After tail validation: {len(incoming)}")

    # Read full history from SharePoint (read-only)
    token         = get_token()
    excel_buffer  = download_excel(token)
    excel_records = read_table(excel_buffer)

    # Merge Excel history with incoming payload in memory
    print("\nMerging records...")
    all_records = normalize_records(excel_records, incoming)
    print(f"  Total unique records: {len(all_records)}")

    # Calculate compliance
    print("\nCalculating compliance...")
    planes = build_compliance(all_records)

    nc   = sum(1 for p in planes if p["status"] == "noncompliant")
    soon = sum(1 for p in planes if p["status"] == "soon")
    ok   = sum(1 for p in planes if p["status"] == "compliant")
    print(f"  Compliant: {ok}  |  Due soon: {soon}  |  Noncompliant: {nc}")

    # Write dashboard JSON
    output = {
        "generated": datetime.now(timezone.utc).isoformat(),
        "planes":    planes,
    }
    with open("crosswinds_data.json", "w") as f:
        json.dump(output, f, indent=2)

    print(f"\nWritten: crosswinds_data.json ({len(planes)} planes)")
    print("=== Done ===")
