#!/usr/bin/env python3
"""
Closeout <-> Debrief Reconciliation
===================================
Foxtrot Aviation Services

When a Commercial Closeout 2.0 webhook (JotForm 222916060752150) is received,
this module extracts the tails and services from the closeout and reconciles
them against the per-program Debrief workbooks
(SharePoint: DataHub > Power Flows > Debriefs > [Program] Debriefs.xlsx).

It checks BOTH directions for the closeout's date + location:
  * every debrief tail/service is reflected in the closeout, and
  * every closeout tail/service is reflected in the debriefs.

Any discrepancies are collected and, if present, an email is drafted via the
Claude API summarizing them.

------------------------------------------------------------------------------
CONFIG FLAGS  (the three decisions you can flip in one line each)
------------------------------------------------------------------------------
"""

import os
import re
import json
import datetime
from collections import defaultdict
import io
import time

from openpyxl import load_workbook

# ----- CONFIG -----------------------------------------------------------------

# (1) At CVG a closeout carries both PSA (field 27) and Envoy (field 45) aircraft.
#     True  -> route each fleet to its own debrief sheet and reconcile both.
#     False -> reconcile only the single primary program for the location.
ROUTE_MULTI_PROGRAM = True

# (2) Comparison strictness.
#     False -> tail + services: flag missing/extra tails AND service mismatches.
#     True  -> tail-level only: flag only missing/extra tails, ignore services.
TAIL_LEVEL_ONLY = False

# (3) Locations that must NOT be processed (skip the whole run if matched).
#     STL AD HOC is a distinct location string from STL; STL (GoJet) IS processed.
SKIP_LOCATIONS = {"DFW", "IAH", "STL AD HOC"}

# ----- DEBRIEF SOURCE ---------------------------------------------------------
# DEBRIEF_SOURCE = "graph"  -> download workbooks from SharePoint via Microsoft
#                              Graph (production).
# DEBRIEF_SOURCE = "local"  -> read from local file paths (testing / mirrors).
DEBRIEF_SOURCE = os.environ.get("DEBRIEF_SOURCE", "graph")

# Local file paths (used when DEBRIEF_SOURCE == "local"). Overridable via env.
DEBRIEF_PATHS = {
    "GoJet": os.environ.get("GOJET_DEBRIEF",  "/mnt/user-data/uploads/GoJet_Debriefs.xlsx"),
    "PSA":   os.environ.get("PSA_DEBRIEF",    "/mnt/user-data/uploads/PSA_Debriefs.xlsx"),
    "Envoy": os.environ.get("ENVOY_DEBRIEF",  "/mnt/user-data/uploads/Envoy_Debriefs.xlsx"),
}

# SharePoint file paths (used when DEBRIEF_SOURCE == "graph"), relative to the
# root of the DataHub Shared Documents drive. All three confirmed.
DEBRIEF_SP_PATHS = {
    "PSA":   "Power Flows/Debriefs/PSA Debriefs.xlsx",
    "Envoy": "Power Flows/Debriefs/Envoy Debriefs.xlsx",
    "GoJet": "Power Flows/Debriefs/GoJet Debriefs.xlsx",
}

# Microsoft Graph / Entra credentials. Same Foxtrot Report Automation app used by
# the compliance trackers. Secret comes from the environment, never hard-coded.
GRAPH_TENANT_ID = os.environ.get("TENANT_ID", "ede0c57f-549f-4a90-9f8c-7ea130346f95")
GRAPH_CLIENT_ID = os.environ.get("CLIENT_ID", "58191600-ab56-4141-bff6-806805fcbff4")
GRAPH_CLIENT_SECRET = os.environ.get("CLIENT_SECRET", "")
GRAPH_DRIVE_ID = os.environ.get(
    "DRIVE_ID", "b!_bzXaIx86kOufgJN3ih-BaDIDthKYuxJkJtLi1Bm5irGjCEnK-VHSpBRRm3_SDKU")

# Real-time race mitigation. Debriefs are completed during the shift and the
# closeout is submitted at end of shift, so overlap is not expected; retries are
# kept as a small safety net against transient Graph hiccups.
GRAPH_FETCH_RETRIES = int(os.environ.get("GRAPH_FETCH_RETRIES", "3"))
GRAPH_FETCH_DELAY_SEC = int(os.environ.get("GRAPH_FETCH_DELAY_SEC", "20"))

DEBRIEF_SHEETS = {"GoJet": "Sheet1", "PSA": "Debriefs", "Envoy": "Envoy General"}
# Envoy DFW debriefs are explicitly ignored (separate 'DFW' sheet AND DFW rows
# inside 'Envoy General' are both excluded).
ENVOY_IGNORE_LOCATIONS = {"DFW"}

# Email config — the drafted body is generic (no named addressee); these populate
# the to/from envelope. Overridable via EMAIL_TO env (comma-separated).
EMAIL_FROM = os.environ.get("EMAIL_FROM", "foxtrot.automation@foxtrotaviation.com")
EMAIL_TO   = [e.strip() for e in os.environ.get(
    "EMAIL_TO", "samuel.kosco@foxtrotaviation.com").split(",") if e.strip()]
CLAUDE_MODEL = "claude-opus-4-7"
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")

# ----- SERVICE VOCABULARY NORMALIZATION --------------------------------------
# Both systems get mapped to a single canonical code set so they can be compared.
# Canonical codes: I, Ex, CC, DSC, CE, ED1, ED2, ED3, ED4, IHC, RON, LAV

# Closeout 'Service Performed' tokens -> canonical
CLOSEOUT_TOKEN_MAP = {
    "I": "I", "INT": "I", "INTERIOR": "I",
    "EX": "Ex", "E": "Ex", "EXTERIOR": "Ex",
    "CC": "CC",
    "DSC": "DSC",
    "CE": "CE",
    "ED1": "ED1", "ED2": "ED2", "ED3": "ED3", "ED4": "ED4",
    "IHC": "IHC",
    "RON": "RON",
    "LAV": "LAV",
}

# Debrief column header -> canonical (only service columns; cols 0-3 are
# Date/Name/Location/Tail and the last is Sub ID).
DEBRIEF_COL_MAP = {
    "Interior Clean (I)": "I",
    "Exterior Clean (E)": "Ex",
    "Cockpit Cleaning (CC)": "CC",
    "Deep Seat Clean (DSC)": "DSC",
    "Carpet Extraction (CE)": "CE",
    "Exterior Detail (ED1)": "ED1",
    "Exterior Detail (ED2)": "ED2",
    "Exterior Detail (ED3)": "ED3",
    "Exterior Detail (ED4)": "ED4",
    "Exterior Detail #1 (ED1)": "ED1",
    "Exterior Detail #2 (ED2)": "ED2",
    "Lav Tank Pressure Washing": "LAV",
    "Interior Heavy Clean (IHC)": "IHC",
    "IHC": "IHC",
    "RON Clean (RON)": "RON",
}


def _canon_closeout_service(raw_service_string):
    """A closeout 'Service Performed' value -> set of canonical codes.

    Handles three shapes:
      * newline-delimited PSA/Envoy tokens: "I\\nEx\\nCC\\nDSC\\nCE"
      * bare "RON"
      * GoJet language with code in parens:
            "Gojet CLN-02 (IHC)\\nGojet CLN-06 (ED1)\\nGojet CLN-04 (CE)"
    """
    codes = set()
    if not raw_service_string:
        return codes
    for line in str(raw_service_string).replace("\r", "").split("\n"):
        line = line.strip()
        if not line:
            continue
        # GoJet: prefer the code inside parentheses if present
        paren = re.search(r"\(([^)]+)\)", line)
        token = paren.group(1) if paren else line
        key = token.strip().upper()
        if key in CLOSEOUT_TOKEN_MAP:
            codes.add(CLOSEOUT_TOKEN_MAP[key])
    return codes


# ----- CLOSEOUT EXTRACTION ----------------------------------------------------
# Reuses the verified field map: 6=Location, 4=Date, 3=Submitter,
# 281=GoJet (tail key 'Tail Number'), 27=PSA ('Dropdown'), 45=Envoy ('Tail Number').

FLEET_FIELDS = [
    ("GoJet", "281", "Tail Number"),
    ("PSA",   "27",  "Dropdown"),
    ("Envoy", "45",  "Tail Number"),
]


def _parse_array(field):
    if not field or not str(field).strip():
        return []
    try:
        return json.loads(field)
    except (json.JSONDecodeError, TypeError):
        return []


def extract_closeout(body):
    """Return dict: location, date(datetime.date), submitter, and
    fleets -> {tail: set(canonical services)}."""
    loc = (body.get("6") or "").strip()
    date_raw = (body.get("4") or "").strip()
    try:
        date = datetime.datetime.strptime(date_raw, "%Y-%m-%d").date()
    except ValueError:
        date = None

    fleets = defaultdict(lambda: defaultdict(set))
    for fleet, fid, tail_key in FLEET_FIELDS:
        for row in _parse_array(body.get(fid)):
            tail = (row.get(tail_key) or "").strip().upper()
            if not tail:
                continue
            fleets[fleet][tail] |= _canon_closeout_service(row.get("Service Performed"))
    return {"location": loc, "date": date, "submitter": (body.get("3") or "").strip(),
            "fleets": {k: dict(v) for k, v in fleets.items()}}


# ----- DEBRIEF EXTRACTION -----------------------------------------------------

def _location_matches(debrief_loc, closeout_loc, fleet):
    """Closeout location is a bare code (DCA, CVG, STL); debrief locations may be
    suffixed (DCA-PSA) or bare (CVG, STL). Match on the bare prefix."""
    if debrief_loc is None:
        return False
    dl = str(debrief_loc).strip().upper()
    cl = closeout_loc.strip().upper()
    base = dl.split("-")[0]            # 'DCA-PSA' -> 'DCA'
    return base == cl


# ----- MICROSOFT GRAPH (SharePoint download) ---------------------------------

class _NonRetryableGraphError(RuntimeError):
    """Graph error that should fail immediately (e.g. 404 path not found)."""


_graph_token_cache = {"token": None, "expires_at": 0.0}


def _graph_token():
    """Obtain (and cache) an app-only Graph access token via client credentials.
    Uses the Foxtrot Report Automation Entra app — same one the trackers use."""
    import requests  # imported lazily so local-mode runs need no network deps

    now = time.time()
    if _graph_token_cache["token"] and now < _graph_token_cache["expires_at"]:
        return _graph_token_cache["token"]

    if not GRAPH_CLIENT_SECRET:
        raise RuntimeError("CLIENT_SECRET not set — cannot authenticate to Graph")

    url = f"https://login.microsoftonline.com/{GRAPH_TENANT_ID}/oauth2/v2.0/token"
    resp = requests.post(url, data={
        "client_id": GRAPH_CLIENT_ID,
        "client_secret": GRAPH_CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    _graph_token_cache["token"] = data["access_token"]
    # refresh a minute early to avoid edge expiry
    _graph_token_cache["expires_at"] = now + int(data.get("expires_in", 3600)) - 60
    return _graph_token_cache["token"]


def _graph_download_workbook_bytes(sp_path):
    """Download a workbook's bytes from the DataHub drive by its path.
    Retries to mitigate the real-time race with Power Automate's debrief writes."""
    import requests

    if not GRAPH_DRIVE_ID:
        raise RuntimeError("DRIVE_ID not set — cannot locate the SharePoint drive")

    # URL-encode the path segments but keep the slashes
    encoded = "/".join(requests.utils.quote(seg) for seg in sp_path.split("/"))
    url = (f"https://graph.microsoft.com/v1.0/drives/{GRAPH_DRIVE_ID}"
           f"/root:/{encoded}:/content")

    last_err = None
    for attempt in range(1, GRAPH_FETCH_RETRIES + 1):
        try:
            headers = {"Authorization": f"Bearer {_graph_token()}"}
            resp = requests.get(url, headers=headers, timeout=60)
            if resp.status_code == 200:
                return resp.content
            if resp.status_code == 404:
                # Genuine path error — don't retry; re-raise past the retry handler
                raise _NonRetryableGraphError(
                    f"404 Not Found downloading '{sp_path}'. Check the SharePoint "
                    f"path (and confirm the GoJet path if this is STL).")
            resp.raise_for_status()
        except _NonRetryableGraphError:
            raise  # fail fast, no retry
        except Exception as e:  # noqa: BLE001 — transient; surface after final attempt
            last_err = e
            if attempt < GRAPH_FETCH_RETRIES:
                time.sleep(GRAPH_FETCH_DELAY_SEC)
    raise RuntimeError(f"Failed to download '{sp_path}' after "
                       f"{GRAPH_FETCH_RETRIES} attempts: {last_err}")


def _open_debrief_workbook(fleet):
    """Return an openpyxl workbook for the fleet, from Graph or local disk."""
    if DEBRIEF_SOURCE == "graph":
        sp_path = DEBRIEF_SP_PATHS[fleet]
        content = _graph_download_workbook_bytes(sp_path)
        return load_workbook(io.BytesIO(content), read_only=True, data_only=True)
    # local
    return load_workbook(DEBRIEF_PATHS[fleet], read_only=True, data_only=True)


def load_debrief_day(fleet, closeout_loc, date):
    """Return {tail: set(canonical services)} for the given fleet/location/date."""
    sheet = DEBRIEF_SHEETS[fleet]
    wb = _open_debrief_workbook(fleet)
    ws = wb[sheet]
    header = next(ws.iter_rows(max_row=1, values_only=True))
    # Map service columns by canonical code
    svc_cols = {i: DEBRIEF_COL_MAP[h] for i, h in enumerate(header)
                if h in DEBRIEF_COL_MAP}

    result = defaultdict(set)
    for r in ws.iter_rows(min_row=2, values_only=True):
        d, name, loc, tail = r[0], r[1], r[2], r[3]
        if isinstance(d, datetime.datetime):
            d = d.date()
        if d != date:
            continue
        if fleet == "Envoy" and str(loc).strip().upper() in ENVOY_IGNORE_LOCATIONS:
            continue  # ignore DFW Envoy rows
        if not _location_matches(loc, closeout_loc, fleet):
            continue
        tail = (str(tail).strip().upper() if tail else "")
        if not tail:
            continue
        for i, code in svc_cols.items():
            val = r[i]
            if isinstance(val, str) and val.strip().lower().startswith("yes"):
                result[tail].add(code)
    wb.close()
    return dict(result)


# ----- TYPO DETECTION ---------------------------------------------------------
# Closeout tails are free-typed (text box); debrief tails come from a dropdown
# and are treated as the source of truth. When a closeout tail has no debrief
# match and vice-versa, check whether the pair is "close enough" to be a typo:
# exactly one character dropped/changed, or two adjacent characters swapped.
# (Damerau-Levenshtein distance == 1, with substring length difference <= 1.)

def _is_one_edit_away(a, b):
    """True if b can be reached from a by exactly one of:
    single deletion, single substitution, or one adjacent transposition.
    (Insertion into a == deletion from b, so this is symmetric.)"""
    if a == b:
        return False  # identical is a match, not a typo
    la, lb = len(a), len(b)
    if abs(la - lb) > 1:
        return False

    if la == lb:
        # substitution (1 differing position) OR adjacent transposition
        diffs = [i for i in range(la) if a[i] != b[i]]
        if len(diffs) == 1:
            return True  # single substitution
        if len(diffs) == 2:
            i, j = diffs
            return j == i + 1 and a[i] == b[j] and a[j] == b[i]  # adjacent swap
        return False

    # length differs by exactly 1 -> single deletion (drop one char from longer)
    longer, shorter = (a, b) if la > lb else (b, a)
    i = j = 0
    skipped = False
    while i < len(longer) and j < len(shorter):
        if longer[i] == shorter[j]:
            i += 1
            j += 1
        elif skipped:
            return False
        else:
            skipped = True
            i += 1  # skip one char in the longer string
    return True  # at most one skip used


def _find_typo_pairs(co_only, db_only):
    """Match unmatched closeout tails against unmatched debrief tails.
    Returns (pairs, leftover_co, leftover_db) where pairs is a list of
    (closeout_tail, debrief_tail). Only pairs that are UNAMBIGUOUS in both
    directions (each is the other's sole near-match) are reported as typos;
    ambiguous or unmatched tails fall through as genuine missing tails."""
    co_list, db_list = list(co_only), list(db_only)
    # candidate near-matches each way
    co_cands = {c: [d for d in db_list if _is_one_edit_away(c, d)] for c in co_list}
    db_cands = {d: [c for c in co_list if _is_one_edit_away(c, d)] for d in db_list}

    pairs, used_co, used_db = [], set(), set()
    for c in co_list:
        cands = co_cands[c]
        if len(cands) == 1:
            d = cands[0]
            # require it to be mutual and unambiguous from the debrief side too
            if len(db_cands[d]) == 1 and db_cands[d][0] == c:
                pairs.append((c, d))
                used_co.add(c)
                used_db.add(d)
    leftover_co = [c for c in co_list if c not in used_co]
    leftover_db = [d for d in db_list if d not in used_db]
    return pairs, leftover_co, leftover_db


# ----- RECONCILIATION ---------------------------------------------------------

def reconcile_fleet(fleet, closeout_tails, debrief_tails):
    """Compare one fleet's closeout tails vs debrief tails for the day.
    Returns a dict of discrepancy lists."""
    disc = {
        "missing_in_debrief": [],   # on closeout, not in debrief, no typo match
        "missing_in_closeout": [],  # in debrief, not on closeout, no typo match
        "probable_typos": [],       # closeout tail ~ debrief tail (1-edit apart)
        "service_mismatches": [],   # tail matched (exact or typo), services differ
    }
    co_set, db_set = set(closeout_tails), set(debrief_tails)
    co_only, db_only = co_set - db_set, db_set - co_set

    # Identify probable typos among the unmatched tails on each side.
    typo_pairs, co_left, db_left = _find_typo_pairs(co_only, db_only)

    # Genuine missing tails (no plausible typo partner)
    for tail in sorted(co_left):
        disc["missing_in_debrief"].append(
            {"tail": tail, "closeout_services": sorted(closeout_tails[tail])})
    for tail in sorted(db_left):
        disc["missing_in_closeout"].append(
            {"tail": tail, "debrief_services": sorted(debrief_tails[tail])})

    # Probable-typo pairs: still reported as an error, flagged as likely a
    # closeout typo. Because they are most likely the same aircraft, compare
    # their services too (unless tail-level-only mode).
    for co_tail, db_tail in sorted(typo_pairs):
        entry = {
            "closeout_tail": co_tail,
            "debrief_tail": db_tail,
            "closeout_services": sorted(closeout_tails[co_tail]),
            "debrief_services": sorted(debrief_tails[db_tail]),
        }
        if not TAIL_LEVEL_ONLY:
            co_svc, db_svc = closeout_tails[co_tail], debrief_tails[db_tail]
            entry["service_match"] = (co_svc == db_svc)
            entry["on_closeout_only"] = sorted(co_svc - db_svc)
            entry["on_debrief_only"] = sorted(db_svc - co_svc)
        disc["probable_typos"].append(entry)

    # Exact-match tails: compare services
    if not TAIL_LEVEL_ONLY:
        for tail in sorted(co_set & db_set):
            co_svc, db_svc = closeout_tails[tail], debrief_tails[tail]
            if co_svc != db_svc:
                disc["service_mismatches"].append({
                    "tail": tail,
                    "on_closeout_only": sorted(co_svc - db_svc),
                    "on_debrief_only": sorted(db_svc - co_svc),
                    "closeout_services": sorted(co_svc),
                    "debrief_services": sorted(db_svc),
                })
    return disc


def reconcile(body):
    """Top-level entry point. Returns a structured report dict.
    report['skipped'] is set when the location is in SKIP_LOCATIONS."""
    co = extract_closeout(body)
    loc, date = co["location"], co["date"]

    if loc.upper() in {s.upper() for s in SKIP_LOCATIONS}:
        return {"skipped": True, "reason": f"Location '{loc}' is in SKIP_LOCATIONS",
                "location": loc, "date": str(date)}
    if date is None:
        return {"skipped": True, "reason": "Closeout date could not be parsed",
                "location": loc, "date": co["date"]}

    fleets_present = {f: t for f, t in co["fleets"].items() if t}
    if not ROUTE_MULTI_PROGRAM and len(fleets_present) > 1:
        # keep only the largest fleet
        primary = max(fleets_present, key=lambda f: len(fleets_present[f]))
        fleets_present = {primary: fleets_present[primary]}

    per_fleet = {}
    has_any = False
    for fleet, co_tails in fleets_present.items():
        db_tails = load_debrief_day(fleet, loc, date)
        d = reconcile_fleet(fleet, co_tails, db_tails)
        per_fleet[fleet] = {
            "closeout_tail_count": len(co_tails),
            "debrief_tail_count": len(db_tails),
            "discrepancies": d,
        }
        if any(d.values()):
            has_any = True

    return {"skipped": False, "location": loc, "date": str(date),
            "submitter": co["submitter"], "fleets": per_fleet,
            "has_discrepancies": has_any}


# ----- EMAIL DRAFTING (Claude API) -------------------------------------------

def draft_discrepancy_email(report):
    """Call the Claude API to draft a short internal email summarizing the
    discrepancies. Returns {'subject':..., 'body':...} or None on failure."""
    if report.get("skipped") or not report.get("has_discrepancies"):
        return None

    try:
        import anthropic
    except ImportError:
        raise RuntimeError("pip install anthropic to enable email drafting")

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    prompt = f"""You are drafting a brief INTERNAL email flagging reconciliation discrepancies between a nightly Commercial Closeout and the program debrief tracker for Foxtrot Aviation Services.

Here is the structured discrepancy report (JSON):

{json.dumps(report, indent=2)}

Write a concise, professional email. Requirements:
- Subject line on the first line prefixed with "Subject: ".
- Do NOT include a salutation/greeting or a named addressee (no "Hi", no "Dear ___") — the recipient list is set separately. Begin directly with the summary sentence.
- Open with one sentence stating the location, date, and that discrepancies were found.
- For each fleet, use a short labeled section. List:
    * Tails on the closeout but missing from the debrief.
    * Tails in the debrief but missing from the closeout.
    * Tails where the service lists differ (name the differing service codes on each side).
    * Probable typos: a closeout tail that does not match any debrief tail but is one character off from a debrief tail that is otherwise unmatched. Report these as an error, BUT explicitly note it is most likely a typo on the closeout, showing both the closeout-entered tail and the likely-correct debrief tail. If the matched services also differ, mention that too.
- Keep it scannable (short lines / simple bullets). No filler and no closing pleasantries; end the body after the last item with no sign-off.
- Do not invent any tails or services beyond what is in the JSON."""

    msg = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=1200,
        messages=[{"role": "user", "content": prompt}],
    )
    text = "".join(b.text for b in msg.content if getattr(b, "type", "") == "text").strip()

    subject = f"Closeout/Debrief discrepancies — {report['location']} {report['date']}"
    body = text
    if text.lower().startswith("subject:"):
        first, _, rest = text.partition("\n")
        subject = first.split(":", 1)[1].strip()
        body = rest.strip()
    return {"subject": subject, "body": body, "to": EMAIL_TO, "from": EMAIL_FROM}


def send_email_via_graph(email):
    """Send a drafted email through Microsoft Graph as EMAIL_FROM.
    Reuses the same Entra app/token as the SharePoint download (sendMail scope).
    `email` is the dict returned by draft_discrepancy_email."""
    import requests

    if not email or not email.get("to"):
        raise RuntimeError("No recipients set (EMAIL_TO) — cannot send.")

    url = (f"https://graph.microsoft.com/v1.0/users/"
           f"{requests.utils.quote(email['from'])}/sendMail")
    payload = {
        "message": {
            "subject": email["subject"],
            "body": {"contentType": "Text", "content": email["body"]},
            "toRecipients": [{"emailAddress": {"address": a}} for a in email["to"]],
        },
        "saveToSentItems": True,
    }
    resp = requests.post(url, headers={
        "Authorization": f"Bearer {_graph_token()}",
        "Content-Type": "application/json",
    }, json=payload, timeout=30)
    if resp.status_code not in (200, 202):
        raise RuntimeError(f"sendMail failed ({resp.status_code}): {resp.text[:300]}")
    return True


# ----- TEXT REPORT (no-API fallback / logging) -------------------------------

def format_report(report):
    if report.get("skipped"):
        return f"SKIPPED: {report['reason']} (location={report.get('location')}, date={report.get('date')})"
    lines = [f"Reconciliation — {report['location']} {report['date']} "
             f"(submitted by {report['submitter']})"]
    if not report["has_discrepancies"]:
        lines.append("  No discrepancies. Closeout and debriefs are in agreement.")
        return "\n".join(lines)
    for fleet, info in report["fleets"].items():
        d = info["discrepancies"]
        if not any(d.values()):
            continue
        lines.append(f"\n  [{fleet}]  closeout tails={info['closeout_tail_count']}, "
                     f"debrief tails={info['debrief_tail_count']}")
        for x in d["missing_in_debrief"]:
            lines.append(f"    - {x['tail']}: on CLOSEOUT, missing from DEBRIEF "
                         f"(services: {', '.join(x['closeout_services']) or 'none'})")
        for x in d["missing_in_closeout"]:
            lines.append(f"    - {x['tail']}: in DEBRIEF, missing from CLOSEOUT "
                         f"(services: {', '.join(x['debrief_services']) or 'none'})")
        for x in d.get("probable_typos", []):
            note = (f"    - {x['closeout_tail']}: on CLOSEOUT, no debrief match "
                    f"-- PROBABLE TYPO of debrief tail {x['debrief_tail']}")
            if x.get("service_match") is False:
                note += (f"; services also differ "
                         f"(closeout-only={x['on_closeout_only'] or '[]'}, "
                         f"debrief-only={x['on_debrief_only'] or '[]'})")
            lines.append(note)
        for x in d["service_mismatches"]:
            lines.append(f"    - {x['tail']}: service mismatch  "
                         f"closeout-only={x['on_closeout_only'] or '[]'}, "
                         f"debrief-only={x['on_debrief_only'] or '[]'}")
    return "\n".join(lines)


def _load_payload():
    """Get the closeout webhook payload for a GitHub Actions run.

    Resolution order:
      1. CLI arg: path to a JSON file (local testing).
      2. CLOSEOUT_PAYLOAD env var containing the JSON string (set by the workflow
         from the repository_dispatch client_payload).
      3. CLOSEOUT_PAYLOAD_FILE env var: path to a JSON file.

    Unwraps to the closeout body dict (the object whose keys are "6", "4", "27"...)
    regardless of how it is nested. GitHub's repository_dispatch caps client_payload
    at 10 top-level properties, so the flow wraps the body under a single
    "closeout" key; this also accepts a "body" wrapper or a bare body.
    """
    import sys
    raw = None
    if len(sys.argv) > 1:
        with open(sys.argv[1]) as fh:
            raw = fh.read()
    elif os.environ.get("CLOSEOUT_PAYLOAD"):
        raw = os.environ["CLOSEOUT_PAYLOAD"]
    elif os.environ.get("CLOSEOUT_PAYLOAD_FILE"):
        with open(os.environ["CLOSEOUT_PAYLOAD_FILE"]) as fh:
            raw = fh.read()
    else:
        raise SystemExit("No payload provided (CLI arg, CLOSEOUT_PAYLOAD, "
                         "or CLOSEOUT_PAYLOAD_FILE).")
    data = json.loads(raw)
    # Peel known wrapper keys until we reach the dict that holds the closeout
    # fields. Order matters: client_payload -> closeout -> body.
    for key in ("client_payload", "closeout", "body"):
        while isinstance(data, dict) and key in data and isinstance(data[key], dict):
            data = data[key]
    return data


def main():
    body = _load_payload()
    report = reconcile(body)

    # Always log a human-readable report to the Actions console.
    print(format_report(report), flush=True)

    if report.get("skipped") or not report.get("has_discrepancies"):
        return  # nothing to email

    email = draft_discrepancy_email(report)
    if email is None:
        return
    # Draft to the log so it's visible even if sending fails.
    print("\n----- DRAFTED EMAIL -----", flush=True)
    print(f"To: {', '.join(email['to'])}\nSubject: {email['subject']}\n", flush=True)
    print(email["body"], flush=True)

    if os.environ.get("SEND_EMAIL", "true").lower() == "true":
        send_email_via_graph(email)
        print("\n[sent via Graph]", flush=True)
    else:
        print("\n[SEND_EMAIL=false — draft only, not sent]", flush=True)


if __name__ == "__main__":
    main()
