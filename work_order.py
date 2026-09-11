#!/usr/bin/env python3
"""
Work-order compliance check
===========================
Foxtrot Aviation Services

Each closeout now uploads its nightly **work order** — a PDF derived from the
commercial compliance trackers that lists, for the tails on that station's shift,
which detailing jobs are overdue or coming due. This module parses that PDF's
extracted text and cross-references it against the debrief to surface two things
the plain closeout↔debrief reconciliation does not:

  1. MISSED PRIORITY — a job the work order flags as overdue, or due within
     WO_DUE_SOON_DAYS (5), that is NOT in the debrief (i.e. a priority job that
     didn't get done / recorded).
  2. UNNECESSARY WORK — a debrief service that the work order did NOT list as due
     or coming due for that tail (not overdue, not in the "due soon" list), i.e.
     work that wasn't needed for compliance. Only compliance-TRACKED services are
     considered (a service the work order vocabulary doesn't know is ignored).
  3. OFF WORK ORDER — a tail that was debriefed (serviced) but isn't on the work
     order's roster at all: a plane worked that wasn't on the shift's work order.

It also detects the wrong-file case: `is_work_order(parsed)` is False when the
uploaded PDF has no Nightly Work Order header (e.g. someone attached the Labor
Pulse Sheet, or a blank/garbage file). The reconciler raises an "invalid work
order" flag on that so the upload error surfaces rather than passing silently.

The parser takes already-extracted text (the reconciler extracts it with pypdf),
so this module has no PDF dependency and is trivially testable. Each work order
self-identifies its fleet in the header line ("Envoy Fleet — Nightly Work Order
…"), so a station with several programs (e.g. CVG PSA + Envoy) uploads one work
order per program and each is matched to its own fleet's debrief.
"""

import re
import datetime

# Jobs "due in N days" with N <= this count as priority for the MISSED check.
# (The work order's own "DUE SOON" bucket is wider — ~7 days — but the missed-job
# threshold is 5. The full DUE SOON list is still used for the UNNECESSARY check.)
WO_DUE_SOON_DAYS = 5

# Work-order fleet header ("<name> Fleet — Nightly Work Order") -> reconciler fleet.
WO_FLEET_MAP = {
    "ENVOY": "Envoy", "PSA": "PSA", "GOJET": "GoJet", "MESA": "Mesa",
    "ULTRA": "Ultra", "BREEZE": "Breeze", "JSX": "JSX", "FRONTIER": "Frontier",
    "REGIONAL": "Regional",
}

# Work-order service name -> canonical code (matches the reconciler's canon set).
# Keys are normalized by _norm_service (lowercased, "#" dropped, spaces collapsed).
WORKORDER_SERVICE_MAP = {
    "interior clean": "I",
    "exterior clean": "Ex",
    "cockpit clean": "CC", "cockpit cleaning": "CC",
    "deep seat clean": "DSC",
    "carpet extraction": "CE",
    "exterior detail 1": "ED1", "exterior detail1": "ED1",
    "exterior detail 2": "ED2", "exterior detail2": "ED2",
    "exterior detail 3": "ED3", "exterior detail 4": "ED4",
    "exterior detail": "ED",   # Mesa uses a single un-numbered Exterior Detail (ED)
    "lav tank pressure wash": "LAV", "lav tank pressure washing": "LAV",
    "interior heavy clean": "IHC",
    "ron clean": "RON", "ron": "RON",
}

# Per-fleet service-name overrides. Same work-order text can mean different
# canonical codes in different fleets: Mesa's "Exterior Detail" is ED, but JSX's
# "Exterior Detail"/"Interior Detail"/"Carpet Extraction" are their own full-name
# codes (matching the JSX debrief columns + _canon_jsx_service on the closeout
# side). canon_service checks the fleet's overrides first, then the global map.
FLEET_SERVICE_OVERRIDES = {
    "JSX": {
        "interior detail":  "Interior Detail",
        "exterior detail":  "Exterior Detail",
        "carpet extraction":"Carpet Extraction",
        "ron cleaning":     "RON", "ron": "RON",
        "biohazard":        "Biohazard",
    },
}

# Compliance-cycle services PER FLEET — the codes each tracker actually manages on
# a DUE-DATE cycle, and therefore can list as "due"/"overdue" on a work order.
# This scopes the "unnecessary work" check ("serviced but not due"): a debriefed
# service is only judged unnecessary if it's a cycle service FOR THAT FLEET.
#
# "Tracked" is per-fleet, not global — the old global set (whole service map minus
# RON/Biohazard) wrongly flagged routine/info-only services as unnecessary:
#   * simple Interior Clean (I) / Exterior Clean (Ex) — routine, never on a cycle,
#     so a debriefed clean false-flagged every time (Sam, 2026-09-11, GSP/PSA);
#   * PSA ED3/ED4 and IHC are info-only (only Mesa puts IHC on a cycle);
#   * RON / Biohazard / EC / ESS / FCD are event-driven, never "due".
# A code belongs here ONLY IF canon_service() can also emit it from a work order —
# a code the parser can't produce would never be in a tail's due set and would
# false-flag every time (e.g. Mesa "Flight Deck" has no work-order vocabulary yet,
# so it is intentionally omitted until the tracker's label is mapped).
FLEET_TRACKED_SERVICES = {
    "Envoy": {"ED1", "ED2"},
    "PSA":   {"CC", "DSC", "CE", "ED1", "ED2", "LAV"},
    "Mesa":  {"IHC", "ED", "DSC", "CE"},
    "GoJet": {"ED1", "ED2", "CE"},
    "JSX":   {"Interior Detail", "Exterior Detail", "Carpet Extraction"},
}
# Fallback for a fleet with no explicit set: the union of all cycle services. It
# still excludes every routine/event code, so it cannot reintroduce the
# clean-flagging bug; in practice only the five fleets above generate work orders.
_ALL_TRACKED = set().union(*FLEET_TRACKED_SERVICES.values())


def tracked_services(fleet):
    """Compliance-cycle service codes for a fleet (falls back to the union)."""
    return FLEET_TRACKED_SERVICES.get(fleet, _ALL_TRACKED)

_DASH = r"[—–-]"                       # em / en / hyphen
# A tail token: 3–8 chars of letters/digits, must contain a digit. Matches full
# N-numbers (N203NN, N80348) and GoJet's bare aircraft numbers (506, 536).
_TAIL_TOK = r"[A-Z0-9]{3,8}"
_TAIL_RE = re.compile(rf"^({_TAIL_TOK})(?:\s+([A-Z]{{2,4}}))?$")
_COMPLIANT_TAIL_RE = re.compile(rf"^({_TAIL_TOK})\s+No additional", re.I)
_OVERDUE_RE = re.compile(rf"^(.*?)\s+{_DASH}\s+(\d+)\s+days?\s+overdue$", re.I)
_DUESOON_RE = re.compile(rf"^(.*?)\s+{_DASH}\s+due in\s+(\d+)\s+days?$", re.I)


def _has_digit(s):
    return any(c.isdigit() for c in s)


def _norm_service(name):
    """Normalize a work-order service name for lookup: lowercase, drop '#',
    collapse whitespace."""
    return " ".join(str(name or "").replace("#", "").split()).lower()


def canon_service(name, fleet=None):
    """Work-order service name -> canonical code, or None if it isn't a known
    (compliance-tracked) service. Fleet-aware: a fleet's overrides win over the
    global map (e.g. JSX 'Exterior Detail' -> 'Exterior Detail', not Mesa's ED)."""
    key = _norm_service(name)
    ov = FLEET_SERVICE_OVERRIDES.get(fleet or "")
    if ov and key in ov:
        return ov[key]
    return WORKORDER_SERVICE_MAP.get(key)


def _fleet_from_header(line):
    m = re.match(r"\s*(.+?)\s+Fleet\b", line, re.I)
    if not m:
        return None
    key = m.group(1).strip().upper()
    return WO_FLEET_MAP.get(key, m.group(1).strip())


def _parse_date(line):
    m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{2,4})", line)
    if not m:
        return None
    mm, dd, yy = (int(x) for x in m.groups())
    if yy < 100:
        yy += 2000
    try:
        return datetime.date(yy, mm, dd)
    except ValueError:
        return None


def parse_work_order(text):
    """Parse extracted work-order text into a structured dict:

        {"fleet": "Envoy", "date": date(2026,8,27),
         "tails": {"N203NN": {"station": "XNA",
                              "overdue":  {"ED2": 4},        # code -> days overdue
                              "due_soon": {}},               # code -> days until due
                   ...},
         "on_shift": {"N203NN", "N208AN", "N211NN"},
         "unknown_services": ["…"]}                          # names not in the map

    Robust to the pypdf line layout (tail+station on one line, each service on its
    own "<name> — <N days overdue | due in N days>" line, sections led by
    NONCOMPLIANT / DUE SOON / COMPLIANT)."""
    lines = [ln.strip() for ln in str(text or "").splitlines() if ln.strip()]
    fleet = date = None
    header_found = False       # the "<Fleet> — Nightly Work Order" header was seen
    tails = {}
    on_shift = set()
    unknown = []
    section = None            # "overdue" | "due_soon" | "compliant" | "roster" | None
    cur = None                # current tail

    for ln in lines:
        if not header_found and re.search(r"\bFleet\b.*Work Order", ln, re.I):
            header_found = True
            fleet = _fleet_from_header(ln)
            date = _parse_date(ln)
            continue
        up = ln.upper()
        if up.startswith("NONCOMPLIANT"):
            section, cur = "overdue", None
            continue
        if up.startswith("DUE SOON"):
            section, cur = "due_soon", None
            continue
        if up.startswith("COMPLIANT"):
            section, cur = "compliant", None
            continue
        if up.startswith("TAILS ON SHIFT"):
            section, cur = "roster", None
            continue
        if up.startswith("GENERATED") or up.startswith("FOXTROT AVIATION"):
            section, cur = None, None       # footer / sub-header
            continue
        if section is None:
            continue

        # The roster line(s) after "Tails on shift" list EVERY tail on the shift
        # (space-separated) — the authoritative on_shift set, incl. compliant tails.
        if section == "roster":
            for tok in ln.split():
                if re.fullmatch(_TAIL_TOK, tok) and _has_digit(tok):
                    on_shift.add(tok.upper())
            continue

        # A service line? (test before the tail pattern — service lines carry " — ")
        mo = _OVERDUE_RE.match(ln)
        ms = _DUESOON_RE.match(ln)
        if mo or ms:
            if cur is None:
                continue
            raw, days = (mo or ms).group(1), int((mo or ms).group(2))
            code = canon_service(raw, fleet)
            if code is None:
                unknown.append(raw.strip())
                continue
            tails[cur]["overdue" if mo else "due_soon"][code] = days
            continue

        # A compliant tail line ("<tail>   No additional services required").
        mc = _COMPLIANT_TAIL_RE.match(ln)
        if mc and _has_digit(mc.group(1)):
            on_shift.add(mc.group(1).upper())
            continue

        # Otherwise a bare tail line (optionally tail + station) in a due section.
        mt = _TAIL_RE.match(ln)
        if mt and _has_digit(mt.group(1)):
            tail = mt.group(1).strip().upper()
            station = (mt.group(2) or "").strip().upper() or None
            on_shift.add(tail)
            if section in ("overdue", "due_soon"):
                cur = tail
                entry = tails.setdefault(tail, {"station": station,
                                                "overdue": {}, "due_soon": {}})
                if station and not entry.get("station"):
                    entry["station"] = station

    return {"fleet": fleet, "date": date, "tails": tails,
            "on_shift": on_shift, "unknown_services": unknown,
            # False when the uploaded PDF isn't a real Nightly Work Order (wrong
            # file — e.g. someone attached the Labor Pulse Sheet instead). The
            # reconciler raises an "invalid work order" flag on this. A genuine
            # work order with everything compliant still has is_work_order=True.
            "is_work_order": header_found}


def is_work_order(parsed):
    """True if the parsed PDF is a genuine Nightly Work Order (its header was
    found). False for the wrong-file case (Labor Pulse Sheet, blank/garbage
    extraction, a non-work-order PDF) — which the reconciler flags."""
    return bool(parsed.get("is_work_order"))


def from_snapshot(snap, fleet=None):
    """Build the SAME structured dict parse_work_order() returns, but from a
    client-side work-order SNAPSHOT instead of extracted PDF text.

    When a tracker generates a work order it also POSTs a JSON snapshot of it to
    the work-order store, keyed by the QR's unique `woId`. A closeout done on a
    phone can then upload a *photo of the QR* rather than the PDF; the reconciler
    decodes the woId, fetches this snapshot, and runs the same evaluate() on it —
    no PDF/OCR round trip. The snapshot carries the SAME human-readable service
    names the PDF prints, so service→code canonicalization runs through
    canon_service() here too: the QR-fetch path and the PDF-parse path produce
    identical structured data (and honor the same per-fleet overrides).

    Snapshot shape (JSON, so lists rather than sets/dicts):
        {"program"/"fleet": "Envoy", "date": "YYYY-MM-DD",
         "onShift": ["N203NN", …],
         "tails": [{"tail": "N203NN", "station": "XNA",
                    "overdue":  [{"name": "Exterior Detail 2", "days": 4}],
                    "dueSoon":  [{"name": "Carpet Extraction", "days": 3}],
                    "never":    ["Interior Clean"]}]}

    A "never" (never-serviced) job is folded into `overdue` with a large day
    count, so it is treated as a missed priority when it isn't in the debrief.
    """
    fleet = fleet or snap.get("fleet") or snap.get("program")
    date = None
    ds = snap.get("date")
    if ds:
        try:
            date = datetime.date.fromisoformat(str(ds)[:10])
        except ValueError:
            date = None
    tails, on_shift, unknown = {}, set(), []

    def _add(entry, bucket, name, days):
        code = canon_service(name, fleet)
        if code is None:
            unknown.append(str(name or ""))
            return
        entry[bucket][code] = int(days or 0)

    for t in (snap.get("tails") or []):
        tail = str(t.get("tail") or "").strip().upper()
        if not tail:
            continue
        on_shift.add(tail)
        entry = tails.setdefault(tail, {"station": (t.get("station") or None),
                                        "overdue": {}, "due_soon": {}})
        for item in (t.get("overdue") or []):
            _add(entry, "overdue", item.get("name"), item.get("days"))
        for name in (t.get("never") or []):
            _add(entry, "overdue", name, 9999)      # never serviced -> maximally overdue
        for item in (t.get("dueSoon") or []):
            _add(entry, "due_soon", item.get("name"), item.get("days"))

    for tail in (snap.get("onShift") or []):
        on_shift.add(str(tail).strip().upper())

    return {"fleet": fleet, "date": date, "tails": tails,
            "on_shift": on_shift, "unknown_services": unknown,
            "is_work_order": True}


def evaluate(work_order, debrief_services_by_tail, due_soon_days=WO_DUE_SOON_DAYS,
             tracked=None):
    """Cross-reference a parsed work order against the debrief.

    debrief_services_by_tail: {canon_tail: set(canonical service codes)} — the
    debrief's serviced jobs for the work order's fleet/date/location (what the
    reconciler already loads per fleet). Tails are upper/stripped.

    tracked: the compliance-cycle service codes to consider for the "unnecessary
    work" check. Defaults to this fleet's set (tracked_services(work_order fleet)),
    so routine/info-only services (simple cleans, RON, Biohazard, PSA IHC/ED3/ED4)
    are never flagged. Pass an explicit set to override.

    Returns three lists of findings (each a dict with tail and a human-readable
    detail; the per-service lists also carry `service`):
      missed         — overdue or due-<=due_soon_days jobs absent from the debrief
      unnecessary    — debriefed cycle services the work order didn't have due
      off_work_order — a tail that was debriefed but isn't on the work order at all
    """
    if tracked is None:
        tracked = tracked_services(work_order.get("fleet"))

    def dserv(tail):
        return {s for s in debrief_services_by_tail.get(tail, set())}

    missed, unnecessary, off_work_order = [], [], []

    # 1) MISSED PRIORITY: overdue, or due within the threshold, not in the debrief.
    for tail, info in work_order["tails"].items():
        done = dserv(tail)
        for code, days in info["overdue"].items():
            if code not in done:
                # A large sentinel (from a snapshot's "never serviced" job) reads
                # as never-serviced rather than an absurd day count.
                never = days >= 9000
                missed.append({"tail": tail, "service": code,
                               "priority": "never serviced" if never else f"{days} days overdue",
                               "detail": "never serviced, not debriefed" if never
                                         else f"overdue {days}d, not debriefed"})
        for code, days in info["due_soon"].items():
            if days <= due_soon_days and code not in done:
                missed.append({"tail": tail, "service": code,
                               "priority": f"due in {days} days",
                               "detail": f"due in {days}d, not debriefed"})

    # 2) UNNECESSARY WORK: a debriefed tracked service the work order did not list
    #    as due (overdue or due-soon) for a tail that is on the work order. Only
    #    tails present on the work order can be judged; untracked services skipped.
    for tail in work_order["on_shift"]:
        info = work_order["tails"].get(tail, {"overdue": {}, "due_soon": {}})
        due_codes = set(info.get("overdue", {})) | set(info.get("due_soon", {}))
        for code in dserv(tail):
            if code in tracked and code not in due_codes:
                unnecessary.append({"tail": tail, "service": code,
                                    "detail": "serviced but not due on the work order"})

    # 3) OFF WORK ORDER: a tail that was debriefed (serviced) but isn't on the
    #    work order's roster at all — a plane worked that wasn't on the shift's
    #    work order. Guarded on a populated roster so an empty/degenerate work
    #    order can't flag every debriefed tail.
    if work_order.get("on_shift"):
        for tail, services in debrief_services_by_tail.items():
            if services and tail not in work_order["on_shift"]:
                off_work_order.append({"tail": tail,
                                       "services": sorted(services),
                                       "detail": "debriefed but not on the work order"})

    return {"missed": missed, "unnecessary": unnecessary,
            "off_work_order": off_work_order}
