# Randomized Quality Audit Program — build plan (v2)

Built on `feat/random-audits` in THREE repos (foxtrot-platform,
sam-kosco/closeout-compare, Foxtrot-Aviation-Services/envoy-compliance-tracker).
Source of truth for intent: Sam's v2 direction (2026-11-02, recorded here),
superseding the September PDF where they differ.

**LAUNCH STATUS (Sam, 2026-11-02): LIVE for PSA and GoJet** — all three
branches merged to main the same day. The psa/gojet tracker pages draw
(`RANDOM_AUDITS = true`), closeout-compare enforces service nights from
2026-11-03 (`RANDOMIZED_FLEETS` start dates — the one-night grace covers
work orders printed before the Pages deploy), and the platform scores the
mandate behind `QUALITY_RANDOM_AUDITS` (sticky: true prod / false staging).
**Envoy is NOT launched**: index.html ships `RANDOM_AUDITS = false`, no
`"Envoy"` entry in `RANDOMIZED_FLEETS`, and the Envoy quality-registry
locations stay unflagged. Launching it = flip those three, nothing else.

**No swap credit (Sam, 2026-11-02, at launch): the stats track ASSIGNED and
MISSED only.** Locations audit plenty of off-list planes for the 50%
coverage goal, so an extra audit proves nothing about the mandate and never
credits it. The work order still instructs the team to audit the next
completed tail and note the swap when a drawn tail no-shows — that note is
the paper trail a human reads; the night still counts as missed on
Administrative Stats.

## Why v2 (Sam, 2026-11-02)

The v1 plan (September) assumed the program would live on the Quality Audits
page and planned its own RM/Director review draw. Three facts moved:

1. **Not every platform audit is on a tracked fleet, and not every tracked
   fleet has a platform audit.** The overlap IS the scope: PSA and GoJet work
   orders relate to the CRJ Verified Clean audit; Envoy work orders to the
   Envoy Verified Clean audit. Nothing else qualifies.
2. **The team's quality goal is now 50% debrief coverage** per location
   (`COVERAGE_GOAL`, live on main since 2026-10-19). The work-order mandate is
   **independent of it** — drawn tails neither satisfy nor replace the 50%;
   they are an extra, non-negotiable floor: if a tail is selected on the work
   order, it needs an audit.
3. **The review draw shipped separately.** Main's `engine/quality_reviews.py`
   (the live 7 AM RM/Director assignment draw, merged 2026-10-20) superseded
   this branch set's review-pool design. This branch set no longer touches the
   Quality Audits page or the review draw AT ALL.
4. **Non-compliance is an ADMIN failure, not a quality metric.** A drawn tail
   that never got its audit surfaces on Location Health → **Administrative
   Stats** (the tab Sam defined as closeout-compare submission errors + budget
   faithfulness + similar), not on the Quality Audits page.

## The program in one paragraph (v2)

Every nightly work order generated on the PSA, Envoy and GoJet tracker pages
draws tails at random that MUST be audited that night — **1 tail when the work
order has 3 or fewer planes, 2 when it has more than 3** (Sam, 2026-11-02).
The drawn tails are printed on the work-order PDF (highlighted in the on-shift
list, an `AUDIT` chip on their service blocks, and a machine-parseable
**QUALITY AUDIT** section), so nobody picks their best-looking plane and local
management cannot change the list. A no-show tail = audit the next tail
completed and note the swap on the audit. Enforcement rides the existing
closeout flow: every closeout that carries a work-order upload — the general
Commercial Closeout AND the location-specific closeouts (DFW/XNA/SGF/LIT/CMH
…) — already goes through closeout-compare, whose parser reads the QUALITY
AUDIT section, persists the assignment to the hub sidecar, and (for launched
fleets) flags a work order with a missing or miscounted section. The platform
joins the sidecar to the SafetyCulture audits it already reads and reports
per-location **assignment coverage** on Administrative Stats.

## Count rule (Sam, 2026-11-02)

```
n planes on the work order ≤ 3  →  1 audit
n planes on the work order > 3  →  2 audits
```

It lives in exactly TWO places that must stay in lockstep: each tracker page's
`raDraw()` and closeout-compare's `_expected_audits()`. (v1's non-decreasing
ladder — all/3/half-floored-at-4 — is superseded.)

## Scope

Only audit types whose program has BOTH a compliance tracker (= a Work Order
tab) AND platform quality audits:

| Work orders | Audit type (quality.AUDITS) | In? |
|---|---|---|
| PSA tracker (CAK/CLT/CVG/DAY/DCA/GSP/ORF/TYS…) | CRJ Verified Clean | **YES** |
| GoJet tracker (STL) | CRJ Verified Clean | **YES** |
| Envoy tracker (CMH/CVG/LIT/SGF/XNA) | Envoy Verified Clean | **YES** |
| Envoy tracker (DFW Envoy IHCs) | Envoy Verified Clean | **YES** — the DFW closeout carries the WO upload, and location-specific closeouts are explicitly in scope (Sam, 2026-11-02) |
| DFW Regional RONs | CRJ Verified Clean | **NO** (Sam: unaffected) |
| Republic (Mesa) tracker | — no audit template | **NO** |
| JSX, Crosswinds trackers | — no platform audits | **NO** |
| Mainline ULTRA / RON | — no work orders | **NO** |

Scope lives in the platform quality registry as `"randomized": "<fleet>"` per
audit-type location — adding a program later is a registry flag, not code.

## The seam: the work-order PDF is the contract (unchanged from v1)

1. **Tracker page (envoy-compliance-tracker)** draws the audit tails at
   generation time and prints them into the PDF: assigned tails highlighted in
   the on-shift list, an `AUDIT` chip beside each assigned tail's service
   block, and the QUALITY AUDIT section (format below). The work-order
   snapshot (`woStoreSnapshot`) carries `audit_required` / `audit_tails` so a
   QR-photo closeout resolves identically to a PDF upload.
2. **closeout-compare** (`work_order.py`) parses the section out of every
   uploaded work order — `collect_work_order_findings` runs on general-form
   AND location-specific closeouts alike — and `_handle_audit_assignment`
   persists it to `Monitoring/audit_assignments.json` unconditionally. For
   fleets in `RANDOMIZED_FLEETS` it adds two finding types: *No audit
   assignment* (section missing — stale or regenerated old-format work order)
   and *Audit count off* (section disagrees with `_expected_audits`).
3. **foxtrot-platform** (`engine/adminstats.py`, NEW in v2) joins the sidecar
   with the SafetyCulture audits `engine/quality.py` already reads and serves
   per-location **assignment coverage** (ASSIGNED tails audited — nothing
   else counts, see the no-swap-credit rule above) to the
   Administrative Stats tab (`#health/stats`). The Quality Audits page is NOT
   touched.

## Machine-parseable PDF section (contract for work_order.py)

```
QUALITY AUDIT -- 2 of 8 planes (random draw)
[ ] N278NN
[ ] N291NN
If an assigned tail doesn't come in, audit the next tail completed and
note the swap on the audit.
```
- ASCII `--` ON PURPOSE: jsPDF/helvetica mangles em-dashes in extracted text.
- Header regex: `QUALITY AUDIT\s*[-–—]+\s*(?:ALL\s+(\d+)|(\d+) of (\d+))`
- Tails: following `\[ \]\s*([A-Z0-9]+)` lines until a non-matching line.
- With the v2 count rule the `ALL n planes` form only occurs at n=1
  (`QUALITY AUDIT -- ALL 1 planes`); the parser treats both forms identically.
- The v1 contract was VERIFIED 2026-09-30 with pypdf (the reconciler's own
  extractor) against generated PSA PDFs; v2 changes only the counts, not the
  format.

## Sidecar schema (unchanged shape)

`Monitoring/audit_assignments.json` — rolling 60 days, keyed
`"<LOC>|<fleet>|<YYYY-MM-DD>"`:
```json
{"CMH|Envoy|2026-11-01": {"required": 2, "on_shift": 8,
  "tails": ["N278NN","N291NN"], "wo_url": "…", "parsed_at": "…"}}
```
Written by closeout-compare (plain-content PUT — ours, not a live workbook),
read by the platform.

## Platform leg (v2): Administrative Stats

`engine/adminstats.py`, gated on `QUALITY_RANDOM_AUDITS` (env, default
false; **at launch a STICKY slot setting — true prod / false staging**,
same reasoning as USER_DIGEST/RAMP_ALERTS):

- `assignment_coverage(akey, lkey, days)` — for each night with a sidecar
  entry, count required tails and
  match submitted audits (awaiting or approved) of that type/location dated
  the assignment night or the day after (the same ±1 the debrief match uses);
  an ASSIGNED tail audited = satisfied, anything else = missed (no swap
  credit — see the launch-status note). Returns nights / required /
  satisfied / missing. NB "assignment coverage" is this Admin Stats figure
  only — the Quality Audits page's DEBRIEF coverage (the 50% goal) is a
  different number and is untouched by this program.
- `/api/health/stats` serves one row per randomized location over the 7/28/90
  windows; `static/health.js` renders the table on `#health/stats` (which
  keeps its under-construction note for the OTHER planned admin metrics —
  closeout submission errors, budget faithfulness).
- Org-chart scoped like the rest of Location Health.

## Gating / rollout

- Tracker pages: `const RANDOM_AUDITS` is the per-tracker launch switch —
  true on psa.html and gojet.html since 2026-11-02, false on index.html
  (Envoy) until that program launches.
- closeout-compare: parser is additive — no section found = no assignment
  recorded = compare behaves exactly as today. `RANDOMIZED_FLEETS` maps each
  LAUNCHED fleet to the first SERVICE DATE enforced (PSA + GoJet:
  2026-11-03 — one night of grace so work orders printed before the Pages
  deploy never false-flag). The sidecar records any section it finds
  regardless, so the platform stats start the first enforced night.
- Platform: `QUALITY_RANDOM_AUDITS` — sticky slot setting since launch,
  true production / false staging. Off = no reads, the Admin Stats tab shows
  only the under-construction card.

Launch sequence (per remaining program — Envoy): flip index.html's
`RANDOM_AUDITS` → add `"Envoy": <first enforced service date>` to
`RANDOMIZED_FLEETS` → add `"randomized": "Envoy"` to its quality-registry
locations (CMH/CVG/LIT/SGF/XNA/DFW under the envoy audit).

## v1 history (for the record)

- v1's four phases were built and verified end-to-end 2026-09-30 → 2026-10-01
  (real PSA PDFs through the reconciler's own extractor; sidecar writes;
  platform draw over live SafetyCulture data in a sandbox). The tracker draw
  and the closeout-compare parser carried into v2 ARE those verified pieces —
  only the count rule changed.
- v1's count rule (1–3 all; 4–5 → 3; 6+ → max(4, ⌈n/2⌉)) superseded
  2026-11-02 by the 1-or-2 rule.
- v1's platform half (review-pool draw, Approvals redesign, quality.js) was
  superseded by main's `engine/quality_reviews.py` (live since 2026-10-20) and
  is dropped; its `assignment_coverage` survives in `engine/adminstats.py`.
  The old platform branch tip (d0de551) holds the full v1 implementation if it
  is ever wanted.
