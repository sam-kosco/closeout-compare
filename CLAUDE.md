# closeout-compare — orientation for Claude Code

Org standards: https://github.com/Foxtrot-Aviation-Services/.github/blob/main/STANDARDS.md
— follow them; this file covers only what is specific to this repo.

Reconciles each night's closeout submission against the per-fleet debrief
workbooks on the Data Hub, and analyses the work order the closeout uploads.
EVENT-DRIVEN, never scheduled: a closeout submission fires JotForm → Power
Automate → `repository_dispatch` → one run of `reconcile.yml` (~22 runs/night,
peaking 2–5 AM ET; 6–9 PM ET is reliably dead). Reads the debrief workbooks,
writes discrepancy and work-order rows through Power Automate, emails the
compliance analysts on every run, and stamps the sidecar the core repo's daily
monitor watches. `README.md` is the plain-English tour; this file is the
detail a session must not get wrong.

## Run
- Manual: Actions → *Closeout Debrief Reconciliation* → *Run workflow*, paste
  the closeout JSON. A manual run has NO dispatch event type, so a
  location-specific payload needs an explicit `"location"` key or it silently
  falls back to CMH (see `_named_key_location`).
- Local, no side effects:
  `DEBRIEF_SOURCE=local SEND_EMAIL=false CLOSEOUT_PAYLOAD="$(cat payload.json)" python closeout_debrief_reconciler.py`
  `SEND_EMAIL=false` is the dry-run switch for EVERY outbound: the email, the
  discrepancy POSTs, the work-order POSTs and the location-stats write all log
  instead of sending. `DEBRIEF_SOURCE=local` reads workbooks from disk and also
  disables the sidecar writes.
- There is no test suite. Changes to parsing or reconciliation logic are
  verified by replaying a real payload with `SEND_EMAIL=false` and diffing the
  console report.

## Data contract
- **Reads** (DataHub drive): `Power Flows/Debriefs/{PSA,Envoy,GoJet,Mesa,Ultra,Breeze,JSX,Frontier,Widebody} Debriefs.xlsx`;
  `Monitoring/WorkOrders/*` (QR snapshots); the uploaded work-order PDFs
  (JotForm URLs, fetched unauthenticated).
- **Writes** (DataHub drive): `Monitoring/closeout_submissions.json` (monitor
  heartbeat), `Monitoring/location_nightly_stats.json`,
  `Monitoring/work_order_findings_posted.json` (resubmission de-dup),
  `Power Flows/Commercial Closeout/iah_dispatch_state.json`. Via Power
  Automate: compliance-table discrepancy rows, and the "Work Order Findings"
  worksheet of `Power Flows/Debriefs/Closeout Compare.xlsx`.
- **Consumers**: the compliance analysts (email); the core repo's daily monitor
  (`jobs/monitor.py` `sweep_closeouts`) reads the submissions sidecar.

## Schedule
None, and none should be added. The trigger of record is the closeout
submission itself. This repo is deliberately absent from the platform
dispatcher's `Monitoring/schedules.json`.

## Secrets
`TENANT_ID` / `CLIENT_ID` / `CLIENT_SECRET` (Entra "Foxtrot Report
Automation"), `DISCREPANCY_WEBHOOK_URL`, `WORKORDER_WEBHOOK_URL` (SAS-signed PA
trigger URLs), `ANTHROPIC_API_KEY` (dormant — see below). Registry:
`SECRETS.md` in the org `.github` repo.

**The PA → GitHub dispatch PAT is the single point of failure and lives inside
the Power Automate flows, not here.** When it lapses, closeouts stop
reconciling with no error anywhere in this repo — the runs simply never start.
The monitor's closeout sweep is what surfaces that, a morning later.

## Conventions / gotchas

- **Two payload shapes, and the location problem.** The general Commercial
  Closeout 2.0 (JotForm `222916060752150`) uses numeric field ids — `6`
  location, `4` date, `3` submitter — and is the form for CAK, CLT, CVG, DAY,
  DCA, GSP, IAH, ORF, STL and TYS. Location-specific forms use named keys
  (`tech`, `date`, `envoy`, `ultra`, `jsx`, …) and carry NO location field, and
  CMH and XNA are indistinguishable by their keys, so the location is derived
  from the dispatch event type (`xna_closeout_submitted` → `XNA`) passed in as
  `CLOSEOUT_EVENT_TYPE`. JSX overrides that with its own `location` key because
  it runs at many stations. Adding a location-specific form means adding its
  event type to `reconcile.yml`'s `types:` list — forget that and the dispatch
  is accepted by GitHub and silently runs nothing.
- **`SKIP_LOCATIONS` applies to the main form only.** `DFW` and `STL AD HOC`
  are skipped there, but a named-key payload is an explicit per-location opt-in
  and bypasses the skip — which is exactly how DFW reconciles through its own
  form. `STL AD HOC` has no dash, so a plain `STL` never matches it.
- **Every fleet is shaped differently and the differences are load-bearing.**
  `DEBRIEF_LAYOUT` carries a 0-based column index per fleet because the
  workbooks disagree: Mesa and the DFW sheet have no Location column at all
  (every row is implicitly IAH / DFW), JSX puts the date at column 3, Ultra and
  Breeze put the tail at column 1. `DEBRIEF_ROW_FILTER` splits DFW_Envoy from
  DFW_Regional by a column VALUE on a shared sheet. Read the comment block
  above `DEBRIEF_LAYOUT` before touching any of it.
- **A service is "performed" per fleet, not globally.** Most fleets want a
  literal `Yes` string; `NUMERIC_SERVICE_FLEETS` (Breeze, JSX) use 1/0. The
  strict Yes-reading is kept for everyone else on purpose, so a stray number in
  some other workbook can't silently flip a service on.
- **Not every mismatch is an error.** `SERVICE_EXEMPT_CLOSEOUT_CODES` (Breeze
  "Other") marks closeout services the debrief has no column for: the tail is
  still checked for presence, but comparing its services would guarantee a
  false mismatch. `TAIL_LEVEL_FLEETS` (Widebody) is the same idea for a
  single-service fleet. Typo pairs — tails one edit apart — are reported as
  probable typos rather than as both a missing and an extra tail.
- **Discrepancies and work-order findings go through Power Automate, not
  Graph.** App-only Graph workbook writes are WAC-blocked on this tenant (403).
  Both webhooks exist for that reason. `write_work_order_findings` is the
  direct-write fallback that only runs when no webhook URL is configured; it
  does not work. Don't "simplify" it back to a Graph write.
- **The discrepancy email is DETERMINISTIC by default and should stay that
  way.** Per STANDARDS, a model may prettify a message but never sits on the
  critical path. The Claude narrative is opt-in behind
  `USE_AI_DISCREPANCY_EMAIL` and is currently OFF, so `ANTHROPIC_API_KEY` is
  dormant; even when on, a failed call falls back to the deterministic builder.
- **Resubmissions are normal and are handled per-surface.** A tech re-sending a
  closeout after fixing a debrief must not duplicate anything: the IAH dispatch
  email is once per service date (`iah_dispatch_state.json`), work-order
  findings are de-duped by work-order URL (`work_order_findings_posted.json`),
  and the submissions sidecar keeps a high-water mark so a correction filed for
  an OLDER night never regresses the newest-night signal. The reconciliation
  email itself DOES go out every time, by design.
- **The submissions sidecar is a contract with another repo.** It feeds
  `sweep_closeouts` in Foxtrot-Aviation-Services/core, whose watch list names
  locations by bare airport code; `_record_closeout_submission` reduces
  `DCA-PSA` to `DCA` so both sides agree. `covered_dates` is a rolling 14
  nights, which is what lets the monitor's Monday scan judge Friday, Saturday
  and Sunday individually. Every reconciled closeout records (Sam, 2026-10-01 —
  it was named-key only before, which left the ten general-form stations
  unwatchable). Shortening the window or changing the key breaks the monitor
  quietly.
- **Bookkeeping must never fail a run.** The sidecar writes, the stats write
  and the work-order check are each wrapped so an exception is logged and the
  reconciliation continues. Keep new side effects in that pattern — but NOT the
  email, which follows the org rule that a swallowed send error is an invisible
  outage.
- **Work orders arrive as PDFs or as a photo of a QR code.** Work orders are
  built on a laptop and closeouts are done on a phone, so the upload is often a
  picture. The tracker POSTs a JSON snapshot keyed by the QR's `woId` to the
  Data Hub; the reconciler decodes the id from the image and fetches the
  snapshot rather than doing OCR. `libzbar0` is an apt package the workflow
  installs — pip alone can't provide it, so dropping that step breaks the photo
  path only, silently, for phone submissions.
- **`WORKORDER_FINDINGS_ENABLED` is a rollout gate**, currently on in the
  workflow. Off, the run does the arrival check only (did a valid work order
  come through) and writes no findings.
- **`feat/random-audits`** carries the Randomized Quality Audit Program, built
  and verified but dark, and coordinated with branches of the same name in
  foxtrot-platform and envoy-compliance-tracker. The three launch together; the
  platform's `RANDOM_AUDITS_PLAN.md` is the anchor doc.

## Known debt
- The repo lives on Sam's personal account and is PUBLIC, while
  `GRAPH_TENANT_ID` / `GRAPH_CLIENT_ID` / `GRAPH_DRIVE_ID` carry real values as
  `os.environ.get` defaults — STANDARDS forbids tenant and drive IDs in a
  public repo. Both are being resolved by the move to the org: make the
  credentials required-env with a loud exit, then transfer and set private.
- Run health in the core monitor suppresses a workflow's failures when a NEWER
  run of the same workflow succeeded. That is right for scheduled pipelines and
  WRONG here, where each run is a different closeout — a successful DFW run
  would bury a failed CMH one. Needs an exemption when this repo joins the org
  and run health starts covering it.
