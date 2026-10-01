# closeout-compare

Reconciles each night's **closeout** against the **debrief** workbooks, and
checks the shift's **work order** while it is at it.

A station finishes its shift and a technician submits a closeout form listing
the aircraft worked and what was done to each. Separately, the debrief workbooks
on SharePoint record the same work, tail by tail, as it happens. The two are
filled in by different people at different moments, so they drift — a tail on
the closeout that nobody debriefed, a debrief nobody closed out, a service
recorded one way and billed another. This repo catches that drift the moment the
closeout lands, instead of at month end when the invoice is already wrong.

Nothing here runs on a schedule. A closeout submission is the trigger: JotForm →
Power Automate → a GitHub `repository_dispatch` → one run of
`.github/workflows/reconcile.yml`. About 22 runs a night, clustered 2–5 AM ET.

## What one run does

1. **Parse the closeout.** Two payload shapes — see
   [Two kinds of closeout](#two-kinds-of-closeout).
2. **Load the debrief** for that location and service date, per fleet, from
   `Power Flows/Debriefs/*.xlsx` on the Data Hub.
3. **Reconcile both directions** — every closeout tail should be in the
   debriefs and every debrief tail should be on the closeout — and compare the
   services on each matched tail. Tails that are one character apart are
   reported as *probable typos* rather than as two separate errors.
4. **Email the result** to the compliance analysts, clean or not. Silence is
   never the answer.
5. **Post each discrepancy** to a Power Automate flow that appends it to a
   compliance table for triage.
6. **Check the work order** — the PDF the closeout uploads, listing which jobs
   were overdue or coming due. Flags priority jobs that were skipped, work that
   was not needed, tails worked that were not on the work order at all, and the
   wrong file being uploaded.
7. **Stamp a heartbeat** so the daily monitor can tell when a station stops
   filing.

## Two kinds of closeout

| | Payload | Location comes from |
|---|---|---|
| **General** — Commercial Closeout 2.0 (JotForm `222916060752150`) | numeric field ids (`6`=location, `4`=date, `3`=submitter, one array per fleet) | field `6` |
| **Location-specific** — one form per station | named keys (`tech`, `date`, and one array per fleet: `envoy`, `ultra`, `jsx`, …) | the dispatch event type, e.g. `xna_closeout_submitted` → `XNA` |

Ten stations are still on the general form: CAK, CLT, CVG, DAY, DCA, GSP, ORF
and TYS (PSA), plus IAH (Mesa) and STL (GoJet). The rest have their own.

The location-specific forms carry no location field, and several are
indistinguishable by their keys alone — CMH and XNA both send just an `envoy`
array — so the location is derived from the `repository_dispatch` event type
the workflow passes in as `CLOSEOUT_EVENT_TYPE`. JSX is the exception: it runs
at many stations and sends its own `location` key, which wins.

## Fleets

Each fleet has its own debrief workbook, sheet, column layout and way of saying
a service was performed. `DEBRIEF_LAYOUT`, `DEBRIEF_COL_MAP` and the
`_canon_*_service` functions hold the differences; everything downstream
compares one canonical code set.

GoJet · PSA · Envoy · Mesa · Ultra · Breeze · JSX · Frontier · Regional ·
Widebody. DFW is the awkward one — four fleets on one form, two of them sharing
a sheet and split by a column value.

## Running it

**Manually, against a real payload.** Actions → *Closeout Debrief
Reconciliation* → *Run workflow*, and paste the closeout JSON. Add a
`"location"` key when testing a location-specific form, since there is no
dispatch event type to derive it from.

**Locally, without sending anything:**

```bash
pip install -r requirements.txt
DEBRIEF_SOURCE=local SEND_EMAIL=false CLOSEOUT_PAYLOAD="$(cat payload.json)" \
  python closeout_debrief_reconciler.py
```

`SEND_EMAIL=false` drafts the email and the records to the console and posts
nothing. `DEBRIEF_SOURCE=local` reads workbooks from disk instead of Graph.

## Layout

| File | |
|---|---|
| `closeout_debrief_reconciler.py` | the pipeline — payload parsing, debrief loading, reconciliation, email, records, sidecars |
| `work_order.py` | pure work-order parser and evaluator; takes extracted text, no PDF dependency |
| `.github/workflows/reconcile.yml` | the one workflow; `repository_dispatch` + manual |

`CLAUDE.md` has the detail: every config flag, the per-fleet quirks, the sidecar
contracts, and the things that have bitten this pipeline in production.
