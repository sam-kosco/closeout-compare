# Randomized Quality Audit Program — build plan

Branch: `feat/random-audits` in THREE repos (this one, sam-kosco/closeout-compare,
Foxtrot-Aviation-Services/envoy-compliance-tracker). Everything ships dark —
people are actively using the Quality page, the work orders, and the closeout
compare, so nothing merges to any main until Sam launches the program.
Source of truth for intent: `Randomized_Quality_Audit_Program.pdf`
(Sam, 2026-09-29) + his clarifications (2026-09-30, recorded here).

## The program in one paragraph

Each location owes audits per night based on its work order. **Count rule
(AMENDED by Sam, 2026-10-01 — strictly non-decreasing): 1–3 planes → all;
4–5 → 3; 6+ → half rounded up, floored at 4 (so 6 → 4).** The original
PDF's "≤5 all / 6+ half" rule dropped from 5 required at 5 planes to 3 at
6; the amendment smooths the ramp (1,2,3,3,3,4,4,4,5,5,6,6,7,7…) and
matches the old rule exactly from 7 planes up. Whenever fewer than all are
owed, the tails to audit are
**drawn at random when the nightly work order is generated** and
printed on it — nobody picks their best-looking plane, and local management
cannot change the list (a no-show tail = audit the next tail completed, note
the swap). Every submitted audit across an RM's locations forms one nightly
pool; **3 are randomly assigned to the RM** to review and approve (~45 min/day
regardless of region size). **Directors get 3 from the combined pool of their
RMs, excluding anything an RM was assigned.** Failed audits go back to the
location for corrective action and a follow-up audit.

## Scope (Sam, 2026-09-30)

Only audit types whose program has BOTH a compliance tracker (= a Work Order
tab) AND platform quality audits:

| Audit type (quality.AUDITS) | Locations | In? |
|---|---|---|
| CRJ Verified Clean | PSA stations (CAK/CLT/CVG/DCA/GSP/ORF/TYS…) | **YES** (PSA tracker) |
| CRJ Verified Clean | STL GoJet | **YES** (GoJet tracker) |
| CRJ Verified Clean | DFW Regional RONs | **NO** (Sam: unaffected) |
| Envoy Verified Clean | CMH/CVG/LIT/SGF/XNA | **YES** (Envoy tracker) |
| Envoy Verified Clean | DFW Envoy IHCs | **DEFAULT YES — CONFIRM WITH SAM** (has tracker + audits; he only excluded the RONs) |
| Mainline ULTRA | all | **NO** (no work order; 1–2/night = audit all, as today) |
| Mainline RON | all | **NO** (Sam: unaffected) |
| Mesa (tracker, no audit template) | IAH | **NO** (nothing to audit against) |

Scope lives in the quality registry as `"randomized": True` per
audit-type location — adding a program later is a registry flag, not code.

## RM pool rules (Sam, 2026-09-30)

- Every RM gets **3 distinct audits** drawn from the pool of ALL audits
  submitted last night across their org-chart locations.
- **RMs sharing locations get disjoint sets** (Brian + Levorn over the same
  locations: each gets 3 DIFFERENT audits). Algorithm: shuffle the RM order,
  each RM samples without replacement from (their pool − audits already
  assigned to anyone); a pool smaller than 3 assigns what exists.
- **Directors draw 3 from the union of their RMs' pools minus everything any
  RM was assigned** — the Director always reviews different work.
- Org chart (engine/orgchart) is the authority for which RMs/Directors cover
  which dists; quality.AUDITS maps location → dist as today.

## The seam: the work-order PDF is the contract

The Work Order Builder is a CLIENT-SIDE tab on each tracker page (no
backend). The closeout form already uploads that night's work-order PDF, and
closeout-compare's `work_order.py` already parses it. So:

1. **Tracker page (envoy-compliance-tracker)** draws the audit tails at
   generation time and prints them into the PDF (and the on-screen preview):
   - tails-on-shift list: assigned tails highlighted,
   - an `AUDIT` tag beside each assigned tail's service section,
   - a new **QUALITY AUDIT** section, machine-parseable (format below).
2. **closeout-compare** (`work_order.py`) parses the QUALITY AUDIT section
   out of the uploaded PDF and the reconciler persists the assignment to the
   hub sidecar `Monitoring/audit_assignments.json`. The nightly compare gains
   one new finding type: *work order required N audits for <date> but the
   section is missing/malformed* (a regenerated old-format PDF, a location
   skipping the draw).
3. **foxtrot-platform** (`engine/quality.py`) joins the sidecar with the
   SafetyCulture audits it already reads:
   - per-location **assignment coverage**: were the assigned tails audited
     (swaps allowed — any-tail fallback counts, flagged "swapped"), and did
     the COUNT meet the rule;
   - the **morning draw job** (dispatcher slot, time TBD ~10:00 ET; runs
     after the overnight audits land in SC) builds RM/Director pools per the
     rules above and persists `Monitoring/audit_reviews.json`;
   - **Approvals tab**: an RM/Director sees "Your 3 assigned reviews" pinned
     first (approve/fail exactly as today); everything else collapses under
     the existing list. Failed audit → existing corrective path (v1: the
     fail itself; follow-up tracking is a later phase).

## Machine-parseable PDF section (contract for work_order.py)

```
QUALITY AUDIT — 4 of 8 planes (random draw)
[ ] N278NN
[ ] N291NN
[ ] N304NN
[ ] N312NN
If an assigned tail doesn't come in, audit the next tail completed and
note the swap on the audit.
```
- The dash is ASCII `--` ON PURPOSE: jsPDF/helvetica mangles em-dashes in
  extracted text (they come out as replacement chars).
- Header regex: `QUALITY AUDIT\s*[-–—]+\s*(?:ALL\s+(\d+)|(\d+) of (\d+))`
- Tails: following `\[ \]\s*([A-Z0-9]+)` lines until a non-matching line.
- When every plane is owed (1–3 planes since the 2026-10-01 amendment):
  section reads `QUALITY AUDIT -- ALL 3 planes` with every tail listed
  (parser treats identically).
- VERIFIED 2026-09-30 with pypdf (the reconciler's extractor) against
  generated PSA PDFs: 8 tails -> "-- 4 of 8" + 4 checkbox tails; 4 tails
  -> "-- ALL 4 planes" + all 4; swap sentence intact in both.

## Sidecar schemas (hub, written by closeout-compare, read by platform)

`Monitoring/audit_assignments.json` — rolling, keyed `"<LOC>|<program>|<YYYY-MM-DD>"`:
```json
{"CMH|Envoy|2026-09-29": {"required": 4, "on_shift": 8,
  "tails": ["N278NN","N291NN","N304NN","N312NN"],
  "wo_url": "…", "parsed_at": "…"}}
```
`Monitoring/audit_reviews.json` — one entry per ET day, written by the
platform draw job:
```json
{"2026-09-29": {"rms": {"57368": ["audit_id", "…", "…"]},
  "directors": {"82632": ["…"]}, "drawn_at": "…"}}
```
Entries older than 60 days pruned on write. Both files are OURS (plain
content PUT via locations._write_ours — not live workbooks).

## Gating / rollout

- Platform: everything behind `QUALITY_RANDOM_AUDITS` env (default false) +
  per-location `randomized` registry flags — merge dark, flip at launch.
  **At launch, set QUALITY_RANDOM_AUDITS as a STICKY slot setting** (true
  on production, false on staging — same as USER_DIGEST/RAMP_ALERTS):
  a staging boot after 9 AM would otherwise run the draw with SEND_EMAIL
  sticky-false, stamp the day in audit_reviews.json, and production would
  then skip it — assignments recorded, nobody emailed.
- Tracker pages: `const RANDOM_AUDITS = false;` at the top of each Work
  Order tab's script — merged dark too; flipping it per-program is the
  per-tracker launch switch (PSA first, Sam decides order).
- closeout-compare: parser is additive — no section found = no assignment
  recorded = compare behaves exactly as today (until trackers launch).

## Build order

1. ✅ envoy-compliance-tracker (2026-09-30, ect branch 9fead96e): draw +
   highlighted shift list + AUDIT chips + QUALITY AUDIT PDF section on
   psa/index(envoy)/gojet pages, snapshot carries audit_required/
   audit_tails; `const RANDOM_AUDITS = true` on the branch. (Mesa
   untouched.)
2. ✅ closeout-compare (2026-09-30, branch 667e5c8): work_order.py parses
   the section (header + checkbox tails, from_snapshot too); reconciler
   `RANDOMIZED_FLEETS` gate, `_handle_audit_assignment` writes the
   sidecar unconditionally + "No audit assignment"/"Audit count off"
   findings; `_expected_audits(n)` — originally `n if n<=5 else
   ceil(n/2)`, since 2026-10-01 `n if n<=3 else 3 if n<=5 else
   max(4, ceil(n/2))` (the non-decreasing amendment, raDraw matches).
3. ✅ platform (2026-10-01, this branch a5e9100): registry flags,
   engine/quality_reviews.py (assignment_coverage, run_draw with
   disjoint RM sets + Director leftovers, draw emails, daemon loop at
   QUALITY_DRAW_HOUR_ET), approvals payload `reviews`, quality.js
   superior/admin/unassigned sections, location overdue chips dropped
   behind `random_audits`. Draw verified over 10 randomized runs
   (scratchpad test_draw.py): disjointness, leftovers-only Directors,
   once-per-day guard, retention prune, extras shapes, swap-credited
   coverage.
4. ✅ Sandbox end-to-end (2026-10-01, scratchpad test_phase4.py + the
   phase-3 sandbox server). REAL artifacts at every seam: the live PSA
   page (RANDOM_AUDITS on, served locally) generated two work orders —
   8 tails drew 4 at random (not the first four), 4 tails printed
   "ALL 4 planes"; both PDFs went through the reconciler's own
   `_pdf_text` + `parse_work_order` (audit section matched the draw
   exactly, zero false findings), `_handle_audit_assignment` with
   RANDOMIZED_FLEETS={"PSA"} wrote the sidecar records
   (`CAK|PSA|2026-10-01`, required/on_shift/tails/wo_url/parsed_at);
   the missing-section and count-off findings fire on launched fleets
   and stay silent for Mesa; and the platform's assignment_coverage
   read that exact sidecar back — 3 assigned tails + 1 swap scored
   satisfied 4/4 (1 swap), one-short scored missing 1. The platform
   review-draw leg ran earlier the same day against LIVE SafetyCulture
   data in the worktree sandbox: 24 disjoint assignments across 5 RMs
   + 3 Directors, Approvals views verified for admin and RM.

**ALL FOUR PHASES DONE — the program is launch-ready.** Remaining is
Sam's launch sequence only: flip `RANDOM_AUDITS` on each tracker page
(PSA first), set the reconciler's `RANDOMIZED_FLEETS`, add
`QUALITY_RANDOM_AUDITS=true` as a STICKY production slot setting, and
merge the three branches. Open question #4 (follow-up tracking after a
fail) stays manual in v1.

## Sam's answers (2026-09-30)

1. DFW Envoy IHCs: **IN**.
2. Draw at **9:00 AM ET, daily** — in-process daily loop on the platform
   (digest pattern: `QUALITY_DRAW_HOUR_ET=9`, gated on
   QUALITY_RANDOM_AUDITS), not a dispatcher workflow.
3. **One email per superior at draw time** (mailer, SEND_EMAIL-gated):
   each assigned audit as Location | audit-type label | Submitter, each
   linking to app.safetyculture.com/inspection/<id>. Plus the Approvals
   tab pinning.
4. Follow-up tracking after a fail: OPEN — v1 assumes manual.

## Approvals redesign: overdue is a PER-RM thing now (Sam, 2026-09-30)

Overdue audits stay a TOPLINE number, but per-LOCATION blocks stop
listing them — accountability moves to the responsible reviewer (the
person the draw assigned). The Approvals tab becomes:

- **Superior (RM/Director) view:**
  1. *Your pending reviews* — every audit assigned to YOU still awaiting
     your approve/fail, **overdue first** (awaiting > APPROVAL_DUE_DAYS),
     then newest. Visually identical rows to today.
  2. *Your reports' overdue* — below, the OVERDUE pending audits assigned
     to reviewers downstream of the viewer (a Director sees their RMs'
     overdue backlogs; an RM with no reporting reviewers sees nothing).
- **Global Admin view:** ALL pending + overdue audits **grouped by
  responsible RM**, overdue-first inside each group.
- Audits with no assignment (pre-launch backlog, non-randomized programs,
  pool leftovers nobody drew) sit in an *Unassigned* group, visible to
  admins and to the org-scoped approvers exactly as today — the old
  behavior is the fallback, not an error.
- Location drill-ins keep their scores/coverage; only the overdue listing
  moves out.
