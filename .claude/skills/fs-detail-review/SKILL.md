---
name: fs-detail-review
description: >
  Tie-out review of the financial statement bridge workbook AND rendered FS PDF that external
  compilers (FS Compilation Partner) deliver during the annual External Audit Firm audit. Numbers-only: confirms every
  dollar amount on the PDF ties to its source (bridge, TB, prior-year FS) and that internal
  cross-references reconcile. Produces an annotated tie-out PDF (mirroring the FY24 final
  FS Compilation Partner convention) plus an Excel exceptions report. Designed to be re-run on each
  new bridge / PDF version.

  Use this skill when the user mentions: "tie out the financials", "tie out the bridge",
  "tie out the FS", "tie-out", "FS Compilation Partner", "annotate the PDF", or drops in a new bridge
  .xlsx + FS .pdf in the Tieout/ folder.
---

# Financial Statement Tie-out

This skill ties every dollar value on the rendered Financial Statements PDF to its source —
bridge workbook → trial balance → prior-year FS — and flags any value that doesn't reconcile.
It is the numbers-only successor to the broader review that previously also covered disclosure
completeness and External Audit Firm-comment coverage (those are out of scope here).

**Outputs:**
- `Tieout/<pdf-name>_TIEOUT.pdf` — clean PDF + tie-out marks overlaid (mirrors the FY24
  final FS Compilation Partner tie-out convention)
- `Tieout/Tieout Exceptions - <version>.xlsx` — Summary + per-lane review tabs (Bridge Ties,
  TB Ties, Mapping Completeness, SOE Rollforward, Internal Ties, PY Ties, Footing) + Exceptions
  + All Records. Each tab is sorted findings-first and color-coded by status.
- `Tieout/.work/` — intermediate artifacts: `inputs.json`, `tie-laneN-*.json`, `ocr-cache.json`,
  page renders

**Key references:**
- [tieout-conventions.md](references/tieout-conventions.md) — the FY24 FS Compilation Partner tie-out
  mark vocabulary that the annotator mirrors
- [bridge-tb-structure.md](references/bridge-tb-structure.md) — bridge column convention,
  the two-block side-by-side tab pattern, per-tab unit detection, sign conventions (read this
  before adapting to a new company's files)
- [exception-vs-noise.md](references/exception-vs-noise.md) — how to tell a real exception from
  a caption/rounding/unit/sign mismatch; the status taxonomy
- [source-of-truth-hierarchy.md](references/source-of-truth-hierarchy.md) — compiler schedule ↔
  subledger ↔ GL reconcile; the fully-amortized-asset gotcha
- [what-reviewers-commonly-catch.md](references/what-reviewers-commonly-catch.md) — abstracted
  checklist of what auditors flag, and how findings get resolved; run the relevant lanes before a
  draft goes out. *(Private deployments also keep a deeper companion catalog of the real reviewer
  comments + resolutions by review round, which is excluded from the public mirror.)*
- [inscope-template-mapping.md](references/inscope-template-mapping.md) — PDF line items ↔
  bridge tab/row mapping
- [company-context.template.md](references/company-context.template.md) — private context — acquisition history,
  subsidiaries, prior-year corrections to watch

## Ten checks (eight tie-out lanes + two source modules)

| Lane | What it ties | Tolerance |
|---|---|---|
| 1. PDF face ↔ Bridge | BS / IS / SOE / SCF rows + FN tables (with secondary number-match to rescue caption changes) | ±$1K simple, ±$5K subtotal |
| 2. Bridge ↔ TB | Bridge BS lines vs rolled-up TB account values (using bridge "TB Mapping" tab) | ±$1K simple, ±$5K subtotal |
| 3. PDF prior-year col ↔ FY24 final FS | every prior-year value on the PDF vs last year's issued FS | exact (any delta = restatement) |
| 4. Internal PDF cross-refs | BS Cash ↔ SCF ending; IS Net Loss ↔ SCF/SOE; BS equity ↔ SOE closing; BS↔FN totals; SCF Begin + Change = End | exact / ±$5K subtotal |
| 5. SOE rollforward | per-class equity columns, cross-foot (xF) of each balance row, vertical roll-forward (F) | ±$5K |
| 6. Footing | every subtotal in every face/FN table re-summed from components (F / xF marks) | ±$5K |
| 7. Mapping completeness | unmapped TB accounts, SUM(mapped)=TB reconciliation, TB integrity, stale mappings, BS identity | $1 / exact |
| 8. PBC ↔ bridge | each mapped footnote/bridge tab back to its supporting PBC workpaper, and (where the PBC carries a GL column) to the GL — source-to-disclosure assurance | ±$5K subtotal |
| 9. Tax provision (FN-07) | statutory recompute, rate-rec footing, total provision ↔ FS tax expense, book-pretax ↔ FS, deferred summary ties + foots | ±$1K |
| 10. Flux review | YoY movements recomputed on the **final** FS, flagged against a threshold; the company's flux-PBC comments carried as directional context; movers with no explanation flagged | analytical (not a tie) |

Tolerances are always applied **in the comparison unit** (typically $K because the PDF is in $K).
Sign-inversion between the bridge (negative for losses) and the PDF (positive in "loss" lines) is
detected and auto-classified as `ties-with-sign-inversion`. For the full noise-vs-exception logic
and the secondary number-match, see [exception-vs-noise.md](references/exception-vs-noise.md).

**Lane 7 (mapping completeness)** is the contextual "did anything fall off the FS / does it all
make sense" check: it flags TB accounts with a balance that aren't mapped in the bridge's
"TB Mapping" tab (a new GL account that silently drops off the statements), reconciles the sum of
mapped accounts to the total TB, verifies the TB nets to zero and the balance sheet balances, and
flags stale mapping targets. It produces TB-account findings only — it does not annotate the PDF.

**Lane 8 (PBC ↔ bridge)** closes the last integrity gap. Lanes 1–2 prove the FS ties the bridge
and the bridge ties the TB, but the bridge's *footnote* schedules (intangibles, debt, leases, tax)
are otherwise taken on faith. Lane 8 ties each mapped footnote back to the PBC workpaper that
supports it, and — where the PBC carries a GL column — to the general ledger. Its flagship check is
the **3-way intangibles reconcile** (compiler's bridge tab ↔ company's Goodwill & Intangibles
rollforward ↔ GL). When **net ties but gross and accumulated amortization each diverge by the same
amount (offsetting)**, it emits a `DIAGNOSTIC` row for the fully-amortized/disposed-asset gotcha —
the compiler still carries assets the GL wrote off (see
[source-of-truth-hierarchy.md](references/source-of-truth-hierarchy.md)). Needs the PBC index
(`build_pbc_index.py`); skipped automatically when no `PBCs/` tree is present. See
[pbc-mapping.md](references/pbc-mapping.md) for the PBC → footnote/bridge-tab map and the
FS-build-vs-audit-evidence split. Produces footnote-level findings only — it does not annotate the PDF.

**Lane 9 (tax provision)** re-performs the provision checks on the tax workpaper the preparer built
(the version the compiler folded into the bridge): statutory tax recomputes from pretax × federal
rate; the statutory→effective rate reconciliation foots to the total provision; total provision ties
the FS income tax expense; the **pretax the provision is built on equals the FS pretax** (the
IS↔FN-07 reconciliation error class — a prior round had a ~$X,XXX.XXK gap here); and the deferred
summary ties the FS deferred table and foots (DTA + DTL − VA = net). See `tie_out_tax_provision.py`.

**Lane 10 (flux review)** is analytical, not a tie-out. The company's flux-review PBC is drafted
early (before the revision rounds), so it goes stale and will *not* tie the final FS — this module
re-baselines it: it recomputes YoY movements on the **final** FS, flags material movers against a
threshold, carries the PBC's comments forward as *directional context* (clearly marked stale, not a
tie), and flags movers with no explanation so they can be drafted. See `build_flux_analysis.py`.

**Provision recompute engine (Increment 1 — permanent differences).** Where Lane 9 ties the
preparer's provision to the FS, this engine *independently recomputes* it and red-boxes drift —
the "recompute, then flag" pattern from Lane 8. `build_tax_treatment_map.py` learns from the last
3 years of provision workbooks (same preparer, same template): for each permanent-difference item
it records recurrence, the (stable) tax-effecting rate, and a **sourcing classification** against
the TB — `tb-exact` (book ties a single GL account to the dollar → recompute straight from
NetSuite), `tb-name-divergent` (account exists but balance differs → an aggregation rule is
needed), or `schedule-derived` (no GL line, e.g. SBC off the cap table) — with a confidence score.
`recompute_tax_provision.py` then recomputes each perm: GL-sourced items get `ties-recomputed`
(fully independent — goodwill amortization, the largest perm, reproduces to the dollar); the rest
get `ties-rate-only` (rate verified, book not yet GL-independent, sourcing gap flagged). Results
land in the **Provision Recompute** tab.

`recompute_deferred_tax.py` (Increment 2) reproduces the FN-07 **deferred** footnote from the
provision's DTA/DTL grouping — every line (NOL, §163j, R&E, intangibles, capitalized commissions,
ROU, …) tied to the FS deferred table, plus the DTA + DTL − VA = net footing — and tracks each
temporary difference across the 3 years, flagging `deferred-dropped` (a balance that was present
then zeroed — confirm it was genuinely released, not a stranded roll-forward line) and
`deferred-new`. Results land in the **Deferred Tax** tab.

`recompute_current_nol.py` (Increment 3) recomputes the **federal taxable-income build** (pretax +
permanent + temporary differences = taxable income) and confirms a loss year carries no federal
cash tax, then tracks **NOL DTA** (federal + state) across the 3 years, flagging a decrease to
explain (utilization / expiry / write-off). Results land in the **Current & NOL** tab.

`recompute_rate_rec.py` (Increment 4 — capstone) assembles the whole statutory→effective rate
reconciliation: statutory tax (recomputed from pretax × rate) + the recomputed permanent
differences (independent) + the remaining reconciling items (state, true-ups, DTA adjustment,
change in VA, rate change, foreign — taken from the preparer's other schedules and flagged
`ties-rate-only`), confirms it **foots to the total provision** and **ties the FS income tax
expense**, and recomputes the effective tax rate. It reports an independence scorecard (items
recomputed from source vs still workpaper-sourced). Results land in the **Rate Rec** tab.

`recompute_state_tax.py` (Increment 5) validates the blended state rate that feeds the provision:
each state's ETR = apportionment factor × state statutory rate (all 52 states foot), the per-state
ETRs sum to the blended state rate, and federal + state × (1 − fed) ties the rate-rec blended rate.
Results land in the **State Tax** tab. The remaining next-tier items are the per-jurisdiction
current split, the NOL vintage roll, and the per-state apportionment factor *from raw sales*
(sourcing / throwback judgment); VA conclusion, UTB, RTP, and transfer pricing stay human-reviewed.
`tax_provision_recompute_poc.py` is the feasibility probe behind this.

**PBC foundation (Phase 0).** `build_pbc_index.py` walks `PBCs/2025 Audit/`, classifies every
file (FS-build vs. the high-volume audit-evidence subset), maps the FS-build PBCs to the footnote /
bridge tab they support, and emits `.work/pbc-index.json`. The **PBC Register** report tab tracks
completeness (which expected FS-build PBCs are present / missing) and flags **possibly-stale** PBCs
whose prep version predates the final bridge (the company's flux-review PBC is the canonical stale
case). The whole `PBCs/` tree and the index are gitignored and dropped from the public mirror.

## Workflow

### Step 1 — Ingest (one command)

```
python scripts/run_tieout.py
```

This orchestrator:
1. Builds `inputs.json` from all sources
2. Runs all eight tie-out lanes → `tie-laneN-*.json` (Lane 8 + PBC index run only when a `PBCs/` tree is present)
3. Renders + OCRs PDF pages with annotations → `_TIEOUT.pdf`
4. Builds the multi-tab Excel exceptions report

### Step 2 — Review the annotated PDF

Open `Tieout/<pdf-name>_TIEOUT.pdf` side-by-side with the clean original. On each face
statement:
- Every current-year value should have a small red **B** mark = tied to bridge
- Every prior-year value should have a small red **PY** mark = tied to FY24 final
- Internal cross-refs marked with green **/**
- **Red boxes around values = exceptions** — investigate these

Conventions match the FY24 final tie-out PDF (`Prior Year Examples/2024/Tieout/vYYYY.M.D_...
Tieout (FINAL).pdf`). See [references/tieout-conventions.md](references/tieout-conventions.md).

### Step 3 — Review the Exceptions report

Open `Tieout/Tieout Exceptions - <version>.xlsx`. The Summary tab shows the breakdown by
status and by lane. The Exceptions tab lists every record requiring human review.

Status legend (full taxonomy in [exception-vs-noise.md](references/exception-vs-noise.md)):
- `exception` — values differ by more than tolerance; investigate
- `restatement` — prior-year column doesn't match last year's issued FS; a real restatement
- `unmapped-account` — (Lane 7) a TB account with a balance not mapped in the bridge "TB
  Mapping" tab — it silently falls off the FS; investigate (likely a new GL account)
- `mapped-to-nothing` / `stale-mapping-target` — (Lane 7) mapping present but no FS target, or
  a target BS line that no longer exists
- `completeness-gap` / `tb-out-of-balance` / `bs-does-not-balance` — (Lane 7) reconciliation /
  integrity / balance-sheet-identity failures
- `missing-on-bridge` — FN table row not auto-matched to bridge tab structure (often a
  labelling mismatch; can be tied manually)
- `missing-on-fy24-pdf` — line not present on prior-year FS (often new disclosure)
- `no-tb-rollup` — bridge line with no TB accounts mapped that is NOT a subtotal (mapping gap)
- `ties-no-tb-needed` — subtotal/equity line with no TB account by design (ties via SOE)
- `ties-caption-changed` — number matches; label differs (confirm the wording change)
- `ties-with-rounding` — within tolerance; treated as tied
- `ties-with-sign-inversion` — magnitude matches but signs differ (FS convention); tied
- `ties-F` / `ties-xF` — (Lanes 5/6) subtotal footed / cross-foot recalculated

## Re-run mode

When a new bridge / PDF lands in `Tieout/`, update the path constants in
`scripts/build_inputs.py` and `scripts/run_tieout.py`, then re-run the orchestrator. The OCR
cache (`.work/ocr-cache.json`) short-circuits the slow OCR pass for unchanged pages — delete it
if the PDF's page contents changed.

## Adapting to a new company / new auditor

This skill encodes one company's structure. To reuse it elsewhere, change these four things
(see [bridge-tb-structure.md](references/bridge-tb-structure.md) for the shape it expects):

1. **Paths** — the input-file constants in `scripts/build_inputs.py` and `scripts/run_tieout.py`
   (the clean PDF, the `.docx` text twin, the bridge `.xlsx`, the consolidated TB, the prior-year
   FS PDF) and the `FY25_VERSION` / output-folder.
2. **GL label aliases** — `LABEL_ALIASES` in `scripts/tie_out_common.py` maps the company's
   chart-of-accounts wording to canonical keys. This is the **biggest** dependency: without
   aliases for the new company's line items, Lane 1 flags most lines as exceptions. Rebuild it
   from the new BS/IS captions.
3. **Bridge tab + TB mapping** — confirm `SECTION_TO_BRIDGE` / `FN_TO_BRIDGE` in
   `tie_out_pdf_to_bridge.py` (Lane 1) match the new bridge's tab names, and that the "TB
   Mapping" tab column headers (`Account`, `FS Mapping BS`, `SCF Mapping`, `FN Mapping (BS)`,
   `FN Mapping (IS)`) match what `load_tb_mapping` in `tie_out_bridge_to_tb.py` expects.
4. **Tickmark vocabulary + colors** — set the auditor's convention in
   [references/tieout-conventions.md](references/tieout-conventions.md) and the mark colors in
   `scripts/annotate_tieout_pdf.py`.

Before sending a draft to the auditor, run all lanes and clear the internal cross-reference
(Lane 4) and mapping-completeness (Lane 7) findings — those pre-empt the consistency comments
auditors raise. See [what-reviewers-commonly-catch.md](references/what-reviewers-commonly-catch.md).

## What this skill does NOT do

(Intentionally narrowed scope, per the user's direction. The *engine* ties numbers; the
*judgment layer* is captured as reference docs, not automated checks.)

- ❌ Automated disclosure completeness / inscope-template coverage
- ❌ Automated External Audit Firm-recurring-themes coverage — but see
  [what-reviewers-commonly-catch.md](references/what-reviewers-commonly-catch.md) and
  [what-reviewers-commonly-catch.md](references/what-reviewers-commonly-catch.md) for the manual
  checklist of what reviewers raise
- ❌ Footnote-narrative review / analytical (year-over-year flux) review — these are human review
- ❌ Automated SuiteQL — but the NetSuite MCP IS used for *manual* source-of-truth verification
  when a compiler schedule diverges from the GL (see
  [source-of-truth-hierarchy.md](references/source-of-truth-hierarchy.md))
- ❌ Memo generation

For the previous broader disclosure-focused review, see `Reviews/` (historical).

## Notes

- Read-only with respect to source files. All outputs go to `Tieout/`.
- The clean PDF is never modified — the annotated version is saved to a `_TIEOUT.pdf` copy.
- The bridge already encodes its own tie-out structure: `Per TB | Topside Entries |
  Adjusted Total | Rounded (000s)`. Lane 2 (bridge ↔ TB) compares the bridge's "Adjusted Total"
  to TB-account rollup totals; any delta represents the bridge's topside entries (which are
  legitimate but should be reviewed).
- SOE rollforward IS tied (Lane 5): per-class equity columns vs the bridge SOE tab, each balance
  row cross-footed (xF), and each year's vertical roll-forward footed (F). SOE closing ↔ BS Total
  Equity is also tied via Lane 4.
- FN tables: face statements are 100% covered; FN tables tie by row label and, when a label
  fails, by the secondary number-match (caption changes become `ties-caption-changed`, not
  exceptions). A genuine value mismatch still surfaces as an exception.
- Lane 7 confirms completeness: if a new GL account is added but not mapped in the bridge, it is
  surfaced as `unmapped-account` rather than silently dropping off the FS.
