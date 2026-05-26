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
- `Tieout/Tieout Exceptions - <version>.xlsx` — Summary tab + Exceptions tab + (optional)
  All Records tab. Every record with status != ties is in Exceptions.
- `Tieout/.work/` — intermediate artifacts: `inputs.json`, `tie-laneN-*.json`, `ocr-cache.json`,
  page renders

**Key references:**
- [tieout-conventions.md](references/tieout-conventions.md) — the FY24 FS Compilation Partner tie-out
  mark vocabulary that the annotator mirrors
- [inscope-template-mapping.md](references/inscope-template-mapping.md) — PDF line items ↔
  bridge tab/row mapping
- [company-context.template.md](references/company-context.template.md) — private context — a recent acquisition,
  subsidiaries, prior-year corrections to watch

## Four tie-out lanes

| Lane | What it ties | Tolerance |
|---|---|---|
| 1. PDF face ↔ Bridge | BS / IS / SOE / SCF rows + FN tables | ±$1K simple, ±$5K subtotal |
| 2. Bridge ↔ TB | Bridge BS lines vs rolled-up TB account values (using bridge "TB Mapping" tab) | ±$1K simple, ±$5K subtotal |
| 3. PDF prior-year col ↔ FY24 final FS | every 2024 value on FY25 PDF vs FY24 final | exact (any delta = restatement) |
| 4. Internal PDF cross-refs | BS Cash ↔ SCF ending; IS Net Loss ↔ SCF Net Loss; BS Total Equity ↔ SOE closing; SCF Begin + Change = End | exact |

Tolerances are always applied **in the comparison unit** (typically $K because PDF is in $K).
Sign-inversion between bridge ($negative for losses) and PDF ($positive in "loss" lines) is
detected and auto-classified as `ties-with-sign-inversion`.

## Workflow

### Step 1 — Ingest (one command)

```
python scripts/run_tieout.py
```

This orchestrator:
1. Builds `inputs.json` from all sources
2. Runs all four tie-out lanes → `tie-laneN-*.json`
3. Renders + OCRs PDF pages with annotations → `_TIEOUT.pdf`
4. Builds Excel exceptions report

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

Status legend:
- `exception` — values differ by more than tolerance; investigate
- `restatement` — FY25 prior-year column doesn't match FY24 final; this is a real restatement
- `missing-on-bridge` — FN table row not auto-matched to bridge tab structure (often a
  labelling mismatch; can be tied manually)
- `missing-on-fy24-pdf` — FY25 line not present on FY24 final (often new disclosure)
- `no-tb-rollup` — bridge subtotal line with no TB accounts mapped (expected for "Total ..."
  lines)
- `ties-with-rounding` — within tolerance; treated as tied
- `ties-with-sign-inversion` — magnitude matches but signs differ (FS convention); tied

## Re-run mode

When a new bridge / PDF lands in `Tieout/`, the user updates the path constants in
`scripts/build_inputs.py` and re-runs the orchestrator. The OCR cache (`.work/ocr-cache.json`)
short-circuits the slow OCR pass for unchanged pages.

## What this skill does NOT do

(Intentionally narrowed scope, per the user's direction.)

- ❌ Disclosure completeness / inscope-template coverage
- ❌ External Audit Firm-recurring-themes coverage check
- ❌ Footnote-narrative review
- ❌ Live SuiteQL to NetSuite
- ❌ Memo generation

For those, see the previous broader review at `Reviews/` (historical).

## Notes

- Read-only with respect to source files. All outputs go to `Tieout/`.
- The clean PDF is never modified — the annotated version is saved to a `_TIEOUT.pdf` copy.
- The bridge already encodes its own tie-out structure: `Per TB | Topside Entries |
  Adjusted Total | Rounded (000s)`. Lane 2 (bridge ↔ TB) compares the bridge's "Adjusted Total"
  to TB-account rollup totals; any delta represents the bridge's topside entries (which are
  legitimate but should be reviewed).
- SOE rollforward tie-out is a known TODO — current year SOE values are not auto-tied because
  the rollforward shape (rows = balance dates, cols = unit classes) is different from
  year-comparison tables. SOE closing balance ↔ BS Total Equity IS tied via Lane 4.
- FN table tie-out is partial — face statements are 100% covered; FN tables tie when their
  row labels match the bridge FN tab. Mismatched labels (e.g., PPE categorization differs
  between FY25 PDF and bridge FN-PPE tab) surface as exceptions.
