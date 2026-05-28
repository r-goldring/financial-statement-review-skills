# Acme Holdings LLC — FS Tie-Out — Claude Code Operating Instructions

> **Note:** This is the public, illustrative version of an in-production financial-
> statement tie-out repo. "Acme Holdings LLC" is a stand-in for the real company.
> Fill in your own entity structure, auditor, and inputs — see [README.md](README.md)
> and the `company-context.template.md` reference.

## What this repo is

A single Claude Code skill, `fs-detail-review`, that ties out a compiled annual
financial-statement PDF against the bridge workbook and trial balance behind it,
and reports exceptions. It is the annual counterpart to the monthly
`month-end-close-skills` repo.

## The skill

| Skill | Trigger phrases | Status |
|-------|----------------|--------|
| `fs-detail-review` | FS tie-out, financial statement tie-out, tie out the statements, bridge to TB, prior-year restatement check, exceptions workbook, annotate the FS PDF | Production |

## Core operating rules

### 1. Tie-out is a review aid, not an authority
The skill flags what doesn't reconcile. A qualified accountant must review every
exception and exercise judgment. Never represent the tie-out output as assurance.

### 2. Ten checks (eight tie-out lanes + tax + flux modules), distinct tolerances
- Lane 1 (PDF ↔ Bridge): ±$1K simple lines, ±$5K subtotals.
- Lane 2 (Bridge ↔ TB): ±$1K.
- Lane 3 (prior-year ↔ last year's issued FS): exact — any delta is a restatement to investigate.
- Lane 4 (internal cross-refs: BS↔SCF cash, IS↔SCF/SOE net loss, BS equity↔SOE, BS↔footnote totals): exact / ±$5K subtotals.
- Lane 5 (SOE rollforward): per-class equity columns, balance-row cross-foot, year roll-forward.
- Lane 6 (footing): every subtotal in every face/footnote table re-summed from its components.
- Lane 7 (mapping completeness): unmapped TB accounts (a new GL account that would fall off the
  FS), SUM(mapped)=TB reconciliation, TB integrity, stale mappings, and the balance-sheet identity.
- Lane 8 (PBC ↔ bridge): each mapped footnote/bridge tab back to its supporting PBC workpaper, and
  to the GL where the PBC carries a GL column — source-to-disclosure assurance. Flagship check is
  the 3-way intangibles reconcile; flags the offsetting gross/accum (fully-amortized-asset) gotcha.
  Runs only when a `PBCs/` tree + the PBC index are present.
- Lane 9 (tax provision, FN-07): statutory recompute, rate-rec footing, total provision ↔ FS tax
  expense, book pretax ↔ FS, deferred summary ties + foots (DTA + DTL − VA = net).
- Lane 10 (flux review): analytical, not a tie. Recomputes YoY movements on the final FS, carries
  the company's (stale) flux-PBC comments forward as directional context, flags movers with no
  explanation. Lanes 8–10 run only when a `PBCs/` tree + the PBC index are present.
Distinguish `ties`, `ties-with-rounding`, `ties-with-sign-inversion`, `ties-caption-changed`, and
a true `exception` (or `unmapped-account` / `restatement`).

### 3. Company context drives nuance
Read `references/company-context.md` (your filled-in copy of the template) before
reasoning about movements. Expected items (a disclosed acquisition, a known
debt amendment, a planned unit issuance) should not be flagged as surprises;
genuinely unexplained deltas should.

### 4. Inputs (point these at your files)
- Current-year compiled FS PDF
- Current-year FS .docx (if available — better text extraction than PDF)
- The bridge workbook (`Per TB`, `Topside Entries`, `Adjusted Total`, `Rounded (000s)` columns)
- Consolidated trial balance + trial balance by subsidiary
- Prior-year final (issued) FS PDF — for restatement detection

### 5. Outputs
- Annotated PDF (`*_TIEOUT.pdf`) with marks: B (bridge), PY (prior year), `/` (internal cross-ref), red box (exception).
- Exceptions workbook with a Summary tab + per-lane review tabs (Bridge Ties, TB Ties, Mapping
  Completeness, SOE Rollforward, Internal Ties, PY Ties, Footing) + Exceptions + All Records;
  each tab is sorted findings-first and color-coded by status.

### 6. ASCII only in any generated CSV/workbook text
Plain ASCII in machine-read outputs; spell out directional words.

### 7. Privacy
`references/company-context.md` and all real workbooks/PDFs/trial balances stay
in your private repo only. The public mirror ships a template + fictional Acme
examples; the sync tooling gitignores and drops the real data.

## Running it

```bash
pip install openpyxl python-docx PyMuPDF easyocr reportlab Pillow
python .claude/skills/fs-detail-review/scripts/run_tieout.py
# flags: --skip-inputs  --skip-lanes  --skip-annotate  --skip-exceptions  --pages N-M
```

See `.claude/skills/fs-detail-review/SKILL.md` for the full workflow and the
reference docs for the bridge taxonomy, findings schema, and tie-out mark
conventions.
