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

### 2. Four lanes, distinct tolerances
- Lane 1 (PDF ↔ Bridge): ±$1K simple lines, ±$5K subtotals.
- Lane 2 (Bridge ↔ TB): ±$1K.
- Lane 3 (prior-year ↔ last year's issued FS): exact — any delta is a restatement to investigate.
- Lane 4 (internal cross-refs): exact.
Distinguish `ties`, `ties-with-rounding`, `ties-with-sign-inversion`, and a true `exception`.

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
- Exceptions workbook (Summary / Exceptions / All Records).

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
