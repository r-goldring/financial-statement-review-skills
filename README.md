# Financial Statement Tie-Out Automation (Claude Code skill)

A production-grade Claude Code skill that ties out a compiled annual financial-
statement PDF against the underlying bridge workbook and trial balance, finds
every value that doesn't reconcile, and produces an annotated PDF plus a
structured exceptions workbook. Built and used by a senior accountant for annual
audit/compilation prep at a multi-entity company on NetSuite.

> Companion repo: for the monthly close (payroll, accruals, intercompany, bank
> reconciliation, flux), see **[month-end-close-skills](https://github.com/r-goldring/month-end-close-skills)**.
> This repo is the *annual* financial-statement-review counterpart.

## What it does

Given your compiled FS PDF, the bridge workbook behind it, and the trial balance,
the skill runs seven tie-out "lanes" and reports a status for every tested value:

| Lane | Checks | Tolerance |
|---|---|---|
| **1. PDF ↔ Bridge** | Every BS / IS / SOE / SCF face value + footnote table figure against the bridge | ±$1K simple, ±$5K subtotal |
| **2. Bridge ↔ Trial Balance** | The bridge's `Adjusted Total` against the rolled-up TB accounts | ±$1K |
| **3. Prior-year ↔ last year's final FS** | Every prior-year column value against the previously-issued statements (catches restatements) | exact |
| **4. Internal cross-references** | BS Cash ↔ SCF ending cash; IS Net Loss ↔ SCF/SOE; BS Equity ↔ SOE; BS ↔ footnote totals; SCF rollforward | exact / ±$5K subtotal |
| **5. SOE rollforward** | Per-class equity columns, balance-row cross-foot, each year's vertical roll-forward | ±$5K |
| **6. Footing** | Every subtotal in every face/footnote table re-summed from its components | ±$5K |
| **7. Mapping completeness** | Unmapped TB accounts (a new GL account that would fall off the FS), SUM(mapped)=TB, TB integrity, stale mappings, balance-sheet identity | exact |

**Outputs:**
- An **annotated PDF** with tie-out marks (B = tied to bridge, PY = tied to prior year, `/` = internal cross-ref, red boxes around exceptions) mirroring standard compilation-review conventions.
- An **exceptions workbook** with a Summary tab + per-lane review tabs (Bridge Ties, TB Ties, Mapping Completeness, SOE Rollforward, Internal Ties, PY Ties, Footing) + Exceptions + All Records, each sorted findings-first and color-coded; per-value status includes `ties`, `ties-with-rounding`, `ties-with-sign-inversion`, `ties-caption-changed`, `exception`, `restatement`, `unmapped-account`, `missing-on-bridge`, etc.

## Who it's for

Controllers, assistant controllers, and senior accountants who tie out annual
financial statements for an audit or compilation and want to automate the
soul-crushing PDF-to-workbook reconciliation. The logic generalizes to any
multi-entity company that compiles statements from a bridge + trial balance.

## Quick start

1. **Install Claude Code** — [docs.claude.com/en/docs/claude-code](https://docs.claude.com/en/docs/claude-code/quickstart).
2. **Clone:**
   ```bash
   git clone https://github.com/r-goldring/financial-statement-review-skills.git
   cd financial-statement-review-skills
   ```
3. **Install Python deps:**
   ```bash
   pip install openpyxl python-docx PyMuPDF easyocr reportlab Pillow
   ```
4. **Look at the example data** in [`examples/`](examples/) — a fictional Acme Holdings LLC bridge, trial balance, FS text, and the resulting exceptions report. It's internally consistent with four deliberate exceptions so you can see what the skill flags.
5. **Customize for your company:**
   - Copy `.claude/skills/fs-detail-review/references/company-context.template.md` to `company-context.md` and fill in your entity structure, auditor, equity classes, debt, and recurring audit themes.
   - Point the pipeline's input paths (in the orchestrator) at your real bridge workbook, trial balance, current-year FS PDF, and prior-year final FS PDF.

## Apply the patterns to your stack

Even if you don't run the code, the approach transfers:

- **Multi-lane reconciliation** — don't just tie the face statements; also tie the bridge to the TB, the prior-year column to last year's *issued* statements (restatement detection), the statements to themselves (internal cross-refs), and — critically — confirm **completeness**: that every TB account with a balance actually maps into the statements, so a new account can't silently fall off.
- **Tolerance with nuance** — distinguish a true exception from rounding and from a sign inversion; apply a looser tolerance to subtotals than to detail lines.
- **Exceptions, not noise** — the win is a report that flags only what genuinely doesn't tie, broken out by type, so a reviewer reads ten real items instead of a thousand green checkmarks.
- **Carry-forward marks** — reuse last year's verified tie-out marks to focus this year's review on what changed.

## Architecture

```
.claude/skills/fs-detail-review/
├── SKILL.md                 # the skill spec + workflow
└── references/              # bridge taxonomy, tie-out conventions, bridge/TB structure,
                             # exception-vs-noise, source-of-truth hierarchy, what-reviewers-
                             # commonly-catch checklist, company-context template
scripts/                     # (under the skill) the tie-out pipeline:
  run_tieout.py              #   orchestrator
  build_inputs.py            #   ingest PDF / docx / bridge / TB -> inputs.json
  extract_*.py               #   robust PDF + bridge extractors
  tie_out_*.py               #   the seven lanes (PDF/bridge, bridge/TB, prior-year,
                             #   internal cross-refs, SOE, footing, mapping completeness)
  annotate_tieout_pdf.py     #   draw marks on the PDF
  build_exceptions_report.py #   the exceptions workbook
examples/                    # fictional Acme Holdings LLC sample data
```

## License

[MIT](LICENSE).

## Disclaimer

Illustrative, built around fictional "Acme Holdings LLC" data. Tie-out is a
review aid, not a substitute for professional judgment — always have a qualified
accountant review the exceptions and the statements. Run against your own data
only after you've reviewed the configuration.
