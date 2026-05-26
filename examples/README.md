# Example data (fictional Acme Holdings LLC)

Internally-consistent sample inputs/outputs so you can see how the tie-out skill
works before pointing it at your own files. All amounts are fictional and in
thousands.

| File | What it is |
|------|------------|
| `sample-fs.txt` | A text stand-in for the rendered FS PDF you tie out against (face statements: BS, IS, SCF summary). In production you point the skill at your real compiled-statements PDF. |
| `sample-fs-bridge.csv` | The bridge: `Per TB` + `Topside Entries` = `Adjusted Total`, then `Rounded (000s)` for FY25 and FY24. This is the workbook the FS values should tie to. |
| `sample-trial-balance.csv` | Account-level trial balance that rolls up (via the `Bridge Line` column) to the bridge's `Per TB` column. |
| `sample-tieout-exceptions.csv` | Example of the exceptions report the skill produces — one row per tested value with a `Status` (ties / exception / restatement). |

## Deliberate exceptions baked in

So the exceptions report has something to show, three values are intentionally
inconsistent:

1. **Operating expenses** — FS shows 64,000 but the bridge shows 65,000 (Lane 1 exception).
2. **Deferred revenue** — bridge shows 15,000 but the TB rolls up to 14,500 (Lane 2 exception).
3. **SCF Net loss** — cash-flow statement shows (4,000) but the income statement shows (5,000) (Lane 4 internal cross-ref exception).
4. **FY24 Total assets** — the FY24 column shows 120,000 vs a prior-year-issued 118,000 (Lane 3 restatement flag).

Everything else ties exactly (or within the rounding/subtotal tolerance).

## Using your own data

Copy `.claude/skills/fs-detail-review/references/company-context.template.md` to
`company-context.md`, fill it in, and point the pipeline's input paths at your
real bridge workbook, trial balance, and rendered FS PDF. See the skill's
`SKILL.md` for the full input list and the four tie-out lanes.
