# PBC → FS-area / footnote / bridge-tab map

The connective tissue between the **PBC source workpapers** and the **bridge / footnotes**
they support. This is what Lane 8 (`tie_out_pbc_to_bridge.py`) uses to know *which* PBC backs
*which* disclosed number, and what the PBC Register uses to track completeness.

`build_pbc_index.py` walks `PBCs/2025 Audit/` (canonical) and applies the rules below to every
file, emitting `.work/pbc-index.json`. **Check `PBCs/2024 Audit/` first** to learn a workpaper's
structure ("what ties where"), then apply to the 2025 version.

## The FS-build vs. audit-evidence split (why most PBCs are NOT tie-out targets)

A full audit PBC set is ~900 files, but only ~30–40 actually *feed* the bridge/footnotes. The
rest is **testing evidence** (bank statements, AP invoices, JE/revenue samples, walkthroughs, IT,
rep letters) — classified for the register but flagged `fs_relevant=false` so it never clutters
the tie-out.

Precedence in `classify()` is **evidence-first**:
1. Any audit-evidence filename marker (`bank statement`, `invoice`, `sample`, `walkthrough`,
   `confirmation`, `je sample`, `soc1`, `rep letter`, …) → evidence.
2. Area in `EVIDENCE_AREAS` (Cash, AP & Accrued, Revenue, Expenses, IT, Walkthroughs/Controls,
   Administrative, General (INT)) → evidence.
3. Area in `FS_BUILD_AREAS` → apply the keyword rules below to map footnote/bridge tab.
4. Anything else → evidence by default.

> **The Chase trap.** "Chase" is both the term-loan lender *and* the company's bank. Without
> evidence-first precedence, 131 Chase bank-statement PDFs mis-classify as `debt` FS-build. The
> area + evidence-keyword guards keep them out.

## FS-build areas (the only areas that yield tie-out targets)

`FS Compilation Requests (FS Compilation Partner) (old)` — **the full FS-build set** (the compiler's input
folder): trial balances, debt schedules, lease schedules, goodwill & intangibles, capitalized
software, capitalized commissions, 401(k), EBITDA, and `9. Tax Provision`. Plus:
`Tax Provision Requests (Tax Provision Preparer)`, `Financial Statements (YE/INT)`, `Goodwill (YE)`,
`Leases (YE/INT)`, `Debt (YE/INT)`, `General (YE)` (going concern + flux).

## Keyword → (category, FS area, footnote, bridge tab)

First match wins; keyword is matched on the full relative path.

| Category | FS area | Footnote | Bridge tab | Trigger keywords |
|---|---|---|---|---|
| tax | Income taxes | FN-07 | FN- Taxes | tax provision, provision memo, deferred tax, rate rec, 163j, NOL |
| goodwill-intangibles | Goodwill & intangibles | FN-03/FN-04 | FN - Intangibles / FN - PPE | goodwill, intangible |
| cap-software | Property & equipment | FN-03 | FN - PPE | capitalized software, cap software, internal-use software |
| cap-commissions | Deferred commissions | BS | BS / SCF | capitalized commission, cap commission |
| lease | Leases | FN-11 | FN - Leases | ASC 842, lease schedule, lease expense, lease(s) |
| debt | Term loans & line of credit | FN-06 | FN - Debt | Chase, loan amort, term loan, credit agreement, effective interest, debt schedule, debt |
| 401k | Employee benefit plan | FN-13 | FN - 401(k) | 401(k), 401k |
| ebitda | Non-GAAP / EBITDA | — | EBITDA Adj Summary | ebitda |
| flux | Analytical / flux review | — | — | flux, account reconciliation and flux |
| going-concern | Going concern | FN-01 | — | going concern |
| gl-detail | General ledger (source) | — | 2025 TB | GL detail, full GL, JE detail, 2025 general ledger |
| trial-balance | Trial balance (source) | — | TB Mapping | trial balance, TB by …, consolidated TB, consolidating BS/IS, detailed TB |
| dept-mapping | Department mapping (source) | — | — | department mapping |
| ppe | Property & equipment | FN-03 | FN - PPE | fixed asset, PPE, FAR, asset RF |

## Highest-value tie-out: the 3-way intangibles reconcile (FN-03/04)

The goodwill & intangibles PBC (`5. Goodwill & Intangibles Reconciliation`, the subledger rec)
is the source of truth for the intangibles footnote. Lane 8 ties:
**compiler schedule (bridge FN-Intangibles tab) ↔ subledger rec (PBC) ↔ GL (NetSuite).**
This automates the FY25 finding by hand — the compiler's gross and accumulated amortization were
each overstated by **~$X.X million** (offsetting, so net tied) because the schedule still carried
fully-amortized developed technology the GL had written off. See [[source-of-truth-hierarchy]]
for the rule, the exact figures, and the gotcha.

## Staleness (carried into the PBC Register + Lane 8)

A PBC's as-of date / version is extracted into the index. A PBC that predates the final bridge
version may legitimately not tie — the register flags it **possibly-stale** for re-confirmation
rather than asserting current. The **flux-review PBC** is the prime example: drafted ≈January,
before the revision rounds, so it is expected NOT to tie the final FS (Phase 3 re-baselines it
rather than tying it).

Related: [[bridge-tab-taxonomy]] (where each bridge tab lives), [[inscope-template-mapping]]
(FS-line → bridge mapping), [[source-of-truth-hierarchy]] (compiler vs. subledger vs. GL).
