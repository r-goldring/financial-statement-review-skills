# Source-of-Truth Hierarchy

**Rule.** When the compiler's (consultant's) workpaper schedule diverges from the company's
own subledger/GL, **the company's records win.** The financial-statement disclosure must
agree with the underlying ledger, not with a workpaper that has drifted from it.

**The three-way reconcile.** For any balance, confirm it ties across all three:
1. **Compiler schedule** (the bridge workbook tab the consultant maintains — feeds the FS)
2. **Subledger** (the company's own reconciliation, e.g., a subledger reconciliation tool)
3. **General ledger** (the system of record, e.g., NetSuite)

If 2 and 3 agree but 1 disagrees, **1 is stale** — correct the compiler's schedule.

## The fully-amortized-asset gotcha (most common cause of divergence)

A compiler schedule that **keeps fully-amortized (or disposed) assets on the books** at gross
cost with an equal, offsetting accumulated-amortization balance will show the **correct net**
but an **overstated gross and accumulated amortization** (each too high by the same amount).
The footnote then *foots internally* (gross − accum = net ties) yet **disagrees with the GL**,
which has written the asset off entirely.

Watch for this whenever:
- net ties but gross and/or accumulated amortization don't tie to the GL/subledger, and
- the difference in gross equals the difference in accumulated amortization (offsetting).

## Worked example (illustrative)

A compiler schedule shows total intangibles of gross **$55K** / accum **$(27K)** / net **$28K**, while the company's subledger and general ledger both show **$50K** / **$(22K)** / **$28K**. Net agrees, but gross and accumulated amortization are each **$5K too high** — a fully-amortized asset (gross **$5K**, accum **$(5K)**, net **$0**) that the GL wrote off during the year but the compiler's schedule still carries. The fix: remove the disposed asset from the schedule so the disclosed gross and accumulated amortization match the GL.

**How to confirm definitively:** pull the GL balance for the relevant account as of period-end
and, if needed, the period's write-off/disposal journal entries (e.g., via the NetSuite MCP).
If GL = subledger ≠ compiler, the compiler schedule is the outlier.

Related: [[exception-vs-noise]] (offsetting gross/accum is a *real* exception, not noise),
[[bridge-tb-structure]] (where the compiler's schedules live in the workbook).
