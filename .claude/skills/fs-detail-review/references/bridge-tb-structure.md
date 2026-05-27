# Bridge & Trial-Balance Structural Conventions

What the tie-out engine assumes about the shape of the bridge workbook and the trial balance.
A new company's files will differ; this is the checklist of what to confirm/adjust.

## Bridge column convention (the build-up of each FS number)

Each bridge face/footnote tab builds a reported number left-to-right:

| Column | Unit | Meaning |
|---|---|---|
| `Per TB` | raw $ | unadjusted trial-balance value pulled from the GL |
| `Topside Entries` (a, e, f…) | raw $ | adjusting entries on top of the TB; bracketed-letter refs point to entry detail |
| `Adjusted Total` | raw $ | `= Per TB + Topside Entries` |
| `Rounded (000s)` | $K | `= round(Adjusted Total / 1000)` — **this is the column that should match the FS** |

For prior-year columns the analogue is `As of <PY date> as filed` → `Misstatement Adjustments`
→ `Adjusted <PY date>`. Current year ties to `Rounded (000s)`; prior year ties to `Adjusted <PY>`.

## The two-block side-by-side tab pattern (important)

Many footnote tabs contain **two blocks laid out side-by-side in the same sheet**:
- a **work area** (left) — the compiler's raw build-up, sometimes with mis-aligned labels, and
- a **formatted disclosure block** (right) — the clean, correctly-labeled version that mirrors
  the FS.

Both blocks carry their own year headers (`2025`, `2024`). A naive extractor that reads *all*
year-labeled columns will pull from both blocks and let the right block overwrite the left
(or vice-versa), corrupting the row values. **Fix:** when multiple same-year header columns are
found on one tab, dedupe to the *leftmost* per year, and rely on the secondary number-search
([[exception-vs-noise]]) to recover anything the chosen block misses. Trust the formatted
disclosure block for the reported figure even if the work-area labels look mis-aligned.

## Per-tab unit detection ($K vs raw $, sometimes mixed)

Units vary **by tab, and occasionally within one tab**:
- Read header rows for "(in thousands)" / "$ in thousands" → `$K`.
- Otherwise fall back to a magnitude heuristic (many 7+ digit values → raw `$1`; values mostly
  in the hundreds-of-thousands with no thousands header → likely `$K`).
- Some tabs are nominally raw `$1` but their top "disclosure" rows are already `$K` — when
  number-matching, try **both** interpretations and take the closer one.
- If a tab is genuinely ambiguous, flag it rather than guessing.

## Sign conventions

- **Trial balance:** credits are **negative**, debits positive (assets/expenses positive;
  liabilities/equity/revenue negative).
- **Bridge / FS:** flip liabilities, equity, and revenue to **positive** for presentation.
- When tying TB ↔ bridge, compare **magnitudes** (or expect a sign flip) — a magnitude match
  with opposite sign is `ties-with-sign-inversion`, not an exception.

## Trial-balance format the loader expects
- An "Account" + balance layout; account names like `NNNNNN - Name`.
- NetSuite TB exports interleave **`Total - …` subtotal rows** and section headers — these are
  NOT leaf accounts; the mapping-completeness lane (Lane 7) filters them so they don't
  double-count or false-flag as unmapped.
- Equity components and AOCI tie via the statement of changes in equity (the SOE roll-forward),
  **not** the TB→BS rollup — they're expected to have no TB account.

Related: [[source-of-truth-hierarchy]], [[exception-vs-noise]].
