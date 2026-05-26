# Real Exception vs. Noise

**Rule.** A number that doesn't match on first pass is usually *noise* (different wording,
rounding, units, or sign convention), not a real exception. Before flagging, try to rescue the
tie by **matching on the number, not the label.** Only a number that can't be reconciled by any
of these is a genuine exception worth a reviewer's time.

## The four noise sources (and how the engine handles each)

1. **Caption / wording change** — same number, different label (e.g., the bridge calls a line
   "Total payments" where the FS says "Outstanding borrowings"; the prior-year FS said
   "…and restricted cash" where the current year dropped it). → **Secondary number-match**:
   when a label lookup fails, search the source for the *value* (with label-context scoring);
   if found, classify `ties-caption-changed` and surface the wording delta for review rather
   than calling it an exception. This rescued ~50 would-be "exceptions" in one round.
2. **Rounding** — the bridge/TB carries more precision than the $K-rounded FS (e.g., bridge
   $241.22K vs FS $241K). → tolerance applied in the *display unit*: `ties-with-rounding`
   within ±$1K simple / ±$5K subtotal.
3. **Units** — a tab is in raw dollars where the FS is in $K (or mixed within one tab). →
   detect and convert per source before comparing; never compare raw-$ to $K. See
   [[bridge-tb-structure]].
4. **Sign convention** — TB stores credits negative; the bridge/FS flips liabilities, equity,
   and revenue positive for presentation. → magnitude match with opposite sign →
   `ties-with-sign-inversion`, not an exception.

## When it IS a real exception
- The value appears **nowhere** in the source within tolerance (after the number search).
- **Offsetting gross/accum that nets correctly but disagrees with the GL** — net-only agreement
  hides a real disclosure error. See [[source-of-truth-hierarchy]].
- A **prior-year column** differs from last year's *issued* financials (restatement) — surface
  even small differences; PY is held to exact match.
- A **subtotal that doesn't foot**, or an internal cross-reference (BS↔SCF cash, BS↔SOE equity,
  IS↔FN) that doesn't agree — these are the comments auditors raise (and did).

## Status taxonomy (what each means for the reviewer)
| Status | Meaning | Reviewer action |
|---|---|---|
| `ties` | exact match | none |
| `ties-with-rounding` | within display-unit tolerance | none |
| `ties-with-sign-inversion` | magnitude matches, sign flipped by convention | none |
| `ties-caption-changed` | number matches; label differs | confirm the wording change is intended |
| `ties-no-tb-needed` | subtotal/equity line with no TB account by design | none |
| `exception` | number doesn't reconcile | investigate — likely real |
| `restatement` | PY column differs from issued PY FS | investigate — likely real |
| `unmapped-account` / `completeness-gap` | a balance never reached the FS | investigate — likely real |

**Bias:** prefer rescuing a tie over raising a false exception. A report full of false positives
gets ignored; a short list of real ones gets acted on.
