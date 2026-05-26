# What Reviewers Commonly Catch (audit-review checklist)

A generic checklist of the categories of issues external auditors and senior reviewers
routinely raise on draft financial statements — distilled from real review cycles. Use it
to self-review a draft *before* it goes to the auditor. Items are grouped by whether an
automated tie-out catches them or whether they need human analytical/judgment review.

> This is an abstracted, anonymized checklist — no client, auditor, or figures. For the
> mechanics of distinguishing a real exception from noise, see `exception-vs-noise.md`;
> for reconciling a compiler's schedule to the underlying ledger, see
> `source-of-truth-hierarchy.md`.

## A. Cross-statement / internal consistency — a tie-out CATCHES these
Run the internal cross-reference + mapping-completeness lanes and clear every item first.
- A balance shown on a face statement disagrees with the same number in a footnote or
  supporting schedule (e.g., income-tax footnote pre-tax loss vs. the income statement).
- Equity unit/share counts in the balance-sheet caption don't match the statement of
  changes in equity (units issued vs. units outstanding).
- A subtotal foots internally but disagrees with the general ledger / subledger.
- The prior-year column wasn't re-tied to last year's issued financials after a restatement.
- A new GL account carries a balance but was never mapped into the statement build — it
  silently falls off the financials.

## B. Source data not yet recorded — fix the cause, not each symptom
- A missing or late journal entry (e.g., stock-based compensation, an accrual) surfaces as
  *several* downstream consistency comments. Confirm all expected entries are booked before
  drafting; one unrecorded entry can generate a chain of review comments.

## C. Analytical / fluctuation review — human review (a tie-out does NOT do this)
- Material year-over-year movements in any line need a documented explanation. Reviewers
  ask "why did this go up/down so much vs. prior year?" Prepare a flux analysis with
  explanations for material P&L and balance-sheet movements in advance.
- A tie-out proves numbers *agree across documents*; it does not explain *why a balance moved*.

## D. Disclosure accuracy / stale carryover language — human review
- Policy and narrative language rolled forward from the prior year must be re-confirmed for
  current-year applicability ("is this sentence still applicable this year?").
- Disclosures for events that ended in a prior period should make clear they no longer apply
  (e.g., a fair-value measurement that lapsed).

## E. Classification / judgment — human review
- Redeemable instruments: confirm redemption terms drive the right classification
  (e.g., mezzanine vs. permanent equity).
- Completeness of year-end accruals (e.g., capital purchases received but not yet invoiced).
- Disclosure of the components of a balance (e.g., debt issuance costs paid).

## F. Presentation / wording — low risk, high frequency
- Terminology used consistently throughout (e.g., the same asset called one name in one
  place and another elsewhere).
- Punctuation, parenthetical labels, table/paragraph ordering.

## G. Procedural / status — not findings
- "Auditor to provide updated opinion," "reviewer to revisit when complete," "pending rest
  of this footnote," "not required for a private company — FYI only."

## How findings typically get resolved (patterns worth reusing)
- **Cross-statement mismatch → correct the caption to match the supporting statement**, not the
  reverse (e.g., a balance-sheet unit/share count fixed to agree with the statement of changes
  in equity).
- **Completeness question → quantify and document the materiality basis** when passing on a
  disclosure (e.g., "$X in AP and $0 in accruals, immaterial — not disclosed"). Record the
  reasoning in the comment thread so it survives review.
- **Terminology question → standardize the term throughout** and keep the financial-statement
  wording in sync with the supporting schedule/bridge.
- **Real numeric exceptions converge over multiple rounds.** Expect a genuine discrepancy to be
  partially corrected, re-reviewed, and only fully tie a draft or two later — track it across
  versions rather than assuming one fix closes it.

## How to use
1. Run the tie-out (all lanes) on the first internal draft; clear every category-A item.
2. Confirm all expected source entries are booked (category B).
3. Prepare flux explanations for material movements (category C).
4. Re-confirm prior-year carryover language (category D) and recurring classification
   judgments (category E).
5. Proofread for terminology and presentation consistency (category F).
