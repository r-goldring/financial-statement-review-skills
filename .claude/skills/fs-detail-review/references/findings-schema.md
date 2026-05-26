# Findings Tracker Schema

The Findings Tracker is an Excel workbook with one row per finding. It is **version-aware** — the same finding can appear across multiple bridge revisions with status updated as resolved. The tracker is the canonical record; the Markdown memo is rendered from it.

## Filename

`Reviews/Findings Tracker - {bridge-version-stamp}.xlsx`

Where `{bridge-version-stamp}` is derived from the bridge filename (e.g., `Partial InScope Updates`, `v5.30`, `v5.30 FINAL`).

## Columns

| Column | Type | Description |
|---|---|---|
| Finding ID | string | `FY25-NNNN`. Sequential within audit year. Stable across versions. |
| Date Identified | date | First date this finding surfaced. Does not change across versions. |
| Bridge Version | string | Bridge filename / version this row reflects. |
| PDF Version | string | FS PDF filename / version this row reflects. Blank if finding is bridge-only. |
| FS Area | enum | BS \| IS \| SCF \| SOE \| FN-{n} \| EBITDA \| TB \| Rollforward \| Process |
| Tab | string | Bridge tab name (or "PDF p.{n}" for PDF-only findings) |
| Cell/Range | string | A1-style reference, e.g., `D14` or `B5:E12`. Blank if narrative. |
| Severity | enum | Critical \| Material \| Minor \| Informational |
| Category | enum | Tie-out \| Disclosure \| Math \| Classification \| FX \| Rollforward \| Subsequent-event \| Prior-year-correction \| External Audit Firm-theme \| Inscope-gap \| Control \| PDF-rendering |
| Description | string | What's wrong. Concise. |
| Evidence | string | Specific dollar amounts, source-of-truth value, observed value, delta. |
| Recommended Action | string | What FS Compilation Partner should do to resolve. |
| Status | enum | Open \| In Discussion \| Resolved \| Deferred \| Not Applicable |
| Resolution Notes | string | What changed when resolved. Blank until resolution. |
| Resolved In Version | string | Bridge version where resolved. Blank until resolved. |
| Resolved Date | date | Date marked resolved. Blank until resolved. |

## Severity definitions

- **Critical** — Must be fixed before issuance. Examples: TB doesn't tie to bridge; PDF doesn't tie dollar-for-dollar to bridge; missing required GAAP disclosure; Recent Acquisition amounts incorrectly consolidated into FY25 BS/IS; rollforward break > materiality.
- **Material** — Should be fixed before issuance. Examples: prior-year footnote omitted without justification; External Audit Firm-recurring-theme item not addressed; rollforward break < materiality but > 5% of account; equity rollforward doesn't explain unit-count change.
- **Minor** — Should be fixed if time permits. Examples: footnote wording inconsistency, presentation polish, footnote cross-reference numbering.
- **Informational** — Worth flagging but not actionable. Examples: new tab added (verify intentional), unusual but explained year-over-year movement.

## Category definitions

- **Tie-out** — Bridge tab disagrees with TB / inscope template / PDF.
- **Disclosure** — Footnote wording, completeness, GAAP requirement.
- **Math** — Sums, percentages, ratios that don't compute.
- **Classification** — Account or line-item classification (e.g., debt vs. equity, current vs. noncurrent).
- **FX** — Foreign-currency translation, FX rate consistency, CTA.
- **Rollforward** — Opening + activity ≠ closing; activity not supported.
- **Subsequent-event** — Post-balance-sheet-date events disclosure (Recent Acquisition is the FY25 case).
- **Prior-year-correction** — Restatement, immaterial correction, or re-classification of prior-year amounts.
- **External Audit Firm-theme** — Item flagged because it's in the FY23/FY24 External Audit Firm recurring-comment set.
- **Inscope-gap** — Required FS Compilation Partner inscope-template table missing or incomplete.
- **Control** — Internal control / SOX-equivalent observation (the FY23+FY24 NetSuite admin access deficiency lives here).
- **PDF-rendering** — PDF page returns no extractable text where content is expected; embedded image; font-encoding issue.

## Status lifecycle

```
Open ──────► In Discussion ──────► Resolved
  │                                    ▲
  ├──► Deferred ──────────────────────┘
  └──► Not Applicable
```

- **Open** — newly identified, not yet acknowledged/discussed
- **In Discussion** — surfaced to FS Compilation Partner / External Audit Firm, awaiting response or fix
- **Resolved** — fixed in a subsequent bridge / PDF version
- **Deferred** — agreed to defer to next year (e.g., immaterial)
- **Not Applicable** — on review, not actually a finding (false-positive from auto-detection)

## Carry-over rules (re-run mode)

When the skill runs against a new bridge version and a prior tracker exists:
1. For each row in prior tracker with status `Open` or `In Discussion`:
   - Re-run the underlying check.
   - If the issue still reproduces → carry forward, update `Bridge Version` and `PDF Version` columns.
   - If the issue no longer reproduces → set `Status = Resolved`, `Resolved In Version = {new version}`, `Resolved Date = today`. Add `Resolution Notes = "Auto-detected resolution; verify"` so the user can confirm.
2. For each new finding from the current run:
   - Append as a new row with new `Finding ID`.
3. Do not modify rows with status `Resolved`, `Deferred`, or `Not Applicable` from prior runs.

## Excel formatting conventions

- Header row: bold, frozen.
- `Severity = Critical` rows: red fill.
- `Severity = Material` rows: orange fill.
- `Status = Resolved` rows: gray strikethrough.
- AutoFilter on all columns.
