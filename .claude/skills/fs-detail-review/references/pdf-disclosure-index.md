# FS PDF Disclosure Index

Observed structure of the rendered Financial Statements PDF. Used in Phase 8 (PDF tie-out) to verify (a) every required section is present, (b) page text extracts cleanly, and (c) numbers tie dollar-for-dollar to the bridge.

## Reference structure (FY2025 PDF, 28 pages)

| Page | Section | Bridge tab(s) | Source data | Tie-out check |
|---|---|---|---|---|
| 1 | Title page | — | — | (informational) |
| 2 | Table of Contents | — | — | TOC pages match actual page numbers |
| 3-4 | Independent Auditor's Report (External Audit Firm) | — | External Audit Firm | Names, dates, opinion language |
| 5 | **Consolidated Balance Sheet** | `BS` / `Balance Sheet` | TB | Every line item ties to TB; subtotals math; comparative columns labeled correctly |
| 6 | **Consolidated Statement of Operations and Comprehensive Loss** | `PL` / `Income Statement` | TB | Every line ties to TB; gross profit math; OCI/CTA |
| 7 | **Consolidated Statement of Changes in Members' Equity** | `SOE` / `Equity Rollforward` | Equity Rollforward | Unit counts roll from prior year; net loss agrees with IS; equity class movements supported |
| 8 | **Consolidated Statement of Cash Flows** | `SCF` / `SOCF` | BS movement + IS non-cash | Operating + investing + financing = net change in cash; reconciles to BS |
| 9-10 | (Note 1 begins, may include divider) | — | — | — |
| 10-16 | **Note 1: Description of Business and Significant Accounting Policies** | (narrative) | — | Going concern paragraph (forward 12 months); concentrations (cash + customer); recently issued ASUs |
| 17 | (continued) — AR / allowance | `FN - AR` | TB AR + allowance | AR rollforward, allowance methodology |
| 17-18 | Fair Value Measurements | `FN - Fair Value` | Members Equity / Related Party | **L3 hierarchy table; methodology; L3 rollforward; sensitivity** [External Audit Firm-theme] |
| 19 | Related-Party Convertible Notes | `FN - Related Party` (currently MISSING in FY25 bridge) | Convertible note ledger | **Note balance; FV change in OI(E); related-party identification** [External Audit Firm-theme] |
| 20 | Members' Equity (Senior Convertible / A-1 / A-2 / B / C / D unit classes) | `FN - Members Equity` | Equity Rollforward | Unit counts authorized vs. issued vs. outstanding; year-over-year changes [External Audit Firm-theme: A-2, C] |
| 21 | (divider or continuation) | — | — | — |
| 22 | Property + Equipment, Intangibles, Goodwill | `FN - PPE`, `FN - Intangibles`, `GW & Int GL detail` | FS Compilation #5 (GW & Int RF), #6 (Cap Software) | Rollforward; useful lives; impairment test result [External Audit Firm-theme: GW + intangibles] |
| 23-24 | (continuation; may include Debt) | `FN - Debt` | FS Compilation #7 (Debt Schedules) | Future maturities; effective interest rate; covenants; classification [External Audit Firm-theme: classification] |
| 25 | Operating Leases | `FN - Leases` | FS Compilation #11 (Lease Schedules) | ROU asset / liability balances; weighted-avg term + discount rate; future minimum payments |
| 26-27 | (continuations: SBC / 401k / Tax / EBITDA narrative) | `FN - SBC`, `FN - 401(k)`, `FN - Taxes` | TB + workpapers | [External Audit Firm-theme: tax provision] |
| 28 | **Subsequent Events** | `FN Sub Events` | a recent acquisition memo | **Recent Acquisition (Feb 2026) — full disclosure of acquisition; any post-close debt/equity events** [External Audit Firm-theme: subsequent event] |

## Required-disclosure checklist

For each item, the skill confirms presence in the PDF:

- [ ] Independent Auditor's Report — External Audit Firm, signed, dated, addressed to member, opinion paragraph, basis-for-opinion paragraph, going-concern reference if applicable, RoM (responsibilities of management), AR (auditor's responsibilities), city/state, year-end date
- [ ] Balance Sheet — comparative columns; current vs. noncurrent classification; commitments-and-contingencies parenthetical reference to footnote
- [ ] Statement of Operations — comparative columns; line-item disclosure of D&A, SBC, change in FV; OCI section
- [ ] Statement of Changes in Members' Equity — opens with prior beginning balance; rolls through net loss + OCI + equity transactions to current year ending; **unit counts by class** for each year-end
- [ ] Statement of Cash Flows — indirect method (starting from net loss); operating, investing, financing sections; supplemental disclosures (interest paid, taxes paid, non-cash items)
- [ ] **Note 1 — Description of Business and Significant Accounting Policies**:
  - [ ] Description of Business
  - [ ] Basis of Presentation (consolidated, GAAP)
  - [ ] Going Concern (forward 12 months from issuance date)
  - [ ] Use of Estimates
  - [ ] Foreign Currency
  - [ ] Cash and Cash Equivalents (incl. restricted cash if applicable)
  - [ ] Accounts Receivable + Allowance
  - [ ] Property and Equipment + capitalized software
  - [ ] Right-of-Use Assets (operating leases)
  - [ ] Goodwill
  - [ ] Intangible Assets, Net
  - [ ] Impairment of Long-Lived Assets
  - [ ] Revenue Recognition (ASC 606) + Deferred Commissions (ASC 340-40)
  - [ ] Stock-Based Compensation
  - [ ] Income Taxes
  - [ ] Concentrations of Credit Risk
  - [ ] Concentrations of Sales Risk
  - [ ] Recently Issued / Adopted Accounting Standards
- [ ] Note 2 — Cash and Cash Equivalents (incl. restricted cash if applicable)
- [ ] Note 3 — Accounts Receivable (rollforward of allowance)
- [ ] Note 4 — Fair Value Measurements (L1/L2/L3 hierarchy + L3 rollforward) [External Audit Firm-theme]
- [ ] Note 5 — Property and Equipment, Net
- [ ] Note 6 — Goodwill and Intangible Assets, Net (incl. annual impairment test result) [External Audit Firm-theme]
- [ ] Note 7 — Accrued Expenses and Other Current Liabilities
- [ ] Note 8 — Debt (term loan, convertible notes, future maturities, effective rate)
- [ ] Note 9 — Operating Leases (US + Poland; ROU + liability + maturity schedule)
- [ ] Note 10 — Members' Equity (unit classes, authorized/issued/outstanding, distribution waterfall)
- [ ] Note 11 — Stock-Based Compensation
- [ ] Note 12 — Related-Party Transactions (incl. convertible notes) [External Audit Firm-theme]
- [ ] Note 13 — Income Taxes (provision components, DTA/DTL, ETR rec, valuation allowance) [External Audit Firm-theme]
- [ ] Note 14 — Employee Benefit Plans (401(k))
- [ ] Note 15 — Commitments and Contingencies
- [ ] Note 16 — Subsequent Events [Recent Acquisition, FY25] [External Audit Firm-theme]

(Footnote numbers are illustrative; actual numbering depends on FS Compilation Partner's structure. The skill compares to prior-year final FS for numbering consistency.)

## Page-rendering verification

For each page, the skill checks:
- Text extraction returns ≥ 50 characters of recognizable text.
- If a page returns < 50 chars (e.g., FY25 PDF currently does on pages 7, 9, 13, 15, 21, 23, 24, 26, 27), generate a **PDF-rendering** finding (Material) so the user can:
  - Confirm the page is intentionally blank (e.g., divider), or
  - Confirm the page contains rasterized content that must be visually reviewed, or
  - Run an OCR pass to recover the text.

For pages where the TOC points to specific content (e.g., page 7 should be Statement of Changes in Members' Equity per the TOC), a blank-text result is **Critical** rather than Material.

## Bridge → PDF tie-out

The skill extracts numbers from the PDF tables (BS, IS, SCF, SOE) and matches them to the corresponding bridge tabs to the dollar. Numbers are matched by:
- Line label fuzzy-matching ("Cash and cash equivalents" ↔ "Cash & cash equivalents")
- Column-period matching (current year column vs. prior year column)
- Rounding to nearest $X,XXX.XX (the FS is presented in thousands per typical FS Compilation Partner convention) — verify presentation unit on the BS/IS header

Any mismatch → **Critical**: the issued document doesn't agree with the working file.
