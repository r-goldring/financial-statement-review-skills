# External Audit Firm Recurring Comment Themes

Compiled from FY2023 and FY2024 External Audit Firm audit results documents and comment-marked draft FS files in `Prior Year Examples/`. Each theme is a coverage check the skill runs during Phase 6 (disclosure completeness) and Phase 7 (risk-based spot checks).

For each theme: what External Audit Firm flagged, what to verify in the current year's bridge + FS, where the source data lives.

---

## 1. Valuation — Level 3 fair value [highest frequency]

**Pattern**: External Audit Firm raised valuation comments in both FY23 and FY24 across multiple Level-3 instruments. FY24 specifically corrected an FY23 valuation error on Class A units and convertible notes (dual-year disclosure).

**Verify in current year:**
- [ ] Fair Value Measurements footnote (typically FN3 or FN4) presents the **Level 1 / Level 2 / Level 3 hierarchy table** with year-end balances.
- [ ] **Methodology disclosed**: Monte Carlo, DCF, or excess-earnings — name the technique.
- [ ] **Significant unobservable inputs** disclosed: discount rate range, volatility, expected term, projected cash flows.
- [ ] **Roll-forward of Level 3 balances** for the year (beginning balance + transfers + gains/losses + purchases/issuances + settlements = ending balance).
- [ ] If material: **sensitivity analysis** or qualitative discussion of input sensitivity.
- [ ] Compare unit counts and per-unit values year-over-year — flag any restated prior-year amounts.

**Bridge tabs**: `Fair Value`, `Members Equity`, `SBC`, `Related Party` (for convertible notes).
**PDF location**: typically pages 17-19 in the FY25 PDF.

---

## 2. Convertible notes (related party) [recurring]

**Pattern**: Related-party convertible notes carried at Level-3 fair value; External Audit Firm focused on completeness and methodology. FY24 disclosed a June 2024 conversion as a subsequent event.

**Verify in current year:**
- [ ] Related-party convertible note balance reconciles to TB and to the lender's confirmation (if any).
- [ ] Change in fair value flowing through Other income (expense) is consistent with valuation methodology.
- [ ] Related-party disclosure footnote names the parties, terms, and conversion features.
- [ ] Any conversion event during the year is disclosed in the equity rollforward AND the related-party footnote.

**Bridge tabs**: `Related Party`, `Debt`, `Members Equity`, `Fair Value`.

---

## 3. Class A unit valuation [recurring]

**Pattern**: Specifically flagged in FY23 and corrected in FY24. Watch for re-emergence.

**Verify in current year:**
- [ ] Class A unit count tied between BS, equity rollforward, and any compensation/SBC notes.
- [ ] Per-unit fair value supported by valuation memo.
- [ ] If unit count changes year over year, equity rollforward explains the activity (issuances, redemptions, exchanges).

**Bridge tabs**: `Members Equity`, `SBC`, `Fair Value`.

---

## 4. Goodwill and intangible assets [recurring]

**Pattern**: AcquiredCo A acquisition's intangibles and goodwill were a focus in FY23 and FY24. Annual impairment testing required.

**Verify in current year:**
- [ ] Goodwill rollforward by acquisition (Acme, AcquiredCo D, AcquiredCo B, AcquiredCo C, AcquiredCo A): opening + additions + impairment = closing.
- [ ] Intangible assets rollforward by class: opening + additions + amortization + impairment + FX = closing.
- [ ] Annual goodwill impairment test performed and documented (memo or disclosure).
- [ ] Useful lives disclosed and consistent with prior year.
- [ ] **Recent Acquisition (FY25)**: must NOT appear in goodwill/intangibles rollforward as a FY25 movement. Should appear in FY26 next year. (See [company-context.template.md](company-context.template.md).)

**Bridge tabs**: `Goodwill`, `Intangibles`, `Business Combinations`.
**PDF location**: typically page 22 in the FY25 PDF.
**Source**: `FS Compilation Requests (FS Compilation Partner)/5. Goodwill & Intangibles Rollforward*.xlsx`.

---

## 5. Tax provision [recurring]

**Pattern**: External Audit Firm comments on calculation adequacy, entity-by-entity status, and accrual vs. cash basis assessments.

**Verify in current year:**
- [ ] Provision tied between bridge `Taxes` tab and TB.
- [ ] Per-entity tax positions disclosed (US federal, US states, Canada, Netherlands, Poland, UK, Uruguay).
- [ ] Deferred tax asset valuation allowance disclosed.
- [ ] Effective tax rate reconciliation (statutory → effective) presented.

**Bridge tabs**: `Taxes`, prior-year `2023 TB` / `2024 TB` for comparability.

---

## 6. Internal control deficiency — NetSuite admin access [persistent]

**Pattern**: Significant Deficiency Letter issued in BOTH FY23 and FY24 for NetSuite admin access by accounting staff (no secondary approval on JEs). Watch FY25 for either resolution or repeat.

**Verify in current year:**
- [ ] Check whether the deficiency was remediated or persists.
- [ ] If persists, expect another Significant Deficiency Letter; ensure management response is prepared.
- [ ] FS itself does not need to disclose this (private company), but the management letter / deficiency letter does.

**Source**: management letter / deficiency letter (separate file from FS).

---

## 7. Going concern [annual]

**Pattern**: No exceptions either year, but supporting memo and management evaluation always required.

**Verify in current year:**
- [ ] Going concern paragraph in Note 1 (Description of Business / Significant Accounting Policies).
- [ ] Forward 12-month evaluation period stated explicitly (FY25 PDF references "December 31, 2026").
- [ ] Going concern memo exists in working papers and supports the conclusion.
- [ ] If conditions changed (e.g., debt covenant amendment, cash position), discussion updated accordingly.

**Bridge tabs**: standalone going-concern memo file, plus narrative in Note 1.
**Source**: prior years had `Going Concern Analysis Support/` folder under FS Compilation Requests.

---

## 8. Prior-period corrections [pattern of dual-year disclosure]

**Pattern**: FY23 audit disclosed an FY22 equity adjustment in the opinion; FY24 audit corrected an FY23 valuation error. External Audit Firm accepts these but documents them prominently.

**Verify in current year:**
- [ ] Any prior-year amounts that have been restated → explicit footnote disclosing the correction, reason, and effect on prior-year statements.
- [ ] Comparability footnote (if any) addresses reclassifications.
- [ ] No silent restatements — every change to a prior-year column should have a paper trail.

**Bridge tabs**: `2024 TB` (prior-year column), `Equity Rollforward`.

---

## 9. EBITDA add-backs [recurring]

**Pattern**: Bridge has dedicated `EBITDA Adj` summary tabs. External Audit Firm scrutinizes what qualifies as a non-recurring add-back.

**Verify in current year:**
- [ ] EBITDA reconciliation (if presented as supplemental info) ties to IS.
- [ ] Add-backs categorized and supported (acquisition-related, severance, legal settlements, etc.).
- [ ] Add-back categories consistent with prior year; new categories explained.

**Bridge tabs**: `EBITDA Adj Summary 2025`, `EBITDA Adj Summary 2024`.
**Source**: `FS Compilation Requests (FS Compilation Partner)/16. 2025 EBITDA Adjustments Departments.xlsx`.

---

## 10. Presentation, classification, comparability

**Pattern**: External Audit Firm comments on debt vs. equity classification, operating vs. non-operating, footnote cross-references, and consistency across periods.

**Verify in current year:**
- [ ] Debt classification (current vs. noncurrent) per scheduled paydowns + covenants.
- [ ] Operating vs. non-operating items consistent year over year.
- [ ] Footnote cross-references in PDF resolve correctly (FN numbers match).
- [ ] Comparative-period column labels and totals match prior-year final FS.

**Bridge tabs**: `Debt`, `BS`, `IS`, all rollforwards.

---

## 11. Subsequent events [annual]

**Pattern**: Always evaluated through issuance date. FY25 has Recent Acquisition (Feb 2026 acquisition).

**Verify in current year:**
- [ ] Subsequent events note evaluated through the right date (issuance date, typically May/June).
- [ ] a recent acquisition fully described: parties, consideration, business acquired, purchase accounting status.
- [ ] Any debt amendments, conversions, equity transactions between Dec 31 and issuance disclosed.
- [ ] Term loan amendment (Chase First Amendment dated 2025.12.19) — note this is **before** year-end so not subsequent; verify it's reflected in the year-end debt position, not the subsequent events note.

**Bridge tabs**: `Subsequent Events`.
**PDF location**: page 28 in the FY25 PDF.

---

## 12. Concentrations of credit risk and revenue [annual]

**Pattern**: Standard private-company disclosure; External Audit Firm checks completeness.

**Verify in current year:**
- [ ] Cash concentration (% with single financial institution above FDIC).
- [ ] Customer concentration (single customer > 10% of revenue or AR).

**PDF location**: Note 1 (Description of Business and Significant Accounting Policies), typically pages 14-16.

---

## Coverage matrix output

The Phase 6 / Phase 7 output includes a coverage matrix for each theme above:

| Theme | Status | Bridge Tab | PDF Page | Notes |
|---|---|---|---|---|
| 1. Valuation Level 3 | ✅ / 🟡 / ⚠️ | ... | ... | ... |
| 2. Convertible notes | ... | ... | ... | ... |
| ... | ... | ... | ... | ... |

✅ addressed | 🟡 partial | ⚠️ gap
