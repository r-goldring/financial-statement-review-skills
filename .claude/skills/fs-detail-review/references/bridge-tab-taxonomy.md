# Bridge Tab Taxonomy

Reference inventory of expected bridge-workbook tabs based on the **FY2024 final bridge** (`vYYYY.M.D_Acme 2024 FS Bridge (FINAL).xlsx`, 48 tabs). Used in Phase 2 (tab inventory) to flag MISSING / NEW / PARTIAL tabs in the current-year bridge.

Section dividers (named `--> ...`) are non-data marker tabs; presence is informational only.

## Section 0 — Inscope-aligned summary tabs (added FY25)

These were added in FY25 to mirror the FS Compilation Partner "InScope Template Tables" naming convention (BS, PS, SOE, SCF). They are the streamlined versions; the longer "Section 1" tabs below are the legacy detailed versions. **The skill verifies BS == Balance Sheet, PL == Income Statement, SOE == Equity Rollforward, SCF == SOCF dollar-for-dollar.**

| Tab | Purpose | Ties to |
|---|---|---|
| BS | Inscope-format Balance Sheet | Inscope template `BS` table; bridge `Balance Sheet` |
| PL | Inscope-format Income Statement (P&L) | Inscope template `PS` table; bridge `Income Statement` |
| SOE | Inscope-format Statement of Equity | Inscope template `SOE` table; bridge `Equity Rollforward` |
| SCF | Inscope-format Statement of Cash Flows | Inscope template `SCF` table; bridge `SOCF` |

## Section 1 — Detailed financial statements (legacy)

| Tab | Purpose | Ties to |
|---|---|---|
| Balance Sheet | Detailed BS with account-level mapping | TB account totals, BS tab (Section 0) |
| Income Statement | Detailed IS with account-level mapping | TB P&L accounts, PL tab (Section 0) |
| Equity Rollforward | Equity activity by class of unit | Members Equity FN, SOE tab (Section 0) |
| SOCF | Statement of Cash Flows (indirect) | BS movement + non-cash items in IS |

## Section 2 — TB layers

| Tab | Purpose | Ties to |
|---|---|---|
| `TB -->` | Section divider | (informational) |
| {YYYY} TB | Current-year trial balance | `Trial Balance/Consolidated TB FY{YY}*.xlsx` |
| {YYYY-1} TB | Prior-year (comparative) trial balance | Prior-year final bridge {YYYY-1} TB |
| {YYYY-2} TB | Prior-prior-year trial balance | Prior-prior bridge / FS for restatement check |
| TB Mapping | Account → FS line mapping | Inscope template + chart of accounts |
| New Accounts | Accounts added since last year | NetSuite COA query |
| {YYYY} IS by Department | Departmental P&L cut | TB by department |
| {YYYY-1} IS by Department | Comparative departmental P&L | Prior-year final bridge |
| {YYYY-2} IS by Department | Prior-prior departmental P&L | (if applicable) |

## Section 3 — Cash flow detail by entity

| Tab | Purpose | Ties to |
|---|---|---|
| `SOCF -->` | Section divider | (informational) |
| CF - Cons | Consolidated cash flow detail | SOCF tab |
| CF - US | US entity cash flow detail | US TB |
| CF - BV | Netherlands (B.V.) cash flow detail | Netherlands TB |
| CF - CAD | Canada cash flow detail | Canada TB |
| CF - UK | UK cash flow detail | UK TB |
| CF - PLN | Poland cash flow detail | Poland TB |
| CF - UY | Uruguay cash flow detail (added FY25) | Uruguay TB |

Critical: `Σ entity cash flows == CF - Cons == SOCF`. Spot-check mid-tab.

## Section 4 — All-entity TB stacks

| Tab | Purpose | Ties to |
|---|---|---|
| FY{YY} TB's - All Entities | Stacked TB rows: account × entity | Σ rows = current-year Consolidated TB |
| FY{YY-1} TB's - All Entities | Prior-year stacked TB | Σ rows = prior-year Consolidated TB |

## Section 5 — Footnote support

| Tab | Purpose | Ties to | External Audit Firm-theme |
|---|---|---|---|
| `FNs -->` | Section divider | (informational) | — |
| FN - AR | Accounts receivable + allowance + concentrations | TB AR accounts | Concentrations |
| FN - Fair Value | Fair value hierarchy table | Members Equity, Related Party FNs | **Valuation L3** |
| FN - PPE | Property + equipment by class, depreciation | TB PPE accounts, FS Compilation #6 (Capitalized Software) | — |
| FN - Intangibles | Intangibles by acquisition + amortization | FS Compilation #5 (Goodwill & Intangibles RF) | **Valuation L3** |
| GW & Int GL detail | Underlying GL detail for goodwill + intangibles | NetSuite GL | — |
| Intangibles Amort Sch {YYYY} | Annual amortization schedule | Intangibles tab | — |
| FN - Business Combinations | Historical + current acquisition disclosures | Business Combination Memo | **Valuation, BCs** |
| FN - Debt | Debt schedule, interest rate, future maturities | FS Compilation #7 (Debt Schedules) | Classification |
| FN - Accrued Expenses | Accrued liability components | TB accrued accounts | — |
| FN - Taxes {YYYY} | Tax provision, DTA/DTL, ETR reconciliation | Provision workpaper | **Tax provision** |
| FN - Taxes {YYYY-1} | Prior-year tax footnote (often kept for comparability) | Prior-year provision | — |
| FN - Members Equity | Equity unit classes, authorized/issued/outstanding | Equity Rollforward | **Class A, Class B** |
| FN - SBC | Stock-based compensation | Equity grants, FS Compilation #13 (Capitalized Commissions if applicable) | — |
| FN - Leases | ASC 842 lease tables (US + Poland), maturity schedule | FS Compilation #11 (Lease Schedules) | — |
| FN - Related Party | Related-party transactions including convertible notes | Convertible note ledger | **Related party, valuation L3** |
| FN - 401(k) | Defined contribution plan disclosure | FS Compilation #12 (401k Contributions) | — |
| FN Sub Events | Subsequent events | a recent acquisition (FY25); any post-close debt amendments | **Subsequent event** |

## Section 6 — EBITDA adjustments

| Tab | Purpose | Ties to | External Audit Firm-theme |
|---|---|---|---|
| `EBTIDA Adj -->` | Section divider (note typo) | (informational) | — |
| EBITDA Adj Summary {YYYY} | Current-year add-back categories | FS Compilation #16 (EBITDA Adj Departments) | **EBITDA add-backs** |
| EBITDA Adj Summary {YYYY-1} | Prior-year add-backs (comparative) | Prior-year final bridge | — |

## Section 7 — Entity-level trial balances

| Tab | Purpose | Ties to |
|---|---|---|
| `Entity TBs -->` | Section divider | (informational) |
| Consolidated | Consolidated TB (after eliminations) | `Trial Balance/Consolidated TB FY{YY}*.xlsx` |
| US | Acme Inc. TB | `Trial Balance/TB by Subsidiary*` US |
| Canada | Acme Canada TB | `Trial Balance/TB by Subsidiary*` Canada |
| Netherlands | Acme Holdings B.V. TB | `Trial Balance/TB by Subsidiary*` Netherlands |
| UK | Acme UK TB | `Trial Balance/TB by Subsidiary*` UK |
| Poland | Acme Poland TB | `Trial Balance/TB by Subsidiary*` Poland |
| Uruguay | Acme Uruguay TB (added FY25) | `Trial Balance/TB by Subsidiary*` Uruguay |

## Tab inventory check — output format

For each tab, classify as:
- ✅ **Present** — expected and populated
- 🟡 **Partial** — expected, present, but blank where data is expected (consistent with "Partial InScope Updates")
- ⚠️ **Missing** — was in prior-year final, not in current bridge
- 🆕 **New** — not in prior-year taxonomy; verify intentional
- 📝 **Renamed** — same purpose, different name (e.g., `SOCF` ↔ `SCF`)

The skill prints a 5-column table: `Section | Tab | Status | Prior-year analog | Notes`.
