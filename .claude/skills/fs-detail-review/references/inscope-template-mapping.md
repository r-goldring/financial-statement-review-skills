# FS Compilation Partner InScope Template Mapping

The FS Compilation Partner "InScope Template Tables" workbook is the canonical structure for what tables must appear in the bridge / final FS. Tab list confirmed from `2025 InScope Template Tables_Copy tables to Bridge as necessary.xlsx`:

`BS, PS, SOE, SCF, AR, PPE Useful Life, Business Combination, Fair Value, Table_9, Table_10, Table_11, Table_12, Table_13, Table_14, Table_15, Table_16, Table_17, Table_18, Table_19, Table_20, Table_21, Table_22, Table_23, Table_24, Table_25, Table_26`

(Table_9 onward are numbered placeholders — the actual title is in cell A1 of each tab. The skill should read A1 to identify each.)

## Required mapping

For each inscope template tab, the bridge must have a corresponding tab AND the rendered FS PDF must have the corresponding section:

| Inscope Tab | Bridge Tab | PDF Section | Notes |
|---|---|---|---|
| BS | `BS` (and `Balance Sheet` for backup) | Balance Sheet | Primary statement |
| PS | `PL` (and `Income Statement` for backup) | Statement of Operations and Comprehensive Loss | "PS" = P&L Statement |
| SOE | `SOE` (and `Equity Rollforward` for backup) | Statement of Changes in Members' Equity | Unit counts critical |
| SCF | `SCF` (and `SOCF` for backup) | Statement of Cash Flows | Indirect method |
| AR | `FN - AR` | Note on Accounts Receivable (incl. allowance rollforward + concentrations) | — |
| PPE Useful Life | `FN - PPE` | Note on Property and Equipment | Useful life by class |
| Business Combination | `FN - Business Combinations` ⚠️ MISSING in current FY25 bridge | Business Combinations note (if applicable) | Even with no current-year acquisition, prior acquisitions may need ongoing disclosure |
| Fair Value | `FN - Fair Value` | Fair Value Measurements note | L1/L2/L3 hierarchy + L3 rollforward [External Audit Firm-theme] |
| Table_9 to Table_26 | (read A1 of each to identify) | (depends on title) | Placeholder tabs — title in cell A1 |

## Verification at runtime

The skill at execution time:
1. Opens `2025 InScope Template Tables*.xlsx`
2. For each tab `Table_9` … `Table_26`: reads cell A1 to get the actual title
3. Confirms the bridge has a tab matching that title (or a recognized synonym)
4. Confirms the FS PDF has a corresponding section/page
5. Any inscope tab without a matching bridge tab → **Inscope-gap finding** (Material)

## Tab titles to expect (typical FS Compilation Partner template, will verify at runtime)

Common tables that appear in `Table_9` through `Table_26`:
- Capitalized software (life by class)
- Goodwill rollforward
- Intangible assets rollforward
- Future amortization (next 5 years + thereafter)
- Debt schedule (future maturities)
- Effective interest rate
- Lease maturity schedule (operating + finance)
- Operating lease cost components (fixed, variable, ST)
- Members' equity unit table (auth/issued/outstanding by class)
- Stock-based compensation expense by classification
- Income tax provision components
- Effective tax rate reconciliation
- Deferred tax assets / liabilities
- 401(k) match expense
- Concentrations (cash, AR, revenue)
- Subsequent events

The skill maps each at runtime by reading A1; this list is illustrative.
