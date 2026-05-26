# NetSuite Spot-Check SuiteQL Snippets

Reusable SuiteQL queries for drilling into NetSuite during Phase 7 (risk-based spot checks). Use these **only** when:
- A tie-out failed in Phase 3 or Phase 4
- A External Audit Firm-recurring-theme item needs evidence beyond the bridge
- The user explicitly asks for live NS data

Default tie-out should use the local TB files (already extracted from NS). Live SuiteQL is for anomaly investigation.

These snippets reference the canonical library at `../../../../Monthly-Accounting/.claude/skills/_shared/netsuite-queries.md` — extend rather than duplicate.

## Subsidiary IDs (from `_shared/subsidiary-constants.md`)

| Entity | Subsidiary ID | Full path |
|---|---|---|
| Consolidated | `-2` | `Acme Holdings LLC.` |
| US (Acme Inc.) | (per `_shared/subsidiary-constants.md`) | `Acme Holdings LLC. : Acme, Inc.` |
| Canada | (per `_shared/subsidiary-constants.md`) | `Acme Holdings LLC. : Acme Canada` |
| Netherlands | (per `_shared/subsidiary-constants.md`) | `Acme Holdings LLC. : Acme Holdings B.V.` |
| Poland | (per `_shared/subsidiary-constants.md`) | `Acme Holdings LLC. : Acme Poland` |
| UK | (per `_shared/subsidiary-constants.md`) | `Acme Holdings LLC. : Acme UK` |
| Uruguay | (per `_shared/subsidiary-constants.md`) | `Acme Holdings LLC. : Acme Uruguay` |

Look up exact internal IDs via `_shared/subsidiary-constants.md` at execution time.

## 1. GL detail by account + period + subsidiary

When a TB tie-out fails on a specific account, pull the GL detail to find the variance:

```sql
SELECT
  t.tranid,
  t.trandate,
  t.memo AS tranmemo,
  tl.memo AS linememo,
  t.type,
  bua.fullname AS account,
  d.name AS department,
  s.fullname AS subsidiary,
  tl.foreignamount,
  tl.amount,
  tl.creditforeignamount,
  tl.debitforeignamount
FROM transaction t
JOIN transactionline tl ON t.id = tl.transaction
JOIN account bua ON tl.expenseaccount = bua.id OR tl.account = bua.id
LEFT JOIN department d ON tl.department = d.id
JOIN subsidiary s ON tl.subsidiary = s.id
WHERE bua.acctnumber = '{account_number}'
  AND tl.subsidiary = {subsidiary_id}
  AND t.trandate BETWEEN '{period_start}' AND '{period_end}'
  AND tl.posting = 'T'
ORDER BY t.trandate, t.tranid;
```

## 2. Beginning-balance check

Verify a TB beginning balance matches the prior period's ending:

```sql
SELECT
  bua.fullname AS account,
  bua.acctnumber,
  s.fullname AS subsidiary,
  SUM(CASE WHEN t.trandate <= '{prior_period_end}' THEN tl.amount ELSE 0 END) AS bb_at_prior_period_end
FROM transaction t
JOIN transactionline tl ON t.id = tl.transaction
JOIN account bua ON tl.account = bua.id
JOIN subsidiary s ON tl.subsidiary = s.id
WHERE bua.acctnumber = '{account_number}'
  AND tl.subsidiary = {subsidiary_id}
  AND tl.posting = 'T'
GROUP BY bua.fullname, bua.acctnumber, s.fullname;
```

## 3. Single transaction lookup by ID

When the bridge cites a specific JE that doesn't tie:

```sql
SELECT
  t.tranid,
  t.trandate,
  t.type,
  t.memo,
  bua.fullname AS account,
  d.name AS department,
  s.fullname AS subsidiary,
  tl.amount,
  tl.memo AS linememo,
  t.createdby,
  t.lastmodifiedby
FROM transaction t
JOIN transactionline tl ON t.id = tl.transaction
JOIN account bua ON tl.account = bua.id
LEFT JOIN department d ON tl.department = d.id
JOIN subsidiary s ON tl.subsidiary = s.id
WHERE t.tranid = '{transaction_id}';
```

## 4. Related-party transactions [External Audit Firm-theme]

For Phase 7 related-party drill-down. Replace `{related_party_entity_ids}` with the IC subsidiary IDs from `_shared/subsidiary-constants.md`:

```sql
SELECT
  t.tranid,
  t.trandate,
  t.type,
  t.memo,
  bua.fullname AS account,
  s.fullname AS subsidiary,
  ent.entityid AS counterparty,
  tl.amount
FROM transaction t
JOIN transactionline tl ON t.id = tl.transaction
JOIN account bua ON tl.account = bua.id
JOIN subsidiary s ON tl.subsidiary = s.id
LEFT JOIN entity ent ON t.entity = ent.id
WHERE t.trandate BETWEEN '{period_start}' AND '{period_end}'
  AND (
    bua.acctnumber LIKE '15%'  -- Intercompany receivables
    OR bua.acctnumber LIKE '24%' -- Intercompany payables / convertible notes
    OR ent.id IN ({related_party_entity_ids})
  )
  AND tl.posting = 'T'
ORDER BY t.trandate;
```

## 5. JEs by user [Internal control deficiency check]

For the persistent FY23/FY24 deficiency around NetSuite admin access — pull JEs created and posted by the same user during the audit period to confirm whether the deficiency persists in FY25:

```sql
SELECT
  t.tranid,
  t.trandate,
  t.memo,
  t.createdby AS created_by_id,
  emp_c.entityid AS created_by_name,
  t.lastmodifiedby AS modified_by_id,
  emp_m.entityid AS modified_by_name,
  t.approvalstatus
FROM transaction t
LEFT JOIN employee emp_c ON t.createdby = emp_c.id
LEFT JOIN employee emp_m ON t.lastmodifiedby = emp_m.id
WHERE t.type = 'Journal'
  AND t.trandate BETWEEN '{period_start}' AND '{period_end}'
  AND t.createdby = t.lastmodifiedby
ORDER BY t.trandate DESC;
```

## 6. New accounts since prior year

For the `New Accounts` bridge tab verification:

```sql
SELECT
  acctnumber,
  fullname,
  accttype,
  description,
  datecreated
FROM account
WHERE datecreated >= '{prior_year_end}'
  AND datecreated <= '{current_year_end}'
ORDER BY acctnumber;
```

## 7. Convertible note balance and activity [External Audit Firm-theme]

```sql
SELECT
  t.trandate,
  t.tranid,
  t.type,
  t.memo,
  bua.fullname AS account,
  tl.amount
FROM transaction t
JOIN transactionline tl ON t.id = tl.transaction
JOIN account bua ON tl.account = bua.id
WHERE bua.fullname LIKE '%Convertible%'
  AND tl.posting = 'T'
  AND t.trandate BETWEEN '{period_start}' AND '{period_end}'
ORDER BY t.trandate;
```

## 8. Subsequent events check (post-close transactions)

For Phase 7 subsequent-events confirmation — what was booked after year-end?

```sql
SELECT
  t.trandate,
  t.tranid,
  t.type,
  t.memo,
  bua.fullname AS account,
  s.fullname AS subsidiary,
  tl.amount
FROM transaction t
JOIN transactionline tl ON t.id = tl.transaction
JOIN account bua ON tl.account = bua.id
JOIN subsidiary s ON tl.subsidiary = s.id
WHERE t.trandate > '{current_year_end}'
  AND t.trandate <= '{report_issuance_date}'
  AND tl.posting = 'T'
  AND ABS(tl.amount) > {materiality_threshold}
ORDER BY t.trandate, ABS(tl.amount) DESC;
```

Material post-close JEs touching balance-sheet accounts may need either subsequent-event disclosure or evaluation as a year-end adjustment (Type I subsequent event).

## Execution

Use the MCP NetSuite tools available in this environment:
- `mcp__claude_ai_NetSuite__ns_runCustomSuiteQL` — execute the SQL
- `mcp__claude_ai_NetSuite__ns_getRecord` — pull a specific transaction
- `mcp__claude_ai_NetSuite__ns_getSubsidiaries` — list subsidiaries with IDs

Always include the `{period_start}` / `{period_end}` parameters explicitly. The fiscal year for Acme Corp is calendar (Dec 31 close).
