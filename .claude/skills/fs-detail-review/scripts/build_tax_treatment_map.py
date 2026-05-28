"""Provision recompute engine — Module A: the account -> tax-treatment map.

Learns, from the last 3 years of the provision workbook (same preparer, same
template), how each permanent-difference reconciling item behaves, and — for the
current year — how cleanly its book amount can be **sourced from the general ledger**.

For each perm item it records:
  - treatment / footnote-grouping letter,
  - recurrence (how many of the 3 years it is non-zero) and rate stability,
  - the median tax-effecting rate (tax_effect / book_amount),
  - the *sourcing classification* against the current-year TB:
        tb-exact         — book ties a single GL account to the dollar (auto-recompute
                           straight from NetSuite),
        tb-name-divergent — an account with a matching name exists but the balance
                           differs (an aggregation / adjustment rule is needed),
        schedule-derived  — no TB line (e.g. SBC comes off the cap-table schedule),
  - a confidence score (recurrence x rate-stability x sourcing).

Output: `.work/tax-treatment-map.json` — consumed by recompute_tax_provision.py.
This is the keystone: it's what lets the recompute pull book amounts from the GL
instead of from the preparer's file, and it flags exactly where it can't.

CLI:
  python build_tax_treatment_map.py <inputs.json> <out tax-treatment-map.json>
       [--fy23 <xlsx>] [--fy24 <xlsx>] [--fy25 <xlsx>]
"""

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from tie_out_common import parse_value  # noqa: E402

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

ROOT = Path(r"C:\path\to\financial-statement-review")
DEFAULTS = {
    "FY23": ROOT / "Prior Year Examples/2023/FS Compilation Requests (FS Compilation Partner)/8. (YE) Acme Corp 2023 Tax Provision_v2024.5.23.xlsx",
    "FY24": ROOT / "PBCs/2024 Audit/FS Compilation Requests (FS Compilation Partner)/Tax Provision/2024 Acme Corp Tax Provision_v2025.5.27.xlsx",
    "FY25": ROOT / "PBCs/2025 Audit/Year End/FS Compilation Requests (FS Compilation Partner) (old)/9. Tax Provision/9. Acme Corp 2025 Tax Provision_v2026.5.21.xlsx",
}

# Name hints to find a candidate GL account when the book amount doesn't match by value.
NAME_HINTS = {
    "Meals": ["meals"],
    "Entertainment": ["entertainment"],
    "Gifts": ["gift"],
    "Goodwill Amortization": ["goodwill expense", "goodwill amort"],
    "Stock Based Compensation": ["stock based compensation", "stock-based compensation", "share-based comp"],
    "Fines & Penalties": ["fines", "penalt"],
    "Expired Contribution": ["charit", "contribution"],
    "Life Insurance Premiums": ["life insurance"],
}


def extract_rate_rec_perms(wb):
    """Return {item: (book_amount, tax_effect)} for the permanent-difference block of
    '6| Rate Rec.' (the rows tax-effected at ~the statutory rate), plus the statutory
    federal rate."""
    rows = [list(r) for r in wb["6| Rate Rec."].iter_rows(values_only=True)]
    perms = {}
    fed_rate = None
    in_perms = False
    for row in rows:
        sec = row[1] if len(row) > 1 and isinstance(row[1], str) else None
        detail = row[2] if len(row) > 2 and isinstance(row[2], str) else None
        if sec and "tax at statutory rate" in sec.lower():
            nums = [parse_value(c) for c in row if parse_value(c) is not None]
            fed_rate = next((n for n in nums if 0 < abs(n) < 1), None)
        if sec and "permanent differences" in sec.lower():
            in_perms = True
            continue
        if sec and sec.lower() in ("fed tax credits", "state taxes"):
            in_perms = False
        if in_perms and detail:
            book = parse_value(row[3]) if len(row) > 3 else None
            te = parse_value(row[4]) if len(row) > 4 else None
            if book is not None or te is not None:
                perms[detail.strip()] = (book, te)
    return perms, fed_rate


def load_tb_accounts(inputs):
    """{account_name: value} from the consolidated TB (raw $)."""
    accts = inputs.get("tb_consolidated", {}).get("accounts", [])
    return [(a["account"], a.get("value")) for a in accts if a.get("value") is not None]


def source_against_tb(item, book, tb):
    """Classify how the book amount sources from the TB. Returns
    (method, account, tb_value)."""
    if book in (None, 0):
        return ("zero-or-none", None, None)
    # Exact value match must be to the dollar — a loose % tolerance lets unrelated
    # accounts with a coincidentally-close balance match (e.g. SBC vs. total marketing).
    # Among any dollar-level matches, prefer one whose name is plausible for the item.
    exact = [(name, val) for name, val in tb if abs(val - book) <= 1.0]
    if exact:
        hints = NAME_HINTS.get(item, [])
        compatible = [(n, v) for n, v in exact if any(h in n.lower() for h in hints)]
        if compatible:
            return ("tb-exact", compatible[0][0], compatible[0][1])
        if not hints:
            # No name expectation for this item — accept the dollar match.
            return ("tb-exact", exact[0][0], exact[0][1])
        # We expected a named account but the only dollar match is name-incompatible →
        # treat as coincidental; fall through to the name-hint search below.
    # name-keyword candidate (value will differ)
    for hint in NAME_HINTS.get(item, []):
        for name, val in tb:
            if hint in name.lower():
                return ("tb-name-divergent", name, val)
    return ("schedule-derived", None, None)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("inputs_json")
    ap.add_argument("out_json")
    for yr, p in DEFAULTS.items():
        ap.add_argument(f"--{yr.lower()}", default=str(p))
    args = ap.parse_args()

    import openpyxl
    inputs = json.loads(Path(args.inputs_json).read_text(encoding="utf-8"))
    tb = load_tb_accounts(inputs)

    # Extract perms per year.
    per_year = {}
    fed_rate = 0.21
    for yr in ("FY23", "FY24", "FY25"):
        path = getattr(args, yr.lower())
        if not Path(path).exists():
            print(f"  (missing {yr}: {path})")
            continue
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        perms, fr = extract_rate_rec_perms(wb)
        wb.close()
        per_year[yr] = perms
        if yr == "FY25" and fr:
            fed_rate = fr

    yrs = list(per_year.keys())
    cur = "FY25" if "FY25" in per_year else yrs[-1]

    all_items = []
    for yr in yrs:
        for k in per_year[yr]:
            if k not in all_items:
                all_items.append(k)

    entries = []
    for item in all_items:
        years_nonzero = 0
        rates = []
        for yr in yrs:
            v = per_year[yr].get(item)
            if v and v[0] not in (None, 0):
                if v[1] is not None:
                    rates.append(v[1] / v[0])
                years_nonzero += 1
        recurrence = years_nonzero
        rate_stable = (len(rates) >= 2 and (max(rates) - min(rates) <= 0.01))
        median_rate = sorted(rates)[len(rates) // 2] if rates else None

        cv = per_year.get(cur, {}).get(item)
        cur_book = cv[0] if cv else None
        cur_te = cv[1] if cv else None
        method, acct, tb_val = source_against_tb(item, cur_book, tb)

        # Confidence: recurring + rate-stable + cleanly TB-sourced => high.
        score = 0.0
        score += {3: 0.4, 2: 0.25, 1: 0.1}.get(recurrence, 0.0)
        score += 0.3 if rate_stable else 0.0
        score += {"tb-exact": 0.3, "tb-name-divergent": 0.1}.get(method, 0.0)
        confidence = ("high" if score >= 0.8 else "medium" if score >= 0.5 else "low")

        entries.append({
            "item": item, "treatment": "permanent",
            "recurrence": f"{recurrence}/{len(yrs)}",
            "rate": round(median_rate, 4) if median_rate is not None else None,
            "rate_stable": rate_stable,
            "cur_book": cur_book, "cur_tax_effect": cur_te,
            "source_method": method, "tb_account": acct, "tb_value": tb_val,
            "confidence": confidence, "score": round(score, 2),
        })

    # New/unmapped check: GL expense accounts with a balance that no perm references.
    # (Informational — most expenses are deductible and rightly absent; this just lists
    #  the perm-mapped accounts so a reviewer can confirm nothing nondeductible is missed.)
    mapped_accts = {e["tb_account"] for e in entries if e["tb_account"]}

    out = {
        "federal_rate": fed_rate,
        "years": yrs,
        "perm_items": entries,
        "mapped_tb_accounts": sorted(a for a in mapped_accts if a),
    }
    Path(args.out_json).write_text(json.dumps(out, indent=2, ensure_ascii=False), encoding="utf-8")

    # Report
    print(f"\n=== Tax-treatment map ({', '.join(yrs)}; fed rate {fed_rate:.0%}) ===")
    print(f"  {'item':<30} {'recur':>6} {'rate':>7} {'stable':>7} {'sourcing':<18} {'conf':<7} account")
    for e in sorted(entries, key=lambda x: (-x["score"], x["item"])):
        if e["recurrence"] == "0/%d" % len(yrs):
            continue
        rate = f"{e['rate']:.0%}" if e["rate"] is not None else "-"
        print(f"  {e['item'][:30]:<30} {e['recurrence']:>6} {rate:>7} "
              f"{str(e['rate_stable']):>7} {e['source_method']:<18} {e['confidence']:<7} "
              f"{(e['tb_account'] or '')[:34]}")
    n_exact = sum(1 for e in entries if e["source_method"] == "tb-exact")
    n_recurring = sum(1 for e in entries if e["recurrence"] == f"{len(yrs)}/{len(yrs)}")
    print(f"\n  perms: {len(entries)} | recurring all yrs: {n_recurring} | "
          f"GL-sourceable (tb-exact): {n_exact} | high-confidence: {sum(1 for e in entries if e['confidence']=='high')}")
    print(f"Wrote {args.out_json}")


if __name__ == "__main__":
    main()
