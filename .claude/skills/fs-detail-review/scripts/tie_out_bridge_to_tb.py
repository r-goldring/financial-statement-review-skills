"""Lane 2: Bridge BS/IS ↔ Trial Balance (Consolidated).

Uses the bridge 'TB Mapping' tab to roll TB accounts up to bridge BS / IS / SCF / FN
line items, then compares the rolled-up TB total to the bridge value.

CLI:
  python tie_out_bridge_to_tb.py <inputs.json> <out.json>

Sign convention:
  TB values are raw $ from NetSuite. Sign reflects accounting convention:
    - Assets / Expenses / Losses: positive (debit balance)
    - Liabilities / Equity / Revenue / Gains: negative (credit balance)
  Bridge BS / PL tabs flip signs for presentation:
    - BS Liabilities & Equity: bridge shows positive
    - PL Revenue: bridge shows positive
  So when comparing: take abs() on both sides or look at magnitude only.
"""

import argparse
import json
import sys
from pathlib import Path
from collections import defaultdict

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

SCRIPT_DIR = Path(__file__).parent
sys.path.insert(0, str(SCRIPT_DIR))

from tie_out_common import parse_value, compare, normalize_label, make_record, is_subtotal_label


# Labels that legitimately have NO TB account because they are computed subtotals,
# equity classes (which tie via the SOE rollforward, not the TB), or accumulated-equity
# components. These get the "ties-no-tb-needed" status rather than "no-tb-rollup",
# so they aren't surfaced as exceptions.
def _is_expected_no_tb(bridge_label, is_subtotal):
    """Return (is_expected, reason) — True if this line has no TB account by design."""
    if is_subtotal:
        return True, "subtotal (computed)"
    s = (bridge_label or "").lower()
    # Equity classes (Senior Converted Units, Class A-1, A-2, B, C, D)
    import re as _re
    if _re.search(r"\b(senior converted|class\s*[a-d](-[12])?)\s+units", s):
        return True, "equity class (ties via SOE)"
    if "accumulated other comprehensive" in s:
        return True, "AOCI (ties via SOE)"
    return False, None


def load_tb_mapping(bridge_tabs):
    """Build dict: account_name → {bs: line, scf: line, fn_bs: line, fn_is: line}."""
    tb_map_tab = next((t for t in bridge_tabs if t["name"] == "TB Mapping"), None)
    if not tb_map_tab:
        raise RuntimeError("Bridge has no 'TB Mapping' tab")
    rows = tb_map_tab["rows"]
    # Find header row (containing 'Account')
    header_idx = None
    headers = None
    for i, row in enumerate(rows[:10]):
        if any(isinstance(c, str) and c.strip() == "Account" for c in row if c is not None):
            header_idx = i
            headers = [(c.strip() if isinstance(c, str) else c) for c in row]
            break
    if header_idx is None:
        # Diagnostic
        print(f"  TB Mapping first 5 rows for diagnosis:")
        for i, row in enumerate(rows[:5]):
            print(f"    r{i}: {[type(c).__name__ + ':' + repr(c)[:30] for c in row[:8]]}")
        raise RuntimeError("TB Mapping tab: no 'Account' header found")

    # Locate columns by name
    col = {}
    for i, h in enumerate(headers):
        if h == "Account":
            col["account"] = i
        elif h and "FS Mapping BS" in str(h):
            col["bs"] = i
        elif h and "SCF Mapping" in str(h):
            col["scf"] = i
        elif h and "FN Mapping (BS)" in str(h):
            col["fn_bs"] = i
        elif h and "FN Mapping (IS)" in str(h):
            col["fn_is"] = i

    mapping = {}
    for row in rows[header_idx + 1:]:
        if len(row) <= col.get("account", 0):
            continue
        account = row[col["account"]] if col["account"] < len(row) else None
        if not account or not str(account).strip():
            continue
        account = str(account).strip()
        mapping[account] = {
            "bs": (str(row[col["bs"]]).strip() if "bs" in col and col["bs"] < len(row) and row[col["bs"]] else None),
            "scf": (str(row[col["scf"]]).strip() if "scf" in col and col["scf"] < len(row) and row[col["scf"]] else None),
            "fn_bs": (str(row[col["fn_bs"]]).strip() if "fn_bs" in col and col["fn_bs"] < len(row) and row[col["fn_bs"]] else None),
            "fn_is": (str(row[col["fn_is"]]).strip() if "fn_is" in col and col["fn_is"] < len(row) and row[col["fn_is"]] else None),
        }
    return mapping


def normalize_account_name(name):
    """Strip leading account number to allow matching '111000 - Cash...' vs 'Cash...'."""
    if not name:
        return ""
    s = str(name).strip()
    # If starts with NNNNNN -
    import re
    m = re.match(r"^\d{4,7}\s*-\s*(.+)$", s)
    if m:
        return m.group(1).strip()
    return s


def rollup_tb_by_bs_line(tb_accounts, tb_mapping):
    """For each unique BS line, sum the TB values of all accounts mapped to it.

    Returns dict: bs_line_label → {accounts_used: [...], sum_value: float}.
    """
    rollup = defaultdict(lambda: {"accounts_used": [], "sum_value": 0.0, "unmapped": False})
    unmapped_accounts = []

    for acct in tb_accounts:
        if acct["value"] is None or abs(acct["value"]) < 0.005:
            continue
        # Find mapping: try exact, then normalized
        mapping = tb_mapping.get(acct["account"])
        if not mapping:
            mapping = tb_mapping.get(normalize_account_name(acct["account"]))
        if not mapping:
            unmapped_accounts.append({"account": acct["account"], "value": acct["value"]})
            continue
        bs_line = mapping.get("bs")
        if bs_line:
            rollup[bs_line]["accounts_used"].append({
                "account": acct["account"], "value": acct["value"]
            })
            rollup[bs_line]["sum_value"] += acct["value"]

    return dict(rollup), unmapped_accounts


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("inputs")
    ap.add_argument("out")
    args = ap.parse_args()

    data = json.loads(Path(args.inputs).read_text(encoding="utf-8"))
    bridge_tabs = data["bridge"]["tabs"]
    tb_accounts = data["tb_consolidated"]["accounts"]
    tb_unit = data["tb_consolidated"]["unit"]

    print(f"Loading TB Mapping...")
    tb_mapping = load_tb_mapping(bridge_tabs)
    print(f"  {len(tb_mapping)} mapped accounts")

    print(f"Rolling up TB (consolidated, {len(tb_accounts)} accounts) by BS line...")
    rollup, unmapped = rollup_tb_by_bs_line(tb_accounts, tb_mapping)
    print(f"  {len(rollup)} BS lines populated")
    print(f"  {len(unmapped)} unmapped accounts")
    if unmapped:
        unmapped_total = sum(a["value"] for a in unmapped)
        print(f"  unmapped total = ${unmapped_total:,.2f}")
        print(f"  first 5 unmapped: {[a['account'][:40] for a in unmapped[:5]]}")

    # Now tie each BS line in the bridge BS tab (2025 column only) to the rollup
    bs_tab = next((t for t in bridge_tabs if t["name"] == "BS"), None)
    if not bs_tab:
        raise RuntimeError("Bridge has no 'BS' tab")
    bridge_unit = bs_tab.get("unit", "ambiguous")

    # Extract bridge BS rows: label → value (2025)
    import re
    bs_rows = bs_tab["rows"]
    year_cols = {}
    header_idx = None
    for i, row in enumerate(bs_rows[:8]):
        row_year_cells = {}
        for c_idx, cell in enumerate(row):
            if c_idx == 0:
                continue
            if cell is not None:
                s = str(cell).strip()
                if re.match(r"^20\d{2}$", s):
                    row_year_cells[c_idx] = s
        if len(row_year_cells) >= 2:
            year_cols = row_year_cells
            header_idx = i
            break
    if not year_cols:
        raise RuntimeError("Could not find year header row in bridge BS")

    # We only tie 2025 (current year); 2024 prior year is handled by lane 3
    col_2025 = next((c for c, y in year_cols.items() if y == "2025"), None)
    if col_2025 is None:
        raise RuntimeError("Bridge BS has no 2025 column")

    bridge_lines = []
    for row in bs_rows[header_idx + 1:]:
        if not row:
            continue
        # Find label
        label = ""
        for c_idx in range(min(4, len(row))):
            if row[c_idx] is not None:
                s = str(row[c_idx]).strip()
                if s:
                    label = s
                    break
        if not label:
            continue
        v = parse_value(row[col_2025]) if col_2025 < len(row) else None
        if v is None:
            continue
        bridge_lines.append({"label": label, "value": v, "is_subtotal": is_subtotal_label(label)})

    # Bridge-BS-label → bridge-TB-Mapping-target-label aliases.
    # The bridge's BS tab uses NEW terminology ("Term loans, current") while the
    # TB Mapping tab still maps accounts to OLD targets ("Notes Payable", "Notes
    # Payable, net of current portion"). These are equivalent.
    BRIDGE_BS_ALIASES = {
        "term loans, current": "notes payable",
        "term loans, net of current portion": "notes payable, net of current portion",
    }

    records = []
    for br in bridge_lines:
        bridge_label = br["label"]
        bridge_value = br["value"]
        # Find matching rollup
        rollup_match = rollup.get(bridge_label)
        # Try fuzzy normalized
        if not rollup_match:
            for k, v in rollup.items():
                if normalize_label(k) == normalize_label(bridge_label):
                    rollup_match = v
                    break
        # Try aliased label (e.g., "Term loans, current" → "Notes Payable")
        if not rollup_match:
            alias_target = BRIDGE_BS_ALIASES.get(bridge_label.lower().strip())
            if alias_target:
                for k, v in rollup.items():
                    if normalize_label(k) == normalize_label(alias_target):
                        rollup_match = v
                        break

        if not rollup_match:
            # Distinguish "no TB needed by design" (subtotals, equity components) from
            # "TB rollup missing for a line that should have accounts" (the latter is a
            # real issue worth flagging).
            expected, reason = _is_expected_no_tb(bridge_label, br["is_subtotal"])
            if expected:
                rec = make_record(
                    lane="bridge_to_tb",
                    pdf_section="BS",
                    pdf_label=bridge_label,
                    pdf_year="2025",
                    pdf_value=bridge_value,
                    source_ref="N/A — no TB rollup needed",
                    source_label=None,
                    source_value=None,
                    status="ties-no-tb-needed",
                    is_subtotal=br["is_subtotal"],
                    notes=f"Expected: {reason}",
                )
            else:
                rec = make_record(
                    lane="bridge_to_tb",
                    pdf_section="BS",
                    pdf_label=bridge_label,
                    pdf_year="2025",
                    pdf_value=bridge_value,
                    source_ref="TB-rollup!consolidated",
                    source_label=None,
                    source_value=None,
                    status="no-tb-rollup",
                    is_subtotal=br["is_subtotal"],
                    notes=f"No TB accounts mapped to bridge BS line '{bridge_label}'",
                )
            records.append(rec)
            continue

        # Compare bridge ($K) to TB rollup ($1, take absolute value)
        # Sign convention: assets positive on both; liabilities negative on TB but positive on bridge
        tb_sum = rollup_match["sum_value"]
        # Take absolute on both for sign-agnostic comparison
        delta, status, tolerance, comparison_unit = compare(
            abs(bridge_value), abs(tb_sum), bridge_unit, tb_unit,
            is_subtotal=br["is_subtotal"],
        )
        accounts_used_brief = [a["account"][:40] for a in rollup_match["accounts_used"][:3]]
        if len(rollup_match["accounts_used"]) > 3:
            accounts_used_brief.append(f"+ {len(rollup_match['accounts_used'])-3} more")
        rec = make_record(
            lane="bridge_to_tb",
            pdf_section="BS",
            pdf_label=bridge_label,
            pdf_year="2025",
            pdf_value=bridge_value,
            source_ref=f"TB-rollup!{len(rollup_match['accounts_used'])} accounts",
            source_label=", ".join(accounts_used_brief),
            source_value=tb_sum,
            comparison_unit=comparison_unit,
            delta=delta,
            tolerance=tolerance,
            status=status,
            is_subtotal=br["is_subtotal"],
            notes=None,
        )
        records.append(rec)

    by_status = {}
    for r in records:
        by_status[r["status"]] = by_status.get(r["status"], 0) + 1
    print(f"\nBy status: {by_status}")
    Path(args.out).write_text(
        json.dumps(records, default=str, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    print(f"Wrote {args.out}")


if __name__ == "__main__":
    main()
