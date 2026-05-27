"""Lane 7: Mapping Completeness & Reasonableness.

Surfaces the *contextual* checks that the numerical tie-out lanes (1-6) don't —
the "did everything make sense / did the bridge roll-ups stay complete / did any
new account fall off the FS" layer:

  1. Unmapped TB accounts — a GL account with a balance that isn't mapped to ANY
     FS line in the bridge's "TB Mapping" tab (a new account that fell through, or
     an account mapped to nothing). These silently drop off the financial statements.
  2. TB-to-FS completeness reconciliation — SUM(all leaf accounts) vs SUM(mapped),
     so the dollars NOT flowing to the FS are quantified in one number.
  3. TB integrity — sum of all leaf accounts should net to ~0 (debits = credits).
  4. Stale bridge mappings — a "TB Mapping" BS target with no matching bridge BS row.
  5. Balance-sheet identity — bridge Total assets = Total liabilities and members' equity.

Findings are emitted as NON-tie statuses into the exceptions report's "Mapping
Completeness" tab. Lane 7 does NOT annotate the PDF (these are TB-account findings,
not PDF-value ties).

CLI:
  python tie_out_mapping_completeness.py <inputs.json> <out.json>
"""

import argparse
import json
import re
import sys
from pathlib import Path

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

SCRIPT_DIR = Path(__file__).parent
sys.path.insert(0, str(SCRIPT_DIR))

from tie_out_common import parse_value, normalize_label, make_record, is_subtotal_label
# Reuse Lane 2's mapping loader, account-name normalizer, and the "no TB needed by
# design" classifier (subtotals + equity classes + AOCI tie via SOE, not the TB rollup).
from tie_out_bridge_to_tb import load_tb_mapping, normalize_account_name, _is_expected_no_tb


def is_tb_subtotal_row(account_name):
    """NetSuite TB exports interleave 'Total - NNNNNN - Name' subtotal rows and
    section headers; these are NOT real leaf accounts. Excluding them avoids
    double-counting and false 'unmapped' flags (the earlier −$377M noise)."""
    s = (account_name or "").strip().lower()
    if not s:
        return True
    if s.startswith("total"):  # "Total - 111000 - ...", "Total Assets", "Total Other Assets"
        return True
    return False


def account_is_mapped(account_name, tb_mapping):
    """Return (is_in_mapping, has_any_target). Try exact name then number-stripped."""
    m = tb_mapping.get(account_name)
    if m is None:
        m = tb_mapping.get(normalize_account_name(account_name))
    if m is None:
        return False, False
    has_target = any(m.get(k) for k in ("bs", "scf", "fn_bs", "fn_is"))
    return True, has_target


def get_bridge_bs_value(bs_tab, target_label):
    """Find a bridge BS row by label and return its 2025-column value."""
    if not bs_tab:
        return None
    rows = bs_tab["rows"]
    year_col = None
    for row in rows[:8]:
        for c_idx, cell in enumerate(row):
            if c_idx == 0 or cell is None:
                continue
            if re.match(r"^20\d{2}$", str(cell).strip()) and str(cell).strip() == "2025":
                year_col = c_idx
                break
        if year_col is not None:
            break
    if year_col is None:
        return None
    target_norm = normalize_label(target_label)
    for row in rows:
        label = ""
        for c_idx in range(min(4, len(row))):
            if row[c_idx] and str(row[c_idx]).strip():
                label = str(row[c_idx]).strip()
                break
        if label and normalize_label(label) == target_norm:
            v = parse_value(row[year_col]) if year_col < len(row) else None
            if v is not None:
                return v
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("inputs")
    ap.add_argument("out")
    args = ap.parse_args()

    data = json.loads(Path(args.inputs).read_text(encoding="utf-8"))
    bridge_tabs = data["bridge"]["tabs"]
    tb_accounts = data["tb_consolidated"]["accounts"]
    tb_unit = data["tb_consolidated"].get("unit", "$1")

    tb_mapping = load_tb_mapping(bridge_tabs)
    print(f"Loaded TB Mapping with {len(tb_mapping)} mapped accounts")
    print(f"Trial balance: {len(tb_accounts)} rows (incl. subtotals)")

    records = []

    # ===== Check 1: unmapped / mapped-to-nothing leaf accounts =====
    total_tb = 0.0          # net sum of ALL leaf accounts (for integrity check)
    total_mapped = 0.0      # net sum of leaf accounts that DO reach the FS
    gross_unmapped = 0.0    # sum of ABS values that fail to flow (no net cancellation)
    n_unmapped = n_mapped_no_target = 0

    for acct in tb_accounts:
        name = acct.get("account")
        val = acct.get("value")
        if val is None:
            continue
        if is_tb_subtotal_row(name):
            continue
        total_tb += val
        in_map, has_target = account_is_mapped(name, tb_mapping)
        if in_map and has_target:
            total_mapped += val
            continue
        # Below the materiality floor: count toward totals but don't raise a finding
        if abs(val) < 1.0:
            continue
        # Genuine gap — emit a finding
        if not in_map:
            status = "unmapped-account"
            note = ("Account has a balance but is ABSENT from the bridge 'TB Mapping' tab — "
                    "it will not flow to the FS. Likely a new GL account added this year.")
            n_unmapped += 1
        else:
            status = "mapped-to-nothing"
            note = ("Account is in 'TB Mapping' but all FS-mapping columns (BS/SCF/FN) are blank — "
                    "balance does not reach the FS.")
            n_mapped_no_target += 1
        gross_unmapped += abs(val)
        records.append(make_record(
            lane="mapping_completeness",
            pdf_section="TB",
            pdf_label=name,
            pdf_value=val,
            source_ref="bridge!TB Mapping",
            source_label=None,
            source_value=None,
            comparison_unit=tb_unit,
            status=status,
            notes=note,
        ))

    # ===== Check 2: TB-to-FS completeness reconciliation =====
    # Use GROSS unmapped (sum of abs) so an unmapped asset and an unmapped liability
    # of equal size can't net to zero and hide two real gaps.
    records.append(make_record(
        lane="mapping_completeness",
        pdf_section="TB",
        pdf_label="TB-to-FS completeness reconciliation (gross unmapped)",
        pdf_value=gross_unmapped,
        source_ref="SUM(|unmapped leaf accounts|)",
        source_label=f"{n_unmapped + n_mapped_no_target} account(s) not flowing to FS",
        source_value=0.0,
        comparison_unit=tb_unit,
        delta=gross_unmapped,
        tolerance=1.0,
        status=("ties" if gross_unmapped < 1.0 else "completeness-gap"),
        is_subtotal=True,
        notes=(f"Net TB(leaf)={total_tb:,.2f}; net mapped={total_mapped:,.2f}. "
               f"GROSS dollars not reaching the FS = {gross_unmapped:,.2f} across "
               f"{n_unmapped + n_mapped_no_target} account(s). Should be $0."),
    ))

    # ===== Check 3: TB integrity (debits = credits) =====
    records.append(make_record(
        lane="mapping_completeness",
        pdf_section="TB",
        pdf_label="TB integrity (debits = credits)",
        pdf_value=total_tb,
        source_ref="SUM(all leaf accounts)",
        source_label="should net to 0",
        source_value=0.0,
        comparison_unit=tb_unit,
        delta=total_tb,
        tolerance=1.0,
        status=("ties" if abs(total_tb) < 1.0 else "tb-out-of-balance"),
        is_subtotal=True,
        notes=("Sum of all leaf-account balances should net to ~0 for a complete TB. "
               "A non-zero net may indicate a balance-sheet-only extract (P&L closed to "
               "equity) rather than an error — review before treating as a finding."),
    ))

    # ===== Check 4: stale bridge mappings (BS targets with no bridge row) =====
    # A target is only "stale" if BOTH (a) no bridge BS row matches it even loosely,
    # AND (b) real dollars actually route to it. This suppresses noise from:
    #   - verbose equity labels (bridge appends unit counts: "Senior Converted Units,
    #     NN,NNN units…" — matched by prefix), and
    #   - dormant $0 targets (Line of credit, Convertible notes - current, Intercompany).
    # The known "Notes Payable"→"Term loans" relabel is resolved via alias (not stale).
    bs_tab = next((t for t in bridge_tabs if t["name"] == "BS"), None)
    bridge_bs_norms = []
    if bs_tab:
        for row in bs_tab["rows"]:
            for c_idx in range(min(4, len(row))):
                cell = row[c_idx]
                if cell and isinstance(cell, str) and cell.strip():
                    bridge_bs_norms.append(normalize_label(cell.strip()))
                    break

    # Known TB-Mapping-target → bridge-BS-label aliases (mirror of Lane 2's relabels)
    TARGET_ALIASES = {
        "notes payable": "term loans, current",
        "notes payable, net of current portion": "term loans, net of current portion",
    }

    def target_matches_bridge(target):
        cands = [target]
        alias = TARGET_ALIASES.get(target.lower().strip())
        if alias:
            cands.append(alias)
        for cand in cands:
            cn = normalize_label(cand)
            for bn in bridge_bs_norms:
                # exact, or bridge label starts with target (verbose equity labels),
                # or target starts with bridge label (rare reverse case)
                if cn == bn or bn.startswith(cn) or cn.startswith(bn):
                    return True
        return False

    # Dollars routed to each BS target (skip subtotals/None)
    target_dollars = {}
    for acct in tb_accounts:
        name, val = acct.get("account"), acct.get("value")
        if val is None or is_tb_subtotal_row(name):
            continue
        m = tb_mapping.get(name) or tb_mapping.get(normalize_account_name(name))
        if m and m.get("bs"):
            target_dollars[m["bs"]] = target_dollars.get(m["bs"], 0.0) + val

    bs_targets = {m["bs"] for m in tb_mapping.values() if m.get("bs")}
    n_stale = 0
    for tgt in sorted(bs_targets):
        # Equity classes / AOCI / subtotals tie via the SOE rollforward (Lane 5), not the
        # TB→BS rollup — a stale-mapping flag on them is not meaningful. Skip (matches
        # Lane 2's _is_expected_no_tb treatment).
        expected_no_tb, _ = _is_expected_no_tb(tgt, is_subtotal_label(tgt))
        if expected_no_tb:
            continue
        if target_matches_bridge(tgt):
            continue
        routed = target_dollars.get(tgt, 0.0)
        if abs(routed) < 1.0:
            continue  # dormant mapping (no real dollars) — not a finding
        n_stale += 1
        records.append(make_record(
            lane="mapping_completeness",
            pdf_section="TB",
            pdf_label=f"Stale mapping target: '{tgt}'",
            pdf_value=routed,
            source_ref="bridge!TB Mapping (FS Mapping BS)",
            source_label=tgt,
            source_value=None,
            comparison_unit=tb_unit,
            status="stale-mapping-target",
            notes=(f"'TB Mapping' routes {routed:,.2f} to this BS line, but no matching row "
                   "exists in the bridge BS tab. Mapping may be stale (line renamed/removed)."),
        ))

    # ===== Check 5: balance-sheet identity =====
    total_assets = get_bridge_bs_value(bs_tab, "Total assets")
    total_liab_eq = get_bridge_bs_value(bs_tab, "Total liabilities and members' equity")
    if total_assets is not None and total_liab_eq is not None:
        bs_delta = total_assets - total_liab_eq
        records.append(make_record(
            lane="mapping_completeness",
            pdf_section="BS",
            pdf_label="Balance-sheet identity (Assets = Liabilities + Equity)",
            pdf_value=total_assets,
            source_ref="bridge!BS!Total liabilities and members' equity",
            source_label="Total liabilities and members' equity",
            source_value=total_liab_eq,
            comparison_unit="$K",
            delta=bs_delta,
            tolerance=5.0,
            status=("ties" if abs(bs_delta) <= 5.0 else "bs-does-not-balance"),
            is_subtotal=True,
            notes="Bridge BS must balance: Total assets = Total liabilities and members' equity.",
        ))

    # ===== Output =====
    by_status = {}
    for r in records:
        by_status[r["status"]] = by_status.get(r["status"], 0) + 1
    print(f"\nMapping completeness: {len(records)} records")
    print(f"By status: {by_status}")
    print(f"  Unmapped accounts: {n_unmapped} | mapped-to-nothing: {n_mapped_no_target} | "
          f"stale targets: {n_stale}")
    print(f"  Total TB (leaf) = {total_tb:,.2f} | mapped = {total_mapped:,.2f} | "
          f"gross unflowed = {gross_unmapped:,.2f}")
    if n_unmapped or n_mapped_no_target:
        print("\n  Unmapped / mapped-to-nothing accounts:")
        for r in records:
            if r["status"] in ("unmapped-account", "mapped-to-nothing"):
                print(f"    [{r['status']:<17}] {str(r['pdf_label'])[:50]:<50} {r['pdf_value']:>16,.2f}")

    Path(args.out).write_text(
        json.dumps(records, default=str, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    print(f"\nWrote {args.out}")


if __name__ == "__main__":
    main()
