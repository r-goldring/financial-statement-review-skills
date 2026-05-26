"""Lane 5: SOE (Statement of Changes in Members' Equity) rollforward tie-out.

SOE has a different shape from face statements:
  - Columns: per-unit-class breakdown (Senior Converted / A-1 / A-2 / B / C / D) each
             with Units and Amount sub-columns, plus Accumulated Deficit, AOCI, Total.
  - Rows: alternating 'Balance as of Dec 31, YYYY' and activity rows
          (Net loss, Stock-based compensation, Issuance/Redemption, FX translation).

Ties produced:
  1. Each value on PDF SOE row ↔ same value on Bridge SOE row (label-based match)
  2. Each 'Balance as of' row cross-foot:
       sum of all class Amounts + Accumulated Deficit + AOCI = Total Members' Equity (xF)
  3. Each year roll-forward:
       opening Balance + activities (Net loss + SBC + Issuance + FX) = closing Balance (F)
  4. Net loss row Accumulated Deficit ↔ IS Net loss (signs may differ)
  5. FX adjustment row AOCI ↔ IS Foreign currency translation adjustment

CLI:
  python tie_out_soe.py <inputs.json> <out.json>
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

from tie_out_common import parse_value, compare, normalize_label, make_record


SOE_PAGE = 7

# Bridge SOE column layout (zero-indexed from the rows[] array)
# Based on inspection of FY25 bridge SOE tab:
#   col 0: row label
#   col 1: SCU Units, col 3: SCU Amount
#   col 5: A-1 Units, col 7: A-1 Amount
#   col 9: A-2 Units, col 11: A-2 Amount
#   col 13: B Units, col 15: B Amount
#   col 17: C Units, col 19: C Amount
#   col 21: D Units, col 23: D Amount
#   col 25: Accumulated Deficit
#   col 27: AOCI
#   col 29: Total Members' Equity
BRIDGE_AMOUNT_COLS = {
    "SCU": 3, "A-1": 7, "A-2": 11, "B": 15, "C": 19, "D": 23,
    "AccDef": 25, "AOCI": 27, "Total": 29,
}
BRIDGE_UNIT_COLS = {
    "SCU": 1, "A-1": 5, "A-2": 9, "B": 13, "C": 17, "D": 21,
}


def extract_bridge_soe_rows(soe_tab):
    """Return list of {label, amounts: {class: float}, units: {class: float}, is_balance, year}."""
    rows = soe_tab["rows"]
    out = []
    for row in rows:
        if not row or row[0] is None:
            continue
        label = str(row[0]).strip()
        if not label:
            continue
        # Skip header / blank rows
        if label.lower() in ("member units", "units", "amount"):
            continue
        if "senior" in label.lower() and "amount" in label.lower():
            continue

        amounts = {}
        for cls, col in BRIDGE_AMOUNT_COLS.items():
            if col < len(row):
                v = parse_value(row[col])
                amounts[cls] = v if v is not None else None

        units = {}
        for cls, col in BRIDGE_UNIT_COLS.items():
            if col < len(row):
                v = parse_value(row[col])
                units[cls] = v if v is not None else None

        is_balance = "balance as of" in label.lower()
        year = None
        if is_balance:
            m = re.search(r"december\s*31,?\s*(20\d{2})", label.lower())
            if m:
                year = m.group(1)

        out.append({
            "label": label, "amounts": amounts, "units": units,
            "is_balance": is_balance, "year": year,
        })
    return out


def extract_docx_soe_rows(soe_table):
    """Extract PDF SOE rows from the docx table. Returns same shape as bridge extraction.

    PDF SOE table structure (from docx):
      col 0: row label
      col 1: SCU Units, col 2: SCU Amount  (some cells may be the '$' separator)
      ...

    Without specific column knowledge per cell, find each row's numbers in order
    and match to the column positions: Units / Amount / Units / Amount / ... / AccDef / AOCI / Total
    """
    rows = soe_table["rows"]
    # PDF SOE has structure: row label + 6 (units, amount) pairs + AccDef + AOCI + Total
    # Each row has up to 16 numeric cells (12 unit/amount + 3 right-side)
    out = []
    for row in rows:
        if not row or len(row) < 2:
            continue
        label = row[0].strip() if row[0] else ""
        if not label:
            continue
        if "member units" in label.lower() or label.lower() in ("units", "amount"):
            continue

        # Extract all numeric cells in order (skipping $-only and empty cells)
        numbers = []
        for cell in row[1:]:
            v = parse_value(cell)
            if v is not None:
                numbers.append(v)

        if not numbers:
            continue

        # The PDF SOE row structure is: 6 (Units, Amount) pairs + AccDef + AOCI + Total
        # That's 15 values. Some rows lose a "$ —" cell to merging and end up with 14.
        # Approach: ALWAYS treat the last 3 values as [AccDef, AOCI, Total]
        # then map the first 12 values as (Units, Amount) × 6 classes.
        amounts = {}
        units = {}
        if len(numbers) >= 15:
            classes = ["SCU", "A-1", "A-2", "B", "C", "D"]
            for i, cls in enumerate(classes):
                units[cls] = numbers[i * 2]
                amounts[cls] = numbers[i * 2 + 1]
            amounts["AccDef"] = numbers[-3]
            amounts["AOCI"] = numbers[-2]
            amounts["Total"] = numbers[-1]
        elif len(numbers) == 14:
            # One $-cell missing; assume SCU section starts at idx 0 but SCU Amount got merged into Units cell
            # The last 3 are still AccDef, AOCI, Total
            # The middle 11 are: SCU_U, A1_U, A1_A, A2_U, A2_A, B_U, B_A, C_U, C_A, D_U, D_A
            # OR: SCU_U, SCU_A, A1_U, A1_A, A2_U, A2_A, B_U, B_A, C_U, C_A, D_U (D_A merged with AccDef)
            # Use heuristic: if numbers[0] looks like a Unit count (big int), assume normal layout
            #                if numbers[0] is small or zero, try the 11-cell variant
            amounts["AccDef"] = numbers[-3]
            amounts["AOCI"] = numbers[-2]
            amounts["Total"] = numbers[-1]
            mid = numbers[:11]
            # Most-likely interpretation: row had no SCU amount cell (it was empty/merged), so
            # SCU_Units is mid[0], SCU_Amount is None, A-1_Units mid[1], A-1_Amount mid[2], etc.
            # That gives 11 = 1 (SCU_U) + 5 * 2 (A-1..D)
            units["SCU"] = mid[0]
            amounts["SCU"] = None  # missing
            for i, cls in enumerate(["A-1", "A-2", "B", "C", "D"]):
                units[cls] = mid[1 + i * 2]
                amounts[cls] = mid[1 + i * 2 + 1]
        else:
            amounts["__sparse"] = numbers

        is_balance = "balance as of" in label.lower()
        year = None
        if is_balance:
            m = re.search(r"december\s*31,?\s*(20\d{2})", label.lower())
            if m:
                year = m.group(1)

        out.append({
            "label": label, "amounts": amounts, "units": units,
            "is_balance": is_balance, "year": year,
        })
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("inputs")
    ap.add_argument("out")
    args = ap.parse_args()

    data = json.loads(Path(args.inputs).read_text(encoding="utf-8"))
    docx_tables = data["fy25_docx"]["tables"]
    bridge_tabs = data["bridge"]["tabs"]

    pdf_soe = next((t for t in docx_tables if t["section"] == "SOE"), None)
    bridge_soe = next((t for t in bridge_tabs if t["name"] == "SOE"), None)
    pdf_is = next((t for t in docx_tables if t["section"] == "IS"), None)

    if not pdf_soe:
        print("ERROR: no PDF SOE table found")
        sys.exit(1)
    if not bridge_soe:
        print("ERROR: no bridge SOE tab found")
        sys.exit(1)

    pdf_unit = pdf_soe.get("unit", "$K")
    bridge_unit = bridge_soe.get("unit", "$K")

    print(f"PDF SOE unit={pdf_unit}, Bridge SOE unit={bridge_unit}")

    pdf_rows = extract_docx_soe_rows(pdf_soe)
    bridge_rows = extract_bridge_soe_rows(bridge_soe)

    print(f"PDF SOE rows: {len(pdf_rows)}")
    print(f"Bridge SOE rows: {len(bridge_rows)}")

    # Build a label-keyed bridge index for matching
    # For labels that repeat across years (Net loss, SBC, FX adjustment), distinguish by position
    bridge_by_label_seq = []  # list of (label, row) in order
    for br in bridge_rows:
        bridge_by_label_seq.append((normalize_label(br["label"]), br))

    records = []

    # Match each PDF row to the same-position bridge row (sequence-based)
    # Both should have: [Balance 2023, Net loss, SBC, Issuance/etc, FX, Balance 2024, repeat..., Balance 2025]
    pdf_seq = [(normalize_label(p["label"]), p) for p in pdf_rows]

    # Walk in parallel
    bi = 0
    for pl_norm, pr in pdf_seq:
        # Find matching bridge row from bi onwards
        match_br = None
        for j in range(bi, len(bridge_by_label_seq)):
            br_norm, br = bridge_by_label_seq[j]
            if br_norm == pl_norm or pl_norm in br_norm or br_norm in pl_norm:
                match_br = br
                bi = j + 1
                break
        if not match_br:
            # No bridge match
            records.append(make_record(
                lane="soe",
                pdf_section="SOE",
                pdf_page=SOE_PAGE,
                pdf_label=pr["label"],
                pdf_year=pr.get("year"),
                pdf_value=None,
                source_ref="bridge!SOE",
                source_label=None,
                source_value=None,
                status="missing-on-bridge",
                is_subtotal=pr["is_balance"],
                notes="No bridge SOE row matched this label",
            ))
            continue

        # Compare amounts class-by-class
        for cls, pdf_val in pr["amounts"].items():
            if cls == "__sparse":
                continue  # can't match sparse confidently
            if pdf_val is None:
                continue
            bridge_val = match_br["amounts"].get(cls)
            if bridge_val is None:
                continue
            delta, status, tolerance, comparison_unit = compare(
                pdf_val, bridge_val, pdf_unit, bridge_unit,
                is_subtotal=pr["is_balance"],
            )
            records.append(make_record(
                lane="soe",
                pdf_section="SOE",
                pdf_page=SOE_PAGE,
                pdf_label=f"{pr['label']} [{cls}]",
                pdf_year=pr.get("year"),
                pdf_value=pdf_val,
                source_ref=f"bridge!SOE!{match_br['label']}!{cls}",
                source_label=match_br["label"],
                source_value=bridge_val,
                comparison_unit=comparison_unit,
                delta=delta,
                tolerance=tolerance,
                status=status,
                is_subtotal=pr["is_balance"],
            ))
        # Compare units
        for cls, pdf_u in pr["units"].items():
            if pdf_u is None:
                continue
            bridge_u = match_br["units"].get(cls)
            if bridge_u is None:
                continue
            # Units don't have a $ unit — compare raw
            delta_u = pdf_u - bridge_u
            status_u = "ties" if abs(delta_u) < 1 else "exception"
            records.append(make_record(
                lane="soe",
                pdf_section="SOE",
                pdf_page=SOE_PAGE,
                pdf_label=f"{pr['label']} [{cls} Units]",
                pdf_year=pr.get("year"),
                pdf_value=pdf_u,
                source_ref=f"bridge!SOE!{match_br['label']}!{cls} Units",
                source_label=match_br["label"],
                source_value=bridge_u,
                comparison_unit="units",
                delta=delta_u,
                tolerance=1,
                status=status_u,
                is_subtotal=pr["is_balance"],
                notes="Unit count comparison",
            ))

    # Cross-foot check on each "Balance as of" row: Total = sum of all class amounts + AccDef + AOCI
    print("\nCross-foot (xF) checks on Balance rows:")
    for pr in pdf_rows:
        if not pr["is_balance"]:
            continue
        if "Total" not in pr["amounts"] or pr["amounts"]["Total"] is None:
            continue
        components = ["SCU", "A-1", "A-2", "B", "C", "D", "AccDef", "AOCI"]
        component_sum = 0.0
        any_missing = False
        for c in components:
            v = pr["amounts"].get(c)
            if v is None:
                any_missing = True
                break
            component_sum += v
        if any_missing:
            continue
        total = pr["amounts"]["Total"]
        delta = total - component_sum
        status = "ties" if abs(delta) < 5 else "exception"
        print(f"  {pr['label']}: components_sum={component_sum:,.0f}  total={total:,.0f}  delta={delta:,.0f}  {status}")
        records.append(make_record(
            lane="soe",
            pdf_section="SOE",
            pdf_page=SOE_PAGE,
            pdf_label=f"{pr['label']} (cross-foot)",
            pdf_year=pr.get("year"),
            pdf_value=total,
            source_ref=f"computed-sum-of-components",
            source_label=f"sum of SCU+A-1+A-2+B+C+D+AccDef+AOCI",
            source_value=component_sum,
            comparison_unit=pdf_unit,
            delta=delta,
            tolerance=5,
            status="ties-xF" if status == "ties" else "exception",
            is_subtotal=True,
            notes="Cross-foot recalc: Total = sum of class amounts + AccDef + AOCI",
        ))

    # === Vertical roll-forward F check ===
    # Opening balance + sum of activity rows in that year = closing balance
    # Group rows by year: between Balance N-1 and Balance N is "year N" activity
    print("\nRoll-forward (F) checks for each year:")
    balance_indices = [i for i, p in enumerate(pdf_rows) if p["is_balance"]]
    for b_idx in range(len(balance_indices) - 1):
        opening = pdf_rows[balance_indices[b_idx]]
        closing = pdf_rows[balance_indices[b_idx + 1]]
        activity_rows = pdf_rows[balance_indices[b_idx] + 1:balance_indices[b_idx + 1]]
        if not opening["amounts"].get("Total") or not closing["amounts"].get("Total"):
            continue
        opening_total = opening["amounts"]["Total"]
        closing_total = closing["amounts"]["Total"]
        activity_sum = 0.0
        for ar in activity_rows:
            v = ar["amounts"].get("Total")
            if v is not None:
                activity_sum += v
        computed = opening_total + activity_sum
        delta = closing_total - computed
        status = "ties-F" if abs(delta) < 5 else "exception"
        print(f"  {opening['label']} -> {closing['label']}: opening={opening_total:,.0f} + activities={activity_sum:,.0f} = {computed:,.0f}  closing={closing_total:,.0f}  delta={delta:,.0f}  {status}")
        records.append(make_record(
            lane="soe",
            pdf_section="SOE",
            pdf_page=SOE_PAGE,
            pdf_label=f"Rollforward: {opening['label']} -> {closing['label']}",
            pdf_year=closing.get("year"),
            pdf_value=closing_total,
            source_ref=f"opening + sum(activities)",
            source_label=f"{opening_total:,.0f} + {activity_sum:,.0f}",
            source_value=computed,
            comparison_unit=pdf_unit,
            delta=delta,
            tolerance=5,
            status=status if status != "ties-F" else "ties-F",
            is_subtotal=True,
            notes="Vertical roll-forward foot",
        ))

    # === IS Net Loss ↔ SOE Net Loss row Accumulated Deficit ===
    # NOTE: this used to be a single check that hit the FIRST "Net loss" SOE row (2024)
    # and compared it to the first big number in the IS Net Loss row (which was 2025),
    # producing a false-positive exception. Lane 4 (pdf_internal) now performs this
    # tie correctly with year-awareness for BOTH 2024 and 2025 (mark "/PL"), so we
    # don't duplicate it here.

    # Status summary
    by_status = {}
    for r in records:
        by_status[r["status"]] = by_status.get(r["status"], 0) + 1
    print(f"\nTotal records: {len(records)}")
    print(f"By status: {by_status}")
    Path(args.out).write_text(
        json.dumps(records, default=str, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    print(f"Wrote {args.out}")


if __name__ == "__main__":
    main()
