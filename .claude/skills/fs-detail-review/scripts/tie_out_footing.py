"""Lane 6: Mathematical footing (F) and cross-foot (xF) verification.

For each subtotal row on any FS table, sum the component rows above it (within
the same section) and verify equals the displayed subtotal. Mark `ties-F` if
foots correctly, `exception` if not.

For tables with multiple data columns (BS, IS, SCF: 2025 and 2024 columns;
SOE: per-class columns; FN tables vary), also verify each column subtotal foots
independently.

CLI:
  python tie_out_footing.py <inputs.json> <out.json>
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

from tie_out_common import parse_value, make_record


# Subtotal definitions: per section, the subtotal label and which previous rows
# (by label keyword) sum into it. None = sum ALL preceding numeric rows since last subtotal.
SUBTOTAL_RULES = {
    "BS": [
        # label_match, mode
        ("Total current assets", "since-last"),
        ("Total assets", "compute"),  # Total current + sum of non-current asset components
        ("Total current liabilities", "since-last"),
        ("Total liabilities", "compute"),  # Total current liab + non-current liab components
        ("Total members' equity", "since-last"),
        ("Total liabilities and members' equity", "compute"),  # Total liab + Total equity
    ],
    "IS": [
        ("Gross profit", "rev-minus-cogs"),
        ("Total operating expenses", "since-last"),
        ("Loss from operations", "compute"),  # Gross profit - Total opex
        ("Total other expense", "since-last"),
        ("Total other income (expense)", "since-last"),
        ("Loss before income taxes", "compute"),  # Loss from ops + Total other
        ("Net loss", "compute"),  # Loss before tax - Tax expense
        ("Comprehensive loss", "compute"),  # Net loss + FX
    ],
    "SCF": [
        ("Net cash and cash equivalents provided by operating activities", "since-last"),
        ("Net cash and cash equivalents used in investing activities", "since-last"),
        ("Net cash and cash equivalents provided by investing activities", "since-last"),
        ("Net cash and cash equivalents used in financing activities", "since-last"),
        ("Net cash and cash equivalents provided by financing activities", "since-last"),
        ("Net increase (decrease) in cash and cash equivalents", "compute"),
        # Cash and cash equivalents at end of year = Begin + Net change (already in Lane 4)
    ],
}


def extract_rows_for_footing(table):
    """Return list of (label, {year: value}) for each row that has values."""
    rows = table.get("rows", [])
    year_cols = {}
    header_row_idx = None
    for i, row in enumerate(rows[:5]):
        cells = {}
        for c_idx, cell in enumerate(row):
            if c_idx == 0:
                continue
            if cell is None:
                continue
            if isinstance(cell, (int, float)) and 2020 <= cell <= 2030 and cell == int(cell):
                cells[c_idx] = str(int(cell))
            elif isinstance(cell, str) and re.match(r"^20\d{2}$", cell.strip()):
                cells[c_idx] = cell.strip()
        if len(cells) >= 2:
            year_cols = cells
            header_row_idx = i
            break
    if not year_cols:
        return [], {}

    out = []
    for row in rows[header_row_idx + 1:]:
        if not row:
            continue
        label = row[0].strip() if row[0] else ""
        if not label:
            continue
        values = {}
        for c_idx, year in year_cols.items():
            if c_idx < len(row):
                v = parse_value(row[c_idx])
                if v is not None:
                    values[year] = v
        out.append({"label": label, "values": values})
    return out, year_cols


def find_row(rows, predicate):
    """Find first row matching predicate."""
    for r in rows:
        if predicate(r["label"]):
            return r
    return None


def find_value(rows, label_match, year):
    """Get a value at row matching label_match, column = year."""
    for r in rows:
        if label_match.lower() in r["label"].lower():
            return r["values"].get(year)
    return None


def foot_since_last(rows, subtotal_idx, year, section=None):
    """Sum values in the same column from the row AFTER the previous subtotal/section break.

    Strategy depends on section:
      - SCF: previous subtotal is another 'Net cash ... activities' row OR a section
             header row that ends with ':' (e.g., 'Cash flows from operating activities:').
      - BS/IS: standard subtotal keywords.
    """
    if section == "SCF":
        # Find the previous row that is either:
        #   - a "Net cash..." subtotal (must START with "net cash" — avoid false-match on
        #     "Adjustments to reconcile net loss to net cash ..." which contains those words)
        #   - a SECTION header (starts with "Cash flows from")
        prev_idx = 0
        for j in range(subtotal_idx - 1, -1, -1):
            lbl = rows[j]["label"].lower().strip()
            if lbl.startswith("net cash") and ("provided by" in lbl or "used in" in lbl):
                prev_idx = j + 1
                break
            if lbl.startswith("cash flows from"):
                prev_idx = j + 1
                break
    else:
        # BS/IS standard logic — but EXCLUDE "net loss" since it's not a true subtotal in SCF context
        prev_idx = 0
        subtotal_keywords = ["total ", "loss from operations", "loss before",
                             "comprehensive loss", "gross profit", "balance as of",
                             "net cash", "net increase", "net decrease"]
        for j in range(subtotal_idx - 1, -1, -1):
            lbl = rows[j]["label"].lower()
            if any(kw in lbl for kw in subtotal_keywords):
                prev_idx = j + 1
                break

    s = 0.0
    components_used = []
    for j in range(prev_idx, subtotal_idx):
        v = rows[j]["values"].get(year)
        if v is not None:
            s += v
            components_used.append((rows[j]["label"], v))
    return s, components_used


def compute_special(section, subtotal_label, rows, year, subtotal_idx=None):
    """Compute a special-formula subtotal (not just sum-since-last)."""
    subtotal_label_l = subtotal_label.lower()
    if section == "BS" and "total assets" in subtotal_label_l and "liabilities" not in subtotal_label_l:
        # Total current assets + sum of all subsequent (non-current) rows up to "Total assets"
        # Find "Total current assets" row
        tca_idx = None
        for i, r in enumerate(rows):
            if r["label"].lower().strip() == "total current assets":
                tca_idx = i
                break
        if tca_idx is None or subtotal_idx is None:
            return None, None
        tca = rows[tca_idx]["values"].get(year)
        if tca is None:
            return None, None
        non_current = 0.0
        for i in range(tca_idx + 1, subtotal_idx):
            v = rows[i]["values"].get(year)
            if v is not None:
                non_current += v
        return tca + non_current, f"Total current ({tca}) + non-current ({non_current})"
    if section == "BS" and "total liabilities" in subtotal_label_l and "members" not in subtotal_label_l and "and" not in subtotal_label_l:
        # Total current liabilities + sum of non-current liability rows
        tcl_idx = None
        for i, r in enumerate(rows):
            if r["label"].lower().strip() == "total current liabilities":
                tcl_idx = i
                break
        if tcl_idx is None or subtotal_idx is None:
            return None, None
        tcl = rows[tcl_idx]["values"].get(year)
        if tcl is None:
            return None, None
        non_current = 0.0
        for i in range(tcl_idx + 1, subtotal_idx):
            v = rows[i]["values"].get(year)
            if v is not None:
                non_current += v
        return tcl + non_current, f"Total current liab ({tcl}) + non-current ({non_current})"
    if section == "BS" and "total liabilities and members" in subtotal_label_l:
        tl = find_value(rows, "Total liabilities", year)
        te = find_value(rows, "Total members' equity", year)
        if tl is not None and te is not None:
            return tl + te, f"Total liabilities ({tl}) + Total members' equity ({te})"
    if section == "IS" and "gross profit" in subtotal_label_l:
        rev = find_value(rows, "Revenue", year)
        cogs = find_value(rows, "Cost of revenue", year)
        if rev is not None and cogs is not None:
            return rev - cogs, f"Revenue ({rev}) - COGS ({cogs})"
    if section == "IS" and "loss from operations" in subtotal_label_l:
        gp = find_value(rows, "Gross profit", year)
        opex = find_value(rows, "Total operating expenses", year)
        if gp is not None and opex is not None:
            return gp - opex, f"Gross profit ({gp}) - Total opex ({opex})"
    if section == "IS" and "loss before income taxes" in subtotal_label_l:
        lfo = find_value(rows, "Loss from operations", year)
        other = (find_value(rows, "Total other expense", year)
                 or find_value(rows, "Total other income", year))
        if lfo is not None and other is not None:
            return lfo + other, f"Loss from ops ({lfo}) + Total other ({other})"
    if section == "IS" and "net loss" in subtotal_label_l and "comprehensive" not in subtotal_label_l:
        lbt = find_value(rows, "Loss before income taxes", year)
        tax = find_value(rows, "Income tax expense", year)
        if lbt is not None and tax is not None:
            return lbt + tax, f"Loss before tax ({lbt}) + Tax expense ({tax})"
    if section == "IS" and "comprehensive loss" in subtotal_label_l:
        nl = find_value(rows, "Net loss", year)
        fx = find_value(rows, "Foreign currency translation", year)
        if nl is not None and fx is not None:
            return nl + fx, f"Net loss ({nl}) + FX ({fx})"
    if section == "SCF" and "net increase" in subtotal_label_l:
        # Net change = sum of all "Net cash ..." subtotals + Effect of exchange rate
        op = find_value(rows, "Net cash and cash equivalents provided by operating activities", year)
        inv = (find_value(rows, "Net cash and cash equivalents used in investing activities", year)
               or find_value(rows, "Net cash and cash equivalents provided by investing activities", year))
        fin = (find_value(rows, "Net cash and cash equivalents used in financing activities", year)
               or find_value(rows, "Net cash and cash equivalents provided by financing activities", year))
        fx = find_value(rows, "Effect of exchange rate", year)
        if op is not None and inv is not None and fin is not None:
            total = op + inv + fin + (fx or 0)
            return total, f"Op ({op}) + Inv ({inv}) + Fin ({fin}) + FX ({fx or 0})"
    return None, None


SECTION_TO_PDF_PAGE = {"BS": 5, "IS": 6, "SOE": 7, "SCF": 8}


# ---------- Generic footing for ANY table with subtotal rows ----------

GENERIC_SUBTOTAL_PATTERNS = [
    "total ",
    "gross profit",
    "loss from operations",
    "loss before",
    "comprehensive loss",
    "comprehensive income",
    "property and equipment, net",
    "intangible assets, net",
    "balance as of",
]

# Patterns that look like subtotals but are actually deductions or detail rows.
# Excluded so we don't falsely sum-and-compare them.
NON_SUBTOTAL_PATTERNS = [
    "less:",
    "less ",
    "valuation allowance",
    "accumulated depreciation",
]


def is_generic_subtotal(label):
    if not label:
        return False
    s = str(label).lower().strip()
    for ns in NON_SUBTOTAL_PATTERNS:
        if s.startswith(ns) or ns in s:
            return False
    for p in GENERIC_SUBTOTAL_PATTERNS:
        if s.startswith(p) or p in s:
            return True
    return False


def find_prev_subtotal_or_header(rows, idx, include_prev_subtotal=False):
    """Walk backwards from idx-1 to find the prior subtotal row or section header.

    If include_prev_subtotal=True, returns the index OF the prior subtotal
    (so it's included in the sum — used for 'Total X carrying amount' or 'X, net'
    patterns where current subtotal = prior subtotal + subsequent adjustments).
    Otherwise returns the index AFTER the prior subtotal.
    """
    for j in range(idx - 1, -1, -1):
        lbl = rows[j]["label"].strip().lower()
        if not lbl:
            continue
        # Section header endings (always exclusive)
        if lbl.endswith(":") and not lbl.startswith("total"):
            return j + 1
        if is_generic_subtotal(rows[j]["label"]):
            return j if include_prev_subtotal else j + 1
    return 0


def is_net_or_cumulative_subtotal(label, rows, idx):
    """Detect subtotals that build on a prior subtotal rather than summing fresh.

    Heuristics:
      - Label contains ', net' (e.g., 'Property and equipment, net')
      - Label starts with 'Net ' (e.g., 'Net deferred tax assets')
      - There's a 'Less:' row between this and the previous subtotal
        (e.g., 'Total long-term term loans' after 'Less: current portion of term loans')
    """
    s = label.lower().strip()
    if ", net" in s or s.endswith(" net"):
        return True
    if s.startswith("net "):
        return True
    # Walk back to find a 'Less:' row before hitting a previous subtotal
    for j in range(idx - 1, -1, -1):
        prev_lbl = rows[j]["label"].lower().strip()
        if not prev_lbl:
            continue
        if prev_lbl.startswith("less:") or prev_lbl.startswith("less "):
            return True
        if is_generic_subtotal(rows[j]["label"]):
            break
    return False


def foot_table_generic(table, section_label, pdf_page):
    """Run generic footing checks on any docx table that has year/numeric columns.

    Identifies subtotal rows by label heuristics, sums components above, and
    compares to displayed value. Emits one record per (subtotal, column).
    """
    rows_data, year_cols = extract_rows_for_footing(table)
    if not rows_data or not year_cols:
        return []

    unit = table.get("unit", "$K")
    records = []
    for i, r in enumerate(rows_data):
        if not is_generic_subtotal(r["label"]):
            continue
        # Skip rows that are just section headers (label ends with ":")
        if r["label"].strip().endswith(":"):
            continue
        # Need at least one value to be a footing-eligible row
        if not r["values"]:
            continue

        # Detect if this subtotal builds on a prior subtotal (net/cumulative pattern)
        cumulative = is_net_or_cumulative_subtotal(r["label"], rows_data, i)
        prev_idx = find_prev_subtotal_or_header(rows_data, i, include_prev_subtotal=cumulative)
        if prev_idx == i:
            continue  # no components

        for year, displayed in r["values"].items():
            if displayed is None:
                continue
            # Sum components in this column
            computed = 0.0
            n_components = 0
            for j in range(prev_idx, i):
                v = rows_data[j]["values"].get(year)
                if v is not None:
                    computed += v
                    n_components += 1
            if n_components == 0:
                continue
            delta = displayed - computed
            tolerance = 2
            status = "ties-F" if abs(delta) <= tolerance else "exception"
            records.append(make_record(
                lane="footing",
                pdf_section=section_label,
                pdf_page=pdf_page,
                pdf_label=f"{r['label']} (F)",
                pdf_year=year,
                pdf_value=displayed,
                source_ref=f"computed:sum-{n_components}-components",
                source_label=f"Sum of {n_components} components since prev subtotal/header",
                source_value=computed,
                comparison_unit=unit,
                delta=delta,
                tolerance=tolerance,
                status=status,
                is_subtotal=True,
                notes=f"Generic foot ({section_label})",
            ))
    return records


def cross_foot_table(table, section_label, pdf_page):
    """Cross-foot check: for tables where a row has multiple year/period columns
    AND a 'Total' column, verify the row's columns sum to the Total.

    Example: SOE 'Balance as of' rows where Total = SCU + A-1 + ... + AccDef + AOCI.

    Currently only fires when a column with header containing 'Total' is detected.
    """
    rows = table.get("rows", [])
    # Find header row containing 'Total' column
    total_col = None
    header_row_idx = None
    for i, row in enumerate(rows[:6]):
        for c_idx, cell in enumerate(row):
            if cell and isinstance(cell, str) and "total" in cell.lower():
                total_col = c_idx
                header_row_idx = i
        if total_col is not None:
            break
    if total_col is None:
        return []

    # For each data row, sum the numeric columns BEFORE the Total column and compare
    unit = table.get("unit", "$K")
    records = []
    for row in rows[header_row_idx + 1:]:
        if not row or len(row) <= total_col:
            continue
        label = row[0].strip() if row[0] else ""
        if not label:
            continue
        # Cross-foot only on "Balance as of" rows for SOE-style tables
        if "balance as of" not in label.lower():
            continue
        components = []
        for c_idx in range(1, total_col):
            v = parse_value(row[c_idx])
            if v is not None:
                components.append(v)
        displayed_total = parse_value(row[total_col])
        if displayed_total is None or not components:
            continue
        computed_total = sum(components)
        delta = displayed_total - computed_total
        tolerance = 5
        status = "ties-xF" if abs(delta) <= tolerance else "exception"
        records.append(make_record(
            lane="footing",
            pdf_section=section_label,
            pdf_page=pdf_page,
            pdf_label=f"{label} (xF)",
            pdf_year=None,
            pdf_value=displayed_total,
            source_ref=f"computed:cross-foot",
            source_label=f"Sum of {len(components)} class columns",
            source_value=computed_total,
            comparison_unit=unit,
            delta=delta,
            tolerance=tolerance,
            status=status,
            is_subtotal=True,
            notes="Cross-foot recalc",
        ))
    return records


# Map docx section -> PDF page hint for FN tables (rough first-page; OK if it spans)
SECTION_TO_FN_PAGE = {
    "FN-01-descriptionofbusiness": 9,
    "FN-02-summaryofsignificantacco": 10,
    "FN-03-propertyandequipmentnet": 17,
    "FN-04-intangibleassetsandgoodw": 18,
    "FN-05-accruedexpensesandotherc": 19,
    "FN-06-termloansandlineofcredit": 19,
    "FN-07-incometaxes": 21,
    "FN-08-membersequity": 23,
    "FN-09-stockbasedcompensation": 24,
    "FN-11-leases": 25,
    "FN-12-relatedparties": 27,
    "FN-13-employeebenefitplan": 27,
    "FN-14-subsequentevents": 28,
}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("inputs")
    ap.add_argument("out")
    args = ap.parse_args()

    data = json.loads(Path(args.inputs).read_text(encoding="utf-8"))
    tables = data["fy25_docx"]["tables"]

    records = []

    # === Face statements: use specific subtotal rules ===
    for section in ["BS", "IS", "SCF"]:
        section_table = next((t for t in tables if t["section"] == section), None)
        if not section_table:
            continue
        rows, year_cols = extract_rows_for_footing(section_table)
        if not rows:
            continue
        section_rules = SUBTOTAL_RULES.get(section, [])
        unit = section_table.get("unit", "$K")
        pdf_page = SECTION_TO_PDF_PAGE.get(section)

        for label_match, mode in section_rules:
            # Find the subtotal row index
            subtotal_idx = None
            subtotal_row = None
            for i, r in enumerate(rows):
                if r["label"].lower().strip() == label_match.lower():
                    subtotal_idx = i
                    subtotal_row = r
                    break
            if subtotal_row is None:
                # Try fuzzy
                for i, r in enumerate(rows):
                    if label_match.lower() in r["label"].lower():
                        subtotal_idx = i
                        subtotal_row = r
                        break
            if subtotal_row is None:
                continue

            for year, displayed in subtotal_row["values"].items():
                if displayed is None:
                    continue
                if mode == "since-last":
                    computed, components = foot_since_last(rows, subtotal_idx, year, section=section)
                    note = f"Sum of {len(components)} components since last subtotal/section"
                elif mode == "compute":
                    computed, components_desc = compute_special(section, label_match, rows, year, subtotal_idx=subtotal_idx)
                    note = components_desc or "Special compute"
                elif mode == "rev-minus-cogs":
                    rev = find_value(rows, "Revenue", year)
                    cogs = find_value(rows, "Cost of revenue", year)
                    if rev is not None and cogs is not None:
                        computed = rev - cogs
                        note = f"Revenue ({rev}) - COGS ({cogs})"
                    else:
                        continue
                else:
                    continue
                if computed is None:
                    continue
                delta = displayed - computed
                # Tolerance: $1K (rounding floor for $K-displayed FS)
                tolerance = 2
                status = "ties-F" if abs(delta) <= tolerance else "exception"
                records.append(make_record(
                    lane="footing",
                    pdf_section=section,
                    pdf_page=pdf_page,
                    pdf_label=f"{label_match} (F)",
                    pdf_year=year,
                    pdf_value=displayed,
                    source_ref=f"computed:{mode}",
                    source_label=note,
                    source_value=computed,
                    comparison_unit=unit,
                    delta=delta,
                    tolerance=tolerance,
                    status=status,
                    is_subtotal=True,
                    notes=f"Footing check ({mode})",
                ))

    # === SOE: cross-foot is handled by tie_out_soe.py (Lane 5) with correct
    # column-detection (Units vs Amount). Skipping generic cross-foot here to
    # avoid duplicate/wrong results.

    # === Footnote tables: generic footing on every FN- table ===
    fn_tables = [t for t in tables if t["section"].startswith("FN-")]
    print(f"Footnote tables found: {len(fn_tables)}")
    for ft in fn_tables:
        section = ft["section"]
        pdf_page = SECTION_TO_FN_PAGE.get(section, 9)
        fn_records = foot_table_generic(ft, section, pdf_page)
        if fn_records:
            print(f"  {section} (table {ft.get('idx')}): {len(fn_records)} footing records")
            records.extend(fn_records)
        # Also try cross-foot if applicable
        xf_records = cross_foot_table(ft, section, pdf_page)
        if xf_records:
            print(f"  {section} (table {ft.get('idx')}): {len(xf_records)} cross-foot records")
            records.extend(xf_records)

    # Status summary
    by_status = {}
    for r in records:
        by_status[r["status"]] = by_status.get(r["status"], 0) + 1
    print(f"\nFooting check: {len(records)} records")
    print(f"By status: {by_status}")
    exceptions_only = [r for r in records if r["status"] == "exception"]
    print(f"\nExceptions ({len(exceptions_only)}):")
    for r in exceptions_only:
        print(f"  ! [{r['pdf_section']} {r['pdf_year']}] {r['pdf_label']:<50}  pdf={r['pdf_value']:>12,.0f}  computed={r['source_value']:>12,.0f}  delta={r['delta']:>8,.0f}")

    Path(args.out).write_text(
        json.dumps(records, default=str, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    print(f"Wrote {args.out}")


if __name__ == "__main__":
    main()
