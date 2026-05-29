"""Lane 1: PDF face statements (FY25) ↔ Bridge workbook.

Reads inputs.json and emits tie records for every line of:
  - PDF Balance Sheet ↔ Bridge "BS" tab
  - PDF Income Statement ↔ Bridge "PL" tab
  - PDF Statement of Changes in Members' Equity ↔ Bridge "SOE" tab
  - PDF Statement of Cash Flows ↔ Bridge "SCF" tab

For each (label, year-column) pair on the PDF side, finds the matching row on the
bridge side by normalized label, then compares values using the tolerance matrix.

CLI:
  python tie_out_pdf_to_bridge.py <inputs.json> <out.json>
"""

import argparse
import json
import sys
import re
from pathlib import Path

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

SCRIPT_DIR = Path(__file__).parent
sys.path.insert(0, str(SCRIPT_DIR))

from tie_out_common import (
    parse_value, compare, normalize_label, is_subtotal_label,
    make_record, fuzzy_match, normalize_unit,
)
from difflib import SequenceMatcher


# Map PDF section → bridge tab name
SECTION_TO_BRIDGE = {
    "BS": "BS",
    "IS": "PL",
    "SOE": "SOE",
    "SCF": "SCF",
}

# Map FY25 PDF FN-section → bridge tab
FN_TO_BRIDGE = {
    "FN-01-descriptionofbusiness": None,  # narrative-only
    "FN-02-summaryofsignificantacco": None,  # policies, mostly narrative
    "FN-03-propertyandequipmentnet": "FN - PPE",
    "FN-04-intangibleassetsandgoodw": "FN - Intangibles",
    "FN-05-accruedexpensesandotherc": "FN - Accrued Expenses",
    "FN-06-termloansandlineofcredit": "FN - Debt",
    # Bridge tab spelling: v5.21 used "FN- Taxes" (typo, no space after FN), the
    # compiler corrected to "FN - Taxes" in the final. Try the corrected form first,
    # then fall back to the typo (handled by bridge_tabs lookup at consumer site).
    "FN-07-incometaxes": "FN - Taxes",
    "FN-08-membersequity": "FN - Members Equity",
    "FN-09-stockbasedcompensation": "FN - SBC",
    "FN-10-commitmentsandcontingenc": None,  # narrative
    "FN-11-leases": "FN - Leases",
    "FN-12-relatedparties": None,  # NO bridge tab (gap — flag separately)
    "FN-13-employeebenefitplan": "FN - 401(k)",
    "FN-14-subsequentevents": "FN Sub Events",
}

# Bridge tab names sometimes change between versions (typo fixes, renames). Try the
# canonical name first; if absent, fall back to a known alias.
BRIDGE_TAB_ALIASES = {
    "FN - Taxes": ["FN- Taxes"],   # v5.21 had the typo "FN- Taxes" (no space)
    "FN - PPE": ["FN- PPE"],
    "FN - Intangibles": ["FN- Intangibles"],
    "FN - Debt": ["FN- Debt"],
    "FN - Accrued Expenses": ["FN- Accrued Expenses"],
    "FN - Members Equity": ["FN- Members Equity"],
    "FN - SBC": ["FN- SBC"],
    "FN - Leases": ["FN- Leases"],
    "FN - 401(k)": ["FN- 401(k)"],
}


def resolve_bridge_tab(bridge_tabs, name):
    """Return the bridge_tabs entry for `name`, trying known aliases as a fallback."""
    if name in bridge_tabs:
        return name, bridge_tabs[name]
    for alias in BRIDGE_TAB_ALIASES.get(name, []):
        if alias in bridge_tabs:
            return alias, bridge_tabs[alias]
    return None, None


def extract_pdf_rows(table):
    """From a .docx table, yield (label, [(col_idx, value, year_label)]) for each row.

    .docx tables come as rows-of-cells; column structure must be detected from header rows.
    Convention observed: row 1 has year labels in cols 1, 3 (e.g., '2025' col 1, '2024' col 3),
    and data rows have label in col 0, values in cols 1 and 3.

    Returns list of (label, value_per_col_dict) where value_per_col_dict maps year-string → float.
    """
    rows = table.get("rows", [])
    if len(rows) < 3:
        return []

    # Find the row containing year strings like '2025', '2024' AS COLUMN HEADERS.
    # Rules to avoid mistaking year-as-row-label (amortization schedules) for headers:
    #   - year cells must be in column index >= 1 (col 0 is reserved for labels)
    #   - at least TWO distinct year cells must be on the same row (header rows compare years)
    year_cols = {}
    header_row_idx = None
    for i, row in enumerate(rows[:5]):
        row_year_cells = {}
        # Col 0 normally holds a label — but if col 0 itself is a 'YYYY' string,
        # this is a stand-alone year-header row (common when the docx was generated
        # from a PDF, where the year row is rendered without a leading label cell).
        # In that case include col 0 too. The "year-as-row-label" trap (e.g. an
        # amortization schedule row ['2026', amount, amount]) is still avoided
        # because such rows have only ONE year cell — we still require >=2.
        col0_is_year = (len(row) > 0 and isinstance(row[0], str)
                        and re.match(r"^20\d{2}$", row[0].strip()))
        for c_idx, cell in enumerate(row):
            if c_idx == 0 and not col0_is_year:
                continue
            if cell and isinstance(cell, str):
                stripped = cell.strip()
                if re.match(r"^20\d{2}$", stripped):
                    row_year_cells[c_idx] = stripped
        if len(row_year_cells) >= 2:
            year_cols = row_year_cells
            header_row_idx = i
            break
    if not year_cols or header_row_idx is None:
        return []

    # If the header row was label-less (col 0 was a year — the carve-out above) but
    # subsequent data rows DO have a label at col 0, the year columns are shifted by
    # one relative to data. Detect by sampling the first data row: if col 0 looks
    # like a non-year text label, shift year_cols indices by +1.
    if 0 in year_cols:
        for sample in rows[header_row_idx + 1:header_row_idx + 6]:
            if not sample or not sample[0] or not isinstance(sample[0], str):
                continue
            s0 = sample[0].strip()
            if s0 and not re.match(r"^20\d{2}$", s0) and not re.match(r"^[\-+]?[\d,.()]+$", s0):
                # data rows have a label column the header didn't — shift years right
                year_cols = {c_idx + 1: yr for c_idx, yr in year_cols.items()}
                break

    out = []
    for row in rows[header_row_idx + 1:]:
        if not row:
            continue
        label = row[0].strip() if len(row) > 0 and row[0] else ""
        if not label:
            continue
        values = {}
        for c_idx, year in year_cols.items():
            if c_idx < len(row):
                v = parse_value(row[c_idx])
                if v is not None:
                    values[year] = v
        if values:
            out.append({"label": label, "values": values, "is_subtotal": is_subtotal_label(label)})
    return out


def extract_bridge_rows(bridge_tab):
    """From a bridge tab, extract (label, {year: value}) records.

    Bridge tabs typically have: col 0 = top-level label, col 1 = sub-item label,
    col 2 = subtotal label, col 3 = 2025 value, col 4 = 2024 value (for face tabs).
    But this varies by tab. Strategy:
      - Find header row containing year strings
      - For data rows, the FIRST non-empty string in cols [0..2] is the label
      - Values are in the columns where header row had years
    """
    rows = bridge_tab.get("rows", [])
    if len(rows) < 3:
        return []

    year_cols = {}
    header_row_idx = None
    for i, row in enumerate(rows[:10]):
        row_year_cells = {}
        # Same col-0 / stand-alone year-header carve-out as the face-statement parser.
        col0 = row[0] if len(row) > 0 else None
        col0_is_year = (isinstance(col0, str) and bool(re.match(r"^20\d{2}$", col0.strip())))
        for c_idx, cell in enumerate(row):
            if c_idx == 0 and not col0_is_year:
                continue
            if cell is None:
                continue
            # Accept both string '2025' and integer 2025 / float 2025.0
            if isinstance(cell, (int, float)) and 2020 <= cell <= 2030 and cell == int(cell):
                row_year_cells[c_idx] = str(int(cell))
                continue
            s = str(cell).strip()
            if re.match(r"^20\d{2}$", s):
                row_year_cells[c_idx] = s
        if len(row_year_cells) >= 2:
            year_cols = row_year_cells
            header_row_idx = i
            break
    if not year_cols or header_row_idx is None:
        return []

    # Bridge tabs may have TWO side-by-side blocks (the FS compiler's work area + formatted
    # disclosure block) — both with their own '2025'/'2024' headers. The extractor
    # would then look up both blocks for each row, and the second block overwrites
    # the first. Dedupe: for each year, keep only the LEFTMOST column. The secondary
    # number-search later catches anything missed by the left block.
    year_cols_dedup = {}
    seen_years = set()
    for c_idx in sorted(year_cols.keys()):
        year = year_cols[c_idx]
        if year not in seen_years:
            year_cols_dedup[c_idx] = year
            seen_years.add(year)
    year_cols = year_cols_dedup

    out = []
    for row in rows[header_row_idx + 1:]:
        if not row:
            continue
        # Find first non-empty label in cols 0-3
        label = ""
        for c_idx in range(min(4, len(row))):
            if row[c_idx] is not None:
                s = str(row[c_idx]).strip()
                if s:
                    label = s
                    break
        if not label:
            continue
        values = {}
        for c_idx, year in year_cols.items():
            if c_idx < len(row):
                v = parse_value(row[c_idx])
                if v is not None:
                    values[year] = v
        if values:
            out.append({"label": label, "values": values, "is_subtotal": is_subtotal_label(label)})
    return out


# Map FY25 PDF doc-section → PDF physical page (based on FY24 layout convention)
SECTION_TO_PDF_PAGE = {
    "BS": 5,
    "IS": 6,
    "SOE": 7,
    "SCF": 8,
}


# Strings that are NOT considered "real" labels for secondary-match scoring
_LABEL_NOISE = {"rounding", "py", "chk", "im", "rx", "(in thousands)", "$ thousands",
                "tied to fq reconciliations", "tied to fq reconciliation",
                "from below", "from intangibles", "from "}


def _is_number_only_text(s):
    """True if s contains only number/currency formatting characters."""
    return bool(re.match(r"^[\(\)\$\s,\-\d\.%]+$", s))


def _looks_like_excel_serial(cell):
    """True if cell is a numeric in the range typical for Excel date serials around 2020-2030."""
    if isinstance(cell, (int, float)):
        return 40000 <= cell <= 50000 and cell == int(cell)
    return False


def find_value_in_bridge_with_context(bridge_tab, target_value_K, is_subtotal, pdf_label):
    """Search the entire bridge tab for a cell matching target_value (in $K), returning the
    best candidate with nearby-label context.

    Approach:
      1. Walk every cell of the bridge tab
      2. Skip cells that look like year headers, excel dates, or non-numeric noise
      3. Convert each cell to $K (if bridge tab is $1 raw, divide by 1000)
      4. Find candidates within tolerance of target_value_K
      5. For each candidate, find the nearest label cell (same row first, then adjacent rows)
      6. Score by label similarity to pdf_label
      7. Return best (score >= 0.45 is considered a meaningful match)

    Returns dict or None:
      {"row_idx", "col_idx", "value_K", "nearby_label", "similarity_score", "all_candidates"}
    """
    if target_value_K is None:
        return None
    rows = bridge_tab.get("rows", [])
    bridge_unit_norm = normalize_unit(bridge_tab.get("unit", "ambiguous"))

    # Tolerance in $K display unit
    simple_tol = 1.0
    sub_tol = 5.0
    tolerance_K = sub_tol if is_subtotal else simple_tol
    # For raw-$ bridge, we need to handle larger absolute deltas because of rounding loss
    if bridge_unit_norm == "$1":
        tolerance_K = max(tolerance_K, 1.0)  # bridge $1 -> $K rounding gives up to $0.5 either way

    # Build a list of all text-label cells indexed by (row, col) for fast nearby-label lookup
    text_cells = {}  # (r, c) -> label string
    for r_idx, row in enumerate(rows):
        for c_idx, cell in enumerate(row):
            if cell is None:
                continue
            if isinstance(cell, (int, float)):
                continue
            s = str(cell).strip()
            if not s:
                continue
            if _is_number_only_text(s):
                continue
            if re.match(r"^20\d{2}$", s):
                continue  # year header
            s_lower = s.lower()
            if s_lower in _LABEL_NOISE:
                continue
            text_cells[(r_idx, c_idx)] = s

    # Skip target=0 entirely (would match every "-" placeholder cell)
    if abs(target_value_K) < 0.5:
        return None

    candidates = []
    for r_idx, row in enumerate(rows):
        for c_idx, cell in enumerate(row):
            if cell is None:
                continue
            if _looks_like_excel_serial(cell):
                continue
            # Skip cells that ARE labels (would only happen for numeric strings)
            v = parse_value(cell)
            if v is None:
                continue
            # Convert bridge value to $K for comparison.
            # Some bridge tabs are nominally $1 (raw) but the disclosure block at the
            # top rows is actually in $K (already rounded). And conversely, FN-PPE is
            # marked $K but a few cells inside still show $1 raw. So always try BOTH
            # interpretations and pick the closer one.
            v_as_K = v
            v_as_1 = v / 1000.0
            # Pick the interpretation closer to target
            d_K = abs(v_as_K - target_value_K)
            d_1 = abs(v_as_1 - target_value_K)
            if d_K <= d_1:
                v_K = v_as_K
                cell_interp = "$K"
            else:
                v_K = v_as_1
                cell_interp = "$1"
            if abs(v_K - target_value_K) > tolerance_K:
                continue
            candidates.append((r_idx, c_idx, v, v_K))

    if not candidates:
        return None

    # For each candidate find nearest label (search same row first, then ±1, ±2 rows)
    def find_nearest_label(r_idx, c_idx):
        # Same row: prefer leftmost label
        same_row_labels = [(c, lbl) for (r, c), lbl in text_cells.items() if r == r_idx]
        if same_row_labels:
            # Prefer label to the left of the value (smaller col index)
            same_row_labels.sort(key=lambda x: (0 if x[0] < c_idx else 1, abs(x[0] - c_idx)))
            return same_row_labels[0][1]
        # Adjacent rows
        for dr in (-1, 1, -2, 2):
            adj_labels = [(c, lbl) for (r, c), lbl in text_cells.items() if r == r_idx + dr]
            if adj_labels:
                adj_labels.sort(key=lambda x: abs(x[0] - c_idx))
                return adj_labels[0][1]
        return None

    pdf_norm = normalize_label(pdf_label)
    # Strip "less:" prefix for both sides — Less: rows have the same canonical line
    pdf_strip = pdf_norm.replace("less_", "").replace("less:", "")

    scored = []
    for r_idx, c_idx, v_raw, v_K in candidates:
        nearby = find_nearest_label(r_idx, c_idx)
        if not nearby:
            score = 0.0
        else:
            nearby_norm = normalize_label(nearby)
            nearby_strip = nearby_norm.replace("less_", "").replace("less:", "")
            if nearby_norm == pdf_norm or nearby_strip == pdf_strip:
                score = 1.0
            else:
                score = SequenceMatcher(None, pdf_norm, nearby_norm).ratio()
        scored.append((score, r_idx, c_idx, v_raw, v_K, nearby))

    # Sort by score desc, then by value-precision (closer to target wins ties)
    scored.sort(key=lambda x: (-x[0], abs(x[4] - target_value_K)))
    best = scored[0]
    score, r_idx, c_idx, v_raw, v_K, nearby = best

    # Decision rule: accept the secondary match if
    #   (a) label similarity >= 0.5  (similar enough that a human would consider it the same line)
    #   OR
    #   (b) value is large/specific (>= $50K display) AND there's only 1 candidate
    accept = score >= 0.5 or (abs(target_value_K) >= 50 and len(candidates) == 1)
    if not accept:
        return None

    return {
        "row_idx": r_idx,
        "col_idx": c_idx,
        "value_K": v_K,
        "value_raw": v_raw,
        "nearby_label": nearby,
        "similarity_score": score,
        "n_candidates": len(candidates),
    }


def extract_fn_row_values(pdf_table):
    """Fallback row extractor for footnote tables that use **column-name headers**
    (e.g. 'Gross Carrying Amount / Accumulated Amortization / Net Carrying Amount')
    instead of year headers.

    Returns [{label, values:[(col_idx, value), ...]}, ...] — every row with at least
    one parseable numeric cell, using the first non-numeric string on the row as the
    label. Column-header rows have no numeric cells and are filtered out automatically.
    """
    out = []
    for row in pdf_table.get("rows", []):
        if not row:
            continue
        label = None
        values = []
        for c_idx, cell in enumerate(row):
            v = parse_value(cell)
            if v is not None:
                values.append((c_idx, v))
            elif label is None and isinstance(cell, str) and cell.strip() and len(cell.strip()) >= 4:
                label = cell.strip()
        # Header rows have no numeric values → skip. Rows missing a label can't tie.
        if label and values:
            out.append({"label": label, "values": values})
    return out


def tie_fn_section_by_values(pdf_table, bridge_tab, section, bridge_name, fn_page=None):
    """Fallback FN tying: when the table has column-name headers (not year headers),
    extract_pdf_rows returns nothing. Instead, for each (label, value) we search the
    bridge tab for a matching value with label-similarity context (same pattern as
    Lane 1's caption-changed secondary match).

    Conservative: requires label similarity >= 0.5 (the same threshold the secondary
    match uses to gate caption-changed acceptance) OR a unique large-value candidate.
    """
    records = []
    rows = extract_fn_row_values(pdf_table)
    for r in rows:
        label = r["label"]
        is_sub = is_subtotal_label(label)
        for c_idx, v in r["values"]:
            if v == 0:
                continue  # zeros match any '-' placeholder; too risky
            match = find_value_in_bridge_with_context(
                bridge_tab, target_value_K=v, is_subtotal=is_sub, pdf_label=label,
            )
            if not match:
                continue
            score = match.get("similarity_score", 0) or 0
            unique_large = (match.get("n_candidates", 0) == 1 and abs(v) >= 50)
            if score < 0.5 and not unique_large:
                continue
            delta = match["value_K"] - v
            tol = 5.0 if is_sub else 1.0
            status = ("ties" if abs(delta) <= tol
                      else "ties-with-rounding" if abs(delta) <= tol * 5
                      else "exception")
            records.append(make_record(
                "pdf_to_bridge",
                pdf_section=section, pdf_page=fn_page,
                pdf_label=label, pdf_year=None, pdf_value=v,
                source_ref=f"bridge!{bridge_tab['name']}",
                source_label=match.get("nearby_label") or label,
                source_value=round(match["value_K"], 1),
                comparison_unit="$K", delta=round(delta, 1),
                tolerance=tol, status=status, is_subtotal=is_sub,
                notes=("FN value-match (column-name-header table) — "
                       f"score={score:.2f}, candidates={match.get('n_candidates', 0)}"),
            ))
    return records


def _try_fn_spillover(pdf_table, bridge_tabs, fn_page_map, current_section):
    """Section-spillover rescue: when a table's assigned section produced 0 ties
    (often because pdf2docx missed a 'Note N' header and tables for the next FN got
    tagged with the previous section), try the OTHER FN bridges and claim the table
    for the one that returns the most ties — provided it clears a minimum-evidence
    threshold so we don't false-positive on a single coincidental value match.

    Returns (records, claimed_label) or ([], None) if no bridge gives enough ties.
    """
    MIN_TIES_TO_CLAIM = 3
    best_records = []
    best_label = None
    best_section = None
    for cand_section, cand_bridge_name in FN_TO_BRIDGE.items():
        if cand_bridge_name is None or cand_section == current_section:
            continue
        _, cand_tab = resolve_bridge_tab(bridge_tabs, cand_bridge_name)
        if cand_tab is None:
            continue
        cand_page = fn_page_map.get(cand_section)
        cand_records = tie_fn_section_by_values(
            pdf_table, cand_tab, cand_section, cand_bridge_name, cand_page,
        )
        ties = sum(1 for r in cand_records if r["status"] in ("ties", "ties-with-rounding"))
        if ties >= MIN_TIES_TO_CLAIM and ties > len(best_records):
            best_records = cand_records
            best_label = cand_bridge_name
            best_section = cand_section
    if best_section:
        # Stamp the claimed section's page on records that don't already have one.
        page = fn_page_map.get(best_section)
        if page is not None:
            for r in best_records:
                if r.get("pdf_page") is None:
                    r["pdf_page"] = page
    return best_records, best_label


def tie_section(pdf_table, bridge_tab, section, source_section_label):
    """Tie one section: a single .docx table ↔ a single bridge tab."""
    records = []
    pdf_rows = extract_pdf_rows(pdf_table)
    bridge_rows = extract_bridge_rows(bridge_tab)
    bridge_by_key = {}
    bridge_labels_raw = []
    for br in bridge_rows:
        k = normalize_label(br["label"])
        bridge_by_key.setdefault(k, []).append(br)
        bridge_labels_raw.append(br["label"])

    pdf_unit = pdf_table.get("unit", "ambiguous")
    bridge_unit = bridge_tab.get("unit", "ambiguous")

    pdf_page = SECTION_TO_PDF_PAGE.get(section, None)

    for pr in pdf_rows:
        # Skip rows where PDF side has NO values (i.e., it's a header row like "Assets" or "Current liabilities:")
        # These were misclassified as "values" by parse_value when cell was empty string.
        if not pr["values"]:
            continue
        pdf_label = pr["label"]
        pdf_key = normalize_label(pdf_label)
        bridge_match = bridge_by_key.get(pdf_key)

        if not bridge_match:
            # Try fuzzy match on the raw label
            fuzzy_label, score = fuzzy_match(pdf_label, bridge_labels_raw, threshold=0.85)
            if fuzzy_label:
                fk = normalize_label(fuzzy_label)
                bridge_match = bridge_by_key.get(fk)

        if not bridge_match:
            for year, pdf_value in pr["values"].items():
                rec = make_record(
                    lane="pdf_to_bridge",
                    pdf_section=section,
                    pdf_page=pdf_page,
                    pdf_label=pdf_label,
                    pdf_year=year,
                    pdf_value=pdf_value,
                    source_ref=f"bridge!{bridge_tab['name']}",
                    source_label=None,
                    source_value=None,
                    status="missing-on-bridge",
                    is_subtotal=pr["is_subtotal"],
                    notes=f"No bridge row matched label '{pdf_label}'",
                )
                records.append(rec)
            continue

        # Use the first match
        bridge_row = bridge_match[0]
        if len(bridge_match) > 1:
            # multiple bridge rows match this label — note ambiguity but still tie to first
            ambiguity = f"multiple-bridge-rows-matched ({len(bridge_match)}): {[b['label'] for b in bridge_match]}"
        else:
            ambiguity = None

        for year, pdf_value in pr["values"].items():
            bridge_value = bridge_row["values"].get(year)
            if bridge_value is None:
                # year column missing on bridge side
                rec = make_record(
                    lane="pdf_to_bridge",
                    pdf_section=section,
                    pdf_page=pdf_page,
                    pdf_label=pdf_label,
                    pdf_year=year,
                    pdf_value=pdf_value,
                    source_ref=f"bridge!{bridge_tab['name']}!{bridge_row['label']}",
                    source_label=bridge_row["label"],
                    source_value=None,
                    status="missing-year-on-bridge",
                    is_subtotal=pr["is_subtotal"],
                    notes=f"Bridge row found but no {year} column (bridge years: {list(bridge_row['values'].keys())})",
                )
                records.append(rec)
                continue

            delta, status, tolerance, comparison_unit = compare(
                pdf_value, bridge_value, pdf_unit, bridge_unit,
                is_subtotal=pr["is_subtotal"],
            )
            rec = make_record(
                lane="pdf_to_bridge",
                pdf_section=section,
                pdf_page=pdf_page,
                pdf_label=pdf_label,
                pdf_year=year,
                pdf_value=pdf_value,
                source_ref=f"bridge!{bridge_tab['name']}",
                source_label=bridge_row["label"],
                source_value=bridge_value,
                comparison_unit=comparison_unit,
                delta=delta,
                tolerance=tolerance,
                status=status,
                is_subtotal=pr["is_subtotal"],
                notes=ambiguity,
            )
            records.append(rec)

    return records


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("inputs", help="inputs.json")
    ap.add_argument("out", help="output JSON")
    ap.add_argument("--fn-page-map", default=None,
                    help="Path to footnote section→PDF page mapping JSON (from detect_fn_pages.py)")
    args = ap.parse_args()

    data = json.loads(Path(args.inputs).read_text(encoding="utf-8"))
    docx_tables = data["fy25_docx"]["tables"]
    bridge_tabs = {t["name"]: t for t in data["bridge"]["tabs"]}

    fn_page_map = {}
    if args.fn_page_map and Path(args.fn_page_map).exists():
        fn_page_map = json.loads(Path(args.fn_page_map).read_text(encoding="utf-8"))
        # Filter to only sections with a valid page
        fn_page_map = {k: v for k, v in fn_page_map.items() if v is not None}
        print(f"  loaded FN page map with {len(fn_page_map)} sections")

    all_records = []

    # === Face statements ===
    for section, bridge_name in SECTION_TO_BRIDGE.items():
        pdf_section_tables = [t for t in docx_tables if t["section"] == section]
        if not pdf_section_tables:
            print(f"  ⚠ No PDF table for section {section}")
            continue
        if bridge_name not in bridge_tabs:
            print(f"  ⚠ No bridge tab '{bridge_name}' for section {section}")
            continue
        # Face statements have one primary table per section
        pdf_table = pdf_section_tables[0]
        bridge_tab = bridge_tabs[bridge_name]
        records = tie_section(pdf_table, bridge_tab, section, bridge_name)
        all_records.extend(records)
        ties = sum(1 for r in records if r["status"] == "ties")
        ties_round = sum(1 for r in records if r["status"] == "ties-with-rounding")
        exc = sum(1 for r in records if r["status"] == "exception")
        miss = sum(1 for r in records if r["status"] in ("missing-on-bridge", "missing-year-on-bridge"))
        print(f"  {section:>4} <-> bridge '{bridge_name}'  total={len(records)}  ties={ties}  rounding={ties_round}  exceptions={exc}  missing={miss}")

    # === Footnote tables ===
    print("\nFootnote tables:")
    for section, bridge_name in FN_TO_BRIDGE.items():
        if bridge_name is None:
            continue
        # Match by FN-NN prefix — section titles depend on the docx loader's truncation
        # of the heading text, which differs between hand-edited docx and pdf2docx-derived
        # docx (e.g. "FN-03-propertyandequipmentnet" vs "FN-03-propertyandequipmentnett").
        # The numeric prefix is canonical, so match on that.
        prefix = "-".join(section.split("-", 2)[:2]) + "-"  # "FN-03-"
        pdf_section_tables = [t for t in docx_tables if t.get("section", "").startswith(prefix)]
        if not pdf_section_tables:
            continue
        resolved_name, bridge_tab = resolve_bridge_tab(bridge_tabs, bridge_name)
        if bridge_tab is None:
            print(f"  ⚠ FN bridge tab '{bridge_name}' not found (for section {section})")
            continue
        if resolved_name != bridge_name:
            print(f"  (resolved bridge tab '{bridge_name}' -> '{resolved_name}' via alias)")
        # Resolve pdf_page from fn_page_map (set later on records produced by tie_section)
        fn_page = fn_page_map.get(section)
        for pdf_table in pdf_section_tables:
            records = tie_section(pdf_table, bridge_tab, section, bridge_name)
            via = "year-hdr"
            # If the table has column-name headers (not year headers) — common in
            # FN disclosures and in pdf2docx-derived docx — extract_pdf_rows returns
            # nothing. Fall back to value-search tying (label + value -> bridge cell
            # via similarity-scored secondary match).
            if not records:
                records = tie_fn_section_by_values(pdf_table, bridge_tab, section, bridge_name, fn_page)
                via = "value-match" if records else "year-hdr"
            # Section-spillover: a single page-break in pdf2docx output can collapse
            # 'Note 8'..'Note 14' headers into the previous FN section's tag (e.g. FN-09
            # SBC tables and FN-11 Leases tables end up tagged FN-07). If this table
            # produced 0 ties to its assigned bridge, try the remaining FN bridges and
            # claim the table for whichever returns >= 3 ties.
            if not records:
                records, claimed = _try_fn_spillover(pdf_table, bridge_tabs, fn_page_map, section)
                if records:
                    via = f"spillover->{claimed}"
            if fn_page is not None and via != "year-hdr" and not via.startswith("spillover"):
                for r in records:
                    if r.get("pdf_page") is None:
                        r["pdf_page"] = fn_page
            elif via.startswith("spillover") and not all(r.get("pdf_page") for r in records):
                # spillover records already have pdf_page from the claimed FN section
                pass
            elif fn_page is not None:
                for r in records:
                    if r.get("pdf_page") is None:
                        r["pdf_page"] = fn_page
            all_records.extend(records)
            if records:
                ties = sum(1 for r in records if r["status"] == "ties")
                exc = sum(1 for r in records if r["status"] == "exception")
                miss = sum(1 for r in records if "missing" in r["status"])
                page_note = f" page={fn_page}" if fn_page else ""
                print(f"  {section:<32} (table {pdf_table['idx']}, {pdf_table['row_count']}r, {via}) <-> '{bridge_name}'  total={len(records)}  ties={ties}  exc={exc}  miss={miss}{page_note}")

    # === Secondary number-only match (rescues caption-changed ties) ===
    # For every non-tie record, search the same bridge tab for the PDF value with
    # label-context disambiguation. If found, reclassify as 'ties-caption-changed'.
    print("\nSecondary number-only match (rescue caption-changed ties):")
    bridge_lookup = {t["name"]: t for t in data["bridge"]["tabs"]}
    reclassified = 0
    reclassified_zero = 0
    non_tie_statuses = {"exception", "missing-on-bridge", "missing-year-on-bridge"}

    # === Zero-value disambiguation pass ===
    # For PDF=0 records that hit 'exception', check if the bridge has a row with the
    # SAME label where the value at the relevant year is 0 / missing. This handles
    # the case where bridge_by_key returns multiple rows (e.g., "Foreign" appears
    # in current/deferred/income-before-tax sections) and the primary match picked
    # the wrong one.
    for rec in all_records:
        if rec["status"] != "exception":
            continue
        pdf_value = rec.get("pdf_value")
        if pdf_value != 0:
            continue
        # Look in the bridge tab for any row with the same label key where the
        # year value is 0 / None / empty
        source_ref = rec.get("source_ref") or ""
        if not source_ref.startswith("bridge!"):
            continue
        tab_name = source_ref.split("!")[1] if "!" in source_ref else None
        if not tab_name or tab_name not in bridge_lookup:
            continue
        bridge_tab = bridge_lookup[tab_name]
        bridge_rows = extract_bridge_rows(bridge_tab)
        pdf_key = normalize_label(rec.get("pdf_label") or "")
        target_year = rec.get("pdf_year")
        for br in bridge_rows:
            if normalize_label(br["label"]) != pdf_key:
                continue
            yv = br["values"].get(target_year)
            # Tie if bridge value is None (empty / no activity) or 0
            if yv is None or yv == 0:
                rec["status"] = "ties"
                rec["source_label"] = br["label"]
                rec["source_value"] = 0.0
                rec["delta"] = 0
                rec["tolerance"] = 1.0
                rec["comparison_unit"] = "$K"
                orig = rec.get("notes") or ""
                rec["notes"] = (
                    f"[zero-value disambiguation: PDF=0 matched bridge row with no "
                    f"{target_year} value (alternate label position)]" + (f" | {orig}" if orig else "")
                )
                reclassified_zero += 1
                break
    print(f"  Reclassified {reclassified_zero} PDF=0 records via label-position disambiguation")

    for rec in all_records:
        if rec["status"] not in non_tie_statuses:
            continue
        pdf_value = rec.get("pdf_value")
        if pdf_value is None or pdf_value == 0:
            continue
        # Find the bridge tab from source_ref like "bridge!FN - PPE" or "bridge!FN - PPE!label"
        source_ref = rec.get("source_ref") or ""
        if not source_ref.startswith("bridge!"):
            continue
        tab_name = source_ref.split("!")[1] if "!" in source_ref else None
        if not tab_name or tab_name not in bridge_lookup:
            continue
        bridge_tab = bridge_lookup[tab_name]
        match = find_value_in_bridge_with_context(
            bridge_tab, pdf_value, rec.get("is_subtotal", False), rec.get("pdf_label") or "",
        )
        if not match:
            continue
        # Update the record in place — preserve original primary-match info in notes
        original_status = rec["status"]
        original_source_label = rec.get("source_label")
        original_source_value = rec.get("source_value")
        bridge_unit_norm = normalize_unit(bridge_tab.get("unit", "ambiguous"))
        # Determine new status: if labels match exactly (score=1.0), 'ties'; else caption-changed
        if match["similarity_score"] >= 0.99:
            new_status = "ties"
        else:
            new_status = "ties-caption-changed"
        rec["status"] = new_status
        rec["source_label"] = match["nearby_label"]
        rec["source_value"] = match["value_raw"]
        rec["comparison_unit"] = "$K"
        rec["delta"] = pdf_value - match["value_K"]
        rec["tolerance"] = 5.0 if rec.get("is_subtotal") else 1.0
        original_note = rec.get("notes") or ""
        rec["notes"] = (
            f"[secondary number match: row {match['row_idx']}, col {match['col_idx']}, "
            f"label='{match['nearby_label']}', score={match['similarity_score']:.2f}, "
            f"candidates={match['n_candidates']}; "
            f"primary status was '{original_status}', primary source='{original_source_label}' "
            f"value={original_source_value}]"
            + (f" | {original_note}" if original_note else "")
        )
        reclassified += 1
    print(f"  Reclassified {reclassified} records via secondary number match")

    # Write
    print(f"\nTotal records: {len(all_records)}")
    by_status = {}
    for r in all_records:
        by_status[r["status"]] = by_status.get(r["status"], 0) + 1
    print(f"By status: {by_status}")
    Path(args.out).write_text(
        json.dumps(all_records, default=str, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    print(f"Wrote {args.out}")


if __name__ == "__main__":
    main()
