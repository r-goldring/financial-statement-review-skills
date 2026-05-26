"""Lane 3: FY25 PDF prior-year (2024) column ↔ FY24 final FS PDF.

For every line on the FY25 PDF that has a 2024 column value, find the matching
line on the FY24 final FS PDF and confirm exact match. Any delta = restatement.

CLI:
  python tie_out_pdf_prior_year.py <inputs.json> <out.json>
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

from tie_out_common import parse_value, compare, normalize_label, make_record, is_subtotal_label, fuzzy_match


# PDF page mapping (FY25 PDF physical pages)
SECTION_TO_PDF_PAGE = {
    "BS": 5, "IS": 6, "SOE": 7, "SCF": 8,
}

# FY24 final FS PDF page mapping
# (verified: FY24 final has BS on page 5, IS on page 6, etc.)
FY24_SECTION_TO_PAGE = {
    "BS": 5, "IS": 6, "SOE": 7, "SCF": 8,
}


def extract_pdf_rows_from_docx_table(table):
    """Same row extraction as lane 1 — pull (label, {year: value}) tuples."""
    rows = table.get("rows", [])
    if len(rows) < 3:
        return []
    year_cols = {}
    header_row_idx = None
    for i, row in enumerate(rows[:5]):
        row_year_cells = {}
        for c_idx, cell in enumerate(row):
            if c_idx == 0:
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


def extract_fy24_pdf_face_lines(fy24_pages, section):
    """Parse the FY24 PDF text (extractable) to get (label → value) for one section.

    The text format from PyMuPDF flattens each row into newline-separated tokens like:
      Cash and cash equivalents
      $
      NN,NNN
      $
      NN,NNN

    Walk lines using a small state machine:
      - State 'pre-table': skip until we find the "As of ..." or "Year Ended ..." header
      - State 'in-table': alternate between label (text) and 2 values (numbers)
      - End-of-table on "See accompanying notes" or similar
    """
    page_num = FY24_SECTION_TO_PAGE.get(section)
    if page_num is None:
        return {}
    page = next((p for p in fy24_pages if p["page_number"] == page_num), None)
    if not page:
        return {}
    text = page["text"]
    lines = [l.strip() for l in text.split("\n") if l.strip()]

    # Locate the header pattern. Skip everything until we see the year header row.
    in_table = False
    label_to_values = {}
    label_parts = []
    values_collected = []

    SKIP_PHRASES = {
        "Assets", "Liabilities and Members' Equity", "Liabilities and Members’ Equity",
        "Current assets:", "Current liabilities:", "Members' equity:", "Members’ equity:",
        "Operating expenses:", "Other income (expense):", "Other comprehensive income (loss):",
        "Cash flows from operating activities:", "Cash flows from investing activities:",
        "Cash flows from financing activities:", "Adjustments to reconcile net loss to net cash and cash equivalents provided by operating activities:",
        "Changes in operating assets and liabilities, net of business acquisitions:",
        "Supplemental disclosures", "Supplemental disclosures for noncash investing and financing activities",
    }
    END_PHRASES = ("See accompanying notes", "Notes to ", "The accompanying notes")

    def commit():
        nonlocal label_parts, values_collected
        if label_parts and values_collected:
            label = " ".join(label_parts).strip()
            if label not in SKIP_PHRASES and label not in label_to_values:
                fy24 = values_collected[0] if len(values_collected) >= 1 else None
                fy23 = values_collected[1] if len(values_collected) >= 2 else None
                label_to_values[label] = {"fy24": fy24, "fy23": fy23}
        label_parts = []
        values_collected = []

    for line in lines:
        # End-of-table detection
        if any(line.startswith(end) for end in END_PHRASES):
            commit()
            break

        if not in_table:
            # Look for table-start signal
            if line in ("As of December 31,", "Year Ended December 31,",
                        "Year Ended December 31, ", "For the Years Ended December 31,"):
                in_table = True
            continue

        # Skip currency symbols
        if line in ("$", "$ "):
            continue
        # Skip year-only lines
        if re.match(r"^20\d{2}$", line):
            continue
        # Skip explicit section-heading phrases
        if line in SKIP_PHRASES:
            label_parts = []  # discard any partial label fragments accumulated
            continue
        # NOTE: Do NOT skip "page-number-like" 1-2 digit strings here. Page numbers
        # in the FY24 PDF appear OUTSIDE the table (top/bottom margins), and the
        # in_table state machine already filters them. Skipping legitimate small
        # values (like $24 disposal-loss, $12 other-assets-change) drops real ties.

        # Multi-line section headers in FY24 PDF — e.g.,
        #   "Adjustments to reconcile net loss to net cash, cash equivalents and"
        #   "restricted cash provided by (used in) operating activities:"
        # are split across two lines and never match a single SKIP_PHRASES entry.
        # Detect them by their trailing colon: any line ending with ":" is a section
        # header; discard any partial label fragments accumulated so far.
        if line.rstrip().endswith(":"):
            label_parts = []
            continue

        # Number test (strict: optional paren, optional $, digits)
        is_number = bool(re.match(r"^\(?\$?-?[\d,]+(?:\.\d+)?\)?$", line))
        if is_number:
            v = parse_value(line)
            if v is not None:
                values_collected.append(v)
                if len(values_collected) >= 2:
                    commit()
        else:
            # If we have already-collected values, commit them and start a new label
            if values_collected:
                commit()
            label_parts.append(line)

    commit()
    return label_to_values


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("inputs")
    ap.add_argument("out")
    args = ap.parse_args()

    data = json.loads(Path(args.inputs).read_text(encoding="utf-8"))
    docx_tables = data["fy25_docx"]["tables"]
    fy24_pages = data["fy24_pdf"]["pages"]

    all_records = []

    for section in ["BS", "IS", "SOE", "SCF"]:
        fy25_table = next((t for t in docx_tables if t["section"] == section), None)
        if not fy25_table:
            print(f"  ⚠ FY25 has no docx table for section {section}")
            continue

        fy25_unit = fy25_table["unit"]
        fy24_lines = extract_fy24_pdf_face_lines(fy24_pages, section)
        if not fy24_lines:
            print(f"  ⚠ FY24 has no face-section data for {section}")
            continue

        fy24_unit = "$K"  # FY24 final FS face statements are in $K
        fy24_labels_raw = list(fy24_lines.keys())

        rows = extract_pdf_rows_from_docx_table(fy25_table)
        ties = 0
        exc = 0
        missing = 0
        for r in rows:
            if "2024" not in r["values"]:
                continue
            fy25_2024_value = r["values"]["2024"]
            # Find matching FY24 row
            label = r["label"]
            label_norm = normalize_label(label)
            match = None
            for fy24_label, vals in fy24_lines.items():
                if normalize_label(fy24_label) == label_norm:
                    match = (fy24_label, vals)
                    break
            if not match:
                # Direction-aware fuzzy: SCF has "beginning of year" / "end of year"
                # lines that mustn't cross-match. If the FY25 label is directional,
                # restrict the FY24 candidate set to labels with the same direction.
                label_low = label.lower()
                if "beginning of year" in label_low or "beginning of period" in label_low:
                    candidates = [l for l in fy24_labels_raw
                                  if "beginning of year" in l.lower()
                                  or "beginning of period" in l.lower()]
                elif "end of year" in label_low or "end of period" in label_low:
                    candidates = [l for l in fy24_labels_raw
                                  if "end of year" in l.lower()
                                  or "end of period" in l.lower()]
                else:
                    candidates = fy24_labels_raw
                fz, score = fuzzy_match(label, candidates, threshold=0.60)
                if fz:
                    match = (fz, fy24_lines[fz])

            if not match:
                # Secondary check: number-only match
                # Look across FY24 lines on the same section for a numeric match within $1K
                number_match = None
                for fy24_label, vals in fy24_lines.items():
                    fy24_v = vals.get("fy24")
                    if fy24_v is None:
                        continue
                    if abs(fy24_v - fy25_2024_value) <= 1.0:
                        number_match = (fy24_label, vals)
                        break
                if number_match:
                    fy24_label, fy24_vals = number_match
                    fy24_value = fy24_vals["fy24"]
                    delta = fy25_2024_value - fy24_value
                    rec = make_record(
                        lane="pdf_prior_year",
                        pdf_section=section,
                        pdf_page=SECTION_TO_PDF_PAGE.get(section),
                        pdf_label=label,
                        pdf_year="2024",
                        pdf_value=fy25_2024_value,
                        source_ref=f"FY24-FS!p{FY24_SECTION_TO_PAGE.get(section)}",
                        source_label=fy24_label,
                        source_value=fy24_value,
                        comparison_unit="$K",
                        delta=delta,
                        tolerance=1.0,
                        status="ties-caption-changed",
                        is_subtotal=r["is_subtotal"],
                        notes=f"Number ties (delta={delta:.0f}); caption changed from '{fy24_label}' to '{label}' — REVIEW for reasonability",
                    )
                    all_records.append(rec)
                    ties += 1
                    continue
                # If PDF value is $0, the line had no activity in 2024 — and not
                # finding it on FY24 likely means it's a NEW disclosure line that
                # didn't exist on FY24 (so was also effectively $0 there). Treat as tied.
                if fy25_2024_value == 0:
                    rec = make_record(
                        lane="pdf_prior_year",
                        pdf_section=section,
                        pdf_page=SECTION_TO_PDF_PAGE.get(section),
                        pdf_label=label,
                        pdf_year="2024",
                        pdf_value=fy25_2024_value,
                        source_ref=f"FY24-FS!p{FY24_SECTION_TO_PAGE.get(section)}",
                        source_label="N/A — line not on FY24 (zero activity in both)",
                        source_value=0.0,
                        comparison_unit="$K",
                        delta=0,
                        tolerance=0,
                        status="ties",
                        is_subtotal=r["is_subtotal"],
                        notes=f"PDF=$0 and no FY24 line matched — likely new disclosure with no FY24 activity",
                    )
                    all_records.append(rec)
                    ties += 1
                    continue
                rec = make_record(
                    lane="pdf_prior_year",
                    pdf_section=section,
                    pdf_page=SECTION_TO_PDF_PAGE.get(section),
                    pdf_label=label,
                    pdf_year="2024",
                    pdf_value=fy25_2024_value,
                    source_ref=f"FY24-FS!p{FY24_SECTION_TO_PAGE.get(section)}",
                    source_label=None,
                    source_value=None,
                    status="missing-on-fy24-pdf",
                    is_subtotal=r["is_subtotal"],
                    notes=f"No FY24 line matched label '{label}'; no FY24 value within $1K either",
                )
                all_records.append(rec)
                missing += 1
                continue

            fy24_label, fy24_vals = match
            fy24_value = fy24_vals["fy24"]
            delta, status, tolerance, comparison_unit = compare(
                fy25_2024_value, fy24_value, fy25_unit, fy24_unit,
                is_subtotal=r["is_subtotal"], kind="prior_year",
            )
            # For prior-year tie-out: any non-tie status promoted to "restatement"
            if status not in ("ties", "missing"):
                if status == "exception":
                    status = "restatement"
            rec = make_record(
                lane="pdf_prior_year",
                pdf_section=section,
                pdf_page=SECTION_TO_PDF_PAGE.get(section),
                pdf_label=label,
                pdf_year="2024",
                pdf_value=fy25_2024_value,
                source_ref=f"FY24-FS!p{FY24_SECTION_TO_PAGE.get(section)}",
                source_label=fy24_label,
                source_value=fy24_value,
                comparison_unit=comparison_unit,
                delta=delta,
                tolerance=tolerance,
                status=status,
                is_subtotal=r["is_subtotal"],
            )
            all_records.append(rec)
            if status == "ties":
                ties += 1
            elif status == "restatement":
                exc += 1
        print(f"  {section:>4} FY25 prior-yr <-> FY24 final  total={len([x for x in all_records if x['pdf_section']==section])}  ties={ties}  restatements={exc}  missing={missing}")

    by_status = {}
    for r in all_records:
        by_status[r["status"]] = by_status.get(r["status"], 0) + 1
    print(f"\nBy status: {by_status}")
    Path(args.out).write_text(
        json.dumps(all_records, default=str, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    print(f"Wrote {args.out}")


if __name__ == "__main__":
    main()
