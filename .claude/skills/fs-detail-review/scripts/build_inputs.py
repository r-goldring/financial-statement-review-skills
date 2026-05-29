"""Build inputs.json — the single source of truth for tie-out.

Aggregates structured data from all six tie-out sources:
  1. FY25 .docx (text source for the clean PDF — PDF has no extractable text)
  2. FY25 clean PDF (page count + page-classification only; OCR happens later in annotator)
  3. FY24 final FS PDF (text-extractable; for prior-year ties)
  4. FY25 bridge workbook (FS Compilation Partner's work file)
  5. Consolidated TB
  6. TB by Subsidiary

Detects display units per source. Hard-fails if any source is missing OR if any source
has a tab/page where unit cannot be determined.

CLI:
  python build_inputs.py <output_inputs.json>
"""

import argparse
import json
import re
import sys
from pathlib import Path
from datetime import datetime

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# Allow importing extract_bridge / extract_fs_pdf from sibling scripts
SCRIPT_DIR = Path(__file__).parent
sys.path.insert(0, str(SCRIPT_DIR))

from extract_bridge import extract_bridge
from extract_fs_pdf import extract_fs_pdf


# ---------- Path constants (FY25 inputs) ----------
ROOT = Path(r"C:\path\to\financial-statement-review")

FY25_PDF = ROOT / "Tieout" / "vfinal Tieou" / "Acme Holdings, LLC 2025 Financial Statements.pdf"
FY25_DOCX = ROOT / "Tieout" / "vfinal Tieou" / "Acme Holdings, LLC 2025 Financial Statements.docx"
FY25_BRIDGE = ROOT / "Tieout" / "vfinal Tieou" / "9. vYYYY.M.D_Acme Corp 2025 FS Bridge_Partial InScope Updates (vF).xlsx"
TB_CONSOLIDATED = ROOT / "Trial Balance" / "Consolidated TB FY25 v4.14.26.xlsx"
TB_BY_SUBSIDIARY = ROOT / "Trial Balance" / "TB by Subsidiary FY25 v4.15.26.xlsx"
FY24_FINAL_PDF = ROOT / "Prior Year Examples" / "2024" / "Financial Statements_FINAL" / "Acme Holdings LLC 2024 Financial Statements.pdf"


# ---------- Unit detection ----------
THOUSANDS_PHRASES = [
    "in thousands", "($ in thousands)", "$ in thousands",
    "amounts in u.s. dollars in thousands",
    "(amounts in u.s. dollars in thousands)",
    "(in thousands except", "(in thousands, except",
]

def detect_unit_from_text(text):
    """Return '$K' if text contains a thousands declaration, else None (unknown)."""
    if not text:
        return None
    lower = text.lower()
    for phrase in THOUSANDS_PHRASES:
        if phrase in lower:
            return "$K"
    return None


def detect_unit_from_values(values, label=""):
    """Heuristic: scan a list of numeric values for magnitude.
    Returns '$K' or '$1' or 'ambiguous'.

    Rules:
      - If max absolute value > N,NNN,NNN → almost certainly $1 (no FS line > $100M in $K display)
      - If max absolute value < NN,NNN AND multiple values present → ambiguous (could be unit count too)
      - If many values have decimal fractions → $1
      - Otherwise: ambiguous
    """
    nums = []
    for v in values:
        try:
            f = float(v)
            if abs(f) > 0.01:
                nums.append(f)
        except (TypeError, ValueError):
            continue
    if not nums:
        return "ambiguous"
    max_abs = max(abs(n) for n in nums)
    has_fractions = sum(1 for n in nums if abs(n - round(n)) > 0.01) > len(nums) * 0.1
    if max_abs > 100_000_000:
        return "$1"
    if has_fractions:
        return "$1"
    if max_abs < 1_000:
        return "ambiguous"
    return "$K"  # default for FS-shaped data


# ---------- 1. FY25 .docx (structured tables + paragraphs) ----------
def load_fy25_docx(path):
    import docx
    from docx.oxml.ns import qn

    doc = docx.Document(str(path))
    body = doc.element.body

    # Walk body to preserve order: each element is paragraph or table
    in_order = []
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            text = "".join([t.text or "" for t in child.iter(qn("w:t"))]).strip()
            in_order.append({"kind": "paragraph", "text": text})
        elif child.tag == qn("w:tbl"):
            rows = []
            for row in child.iterchildren(qn("w:tr")):
                cells = []
                for cell in row.iterchildren(qn("w:tc")):
                    cell_text = "".join([t.text or "" for t in cell.iter(qn("w:t"))]).strip()
                    cells.append(cell_text)
                rows.append(cells)
            in_order.append({"kind": "table", "rows": rows})

    # Group elements by FS section (BS / IS / SOE / SCF / Note 1 / Note 2 ...)
    # Section detection: look for FS-section headers
    section_markers = [
        ("Consolidated Balance Sheets", "BS"),
        ("Consolidated Statements of Operations", "IS"),
        ("Consolidated Statements of Changes in Members", "SOE"),
        ("Consolidated Statements of Cash Flows", "SCF"),
        ("Notes to Consolidated Financial Statements", "Notes-header"),
    ]
    # Note N detection: paragraph like "1. Summary of Significant Accounting Policies"
    note_pat = re.compile(r"^(\d{1,2})\.\s+(.+)")

    current_section = "cover"
    in_notes = False  # latches once we hit the Notes header; never resets
    elements_with_section = []
    for el in in_order:
        if el["kind"] == "paragraph":
            text = el["text"]
            matched_marker = False
            for marker, name in section_markers:
                if marker.lower() in text.lower() and len(text) < 200:
                    current_section = name
                    matched_marker = True
                    if name == "Notes-header":
                        in_notes = True
                    break
            if not matched_marker and in_notes:
                # Match footnote heading like "1. Description of Business" or "(1) ..."
                m = note_pat.match(text)
                m2 = None if m else re.match(r"^\((\d{1,2})\)\s+(.+)", text)
                if m or m2:
                    g = m or m2
                    num = g.group(1).zfill(2)
                    title = re.sub(r'[^a-z]', '', g.group(2).lower())[:24]
                    if title and len(text) < 200:
                        current_section = f"FN-{num}-{title}"
        el = dict(el)
        el["section"] = current_section
        elements_with_section.append(el)

    # Filter to keep only tables + identifying paragraphs
    tables_out = []
    for idx, el in enumerate(elements_with_section):
        if el["kind"] == "table":
            # Find the nearest preceding non-empty paragraph as title-context
            title = ""
            unit_paragraph = ""
            for j in range(idx - 1, max(-1, idx - 8), -1):
                p = elements_with_section[j]
                if p["kind"] == "paragraph" and p["text"]:
                    if not title:
                        title = p["text"][:100]
                    if "thousand" in p["text"].lower() or "thousand" in p["text"]:
                        unit_paragraph = p["text"]
                    if title and unit_paragraph:
                        break
            # Detect unit
            unit = detect_unit_from_text(unit_paragraph) or detect_unit_from_text(title)
            if not unit:
                # Try values
                all_vals = []
                for row in el["rows"]:
                    for cell in row:
                        m = re.search(r"\(?[\$\-]?[\d,]+(?:\.\d+)?\)?", cell or "")
                        if m:
                            s = m.group().replace("$", "").replace(",", "").strip()
                            is_neg = s.startswith("(") and s.endswith(")")
                            s = s.strip("()")
                            try:
                                v = float(s)
                                if is_neg:
                                    v = -v
                                if abs(v) > 100:  # skip page numbers
                                    all_vals.append(v)
                            except ValueError:
                                continue
                unit = detect_unit_from_values(all_vals, label=title)
            tables_out.append({
                "idx": idx,
                "section": el["section"],
                "title": title,
                "unit": unit,
                "unit_source": "header" if unit_paragraph else ("title" if detect_unit_from_text(title) else "magnitude"),
                "rows": el["rows"],
                "row_count": len(el["rows"]),
            })

    paragraphs_out = []
    for idx, el in enumerate(elements_with_section):
        if el["kind"] == "paragraph" and el["text"]:
            paragraphs_out.append({
                "idx": idx,
                "section": el["section"],
                "text": el["text"],
            })

    # pdf2docx (used when the source is a final PDF instead of a hand-edited docx)
    # emits the dollar sign on $-prefixed amounts as a standalone cell — so a row like
    # "Cash $X,XXX.XX $X,XXX.XX" becomes 5 cells: [label, '$', 'NN,NNN', '$', 'NN,NNN'],
    # while non-$ rows stay 3 cells. The lane parsers expect uniform shape, so drop
    # standalone single-character marker cells ('$', '(', ')'). Harmless on a
    # hand-edited docx (those cells don't exist there).
    _MARKERS = {"$", "(", ")", ""}
    for t in tables_out:
        cleaned = []
        for row in t.get("rows", []):
            cleaned.append([c for c in row if str(c).strip() not in _MARKERS])
        t["rows"] = cleaned
        t["row_count"] = len(cleaned)

    # Face statements (BS/IS/SOE/SCF) are logically one table per section, but if the
    # docx was generated from a PDF (e.g. pdf2docx) the table can be split at page
    # boundaries into several. Lanes 1/3/4/6 take the first matching table per
    # section, so a split breaks them. Merge those rows back into one logical table.
    # Footnote tables stay separate — each FN disclosure is its own table.
    FACE_SECTIONS = {"BS", "IS", "SOE", "SCF"}
    merged_out = []
    by_section = {}
    for t in tables_out:
        sec = t.get("section")
        if sec in FACE_SECTIONS:
            by_section.setdefault(sec, []).append(t)
        else:
            merged_out.append(t)
    for sec, parts in by_section.items():
        rows = [r for p in parts for r in p.get("rows", [])]
        first = parts[0]
        merged_out.append({**first, "rows": rows, "row_count": len(rows),
                           "merged_from": [p["idx"] for p in parts] if len(parts) > 1 else None})
    # Keep emission ordered by original idx for stability.
    merged_out.sort(key=lambda t: t.get("idx", 0))

    return {
        "path": str(path),
        "tables": merged_out,
        "paragraphs": paragraphs_out,
        "section_summary": _summarize_sections(elements_with_section),
    }


def _summarize_sections(elements):
    sections = {}
    for el in elements:
        sec = el.get("section", "unknown")
        sections.setdefault(sec, {"paragraphs": 0, "tables": 0})
        if el["kind"] == "paragraph":
            sections[sec]["paragraphs"] += 1
        else:
            sections[sec]["tables"] += 1
    return sections


# ---------- 2. FY25 PDF (page count + classification only) ----------
def load_fy25_pdf(path):
    """No text on this PDF — just record page count and rect dims.
    Page classification (BS, IS, etc.) is inferred from FY25 docx's section markers
    AND validated against the TOC (BS on doc page 4 → PDF page 5).
    """
    import fitz
    doc = fitz.open(str(path))
    pages = []
    for i in range(len(doc)):
        p = doc[i]
        pages.append({
            "page_number": i + 1,
            "width": p.rect.width,
            "height": p.rect.height,
            "rotation": p.rotation,
            "drawing_count": len(p.get_drawings()),
        })
    doc.close()
    # FY25 standard layout (mirroring FY24):
    #   PDF p.1 = Cover, p.2 = TOC, p.3 = Auditor's Report, p.4 = Auditor's continuation/blank,
    #   p.5 = BS, p.6 = IS, p.7 = SOE, p.8 = SCF, p.9+ = Notes
    classification = {
        1: "cover", 2: "toc", 3: "auditor", 4: "auditor",
        5: "BS", 6: "IS", 7: "SOE", 8: "SCF",
    }
    for p in pages:
        n = p["page_number"]
        p["kind"] = classification.get(n, f"FN-page-{n:02d}")
    return {"path": str(path), "page_count": len(pages), "pages": pages}


# ---------- 3. FY24 final FS PDF (text extraction) ----------
def load_fy24_final_pdf(path):
    """Extract text per page; detect unit per page from header text."""
    raw = extract_fs_pdf(str(path), min_chars=50)
    pages_out = []
    for p in raw["pages"]:
        unit = detect_unit_from_text(p["text"][:500])  # check first ~500 chars (header area)
        if unit is None:
            # fall back: scan numbers
            nums = []
            for tok in re.findall(r"\(?\$?-?[\d,]+(?:\.\d+)?\)?", p["text"]):
                clean = tok.replace("$", "").replace(",", "")
                is_neg = clean.startswith("(") and clean.endswith(")")
                clean = clean.strip("()")
                try:
                    v = float(clean)
                    if is_neg:
                        v = -v
                    if abs(v) > 100 and not (2020 <= v <= 2030 and v == int(v)):
                        nums.append(v)
                except ValueError:
                    continue
            unit = detect_unit_from_values(nums)
        pages_out.append({
            "page_number": p["page_number"],
            "char_count": p["char_count"],
            "unit": unit,
            "text": p["text"],
        })
    return {"path": str(path), "page_count": raw["page_count"], "pages": pages_out}


# ---------- 4. Bridge workbook ----------
BRIDGE_FACE_TABS = {"BS", "PL", "SOE", "SCF", "Balance Sheet", "Income Statement", "Equity Rollforward", "SOCF"}

def load_bridge(path):
    raw = extract_bridge(str(path))
    tabs_out = []
    for name in raw["sheet_names"]:
        sheet = raw["sheets"][name]
        rows = sheet["rows"]
        # Detect unit
        header_lines = []
        for row in rows[:10]:
            line = " ".join(str(c) for c in row if c is not None)
            header_lines.append(line)
        header_text = " | ".join(header_lines)
        unit = detect_unit_from_text(header_text)
        # Look for explicit unit column markers
        all_text_in_top = header_text.lower()
        has_rounded_000 = "rounded (000s)" in all_text_in_top or "rounded(000s)" in all_text_in_top or "(in thousands)" in all_text_in_top or "(000s)" in all_text_in_top
        if not unit and has_rounded_000:
            # Bridge has BOTH raw $ and $K columns — primary unit varies but $K column exists
            unit = "$1-with-$K-column"  # special: structured per FY24 convention
        if not unit:
            # Magnitude check
            nums = []
            for row in rows:
                for c in row:
                    try:
                        v = float(c)
                        if abs(v) > 0.01:
                            nums.append(v)
                    except (TypeError, ValueError):
                        continue
            unit = detect_unit_from_values(nums)
        tabs_out.append({
            "name": name,
            "unit": unit,
            "max_row": sheet["max_row"],
            "max_col": sheet["max_col"],
            "rows": rows,
        })
    return {
        "path": str(path),
        "strategy": raw["strategy"],
        "tab_count": len(tabs_out),
        "tabs": tabs_out,
    }


# ---------- 5. Consolidated TB ----------
def load_tb_consolidated(path):
    raw = extract_bridge(str(path))
    sheet = raw["sheets"]["TrialBalance"]
    rows = sheet["rows"]
    # Find header row (contains "Account" and "Total" or similar)
    header_idx = None
    for i, row in enumerate(rows[:30]):
        joined = " ".join(str(c) for c in row if c is not None).lower()
        if "account" in joined and ("total" in joined or "balance" in joined):
            header_idx = i
            break
    accounts = []
    if header_idx is not None:
        for row in rows[header_idx + 1:]:
            if not row or all(c is None or (isinstance(c, str) and not c.strip()) for c in row):
                continue
            label = row[0] if len(row) > 0 else None
            value = row[1] if len(row) > 1 else None
            if label is None or not str(label).strip():
                continue
            try:
                v = float(value) if value is not None else None
            except (TypeError, ValueError):
                v = None
            accounts.append({"account": str(label).strip(), "value": v})
    return {
        "path": str(path),
        "unit": "$1",  # TB always raw $
        "header_row_idx": header_idx,
        "account_count": len(accounts),
        "accounts": accounts,
        "_raw_sheets": list(raw["sheet_names"]),
    }


# ---------- 6. TB by Subsidiary ----------
def load_tb_by_subsidiary(path):
    raw = extract_bridge(str(path))
    entities = {}
    for name in raw["sheet_names"]:
        sheet = raw["sheets"][name]
        rows = sheet["rows"]
        # Same header-detection strategy
        header_idx = None
        for i, row in enumerate(rows[:30]):
            joined = " ".join(str(c) for c in row if c is not None).lower()
            if "account" in joined and ("total" in joined or "balance" in joined or "amount" in joined):
                header_idx = i
                break
        accounts = []
        if header_idx is not None:
            for row in rows[header_idx + 1:]:
                if not row or all(c is None or (isinstance(c, str) and not c.strip()) for c in row):
                    continue
                label = row[0] if len(row) > 0 else None
                value = row[1] if len(row) > 1 else None
                if label is None or not str(label).strip():
                    continue
                try:
                    v = float(value) if value is not None else None
                except (TypeError, ValueError):
                    v = None
                accounts.append({"account": str(label).strip(), "value": v})
        entities[name] = {"unit": "$1", "header_row_idx": header_idx, "account_count": len(accounts), "accounts": accounts}
    return {"path": str(path), "entities": entities}


# ---------- Main ----------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("output", help="Path to write inputs.json")
    ap.add_argument("--skip-fy25-pdf", action="store_true", help="skip FY25 PDF load (debug)")
    args = ap.parse_args()

    out_path = Path(args.output)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # Hard-fail if any input missing
    inputs = {
        "FY25_PDF": FY25_PDF,
        "FY25_DOCX": FY25_DOCX,
        "FY25_BRIDGE": FY25_BRIDGE,
        "TB_CONSOLIDATED": TB_CONSOLIDATED,
        "TB_BY_SUBSIDIARY": TB_BY_SUBSIDIARY,
        "FY24_FINAL_PDF": FY24_FINAL_PDF,
    }
    for name, p in inputs.items():
        if not p.exists():
            print(f"MISSING: {name} = {p}", file=sys.stderr)
            sys.exit(1)
        print(f"  found  {name:20s}  {p}")

    print("\nLoading FY25 .docx (tables + paragraphs)...")
    fy25_docx = load_fy25_docx(FY25_DOCX)
    print(f"  {len(fy25_docx['tables'])} tables, {len(fy25_docx['paragraphs'])} paragraphs")
    print(f"  sections detected: {sorted(fy25_docx['section_summary'].keys())}")

    print("\nLoading FY25 clean PDF (page meta only)...")
    if args.skip_fy25_pdf:
        fy25_pdf = {"path": str(FY25_PDF), "page_count": 0, "pages": [], "_skipped": True}
    else:
        fy25_pdf = load_fy25_pdf(FY25_PDF)
        print(f"  {fy25_pdf['page_count']} pages")

    print("\nLoading FY24 final FS PDF (text)...")
    fy24_pdf = load_fy24_final_pdf(FY24_FINAL_PDF)
    units = sorted(set(p["unit"] for p in fy24_pdf["pages"] if p["unit"]))
    print(f"  {fy24_pdf['page_count']} pages, units found: {units}")

    print("\nLoading FY25 bridge workbook...")
    bridge = load_bridge(FY25_BRIDGE)
    units_per_tab = {t["name"]: t["unit"] for t in bridge["tabs"]}
    print(f"  {bridge['tab_count']} tabs, strategy: {bridge['strategy']}")
    print(f"  unit summary:")
    for n, u in units_per_tab.items():
        print(f"    {n:35s}  {u}")

    print("\nLoading Consolidated TB...")
    tb_cons = load_tb_consolidated(TB_CONSOLIDATED)
    print(f"  {tb_cons['account_count']} accounts (header at row {tb_cons['header_row_idx']})")

    print("\nLoading TB by Subsidiary...")
    tb_sub = load_tb_by_subsidiary(TB_BY_SUBSIDIARY)
    for entity, data in tb_sub["entities"].items():
        print(f"  {entity:20s}  {data['account_count']} accounts")

    # Check for ambiguous units
    ambiguous = []
    for tab in bridge["tabs"]:
        if tab["unit"] == "ambiguous":
            ambiguous.append(f"bridge tab: {tab['name']}")
    for table in fy25_docx["tables"]:
        if table["unit"] == "ambiguous":
            ambiguous.append(f"fy25-docx table: {table['title'][:50]}")
    for p in fy24_pdf["pages"]:
        if p["unit"] == "ambiguous":
            ambiguous.append(f"fy24-pdf page: {p['page_number']}")

    if ambiguous:
        print("\n⚠ Ambiguous units (not hard-failing — review and tag):")
        for a in ambiguous:
            print(f"  - {a}")

    # Write JSON
    result = {
        "metadata": {
            "built_at": datetime.utcnow().isoformat(),
            "fy25_pdf": str(FY25_PDF),
            "fy25_docx": str(FY25_DOCX),
            "fy25_bridge": str(FY25_BRIDGE),
            "tb_consolidated": str(TB_CONSOLIDATED),
            "tb_by_subsidiary": str(TB_BY_SUBSIDIARY),
            "fy24_final_pdf": str(FY24_FINAL_PDF),
            "ambiguous_units": ambiguous,
        },
        "fy25_docx": fy25_docx,
        "fy25_pdf": fy25_pdf,
        "fy24_pdf": fy24_pdf,
        "bridge": bridge,
        "tb_consolidated": tb_cons,
        "tb_subsidiary": tb_sub,
    }

    print(f"\nWriting {out_path} ...")
    out_path.write_text(
        json.dumps(result, default=str, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    print(f"  size: {out_path.stat().st_size:,} bytes")


if __name__ == "__main__":
    main()
