"""Detect which PDF page each footnote section starts on.

Strategy:
  - Render pages 9-28 of the FY25 PDF to PNG (re-uses OCR cache if available)
  - OCR each page
  - Find footnote section headings: text matching r'^\d{1,2}\.\s+' patterns
  - Map docx section name → starting PDF page

Writes a JSON: {section_name: pdf_page, ...} to enable Lane 1 to set pdf_page on FN records.

CLI:
  python detect_fn_pages.py <pdf_path> <out_json> [--ocr-cache <path>]
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


# Map docx section name → list of keywords to search for in OCR text
# (first match wins; section headers tend to be like "1. Description of Business")
SECTION_KEYWORDS = {
    "FN-01-descriptionofbusiness":    ["description of business"],
    "FN-02-summaryofsignificantacco": ["summary of significant accounting policies"],
    # NOTE: OCR may transcribe ", net" as "; net" or "net" — list variants
    "FN-03-propertyandequipmentnet":  ["property and equipment; net", "property and equipment, net",
                                       "property and equipment net"],
    "FN-04-intangibleassetsandgoodw": ["intangible assets and goodwill"],
    "FN-05-accruedexpensesandotherc": ["accrued expenses and other current liabilities"],
    "FN-06-termloansandlineofcredit": ["term loans and line of credit"],
    "FN-07-incometaxes":              ["income taxes"],
    "FN-08-membersequity":            ["members' equity", "members equity"],
    "FN-09-stockbasedcompensation":   ["stock-based compensation", "stock based compensation"],
    "FN-10-commitmentsandcontingenc": ["commitments and contingencies"],
    "FN-11-leases":                   ["leases"],
    "FN-12-relatedparties":           ["related parties"],
    "FN-13-employeebenefitplan":      ["employee benefit plan"],
    "FN-14-subsequentevents":         ["subsequent events"],
}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf_path")
    ap.add_argument("out_json")
    ap.add_argument("--ocr-cache", default=None)
    ap.add_argument("--start-page", type=int, default=9, help="First page to scan")
    ap.add_argument("--dpi", type=int, default=200)
    args = ap.parse_args()

    import fitz
    import easyocr

    doc = fitz.open(args.pdf_path)
    last_page = len(doc)
    print(f"PDF has {last_page} pages; scanning from page {args.start_page}")

    cache = {}
    cache_path = Path(args.ocr_cache) if args.ocr_cache else None
    if cache_path and cache_path.exists():
        cache = json.loads(cache_path.read_text(encoding="utf-8"))
        print(f"Loaded OCR cache with {len(cache)} pages")

    print("Loading easyocr...")
    reader = easyocr.Reader(["en"], gpu=False, verbose=False)
    print()

    page_to_text = {}
    for pg in range(args.start_page, last_page + 1):
        cache_key = f"page-{pg}"
        if cache_key in cache:
            tokens = cache[cache_key]
            page_text = " ".join(t["text"] for t in tokens)
            print(f"  p.{pg:>2}  (cached, {len(tokens)} tokens)")
        else:
            page = doc[pg - 1]
            pix = page.get_pixmap(dpi=args.dpi)
            png_bytes = pix.tobytes("png")
            results = reader.readtext(png_bytes, paragraph=False)
            tokens = [
                {"bbox": [[float(p[0]), float(p[1])] for p in bbox],
                 "text": str(text), "conf": float(conf)}
                for bbox, text, conf in results
            ]
            cache[cache_key] = tokens
            page_text = " ".join(t["text"] for t in tokens)
            print(f"  p.{pg:>2}  OCR ran ({len(tokens)} tokens)")
        page_to_text[pg] = page_text.lower()

    doc.close()

    # Save cache
    if cache_path:
        cache_path.write_text(json.dumps(cache, indent=2), encoding="utf-8")

    # Find each FN section's first page
    section_to_page = {}
    for section, keywords in SECTION_KEYWORDS.items():
        found_page = None
        for pg in sorted(page_to_text.keys()):
            text = page_to_text[pg]
            for kw in keywords:
                # Pattern: "(N)." or "N." preceded by space, followed by keyword
                # Robust to OCR variability: just check if keyword appears
                # but prefer earliest occurrence
                if kw in text:
                    # Confirm there's a section header nearby (a "N." pattern)
                    # Look for "N." or "(N)" within ~150 chars before keyword
                    idx = text.index(kw)
                    snippet = text[max(0, idx - 150):idx]
                    if re.search(r"\(\d{1,2}\)\s*$", snippet) or re.search(r"\d{1,2}\.\s*$", snippet) or len(snippet) < 30:
                        found_page = pg
                        break
                    # Also accept if it's near start of page
                    if idx < 100:
                        found_page = pg
                        break
            if found_page:
                break
        section_to_page[section] = found_page

    # Print results
    print("\nFootnote section → PDF page:")
    for section, page in section_to_page.items():
        if page is None:
            print(f"  {section:<35s}  ⚠ NOT FOUND")
        else:
            print(f"  {section:<35s}  p.{page}")

    Path(args.out_json).write_text(
        json.dumps(section_to_page, indent=2),
        encoding="utf-8",
    )
    print(f"\nWrote {args.out_json}")


if __name__ == "__main__":
    main()
