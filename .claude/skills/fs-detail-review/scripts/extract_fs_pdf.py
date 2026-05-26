"""Robust FS PDF extractor.

Strategies, in order:
  1. PyMuPDF (fitz) — handles CID-encoded fonts that pdfplumber can't decode
  2. Mark page as "needs OCR" if extracted text is < min_chars (does not run OCR
     by default; surfaces the page list so the human can decide)

Outputs JSON:
  {
    "path": str,
    "page_count": int,
    "pages": [
      {"page_number": int, "char_count": int, "text": str, "needs_review": bool},
      ...
    ],
    "low_text_pages": [int, ...],
    "tables": [   # only the heuristic line-grouping; for real table extraction call extract_tables=True
      ...
    ]
  }

CLI:
  python extract_fs_pdf.py <fs.pdf>                            # summary
  python extract_fs_pdf.py <fs.pdf> --json <out>               # full JSON
  python extract_fs_pdf.py <fs.pdf> --page N                   # one page text
  python extract_fs_pdf.py <fs.pdf> --pages-needing-review     # list low-text pages
  python extract_fs_pdf.py <fs.pdf> --min-chars 50             # custom threshold
"""

import argparse
import json
import sys
import io
from pathlib import Path

# UTF-8 stdout on Windows so console doesn't choke on em-dashes / smart quotes
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")


def extract_fs_pdf(path, min_chars=50, with_blocks=False, render_dir=None, ocr=False, dpi=150):
    """Extract text from a financial-statement PDF.

    Pages with text below `min_chars` are flagged for review. If `render_dir`
    is set, those pages are also rendered to PNG (for human visual review or
    downstream OCR).

    If `ocr=True` and pytesseract is installed, low-text pages are OCR'd and
    the recovered text is added to the page record under `text_ocr`.
    """
    import fitz
    doc = fitz.open(path)
    pages = []
    low_text_pages = []
    for i in range(len(doc)):
        page = doc[i]
        text = page.get_text() or ""
        char_count = len(text.strip())
        needs_review = char_count < min_chars
        record = {
            "page_number": i + 1,
            "char_count": char_count,
            "text": text,
            "needs_review": needs_review,
            "image_count": len(page.get_images()),
            "drawing_count": len(page.get_drawings()),
        }
        if with_blocks:
            blocks = page.get_text("dict").get("blocks", [])
            record["block_count"] = len(blocks)
            record["blocks"] = blocks

        if needs_review:
            low_text_pages.append(i + 1)
            if render_dir:
                from pathlib import Path
                Path(render_dir).mkdir(parents=True, exist_ok=True)
                pix = page.get_pixmap(dpi=dpi)
                png_path = Path(render_dir) / f"page-{i+1:02d}.png"
                pix.save(str(png_path))
                record["rendered_png"] = str(png_path)
            if ocr:
                try:
                    import pytesseract
                    from PIL import Image
                    if "rendered_png" not in record:
                        pix = page.get_pixmap(dpi=dpi)
                        record["_pix"] = pix
                        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                    else:
                        img = Image.open(record["rendered_png"])
                    record["text_ocr"] = pytesseract.image_to_string(img)
                except ImportError:
                    record["ocr_error"] = "pytesseract not installed; pip install pytesseract + install tesseract binary"
                except Exception as e:
                    record["ocr_error"] = f"{type(e).__name__}: {e}"

        pages.append(record)
    doc.close()

    return {
        "path": str(path),
        "page_count": len(pages),
        "pages": pages,
        "low_text_pages": low_text_pages,
        "min_chars_threshold": min_chars,
    }


def find_numbers_on_page(page_text):
    """Heuristic: find dollar-amount-shaped tokens for tie-out.

    Matches:
      NN,NNN    or    ($X,XXX.XX)    or    NN,NNN.56    or    ($X,XXX.XX)
    Returns list of floats (parentheses → negative).
    """
    import re
    pattern = re.compile(r"\(?\$?-?[\d,]+(?:\.\d+)?\)?")
    nums = []
    for tok in pattern.findall(page_text):
        clean = tok.replace("$", "").replace(",", "")
        is_neg = clean.startswith("(") and clean.endswith(")")
        clean = clean.strip("()")
        if clean in ("", "-", "."):
            continue
        try:
            v = float(clean)
            if is_neg:
                v = -v
            # Drop trivial integers (page numbers, footnote refs, years 2020-2030)
            if abs(v) < 100 or (2020 <= v <= 2030 and v == int(v)):
                continue
            nums.append(v)
        except ValueError:
            continue
    return nums


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("path", help="Path to PDF")
    ap.add_argument("--json", help="Write full JSON dump (pages + text)")
    ap.add_argument("--page", type=int, help="Print text of one page (1-indexed)")
    ap.add_argument("--pages-needing-review", action="store_true",
                    help="Print only the page numbers that returned low text")
    ap.add_argument("--min-chars", type=int, default=50,
                    help="Threshold (chars) below which a page is flagged needs_review")
    ap.add_argument("--numbers", action="store_true",
                    help="Print dollar-amount tokens found on each page")
    ap.add_argument("--render-dir", help="Directory to render low-text pages to PNG")
    ap.add_argument("--ocr", action="store_true", help="OCR low-text pages (needs pytesseract)")
    ap.add_argument("--dpi", type=int, default=150)
    args = ap.parse_args()

    result = extract_fs_pdf(args.path, min_chars=args.min_chars,
                            render_dir=args.render_dir, ocr=args.ocr, dpi=args.dpi)

    if args.json:
        Path(args.json).write_text(json.dumps(result, indent=2, ensure_ascii=False))
        print(f"Wrote {args.json}")
        return

    if args.pages_needing_review:
        for p in result["low_text_pages"]:
            print(p)
        return

    if args.page is not None:
        for p in result["pages"]:
            if p["page_number"] == args.page:
                print(p["text"])
                return
        print(f"Page {args.page} not found", file=sys.stderr)
        sys.exit(1)

    if args.numbers:
        for p in result["pages"]:
            nums = find_numbers_on_page(p["text"])
            if nums:
                print(f"Page {p['page_number']}: {nums}")
        return

    print(f"File: {result['path']}")
    print(f"Pages: {result['page_count']}")
    print(f"Threshold for 'needs review': < {args.min_chars} chars")
    print(f"\nPer-page char counts:")
    for p in result["pages"]:
        flag = "  ⚠ needs review" if p["needs_review"] else ""
        print(f"  p.{p['page_number']:>3}: {p['char_count']:>6} chars{flag}")
    if result["low_text_pages"]:
        print(f"\n{len(result['low_text_pages'])} page(s) need review: {result['low_text_pages']}")


if __name__ == "__main__":
    main()
