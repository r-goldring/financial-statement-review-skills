"""Extract red annotations from the FY24 final tieout PDF and carry forward
to the FY25 PDF.

Approach:
  1. Extract every red-colored text span and red drawing (lines/arrows) from FY24.
  2. For each red text span, find the nearest BLACK text "anchor" (within ~80pt
     left/above) — this gives us the context (e.g., a value like 'NN,NNN' or a
     label like 'Cash and cash equivalents').
  3. On FY25, OCR the corresponding page and search for the same anchor text.
  4. If anchor found:
       - If the value/text near the FY25 anchor matches FY24 → carry forward in RED
       - If the value/text changed → carry forward in BLUE (review)
  5. Drawings (lines/arrows): also carry forward, anchored to a nearby text span.

Output: a JSON list of tie-records that the existing annotator can consume.
Each record has lane="carry_forward" and either color_override="red" or "blue".

CLI:
  python carry_forward_fy24_marks.py <fy24_tieout.pdf> <fy25_clean.pdf>
                                     <out_records.json>
                                     [--ocr-cache <path>]
"""

import argparse
import json
import sys
from pathlib import Path

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

SCRIPT_DIR = Path(__file__).parent
sys.path.insert(0, str(SCRIPT_DIR))


def is_red(color_int):
    """Detect FY24's red annotation color (around #db261c)."""
    r = (color_int >> 16) & 0xff
    g = (color_int >> 8) & 0xff
    b = color_int & 0xff
    return r > 150 and g < 100 and b < 100


def is_red_rgb(rgb):
    """For drawings — stroke/fill colors are 0-1 floats."""
    if not rgb or len(rgb) < 3:
        return False
    r, g, b = rgb[0], rgb[1], rgb[2]
    return r > 0.5 and g < 0.4 and b < 0.4


def extract_fy24_red_marks(pdf_path):
    """Return list of {page, text, bbox, kind} for each red text span on each page."""
    import fitz
    doc = fitz.open(pdf_path)
    marks = []
    for page_num in range(1, len(doc) + 1):
        page = doc[page_num - 1]
        blocks = page.get_text("dict")["blocks"]
        page_red = []
        for block in blocks:
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    color = span.get("color", 0)
                    if not is_red(color):
                        continue
                    text = span.get("text", "").strip()
                    if not text:
                        continue
                    bbox = span.get("bbox")  # (x0, y0, x1, y1)
                    if not bbox:
                        continue
                    page_red.append({
                        "text": text,
                        "bbox": bbox,
                        "size": span.get("size"),
                    })
        # Capture black-text spans for context lookup
        page_black = []
        for block in blocks:
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    color = span.get("color", 0)
                    if is_red(color):
                        continue
                    text = span.get("text", "").strip()
                    if not text:
                        continue
                    bbox = span.get("bbox")
                    if not bbox:
                        continue
                    page_black.append({"text": text, "bbox": bbox})

        # For each red mark, find the nearest black text anchor
        for rm in page_red:
            rx0, ry0, rx1, ry1 = rm["bbox"]
            ry_center = (ry0 + ry1) / 2
            rx_center = (rx0 + rx1) / 2
            best_anchor = None
            best_score = float("inf")
            for bs in page_black:
                bx0, by0, bx1, by1 = bs["bbox"]
                by_center = (by0 + by1) / 2
                bx_center = (bx0 + bx1) / 2
                # Anchor preference: on the same row, slightly to the left
                if abs(by_center - ry_center) > 8:
                    continue  # not on same row
                # Score = horizontal distance; prefer text just to the LEFT of mark
                if bx1 > rx0 - 2 and bx0 < rx0 + 2:
                    continue  # overlaps; skip
                dx = rx0 - bx1  # positive = anchor is to the left
                if dx < 0:
                    continue  # anchor is to the right of mark — not preferred
                if dx < best_score:
                    best_score = dx
                    best_anchor = bs

            rm["anchor"] = best_anchor
            rm["anchor_score"] = best_score if best_anchor else None

        # Also extract red drawings (lines/arrows)
        page_drawings = []
        for d in page.get_drawings():
            stroke = d.get("color")
            fill = d.get("fill")
            stroke_red = is_red_rgb(stroke)
            fill_red = is_red_rgb(fill)
            if not (stroke_red or fill_red):
                continue
            page_drawings.append({
                "rect": list(d.get("rect", (0, 0, 0, 0))),
                "items": [
                    {
                        "type": item[0],
                        "points": [list(item[i]) for i in range(1, len(item))
                                   if hasattr(item[i], 'x')]
                    }
                    for item in d.get("items", [])
                ],
                "stroke": list(stroke) if stroke else None,
            })

        marks.append({
            "page": page_num,
            "page_width": page.rect.width,
            "page_height": page.rect.height,
            "rotation": page.rotation,
            "red_text": page_red,
            "red_drawings": page_drawings,
        })

    doc.close()
    return marks


def find_anchor_in_ocr(anchor_text, ocr_tokens):
    """Find the OCR token matching anchor_text (or its prefix).

    Returns list of (bbox_pixels, text, conf) of matches.
    """
    if not anchor_text:
        return []
    needle = anchor_text.strip().lower()
    if not needle:
        return []
    matches = []
    # Try exact, then substring
    for t in ocr_tokens:
        text = t["text"].strip().lower()
        if not text:
            continue
        if text == needle or text.startswith(needle) or needle in text:
            matches.append(t)
    # Also try multi-word match: split anchor into words, find OCR tokens that contain each word
    if not matches and " " in needle:
        words = needle.split()
        for t in ocr_tokens:
            text = t["text"].strip().lower()
            if all(w in text for w in words):
                matches.append(t)
    return matches


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("fy24_pdf")
    ap.add_argument("fy25_pdf")
    ap.add_argument("out_json")
    ap.add_argument("--ocr-cache", required=True,
                    help="OCR cache for FY25 PDF (will be used to match anchors)")
    ap.add_argument("--page-shift", type=int, default=0,
                    help="If FY25 has pages offset from FY24 (e.g., new cover page)")
    args = ap.parse_args()

    cache = {}
    if Path(args.ocr_cache).exists():
        cache = json.loads(Path(args.ocr_cache).read_text(encoding="utf-8"))
    print(f"Loaded FY25 OCR cache with {len(cache)} pages")

    print(f"Extracting FY24 red annotations from {args.fy24_pdf}...")
    fy24_marks = extract_fy24_red_marks(args.fy24_pdf)
    total_red_text = sum(len(p["red_text"]) for p in fy24_marks)
    total_red_drawings = sum(len(p["red_drawings"]) for p in fy24_marks)
    print(f"  FY24: {total_red_text} red text spans, {total_red_drawings} red drawings across {len(fy24_marks)} pages")

    # Open FY25 to get page dims (for anchor coord scaling)
    import fitz
    fy25_doc = fitz.open(args.fy25_pdf)

    output_records = []
    for fy24_page in fy24_marks:
        fy24_page_num = fy24_page["page"]
        fy25_page_num = fy24_page_num + args.page_shift
        if fy25_page_num < 1 or fy25_page_num > len(fy25_doc):
            continue
        cache_key = f"page-{fy25_page_num}"
        ocr_tokens = cache.get(cache_key)
        if not ocr_tokens:
            # Skip pages where we don't have OCR yet
            continue

        # FY24 page rect (for relative coord)
        fy24_w = fy24_page["page_width"]
        fy24_h = fy24_page["page_height"]
        fy25_page = fy25_doc[fy25_page_num - 1]
        fy25_w = fy25_page.rect.width
        fy25_h = fy25_page.rect.height

        # Identify the OCR scale: convert FY25 OCR pixel coords to PDF pt coords.
        # The OCR cache was at DPI=200; so scale = 72/200.
        ocr_scale = 72.0 / 200.0

        for rm in fy24_page["red_text"]:
            text = rm["text"]
            anchor = rm.get("anchor")
            anchor_text = anchor["text"] if anchor else None
            anchor_bbox_fy24 = anchor["bbox"] if anchor else None

            if not anchor_text or not anchor_bbox_fy24:
                # No anchor — fall back to absolute coords (with page-size scaling)
                # Compute relative position
                rx0, ry0, rx1, ry1 = rm["bbox"]
                rel_x = (rx0 + rx1) / 2 / fy24_w
                rel_y = (ry0 + ry1) / 2 / fy24_h
                fy25_x = rel_x * fy25_w
                fy25_y = rel_y * fy25_h
                output_records.append({
                    "lane": "carry_forward",
                    "kind": "text",
                    "pdf_page": fy25_page_num,
                    "pdf_label": f"carry-fwd: {text}",
                    "mark_text": text,
                    "x_pt": fy25_x,
                    "y_pt": fy25_y,
                    "color": "blue",  # no anchor confirmed — flag as needs review
                    "status": "carry-fwd-no-anchor",
                    "rotation": fy24_page.get("rotation", 0),
                })
                continue

            # Find the anchor in FY25 OCR
            anchor_matches = find_anchor_in_ocr(anchor_text, ocr_tokens)
            if not anchor_matches:
                # Anchor not found on FY25 → BLUE (something changed)
                rx0, ry0, rx1, ry1 = rm["bbox"]
                rel_x = (rx0 + rx1) / 2 / fy24_w
                rel_y = (ry0 + ry1) / 2 / fy24_h
                fy25_x = rel_x * fy25_w
                fy25_y = rel_y * fy25_h
                output_records.append({
                    "lane": "carry_forward",
                    "kind": "text",
                    "pdf_page": fy25_page_num,
                    "pdf_label": f"carry-fwd: {text}",
                    "mark_text": text,
                    "x_pt": fy25_x,
                    "y_pt": fy25_y,
                    "color": "blue",
                    "status": "carry-fwd-anchor-missing",
                    "anchor_text": anchor_text,
                    "rotation": fy24_page.get("rotation", 0),
                })
                continue

            # Anchor found! Pick the first match (TODO: better disambiguation)
            anchor_match = anchor_matches[0]
            am_bbox = anchor_match["bbox"]
            # Compute FY24 offset of mark from anchor
            arx0_24, ary0_24, arx1_24, ary1_24 = anchor_bbox_fy24
            anchor_cx_24 = (arx0_24 + arx1_24) / 2
            anchor_cy_24 = (ary0_24 + ary1_24) / 2
            rx0, ry0, rx1, ry1 = rm["bbox"]
            offset_x = (rx0 + rx1) / 2 - anchor_cx_24
            offset_y = (ry0 + ry1) / 2 - anchor_cy_24

            # FY25 anchor pt position (convert OCR pixel to pt)
            am_xs = [p[0] for p in am_bbox]
            am_ys = [p[1] for p in am_bbox]
            anchor_x_px = sum(am_xs) / len(am_xs)
            anchor_y_px = sum(am_ys) / len(am_ys)
            anchor_x_pt = anchor_x_px * ocr_scale
            anchor_y_pt = anchor_y_px * ocr_scale

            # Apply FY24 offset to anchor on FY25
            fy25_x = anchor_x_pt + offset_x
            fy25_y = anchor_y_pt + offset_y

            # Determine color: RED if anchor text matches FY24 exactly (no change to surrounding)
            # BLUE if anchor differs.
            anchor_match_text = anchor_match["text"].strip().lower()
            fy24_anchor_text = anchor_text.strip().lower()
            if anchor_match_text == fy24_anchor_text:
                color = "red"
                status = "carry-fwd-verified"
            else:
                color = "blue"
                status = "carry-fwd-text-differs"

            output_records.append({
                "lane": "carry_forward",
                "kind": "text",
                "pdf_page": fy25_page_num,
                "pdf_label": f"carry-fwd: {text} (anchor: {anchor_text})",
                "mark_text": text,
                "x_pt": fy25_x,
                "y_pt": fy25_y,
                "color": color,
                "status": status,
                "anchor_text": anchor_text,
                "anchor_fy25_text": anchor_match["text"],
                "rotation": fy24_page.get("rotation", 0),
            })

        # Drawings: extract as carry-fwd records with relative coords
        for d in fy24_page["red_drawings"]:
            # Use the rect midpoint as anchor; carry forward at same proportional position
            rect = d["rect"]
            if not rect or len(rect) != 4:
                continue
            x0, y0, x1, y1 = rect
            # Relative position on page
            rel_x0 = x0 / fy24_w
            rel_y0 = y0 / fy24_h
            rel_x1 = x1 / fy24_w
            rel_y1 = y1 / fy24_h
            output_records.append({
                "lane": "carry_forward",
                "kind": "drawing",
                "pdf_page": fy25_page_num,
                "items": d["items"],
                "stroke": d["stroke"],
                "fy24_rect": rect,
                "fy24_page_w": fy24_w,
                "fy24_page_h": fy24_h,
                "color": "red",  # drawings always carry forward as red (no easy way to verify)
                "status": "carry-fwd-drawing",
                "rotation": fy24_page.get("rotation", 0),
            })

    fy25_doc.close()
    # Summary
    by_color = {}
    by_status = {}
    for r in output_records:
        by_color[r["color"]] = by_color.get(r["color"], 0) + 1
        by_status[r["status"]] = by_status.get(r["status"], 0) + 1
    print(f"\nCarry-forward records: {len(output_records)}")
    print(f"  By color: {by_color}")
    print(f"  By status: {by_status}")

    Path(args.out_json).write_text(
        json.dumps(output_records, default=str, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    print(f"Wrote {args.out_json}")


if __name__ == "__main__":
    main()
