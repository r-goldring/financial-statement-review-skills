"""Annotate the FY25 clean PDF with tie-out marks.

Reads:
  - The clean PDF
  - All tie-record JSON files (one per lane)

Process:
  - Group records by pdf_page
  - For each page, render → OCR (easyocr) → bbox map
  - For each tied record, find the OCR bbox matching the pdf_value, draw a mark
  - For exceptions, draw a red box + "!" mark
  - At top of each face-statement page, draw "FS Bridge" / "PY" column tags

Mark vocabulary (per references/tieout-conventions.md):
  - 'B'  → tied to bridge (lane 1 ties)
  - 'TB' → tied to TB (lane 2 ties)
  - 'PY' → tied to prior year FY24 (lane 3 ties)
  - '/'  → internal cross-ref (lane 4 ties)
  - red box + '!' → exception (any status != ties / ties-with-rounding / ties-with-sign-inversion)

CLI:
  python annotate_tieout_pdf.py <pdf_path> <out_pdf> <tie_records.json> [<tie_records.json> ...]
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


# Lane → mark glyph and color
LANE_MARK = {
    "pdf_to_bridge":    ("B",  (0.85, 0.13, 0.15)),   # red
    "bridge_to_tb":     ("TB", (0.40, 0.15, 0.60)),   # purple (different from bridge)
    "pdf_prior_year":   ("PY", (0.85, 0.13, 0.15)),   # red
    "pdf_internal":     ("/",  (0.10, 0.55, 0.18)),   # green (internal cross-ref)
    "soe":              ("B",  (0.85, 0.13, 0.15)),   # SOE values tie to bridge
    "footing":          ("F",  (0.85, 0.13, 0.15)),   # foot/cross-foot mark
}

# Status-based override for specific marks
STATUS_MARK_OVERRIDE = {
    "ties-F":  ("F",  (0.85, 0.13, 0.15)),
    "ties-xF": ("xF", (0.85, 0.13, 0.15)),
    "ties-caption-changed": ("PY*", (0.90, 0.55, 0.10)),  # orange PY* = ties via number, caption changed
}

EXCEPTION_COLOR = (0.10, 0.30, 0.85)  # vivid blue — easy to spot against the red tickmark fill
TIE_TEXT_COLOR = (0.85, 0.13, 0.15)  # softer red (FY24-style red)
CARRY_FORWARD_RED = (0.85, 0.13, 0.15)
CARRY_FORWARD_BLUE = (0.10, 0.30, 0.85)  # used for "needs review" carry-forwards (when retained)


def value_to_search_strings(v):
    """Return a list of plausible OCR string representations of a numeric value.

    OCR returns the visible text. PDF formats numbers as:
      $X,XXX.XX   /  NN,NNN   /  ($X,XXX.XX)   /  $($X,XXX.XX)
    We generate candidate strings to match against OCR text.
    """
    if v is None:
        return []
    out = set()
    if v == int(v):
        n = int(v)
    else:
        n = v
    abs_n = abs(n)
    # Several styles
    if isinstance(n, int):
        out.add(f"{abs_n:,}")
        out.add(f"${abs_n:,}")
        out.add(f"$ {abs_n:,}")
        if n < 0:
            out.add(f"({abs_n:,})")
            out.add(f"$({abs_n:,})")
            out.add(f"$ ({abs_n:,})")
    else:
        out.add(f"{abs_n:,.0f}")
        out.add(f"{abs_n:,.1f}")
        out.add(f"${abs_n:,.0f}")
    return list(out)


def ocr_page(pix_or_path, reader):
    """Run easyocr on a rendered page. Returns list of (bbox_poly, text, conf)."""
    results = reader.readtext(pix_or_path, paragraph=False)
    return results


def detect_year_column_x(ocr_results):
    """Find the X-pixel center positions of '2025' and '2024' year-column headers.

    Returns dict like {'2025': x_center_px, '2024': x_center_px} or {} if not found.
    """
    year_positions = {}
    for bbox, text, conf in ocr_results:
        t = text.strip()
        if t in ("2025", "2024", "2023") and conf > 0.5:
            xs = [p[0] for p in bbox]
            ys = [p[1] for p in bbox]
            # Only accept year headers in the TOP portion of the page (typically y < 400 at 200 DPI)
            y_avg = sum(ys) / len(ys)
            if y_avg > 600:  # too far down — probably year-as-row-label, not a column header
                continue
            # Prefer the first/topmost occurrence if duplicates
            if t not in year_positions:
                year_positions[t] = sum(xs) / len(xs)
    return year_positions


def find_ocr_match(ocr_results, search_strings, label_hint=None, year_col_x=None, year=None):
    """Return the best OCR token whose text matches any search string.

    Filtering strategy:
      1. Find all OCR tokens matching one of the search_strings
      2. If year_col_x + year are given, filter to matches whose X is near the year column's X
      3. Among remaining, prefer matches in the same ROW as the label_hint
    """
    matches = []
    for bbox, text, conf in ocr_results:
        t = text.strip()
        for s in search_strings:
            if t == s or t == s + " " or t.replace(" ", "") == s.replace(" ", ""):
                matches.append((bbox, text, conf))
                break

    if not matches:
        return None

    # Year-column X constraint
    if year and year_col_x and year in year_col_x:
        target_x = year_col_x[year]
        # Accept matches within +/- 100 px of the year column center (numeric column is usually narrow)
        filtered = []
        for m in matches:
            bbox = m[0]
            xs = [p[0] for p in bbox]
            x_center = sum(xs) / len(xs)
            if abs(x_center - target_x) <= 100:
                filtered.append(m)
        if filtered:
            matches = filtered

    if len(matches) == 1 or not label_hint:
        return matches[0]

    # Multiple matches — find the one closest in Y to a row containing label_hint
    label_hint_lower = label_hint.lower()
    label_rows = []
    for bbox, text, conf in ocr_results:
        if label_hint_lower in text.lower():
            ys = [p[1] for p in bbox]
            cy = (min(ys) + max(ys)) / 2
            label_rows.append(cy)
    if not label_rows:
        return matches[0]

    def y_center(bbox):
        ys = [p[1] for p in bbox]
        return (min(ys) + max(ys)) / 2

    def dist(m):
        cy = y_center(m[0])
        return min(abs(cy - ly) for ly in label_rows)

    matches.sort(key=dist)
    return matches[0]


def annotate_page(page, pix_dpi, ocr_results, records_for_page):
    """Draw annotations for one PDF page.

    page: fitz Page object
    pix_dpi: DPI at which OCR was performed (for coord conversion)
    ocr_results: list of (bbox_poly, text, conf)
    records_for_page: list of (idx, tie-record-dict) tuples

    Returns: (marks_drawn, set of record indices that were drawn)
    """
    import fitz
    scale = 72.0 / pix_dpi  # px -> pt

    # For rotated pages (rotation != 0), OCR returns coords in the RENDERED view,
    # but PyMuPDF's insert_text expects coords in the UNROTATED page system.
    # Use derotation_matrix to transform.
    rotation = page.rotation
    derot = page.derotation_matrix if rotation else None
    text_rotation = rotation

    # Pre-detect year-column X positions to constrain matches by column
    year_col_x = detect_year_column_x(ocr_results)

    placements = []
    for idx, rec in records_for_page:
        v = rec.get("pdf_value")
        if v is None:
            continue
        search_strs = value_to_search_strings(v)
        if not search_strs:
            continue
        label_hint = rec.get("pdf_label", "")
        label_hint_short = " ".join(label_hint.split()[:4]) if label_hint else ""
        # Year hint: from record's pdf_year. For records with no year (e.g., footing checks
        # for both columns simultaneously), don't constrain by column.
        year_hint = rec.get("pdf_year")
        match = find_ocr_match(
            ocr_results, search_strs,
            label_hint=label_hint_short,
            year_col_x=year_col_x if year_col_x else None,
            year=year_hint,
        )
        if not match:
            continue
        bbox, text, conf = match
        xs = [p[0] for p in bbox]
        ys = [p[1] for p in bbox]
        right_x_px = max(xs)
        top_y_px = min(ys)
        bot_y_px = max(ys)
        # Convert pixel → pt (in the rendered/rotated view)
        place_x_view = (right_x_px + 6) * scale
        place_y_view = (top_y_px + (bot_y_px - top_y_px) * 0.75) * scale
        # Transform from rendered view coords to unrotated page coords if needed
        if derot:
            pt_view = fitz.Point(place_x_view, place_y_view)
            pt_unrot = pt_view * derot
            place_x_pt, place_y_pt = pt_unrot.x, pt_unrot.y
        else:
            place_x_pt, place_y_pt = place_x_view, place_y_view
        placements.append((idx, rec, place_x_pt, place_y_pt, bbox, text_rotation))

    # Draw marks
    marks_drawn = 0
    drawn_ids = set()
    placements_by_pos = {}
    for idx, rec, x, y, ocr_bbox, txt_rot in placements:
        key = (round(x / 4), round(y / 4))
        placements_by_pos.setdefault(key, []).append((idx, rec, x, y, ocr_bbox, txt_rot))

    for bucket in placements_by_pos.values():
        # Offset direction depends on rotation: for landscape (90°), stack marks vertically
        offset = 0
        for idx, rec, x, y, ocr_bbox, txt_rot in bucket:
            lane = rec["lane"]
            status = rec["status"]
            tie_statuses = ("ties", "ties-with-rounding", "ties-with-sign-inversion",
                            "ties-F", "ties-xF", "ties-caption-changed")
            is_exception = status not in tie_statuses

            # Special PY-untied case: prior_year lane with restatement or missing-on-fy24-pdf
            # should be drawn as a BLUE "PY" mark (not a red box + !), so the user can
            # spot untied PY ticks at a glance during PDF review.
            py_untied = (lane == "pdf_prior_year"
                        and status in ("restatement", "missing-on-fy24-pdf"))

            # Internal cross-ref records carry a `mark` field (e.g., "/SoCF", "/FN 3", "/PL").
            # Draw the specific mark text; blue if exception (untied), red if ties.
            internal_untied = (lane == "pdf_internal" and is_exception)

            # For rotated pages, compute the offset in rendered view then transform.
            if txt_rot == 90:
                place_x = x
                place_y = y + offset
            else:
                place_x = x + offset
                place_y = y

            if py_untied:
                # Draw PY in blue — flag as needs review but don't red-box (no real
                # math exception, just a label/caption mismatch most of the time)
                page.insert_text(
                    fitz.Point(place_x, place_y), "PY",
                    fontname="helv", fontsize=7,
                    color=CARRY_FORWARD_BLUE,
                    rotate=txt_rot,
                )
                marks_drawn += 1
                drawn_ids.add(idx)
                offset += 8
                continue

            # Internal cross-ref: draw the specific mark text (e.g., "/SoCF", "/FN 3", "/PL")
            if lane == "pdf_internal":
                mark_text = rec.get("mark", "/")
                mark_color = CARRY_FORWARD_BLUE if internal_untied else CARRY_FORWARD_RED
                page.insert_text(
                    fitz.Point(place_x, place_y), mark_text,
                    fontname="helv", fontsize=6,
                    color=mark_color,
                    rotate=txt_rot,
                )
                marks_drawn += 1
                drawn_ids.add(idx)
                # Internal marks are longer text — larger offset to avoid overlap
                offset += 14
                continue

            if is_exception:
                # Highlight the offending value with a red box around its bbox so
                # real exceptions are visually prominent (vs the small red ! alone).
                if ocr_bbox:
                    xs = [p[0] for p in ocr_bbox]
                    ys = [p[1] for p in ocr_bbox]
                    bx0, by0, bx1, by1 = min(xs), min(ys), max(xs), max(ys)
                    pad = 1.5
                    box_rect = fitz.Rect(bx0 - pad, by0 - pad, bx1 + pad, by1 + pad)
                    if page.rotation != 0:
                        derot = page.derotation_matrix
                        tl = fitz.Point(box_rect.x0, box_rect.y0) * derot
                        br = fitz.Point(box_rect.x1, box_rect.y1) * derot
                        box_rect = fitz.Rect(min(tl.x, br.x), min(tl.y, br.y),
                                              max(tl.x, br.x), max(tl.y, br.y))
                    try:
                        page.draw_rect(box_rect, color=EXCEPTION_COLOR, width=1.2)
                    except Exception:
                        pass
                page.insert_text(
                    fitz.Point(place_x, place_y), "!",
                    fontname="helv", fontsize=9, color=EXCEPTION_COLOR,
                    rotate=txt_rot,
                )
            else:
                if status in STATUS_MARK_OVERRIDE:
                    glyph, color = STATUS_MARK_OVERRIDE[status]
                else:
                    glyph, color = LANE_MARK.get(lane, ("?", TIE_TEXT_COLOR))
                page.insert_text(
                    fitz.Point(place_x, place_y), glyph,
                    fontname="helv", fontsize=7, color=color,
                    rotate=txt_rot,
                )
            marks_drawn += 1
            drawn_ids.add(idx)
            offset += 8
    return marks_drawn, drawn_ids


def build_footing_status_per_page(all_records):
    """Build per-page footing status: page -> ("ties" | "exception" | "unknown").

    Used to override carry-forward F-mark colors based on actual math verification:
      - If page has Lane 6/Lane 5 footing records AND all of them tie → RED
      - If any of them fail → BLUE
      - If no footing records on that page → keep original color
    """
    page_records = {}
    for r in all_records:
        if r.get("lane") not in ("footing", "soe"):
            continue
        # SOE records use "ties" status too, but we care specifically about F / xF
        status = r.get("status", "")
        if status not in ("ties-F", "ties-xF", "exception"):
            continue
        # Only count exceptions that look like footing checks (pdf_label ends with "(F)" or "(xF)")
        if status == "exception":
            label = r.get("pdf_label", "")
            if not (label.endswith("(F)") or label.endswith("(xF)") or "foot" in label.lower() or "cross-foot" in label.lower()):
                continue
        page = r.get("pdf_page")
        if not page:
            continue
        page_records.setdefault(page, []).append(status)

    page_status = {}
    for page, statuses in page_records.items():
        if any(s == "exception" for s in statuses):
            page_status[page] = "exception"
        else:
            page_status[page] = "ties"
    return page_status


def draw_carry_forward(doc, all_records, drawn_record_ids, drop_unverified=False):
    """Draw carry-forward records (text marks and drawings ported from a prior tieout PDF).

    Each carry-forward text record has pre-computed (x_pt, y_pt) and a color flag.
    Each carry-forward drawing record has prior-PDF items (lines) to re-render.

    Override rule: F / xF / V marks on a page where Lane 6 confirmed footing
    are drawn RED regardless of anchor verification.

    If `drop_unverified` is True, carry-forward records whose anchor didn't verify
    (color == 'blue') are skipped entirely — useful when the source tieout PDF has
    drifted enough from the target that unverified marks are mostly noise.
    """
    import fitz
    marks_drawn = 0

    # Build per-page footing status from Lane 5/6 records
    page_footing_status = build_footing_status_per_page(all_records)

    for idx, rec in enumerate(all_records):
        if rec.get("lane") != "carry_forward":
            continue
        # When the source tieout has drifted from the target, the unverified
        # 'needs review' (blue) marks land at approximate positions and clutter the
        # output. Skipping them keeps the annotated PDF aligned and concise.
        if drop_unverified and rec.get("color") == "blue":
            continue
        page_num = rec.get("pdf_page")
        if not page_num or page_num < 1 or page_num > len(doc):
            continue
        page = doc[page_num - 1]

        # Pick color based on status
        color_str = rec.get("color", "red")
        # Footing override: F / xF / V marks reflect math, not text.
        # If Lane 5/6 confirmed math foots on this page → RED.
        # If Lane 5/6 found a footing failure → BLUE.
        # If no footing records for the page → keep original color.
        mark_text = rec.get("mark_text", "").strip()
        if mark_text in ("F", "xF", "V"):
            status = page_footing_status.get(page_num)
            if status == "ties":
                color_str = "red"
            elif status == "exception":
                color_str = "blue"
        color = CARRY_FORWARD_BLUE if color_str == "blue" else CARRY_FORWARD_RED

        kind = rec.get("kind", "text")
        rotation = rec.get("rotation", 0)

        if kind == "text":
            text = rec.get("mark_text", "")
            if not text:
                continue
            x_pt = rec.get("x_pt")
            y_pt = rec.get("y_pt")
            if x_pt is None or y_pt is None:
                continue
            # If the page is rotated, transform coords via derotation matrix
            if page.rotation != 0:
                derot = page.derotation_matrix
                pt = fitz.Point(x_pt, y_pt) * derot
                x_pt, y_pt = pt.x, pt.y
            try:
                page.insert_text(
                    fitz.Point(x_pt, y_pt), text,
                    fontname="helv",
                    fontsize=8,
                    color=color,
                    rotate=page.rotation,
                )
                marks_drawn += 1
                drawn_record_ids.add(idx)
            except Exception as e:
                # Out-of-bounds or invalid coord — skip
                pass

        elif kind == "drawing":
            items = rec.get("items", [])
            fy24_w = rec.get("fy24_page_w", 612)
            fy24_h = rec.get("fy24_page_h", 792)
            fy25_w = page.rect.width
            fy25_h = page.rect.height
            sx = fy25_w / fy24_w if fy24_w else 1
            sy = fy25_h / fy24_h if fy24_h else 1
            try:
                shape = page.new_shape()
                for item in items:
                    t = item.get("type")
                    pts = item.get("points", [])
                    if t == "l" and len(pts) >= 2:
                        p0 = fitz.Point(pts[0][0] * sx, pts[0][1] * sy)
                        p1 = fitz.Point(pts[1][0] * sx, pts[1][1] * sy)
                        if page.rotation != 0:
                            derot = page.derotation_matrix
                            p0 = p0 * derot
                            p1 = p1 * derot
                        shape.draw_line(p0, p1)
                shape.finish(color=color, width=0.8)
                shape.commit()
                marks_drawn += 1
                drawn_record_ids.add(idx)
            except Exception:
                pass

    return marks_drawn


def draw_column_tags(page, page_kind):
    """Draw 'FS Bridge' / 'PY' column header tags on face statement pages."""
    import fitz
    # Skip SOE — its layout (per-class columns) is too different for simple tags.
    if page_kind not in ("BS", "IS", "SCF"):
        return
    # Place tags at top of page above where values appear
    # Conventionally: current-year column is around 60-70% page width, prior year ~85-90%
    pw = page.rect.width
    ph = page.rect.height
    cy_x = pw * 0.65
    py_x = pw * 0.85
    tag_y = ph * 0.18
    page.insert_text(fitz.Point(cy_x, tag_y), "FS Bridge",
                     fontname="helv", fontsize=10, color=TIE_TEXT_COLOR)
    page.insert_text(fitz.Point(py_x, tag_y), "PY",
                     fontname="helv", fontsize=10, color=TIE_TEXT_COLOR)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf_in", help="Path to clean FY25 PDF")
    ap.add_argument("pdf_out", help="Path to write annotated PDF")
    ap.add_argument("tie_records", nargs="+", help="One or more tie-records JSON files")
    ap.add_argument("--dpi", type=int, default=200, help="Render DPI for OCR (default 200)")
    ap.add_argument("--ocr-cache", default=None, help="Cache OCR results to this JSON path")
    ap.add_argument("--pages", default=None, help="Restrict to these pages (comma-separated)")
    ap.add_argument("--drop-unverified-carry-forward", action="store_true",
                    help="Skip carry-forward records whose anchor text didn't verify in "
                         "the target PDF (color=='blue'). Useful when the carry-forward "
                         "source is a draft that has drifted from the target — keeps the "
                         "annotated PDF clean of misaligned 'needs review' marks.")
    args = ap.parse_args()

    import fitz

    # Load all tie records
    all_records = []
    for p in args.tie_records:
        recs = json.loads(Path(p).read_text(encoding="utf-8"))
        all_records.extend(recs)
        print(f"  loaded {len(recs)} records from {Path(p).name}")
    print(f"  total: {len(all_records)} records\n")

    # Group by page. For footnote pages (>= 9), also add the record to adjacent pages
    # because FN tables can span 1-2 pages beyond the section heading. The annotator's
    # value-matching will only place the mark on the page where the value actually appears.
    by_page = defaultdict(list)
    # Track which (record_id, page) pairs we've already added so we don't double-draw
    seen = set()
    for idx, r in enumerate(all_records):
        p = r.get("pdf_page")
        if not p:
            continue
        candidate_pages = [p]
        if p >= 9:
            candidate_pages += [p + 1, p + 2]
        for cp in candidate_pages:
            key = (idx, cp)
            if key not in seen:
                by_page[cp].append((idx, r))
                seen.add(key)

    print(f"Records per page (top 10):")
    sorted_pages = sorted(by_page.items())
    for pg, recs in sorted_pages[:10]:
        tie_set = ("ties", "ties-with-rounding", "ties-with-sign-inversion",
                   "ties-F", "ties-xF", "ties-caption-changed")
        ties = sum(1 for _, r in recs if r["status"] in tie_set)
        exc = sum(1 for _, r in recs if r["status"] not in tie_set + ("missing-on-bridge", "missing-year-on-bridge", "missing-on-fy24-pdf", "no-tb-rollup"))
        miss = sum(1 for _, r in recs if "missing" in r["status"] or "no-tb" in r["status"])
        print(f"  p.{pg:>2}: {len(recs)} records (ties={ties}, exceptions={exc}, missing-source={miss})")
    print()

    # Pages to process: include the assigned pages PLUS their adjacent pages (footnote
    # tables often span 1-2 pages beyond the section heading).
    if args.pages:
        target_pages = set(int(x) for x in args.pages.split(","))
    else:
        target_pages = set(by_page.keys())
        # Add adjacent pages for footnote sections (any page >= 9)
        for pg in list(target_pages):
            if pg >= 9:
                target_pages.add(pg + 1)
                target_pages.add(pg + 2)

    # Set up OCR
    print("Loading easyocr...")
    import easyocr
    reader = easyocr.Reader(["en"], gpu=False, verbose=False)
    print("  ready\n")

    # Open the PDF
    doc = fitz.open(args.pdf_in)

    # OCR cache
    ocr_cache_path = Path(args.ocr_cache) if args.ocr_cache else None
    ocr_cache = {}
    if ocr_cache_path and ocr_cache_path.exists():
        ocr_cache = json.loads(ocr_cache_path.read_text(encoding="utf-8"))
        print(f"  loaded OCR cache with {len(ocr_cache)} pages")

    total_marks = 0
    drawn_record_ids = set()  # avoid drawing the same record twice when it's on multiple candidate pages

    # First pass: draw all CARRY_FORWARD records (text + drawings) — these don't need OCR
    # because they have pre-computed coords.
    cf_drawn = draw_carry_forward(doc, all_records, drawn_record_ids,
                                   drop_unverified=args.drop_unverified_carry_forward)
    total_marks += cf_drawn
    print(f"  carry-forward marks drawn: {cf_drawn}")

    # Second pass: tie-out marks (need OCR for value-bbox)
    for page_num in sorted(target_pages):
        if page_num not in by_page:
            continue
        records_for_page = by_page[page_num]
        if not records_for_page:
            continue
        # Filter out records already drawn elsewhere
        records_for_page = [(idx, r) for idx, r in records_for_page if idx not in drawn_record_ids]
        if not records_for_page:
            continue
        # Defensive: records sourced from a different-paginated PDF (e.g. a prior
        # tieout with N+1 pages) can reference page indexes past the target doc.
        if page_num < 1 or page_num > len(doc):
            continue
        page = doc[page_num - 1]
        cache_key = f"page-{page_num}"
        if cache_key in ocr_cache:
            ocr_results = [(tuple(map(tuple, r["bbox"])), r["text"], r["conf"]) for r in ocr_cache[cache_key]]
            print(f"  p.{page_num:>2}  (cached OCR, {len(ocr_results)} tokens)")
        else:
            print(f"  p.{page_num:>2}  OCR running...", end="", flush=True)
            pix = page.get_pixmap(dpi=args.dpi)
            png_bytes = pix.tobytes("png")
            import io as _io
            ocr_results = reader.readtext(png_bytes, paragraph=False)
            print(f"  {len(ocr_results)} tokens")
            if ocr_cache_path:
                ocr_cache[cache_key] = [
                    {
                        "bbox": [[float(p[0]), float(p[1])] for p in bbox],
                        "text": str(text),
                        "conf": float(conf),
                    }
                    for bbox, text, conf in ocr_results
                ]

        page_kind = None
        if page_num in (5, 6, 7, 8):
            page_kind = {5: "BS", 6: "IS", 7: "SOE", 8: "SCF"}.get(page_num)

        # Draw column tags first (so they're behind the marks if overlap)
        if page_kind:
            draw_column_tags(page, page_kind)

        # Draw the cell-level marks
        marks, drawn_ids = annotate_page(page, args.dpi, ocr_results, records_for_page)
        drawn_record_ids.update(drawn_ids)
        total_marks += marks
        print(f"        marks drawn: {marks}")

    # Save OCR cache
    if ocr_cache_path:
        ocr_cache_path.write_text(json.dumps(ocr_cache, indent=2), encoding="utf-8")

    # Save annotated PDF
    Path(args.pdf_out).parent.mkdir(parents=True, exist_ok=True)
    doc.save(args.pdf_out, garbage=4, deflate=True)
    doc.close()
    print(f"\nTotal marks drawn: {total_marks}")
    print(f"Wrote {args.pdf_out}")


if __name__ == "__main__":
    main()
