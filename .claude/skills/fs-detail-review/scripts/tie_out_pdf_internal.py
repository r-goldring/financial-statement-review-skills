"""Lane 4: Internal PDF cross-references.

Verifies that values appearing on multiple pages/tables within the FY25 PDF agree.
Mirrors FY24 tickmark convention:
  /SoCF     -> ties to Statement of Cash Flows
  /SoSE     -> ties to Statement of Changes in Members' Equity
  /PL       -> ties to Income Statement (Profit & Loss)
  /BS       -> ties to Balance Sheet
  /FN N     -> ties to Footnote N

Each record drawn as a tickmark on the SOURCE side, with text indicating where it ties to.

CLI:
  python tie_out_pdf_internal.py <inputs.json> <out.json>
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

from tie_out_common import parse_value, compare, make_record


SECTION_TO_PDF_PAGE = {"BS": 5, "IS": 6, "SOE": 7, "SCF": 8}

# These come from .work/fn-page-map.json — hardcoded for the FY25 vYYYY.M.D PDF
FN_TO_PDF_PAGE = {
    "FN-01": 9, "FN-02": 9, "FN-03": 17, "FN-04": 18, "FN-05": 19,
    "FN-06": 19, "FN-07": 21, "FN-08": 7, "FN-09": 24, "FN-10": 25,
    "FN-11": 25, "FN-12": 27, "FN-13": 27, "FN-14": 28,
}


def find_row(table, label_matcher):
    """Return the first row dict (label, values, row_index) matching the given matcher fn."""
    rows = table.get("rows", [])
    year_cols = {}
    for i, row in enumerate(rows[:5]):
        cells = {}
        for c_idx, cell in enumerate(row):
            if c_idx == 0:
                continue
            if cell and isinstance(cell, str) and re.match(r"^20\d{2}$", cell.strip()):
                cells[c_idx] = cell.strip()
        if len(cells) >= 2:
            year_cols = cells
            break
    for row_idx, row in enumerate(rows):
        if not row:
            continue
        label = row[0].strip() if len(row) > 0 and row[0] else ""
        if not label:
            continue
        if label_matcher(label):
            values = {}
            for c_idx, year in year_cols.items():
                if c_idx < len(row):
                    v = parse_value(row[c_idx])
                    if v is not None:
                        values[year] = v
            return {"label": label, "values": values, "row_idx": row_idx, "raw_row": row}
    return None


def find_fn_table(tables, fn_prefix, title_hint=None):
    """Return FN table(s) matching prefix; if title_hint given, only matching tables."""
    matches = [t for t in tables if t["section"].startswith(fn_prefix)]
    if title_hint:
        matches = [t for t in matches if title_hint.lower() in (t.get("title") or "").lower()]
    return matches


def get_total_col_index(table):
    """Return column index containing 'Total' header in the SOE table — the last numeric column."""
    rows = table.get("rows", [])
    if not rows:
        return None
    # SOE first row has class names, second row has Units/Amount headers
    # The total column is usually the last one with "Amount" or "Total"
    header_row_0 = rows[0] if rows else []
    header_row_1 = rows[1] if len(rows) > 1 else []
    for i in reversed(range(len(header_row_0))):
        cell = (header_row_0[i] or "").lower()
        if "total" in cell:
            return i
    # Fallback: last column
    return len(header_row_1) - 1 if header_row_1 else None


def soe_balance_row(soe_table, target_date_substr):
    """Find SOE row by date label, e.g., 'December 31, 2025'."""
    for r_idx, row in enumerate(soe_table.get("rows", [])):
        if not row:
            continue
        label = (row[0] or "").strip()
        if target_date_substr in label and "balance" in label.lower():
            return {"row_idx": r_idx, "raw_row": row, "label": label}
    return None


def soe_last_numeric(row):
    """Last meaningful numeric value in a row (the Total column)."""
    for cell in reversed(row):
        v = parse_value(cell)
        if v is not None and abs(v) > 100:
            return v
    return None


def soe_amount_for_class(row, class_amount_col):
    """Return the parsed value at the specified column index in an SOE row."""
    if class_amount_col >= len(row):
        return None
    return parse_value(row[class_amount_col])


def soe_row_total(row):
    """Return Total column value (last numeric column)."""
    return soe_last_numeric(row)


def fn03_net_ppe_for_year(fn03_table, year_col):
    """Return net property and equipment value at year_col."""
    for row in fn03_table.get("rows", []):
        label = (row[0] or "").strip().lower()
        if label.startswith("property and equipment, net"):
            return parse_value(row[year_col]) if year_col < len(row) else None
    return None


def make_internal_record(pdf_section, pdf_page, pdf_label, pdf_year, pdf_value,
                        target_section, target_page, target_label, target_value,
                        mark, notes, is_subtotal=True):
    """Construct a Lane 4 record with the right mark in the source_ref."""
    if pdf_value is None or target_value is None:
        return None
    delta, status, tol, cu = compare(pdf_value, target_value, "$K", "$K",
                                     is_subtotal=is_subtotal, kind="internal")
    rec = make_record(
        lane="pdf_internal",
        pdf_section=pdf_section,
        pdf_page=pdf_page,
        pdf_label=pdf_label,
        pdf_year=pdf_year,
        pdf_value=pdf_value,
        source_ref=f"PDF!p{target_page}-{target_section}",
        source_label=target_label,
        source_value=target_value,
        comparison_unit=cu, delta=delta, tolerance=tol,
        status=status, is_subtotal=is_subtotal,
        notes=notes,
    )
    rec["mark"] = mark
    rec["target_section"] = target_section
    rec["target_page"] = target_page
    return rec


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("inputs")
    ap.add_argument("out")
    args = ap.parse_args()

    data = json.loads(Path(args.inputs).read_text(encoding="utf-8"))
    tables = data["fy25_docx"]["tables"]

    bs = next((t for t in tables if t["section"] == "BS"), None)
    is_t = next((t for t in tables if t["section"] == "IS"), None)
    soe = next((t for t in tables if t["section"] == "SOE"), None)
    scf = next((t for t in tables if t["section"] == "SCF"), None)

    records = []

    # ==========================================================================
    # === BS <-> SCF (Cash rollforward) ===
    # ==========================================================================
    if bs and scf:
        bs_cash = find_row(bs, lambda l: l.lower().startswith("cash and cash equivalents") and "supplemental" not in l.lower())
        scf_end_cash = find_row(scf, lambda l: ("end of year" in l.lower() or "end of period" in l.lower()) and "cash" in l.lower())
        scf_begin_cash = find_row(scf, lambda l: ("beginning of year" in l.lower() or "beginning of period" in l.lower()) and "cash" in l.lower())

        if bs_cash and scf_end_cash:
            for year in ("2025", "2024"):
                bs_v = bs_cash["values"].get(year)
                scf_v = scf_end_cash["values"].get(year)
                rec = make_internal_record(
                    "BS", SECTION_TO_PDF_PAGE["BS"], f"Cash and cash equivalents ({year})", year, bs_v,
                    "SCF", SECTION_TO_PDF_PAGE["SCF"], scf_end_cash["label"], scf_v,
                    "/SoCF",
                    f"BS Cash {year} = SCF Ending Cash {year}",
                )
                if rec:
                    records.append(rec)

        # BS PY Cash (ending of 2024) = SCF Beginning of 2025
        if bs_cash and scf_begin_cash:
            bs_2024 = bs_cash["values"].get("2024")
            scf_2025_begin = scf_begin_cash["values"].get("2025")
            rec = make_internal_record(
                "SCF", SECTION_TO_PDF_PAGE["SCF"], "Cash beginning of 2025", "2025", scf_2025_begin,
                "BS", SECTION_TO_PDF_PAGE["BS"], "BS Cash 2024 ending", bs_2024,
                "/BS",
                "SCF 2025 Beginning Cash = BS 2024 Ending Cash",
            )
            if rec:
                records.append(rec)

        # SCF rollforward: Begin + Net change = End
        scf_change = find_row(scf, lambda l: "net increase" in l.lower() or "net decrease" in l.lower() or "net change in cash" in l.lower())
        if scf_change and scf_begin_cash and scf_end_cash:
            for year in ("2025", "2024"):
                ch = scf_change["values"].get(year)
                bg = scf_begin_cash["values"].get(year)
                ed = scf_end_cash["values"].get(year)
                if ch is not None and bg is not None and ed is not None:
                    computed = bg + ch
                    delta, status, tol, cu = compare(ed, computed, "$K", "$K",
                                                    is_subtotal=True, kind="internal")
                    records.append({**make_record(
                        lane="pdf_internal",
                        pdf_section="SCF",
                        pdf_page=SECTION_TO_PDF_PAGE["SCF"],
                        pdf_label=f"SCF rollforward: Begin + Change = End ({year})",
                        pdf_year=year,
                        pdf_value=ed,
                        source_ref=f"PDF!p{SECTION_TO_PDF_PAGE['SCF']}-SCF-rollforward",
                        source_label=f"({bg:,.0f} begin + {ch:,.0f} change)",
                        source_value=computed,
                        comparison_unit=cu, delta=delta, tolerance=tol,
                        status=status, is_subtotal=True,
                        notes="Internal SCF rollforward consistency",
                    ), "mark": "F", "target_section": "SCF", "target_page": SECTION_TO_PDF_PAGE["SCF"]})

    # ==========================================================================
    # === IS <-> SCF (Net loss) ===
    # ==========================================================================
    if is_t and scf:
        is_net_loss = find_row(is_t, lambda l: l.lower().strip() == "net loss")
        scf_net_loss = find_row(scf, lambda l: l.lower().strip() == "net loss")
        if is_net_loss and scf_net_loss:
            for year in ("2025", "2024"):
                is_v = is_net_loss["values"].get(year)
                scf_v = scf_net_loss["values"].get(year)
                if is_v is not None and scf_v is not None:
                    rec = make_internal_record(
                        "SCF", SECTION_TO_PDF_PAGE["SCF"], f"Net loss ({year})", year, abs(scf_v),
                        "IS", SECTION_TO_PDF_PAGE["IS"], f"Net loss (IS) {year}", abs(is_v),
                        "/PL",
                        f"SCF Net Loss {year} = IS Net Loss {year}",
                    )
                    if rec:
                        records.append(rec)

    # ==========================================================================
    # === IS <-> SCF (Depreciation & amortization) ===
    # ==========================================================================
    if is_t and scf:
        is_da = find_row(is_t, lambda l: l.lower().strip() == "depreciation and amortization")
        scf_da = find_row(scf, lambda l: l.lower().strip() == "depreciation and amortization")
        if is_da and scf_da:
            for year in ("2025", "2024"):
                is_v = is_da["values"].get(year)
                scf_v = scf_da["values"].get(year)
                rec = make_internal_record(
                    "SCF", SECTION_TO_PDF_PAGE["SCF"], f"Depreciation and amortization ({year})", year, scf_v,
                    "IS", SECTION_TO_PDF_PAGE["IS"], f"Depreciation and amortization (IS) {year}", is_v,
                    "/PL",
                    f"SCF D&A {year} = IS D&A {year}",
                )
                if rec:
                    records.append(rec)

    # ==========================================================================
    # === BS <-> SOE (Members' Equity components, closing balances) ===
    # ==========================================================================
    if bs and soe:
        # Total members' equity
        bs_total_equity = find_row(bs, lambda l: "total members" in l.lower() and "equity" in l.lower() and "liabilities" not in l.lower())
        soe_closing_2025 = soe_balance_row(soe, "December 31, 2025")
        soe_closing_2024 = soe_balance_row(soe, "December 31, 2024")

        if bs_total_equity and soe_closing_2025:
            bs_eq = bs_total_equity["values"].get("2025")
            soe_total = soe_row_total(soe_closing_2025["raw_row"])
            rec = make_internal_record(
                "BS", SECTION_TO_PDF_PAGE["BS"], "Total members' equity (2025)", "2025", bs_eq,
                "SOE", SECTION_TO_PDF_PAGE["SOE"], "Balance as of December 31, 2025 (Total)", soe_total,
                "/SoSE",
                "BS Total Members' Equity 2025 = SOE closing 2025 Total",
            )
            if rec:
                records.append(rec)

        if bs_total_equity and soe_closing_2024:
            bs_eq = bs_total_equity["values"].get("2024")
            soe_total = soe_row_total(soe_closing_2024["raw_row"])
            rec = make_internal_record(
                "BS", SECTION_TO_PDF_PAGE["BS"], "Total members' equity (2024)", "2024", bs_eq,
                "SOE", SECTION_TO_PDF_PAGE["SOE"], "Balance as of December 31, 2024 (Total)", soe_total,
                "/SoSE",
                "BS Total Members' Equity 2024 = SOE closing 2024 Total",
            )
            if rec:
                records.append(rec)

        # Individual equity classes — match by column position
        # SOE column layout (observed from raw row data; balance rows are len=28-29):
        #   col 0: row label
        #   col 1: SCU Units      col 3: SCU Amount
        #   col 5: A-1 Units      col 7: A-1 Amount
        #   col 9: A-2 Units      col 11: A-2 Amount
        #   col 13: B Units       col 15: B Amount
        #   col 17: C Units       col 19: C Amount
        #   col 21: D Units       col 23: D Amount
        #   col 25: Accumulated Deficit
        #   col 26: AOCI
        #   col 27 or 28: Total (drifts; use soe_last_numeric)
        bs_to_soe_class = [
            (lambda l: "senior converted units" in l.lower() and "redemption" not in l.lower(), 3, "Senior Converted Units (Amount)"),
            (lambda l: l.lower().startswith("class a-1"), 7, "Class A-1 (Amount)"),
            (lambda l: l.lower().startswith("class a"), 11, "Class A (Amount)"),
            (lambda l: l.lower().startswith("class b units"), 15, "Class B (Amount)"),
            (lambda l: l.lower().startswith("class c units"), 19, "Class B (Amount)"),
            (lambda l: l.lower().startswith("class d units"), 23, "Class D (Amount)"),
            (lambda l: l.lower().startswith("accumulated deficit"), 25, "Accumulated deficit"),
            (lambda l: "accumulated other comprehensive" in l.lower(), 26, "AOCI"),
        ]

        for matcher, soe_col, soe_label in bs_to_soe_class:
            bs_row = find_row(bs, matcher)
            if not bs_row:
                continue
            for year, soe_balance in (("2025", soe_closing_2025), ("2024", soe_closing_2024)):
                if not soe_balance:
                    continue
                bs_v = bs_row["values"].get(year)
                soe_v = soe_amount_for_class(soe_balance["raw_row"], soe_col)
                rec = make_internal_record(
                    "BS", SECTION_TO_PDF_PAGE["BS"], f"{bs_row['label'][:40]} ({year})", year, bs_v,
                    "SOE", SECTION_TO_PDF_PAGE["SOE"], f"{soe_label} {year}", soe_v,
                    "/SoSE",
                    f"BS {soe_label} {year} = SOE closing {year}",
                )
                if rec:
                    records.append(rec)

    # ==========================================================================
    # === SOE <-> IS (Net loss row) ===
    # ==========================================================================
    if soe and is_t:
        # SOE Net loss row — Total column = IS Net loss
        rows = soe.get("rows", [])
        # 2025 Net loss in SOE is after "Balance as of December 31, 2024" row
        # 2024 Net loss in SOE is between "Balance as of December 31, 2023" and "Balance as of December 31, 2024"
        idx_2023 = next((i for i, r in enumerate(rows) if r and r[0] and "December 31, 2023" in r[0]), None)
        idx_2024 = next((i for i, r in enumerate(rows) if r and r[0] and "December 31, 2024" in r[0]), None)
        idx_2025 = next((i for i, r in enumerate(rows) if r and r[0] and "December 31, 2025" in r[0]), None)

        is_net_loss = find_row(is_t, lambda l: l.lower().strip() == "net loss")

        if idx_2023 is not None and idx_2024 is not None and is_net_loss:
            # Net loss for 2024 is between idx_2023 and idx_2024
            for j in range(idx_2023 + 1, idx_2024):
                row = rows[j]
                if row and row[0] and "net loss" in (row[0] or "").lower():
                    soe_total = soe_last_numeric(row)
                    if soe_total is not None:
                        is_v = is_net_loss["values"].get("2024")
                        if is_v is not None:
                            rec = make_internal_record(
                                "SOE", SECTION_TO_PDF_PAGE["SOE"], "Net loss 2024 (SOE)", "2024", abs(soe_total),
                                "IS", SECTION_TO_PDF_PAGE["IS"], "Net loss (IS) 2024", abs(is_v),
                                "/PL",
                                "SOE Net Loss 2024 (Total col) = IS Net Loss 2024",
                            )
                            if rec:
                                records.append(rec)
                    break

        if idx_2024 is not None and idx_2025 is not None and is_net_loss:
            for j in range(idx_2024 + 1, idx_2025):
                row = rows[j]
                if row and row[0] and "net loss" in (row[0] or "").lower():
                    soe_total = soe_last_numeric(row)
                    if soe_total is not None:
                        is_v = is_net_loss["values"].get("2025")
                        if is_v is not None:
                            rec = make_internal_record(
                                "SOE", SECTION_TO_PDF_PAGE["SOE"], "Net loss 2025 (SOE)", "2025", abs(soe_total),
                                "IS", SECTION_TO_PDF_PAGE["IS"], "Net loss (IS) 2025", abs(is_v),
                                "/PL",
                                "SOE Net Loss 2025 (Total col) = IS Net Loss 2025",
                            )
                            if rec:
                                records.append(rec)
                    break

    # ==========================================================================
    # === SOE <-> SCF (Stock-based compensation, Redemption) ===
    # ==========================================================================
    if soe and scf:
        scf_sbc = find_row(scf, lambda l: l.lower().strip() == "stock-based compensation")
        scf_redemption = find_row(scf, lambda l: "redemption of senior converted units" in l.lower())
        rows = soe.get("rows", [])
        idx_2023 = next((i for i, r in enumerate(rows) if r and r[0] and "December 31, 2023" in r[0]), None)
        idx_2024 = next((i for i, r in enumerate(rows) if r and r[0] and "December 31, 2024" in r[0]), None)
        idx_2025 = next((i for i, r in enumerate(rows) if r and r[0] and "December 31, 2025" in r[0]), None)

        # SBC rows (2024 = between 23 and 24 balance; 2025 = between 24 and 25 balance)
        if scf_sbc and idx_2023 is not None and idx_2024 is not None:
            for j in range(idx_2023 + 1, idx_2024):
                row = rows[j]
                if row and row[0] and "stock-based compensation" in row[0].lower():
                    soe_sbc = soe_last_numeric(row)
                    scf_v = scf_sbc["values"].get("2024")
                    if soe_sbc is not None and scf_v is not None:
                        rec = make_internal_record(
                            "SOE", SECTION_TO_PDF_PAGE["SOE"], "Stock-based compensation 2024 (SOE)", "2024", soe_sbc,
                            "SCF", SECTION_TO_PDF_PAGE["SCF"], "SBC (SCF) 2024", scf_v,
                            "/SoCF",
                            "SOE SBC 2024 = SCF SBC 2024",
                        )
                        if rec:
                            records.append(rec)
                    break

        if scf_sbc and idx_2024 is not None and idx_2025 is not None:
            for j in range(idx_2024 + 1, idx_2025):
                row = rows[j]
                if row and row[0] and "stock-based compensation" in row[0].lower():
                    soe_sbc = soe_last_numeric(row)
                    scf_v = scf_sbc["values"].get("2025")
                    if soe_sbc is not None and scf_v is not None:
                        rec = make_internal_record(
                            "SOE", SECTION_TO_PDF_PAGE["SOE"], "Stock-based compensation 2025 (SOE)", "2025", soe_sbc,
                            "SCF", SECTION_TO_PDF_PAGE["SCF"], "SBC (SCF) 2025", scf_v,
                            "/SoCF",
                            "SOE SBC 2025 = SCF SBC 2025",
                        )
                        if rec:
                            records.append(rec)
                    break

        # Redemption of Senior Converted Units (2025 only)
        if scf_redemption and idx_2024 is not None and idx_2025 is not None:
            for j in range(idx_2024 + 1, idx_2025):
                row = rows[j]
                if row and row[0] and "redemption of senior converted units" in row[0].lower():
                    soe_redemption = soe_last_numeric(row)
                    scf_v = scf_redemption["values"].get("2025")
                    if soe_redemption is not None and scf_v is not None:
                        rec = make_internal_record(
                            "SOE", SECTION_TO_PDF_PAGE["SOE"], "Redemption SCU 2025 (SOE)", "2025",
                            abs(soe_redemption),
                            "SCF", SECTION_TO_PDF_PAGE["SCF"], "Redemption SCU (SCF) 2025",
                            abs(scf_v),
                            "/SoCF",
                            "SOE Redemption SCU 2025 = SCF Redemption SCU 2025",
                        )
                        if rec:
                            records.append(rec)
                    break

    # ==========================================================================
    # === BS <-> FN-03 (Property and Equipment) ===
    # ==========================================================================
    if bs:
        fn03_tables = find_fn_table(tables, "FN-03", "composition of property")
        if fn03_tables:
            fn03 = fn03_tables[0]
            bs_ppe = find_row(bs, lambda l: l.lower().startswith("property and equipment, net"))
            # FN-03 net row contains both 2025 and 2024 columns
            fn03_net = find_row(fn03, lambda l: l.lower().startswith("property and equipment, net"))
            if bs_ppe and fn03_net:
                for year in ("2025", "2024"):
                    bs_v = bs_ppe["values"].get(year)
                    fn_v = fn03_net["values"].get(year)
                    rec = make_internal_record(
                        "BS", SECTION_TO_PDF_PAGE["BS"], f"Property and equipment, net ({year})", year, bs_v,
                        "FN-03", FN_TO_PDF_PAGE["FN-03"], f"FN-03 Net PP&E {year}", fn_v,
                        "/FN 3",
                        "BS PP&E net = FN-03 net PP&E",
                    )
                    if rec:
                        records.append(rec)

    # ==========================================================================
    # === BS <-> FN-04 (Intangibles, Goodwill) ===
    # ==========================================================================
    if bs:
        fn04_tables = find_fn_table(tables, "FN-04", "summary of the company")
        # There are 2 FN-04 tables — one for 2025, one for 2024 (parallel structure)
        if len(fn04_tables) >= 2:
            fn04_2025, fn04_2024 = fn04_tables[0], fn04_tables[1]
            # FN-04 column layout (observed): col 0=label, col 1=Gross, col 3=Accum Amort, col 5=Net
            net_col = 5
            bs_intangibles = find_row(bs, lambda l: l.lower().startswith("intangible assets, net"))
            bs_goodwill = find_row(bs, lambda l: l.lower().startswith("goodwill"))

            # Find "Total intangible assets" and "Goodwill" rows in each FN-04 table
            def fn04_total_intangibles(fn04t):
                for row in fn04t.get("rows", []):
                    if row and row[0] and row[0].strip().lower().startswith("total intangible assets"):
                        return parse_value(row[net_col]) if net_col < len(row) else None
                return None

            def fn04_total_goodwill(fn04t):
                for row in fn04t.get("rows", []):
                    if row and row[0] and "total goodwill" in row[0].strip().lower():
                        return parse_value(row[net_col]) if net_col < len(row) else None
                return None

            int_2025 = fn04_total_intangibles(fn04_2025)
            int_2024 = fn04_total_intangibles(fn04_2024)
            gw_2025 = fn04_total_goodwill(fn04_2025)
            gw_2024 = fn04_total_goodwill(fn04_2024)

            if bs_intangibles:
                if int_2025 is not None:
                    rec = make_internal_record(
                        "BS", SECTION_TO_PDF_PAGE["BS"], "Intangible assets, net (2025)", "2025",
                        bs_intangibles["values"].get("2025"),
                        "FN-04", FN_TO_PDF_PAGE["FN-04"], "FN-04 Total intangibles net 2025", int_2025,
                        "/FN 4",
                        "BS Intangibles net 2025 = FN-04 total net 2025",
                    )
                    if rec:
                        records.append(rec)
                if int_2024 is not None:
                    rec = make_internal_record(
                        "BS", SECTION_TO_PDF_PAGE["BS"], "Intangible assets, net (2024)", "2024",
                        bs_intangibles["values"].get("2024"),
                        "FN-04", FN_TO_PDF_PAGE["FN-04"], "FN-04 Total intangibles net 2024", int_2024,
                        "/FN 4",
                        "BS Intangibles net 2024 = FN-04 total net 2024",
                    )
                    if rec:
                        records.append(rec)

            if bs_goodwill:
                if gw_2025 is not None:
                    rec = make_internal_record(
                        "BS", SECTION_TO_PDF_PAGE["BS"], "Goodwill, net (2025)", "2025",
                        bs_goodwill["values"].get("2025"),
                        "FN-04", FN_TO_PDF_PAGE["FN-04"], "FN-04 Total goodwill net 2025", gw_2025,
                        "/FN 4",
                        "BS Goodwill 2025 = FN-04 goodwill net 2025",
                    )
                    if rec:
                        records.append(rec)
                if gw_2024 is not None:
                    rec = make_internal_record(
                        "BS", SECTION_TO_PDF_PAGE["BS"], "Goodwill, net (2024)", "2024",
                        bs_goodwill["values"].get("2024"),
                        "FN-04", FN_TO_PDF_PAGE["FN-04"], "FN-04 Total goodwill net 2024", gw_2024,
                        "/FN 4",
                        "BS Goodwill 2024 = FN-04 goodwill net 2024",
                    )
                    if rec:
                        records.append(rec)

    # ==========================================================================
    # === BS <-> FN-05 (Accrued Expenses) ===
    # ==========================================================================
    if bs:
        fn05_tables = find_fn_table(tables, "FN-05", "accrued expenses")
        if fn05_tables:
            fn05 = fn05_tables[0]
            bs_accrued = find_row(bs, lambda l: l.lower().startswith("accrued expenses and other"))
            fn05_total = find_row(fn05, lambda l: l.lower().startswith("total accrued expenses"))
            if bs_accrued and fn05_total:
                for year in ("2025", "2024"):
                    bs_v = bs_accrued["values"].get(year)
                    fn_v = fn05_total["values"].get(year)
                    rec = make_internal_record(
                        "BS", SECTION_TO_PDF_PAGE["BS"], f"Accrued expenses ({year})", year, bs_v,
                        "FN-05", FN_TO_PDF_PAGE["FN-05"], f"FN-05 Total accrued {year}", fn_v,
                        "/FN 5",
                        "BS Accrued = FN-05 total accrued",
                    )
                    if rec:
                        records.append(rec)

    # ==========================================================================
    # === BS <-> FN-06 (Term Loans) ===
    # ==========================================================================
    if bs:
        fn06_tables = find_fn_table(tables, "FN-06", "term loans and line of credit obligations")
        if fn06_tables:
            fn06 = fn06_tables[0]
            bs_loans_current = find_row(bs, lambda l: l.lower().startswith("term loans, current"))
            bs_loans_lt = find_row(bs, lambda l: l.lower().startswith("term loans, net of current"))
            # FN-06: "Less: current portion of term loans" gives current; "Total long-term term loans" gives LT
            # But the FN-06 "Less: current portion" is negative — use abs
            fn06_current = find_row(fn06, lambda l: "current portion of term loans" in l.lower() and l.startswith("Less:"))
            fn06_lt = find_row(fn06, lambda l: l.lower().startswith("total long-term term loans"))

            if bs_loans_current and fn06_current:
                for year in ("2025", "2024"):
                    bs_v = bs_loans_current["values"].get(year)
                    fn_v = abs(fn06_current["values"].get(year)) if fn06_current["values"].get(year) is not None else None
                    rec = make_internal_record(
                        "BS", SECTION_TO_PDF_PAGE["BS"], f"Term loans, current ({year})", year, bs_v,
                        "FN-06", FN_TO_PDF_PAGE["FN-06"], f"FN-06 current portion {year}", fn_v,
                        "/FN 6",
                        "BS Term loans current = FN-06 current portion",
                    )
                    if rec:
                        records.append(rec)

            if bs_loans_lt and fn06_lt:
                for year in ("2025", "2024"):
                    bs_v = bs_loans_lt["values"].get(year)
                    fn_v = fn06_lt["values"].get(year)
                    rec = make_internal_record(
                        "BS", SECTION_TO_PDF_PAGE["BS"], f"Term loans, net of current ({year})", year, bs_v,
                        "FN-06", FN_TO_PDF_PAGE["FN-06"], f"FN-06 long-term {year}", fn_v,
                        "/FN 6",
                        "BS Term loans LT = FN-06 long-term",
                    )
                    if rec:
                        records.append(rec)

    # ==========================================================================
    # === IS <-> FN-07 (Income Tax) ===
    # ==========================================================================
    if is_t:
        fn07_tables = find_fn_table(tables, "FN-07", "components of the provision for income tax")
        if fn07_tables:
            fn07 = fn07_tables[0]
            is_tax = find_row(is_t, lambda l: l.lower().strip() == "income tax expense")
            fn07_tax = find_row(fn07, lambda l: l.lower().strip() == "income tax expense")
            if is_tax and fn07_tax:
                for year in ("2025", "2024"):
                    is_v = is_tax["values"].get(year)
                    fn_v = fn07_tax["values"].get(year)
                    # IS shows negative (expense), FN shows positive — use abs
                    rec = make_internal_record(
                        "IS", SECTION_TO_PDF_PAGE["IS"], f"Income tax expense ({year})", year,
                        abs(is_v) if is_v is not None else None,
                        "FN-07", FN_TO_PDF_PAGE["FN-07"], f"FN-07 Income tax expense {year}",
                        abs(fn_v) if fn_v is not None else None,
                        "/FN 7",
                        "IS Tax expense = FN-07 tax expense",
                    )
                    if rec:
                        records.append(rec)

        # Loss before income taxes — IS vs FN-07 components of (loss) before tax
        fn07_pretax_tables = find_fn_table(tables, "FN-07", "components of income (loss) before income taxes")
        if fn07_pretax_tables:
            fn07p = fn07_pretax_tables[0]
            is_pretax = find_row(is_t, lambda l: l.lower().strip() == "loss before income taxes")
            # FN-07 pre-tax total row has no label — find_row won't match it,
            # so we walk rows manually using the year-col map from the header.
            year_cols = {}
            for r in fn07p.get("rows", [])[:5]:
                cells = {c: cell.strip() for c, cell in enumerate(r) if cell and isinstance(cell, str) and re.match(r"^20\d{2}$", cell.strip())}
                if len(cells) >= 2:
                    year_cols = cells
                    break
            fn07_pretax_total = None
            for r_idx, row in enumerate(fn07p.get("rows", [])):
                # Skip the first 3 header rows (date header, year header, units header)
                if r_idx < 3:
                    continue
                if row and (not row[0] or row[0].strip() == "") and len(row) > 1:
                    candidate_vals = {y: parse_value(row[c]) for c, y in year_cols.items() if c < len(row)}
                    v25 = candidate_vals.get("2025")
                    # The total row has values > $1M abs (i.e., $1000K in display)
                    if v25 is not None and abs(v25) > 1000:
                        fn07_pretax_total = {"label": "Pre-tax total (FN-07)",
                                             "values": candidate_vals}
                        break
            if is_pretax and fn07_pretax_total:
                for year in ("2025", "2024"):
                    is_v = is_pretax["values"].get(year)
                    fn_v = fn07_pretax_total["values"].get(year)
                    rec = make_internal_record(
                        "IS", SECTION_TO_PDF_PAGE["IS"], f"Loss before income taxes ({year})", year,
                        abs(is_v) if is_v is not None else None,
                        "FN-07", FN_TO_PDF_PAGE["FN-07"], f"FN-07 Pre-tax (loss) {year}",
                        abs(fn_v) if fn_v is not None else None,
                        "/FN 7",
                        "IS Loss before tax = FN-07 pre-tax loss",
                    )
                    if rec:
                        records.append(rec)

    # ==========================================================================
    # === BS <-> FN-11 (Lease Liabilities) ===
    # ==========================================================================
    if bs:
        fn11_tables = find_fn_table(tables, "FN-11", "maturities")
        if fn11_tables:
            fn11 = fn11_tables[0]
            bs_lease_current = find_row(bs, lambda l: l.lower().startswith("lease liabilities, current"))
            bs_lease_lt = find_row(bs, lambda l: l.lower().startswith("lease liabilities, net of current"))
            fn11_current = find_row(fn11, lambda l: l.lower().startswith("operating lease liabilities, current"))
            fn11_lt = find_row(fn11, lambda l: l.lower().startswith("operating lease liabilities, noncurrent"))

            # FN-11 maturities table only has 2025 column
            if bs_lease_current and fn11_current:
                bs_v = bs_lease_current["values"].get("2025")
                # FN-11 maturities row values are in col 1
                fn_row = fn11_current["raw_row"]
                fn_v = parse_value(fn_row[1]) if len(fn_row) > 1 else None
                rec = make_internal_record(
                    "BS", SECTION_TO_PDF_PAGE["BS"], "Lease liabilities, current (2025)", "2025", bs_v,
                    "FN-11", FN_TO_PDF_PAGE["FN-11"], "FN-11 Op lease liab, current 2025", fn_v,
                    "/FN 11",
                    "BS Lease liab current = FN-11 op lease liab current",
                )
                if rec:
                    records.append(rec)

            if bs_lease_lt and fn11_lt:
                bs_v = bs_lease_lt["values"].get("2025")
                fn_row = fn11_lt["raw_row"]
                fn_v = parse_value(fn_row[1]) if len(fn_row) > 1 else None
                rec = make_internal_record(
                    "BS", SECTION_TO_PDF_PAGE["BS"], "Lease liabilities, net of current (2025)", "2025", bs_v,
                    "FN-11", FN_TO_PDF_PAGE["FN-11"], "FN-11 Op lease liab, noncurrent 2025", fn_v,
                    "/FN 11",
                    "BS Lease liab LT = FN-11 op lease liab noncurrent",
                )
                if rec:
                    records.append(rec)

    # ==========================================================================
    # === Internal: IS Loss from operations + Total other expense = Loss before tax ===
    # ==========================================================================
    if is_t:
        is_loss_ops = find_row(is_t, lambda l: l.lower().strip() == "loss from operations")
        is_other = find_row(is_t, lambda l: "total other" in l.lower() and "expense" in l.lower())
        is_loss_pretax = find_row(is_t, lambda l: l.lower().strip() == "loss before income taxes")
        if is_loss_ops and is_other and is_loss_pretax:
            for year in ("2025", "2024"):
                lo = is_loss_ops["values"].get(year)
                ot = is_other["values"].get(year)
                pt = is_loss_pretax["values"].get(year)
                if lo is not None and ot is not None and pt is not None:
                    computed = lo + ot
                    delta, status, tol, cu = compare(pt, computed, "$K", "$K",
                                                    is_subtotal=True, kind="internal")
                    records.append({**make_record(
                        lane="pdf_internal",
                        pdf_section="IS",
                        pdf_page=SECTION_TO_PDF_PAGE["IS"],
                        pdf_label=f"IS Loss from ops + Total other = Loss before tax ({year})",
                        pdf_year=year,
                        pdf_value=pt,
                        source_ref=f"PDF!p{SECTION_TO_PDF_PAGE['IS']}-IS-pretax-rollforward",
                        source_label=f"({lo:,.0f} ops + {ot:,.0f} other)",
                        source_value=computed,
                        comparison_unit=cu, delta=delta, tolerance=tol,
                        status=status, is_subtotal=True,
                        notes="IS pretax components consistency",
                    ), "mark": "F", "target_section": "IS", "target_page": SECTION_TO_PDF_PAGE["IS"]})

    # ==========================================================================
    # Output
    # ==========================================================================
    by_status = {}
    by_mark = {}
    for r in records:
        by_status[r["status"]] = by_status.get(r["status"], 0) + 1
        by_mark[r.get("mark", "?")] = by_mark.get(r.get("mark", "?"), 0) + 1

    print(f"Internal cross-refs: {len(records)} records")
    print(f"By status: {by_status}")
    print(f"By mark:   {by_mark}")
    for r in records:
        marker = "✓" if r["status"] in ("ties", "ties-with-rounding") else "!"
        mark = r.get("mark", "")
        pdf_v = r.get("pdf_value")
        src_v = r.get("source_value")
        print(f"  {marker} [{r['pdf_section']:>4}] {mark:>7} {r['pdf_label'][:50]:<50}  pdf={pdf_v}  src={src_v}  status={r['status']}")

    Path(args.out).write_text(
        json.dumps(records, default=str, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    print(f"Wrote {args.out}")


if __name__ == "__main__":
    main()
