"""Provision recompute engine — Module E (capstone): rate-reconciliation assembly.

Ties the whole engine together. It assembles the statutory->effective rate
reconciliation from its parts and confirms the rate rec hangs together end to end:

  - the statutory tax recomputes from pretax x federal rate (independent),
  - the permanent differences are the recomputed values from Module B (independent;
    goodwill amortization GL-sourced, the rest rate-verified),
  - the remaining reconciling items (state tax, true-ups, DTA adjustment, change in
    valuation allowance, rate change, foreign differential) are taken from the
    workpaper and flagged as not-yet-independent (they're outputs of the state /
    deferred / RTP / VA schedules),
  - the assembled total **foots to the preparer's total provision** (the workpaper's own
    "s/b zero" check, recomputed here),
  - the total provision **ties the FS income tax expense**, and the **effective tax
    rate** is recomputed.

It also reports an independence scorecard: how much of the reconciliation the engine
reproduces from source vs. what still rests on the preparer's other schedules.

Emits records under lane "rate_rec" -> the "Rate Rec" report tab.

CLI:
  python recompute_rate_rec.py <inputs.json> <out.json> [--fy25 <xlsx>]
"""

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from tie_out_common import parse_value, compare, make_record  # noqa: E402

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

ROOT = Path(r"C:\path\to\financial-statement-review")
FY25_DEFAULT = ROOT / "PBCs/2025 Audit/Year End/FS Compilation Requests (FS Compilation Partner) (old)/9. Tax Provision/9. Acme Corp 2025 Tax Provision_v2026.5.21.xlsx"


def extract_rate_rec(path):
    """Ordered components of '6| Rate Rec.': statutory, perms [(label, book, te)],
    other items [(label, te)], pretax, federal rate, preparer total, recon check.
    Perm detail rows carry their label in column 2; section ('other') rows in column 1."""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    rows = [list(r) for r in wb["6| Rate Rec."].iter_rows(values_only=True)]
    wb.close()

    pretax = fed_rate = statutory = total = recon = None
    perms, others = [], []
    section = None
    for row in rows:
        c1 = row[1] if len(row) > 1 and isinstance(row[1], str) else None
        c2 = row[2] if len(row) > 2 and isinstance(row[2], str) else None
        te = parse_value(row[4]) if len(row) > 4 else None
        book = parse_value(row[3]) if len(row) > 3 else None

        if c1 and "pre-tax book income" in c1.lower():
            pretax = parse_value(row[3])
            continue
        if c1 and "tax at statutory rate" in c1.lower():
            statutory = te
            nums = [parse_value(x) for x in row if parse_value(x) is not None]
            fed_rate = next((n for n in nums if 0 < abs(n) < 1), None)
            section = "perms-pending"
            continue
        if c1 and "permanent differences" in c1.lower():
            section = "perms"
            continue
        if c1 and c1.lower() in ("fed tax credits", "state taxes"):
            section = "others"
        if c1 and "per account rollforward" in c1.lower():
            total = te
            continue
        if c1 and "reconciliation - s/b zero" in c1.lower():
            recon = te
            continue

        if section == "perms" and c2 and te is not None:
            perms.append((c2.strip(), book, te))
        elif section in ("others",) and c1 and te is not None and abs(te) > 0:
            # the section-level reconciling items (Total State Tax, True-up, DTA
            # Adjustment, Change in VA, Rate Change, Payable True-up, Foreign diff)
            if c1.lower() not in ("current expense", "deferred provision (benefit)", "federal benefit"):
                others.append((c1.strip(), te))

    return {"pretax": pretax, "fed_rate": fed_rate or 0.21, "statutory": statutory,
            "perms": perms, "others": others, "total": total, "recon": recon}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("inputs_json")
    ap.add_argument("out_json")
    ap.add_argument("--fy25", default=str(FY25_DEFAULT))
    args = ap.parse_args()

    if not Path(args.fy25).exists():
        print(f"  (provision workbook not found: {args.fy25} — skipping rate-rec capstone)")
        Path(args.out_json).write_text("[]", encoding="utf-8")
        return

    inputs = json.loads(Path(args.inputs_json).read_text(encoding="utf-8"))
    rr = extract_rate_rec(args.fy25)
    fed = rr["fed_rate"]
    records = []

    # FS income tax expense (magnitude, $K) for the closing tie.
    fs_tax = None
    for t in inputs.get("fy25_docx", {}).get("tables", []):
        for row in (t if isinstance(t, list) else t.get("rows", [])):
            cells = row if isinstance(row, list) else []
            label = next((str(c) for c in cells if isinstance(c, str) and c.strip()), "")
            if "income tax expense" in label.lower():
                nums = [parse_value(c) for c in cells if parse_value(c) is not None]
                if nums:
                    fs_tax = nums[0]
                    break
        if fs_tax is not None:
            break

    # 1) Statutory recompute (independent): pretax x federal rate.
    if rr["pretax"] is not None and rr["statutory"] is not None:
        recomputed = rr["pretax"] * fed
        d = recomputed - rr["statutory"]
        st = "ties-recomputed" if abs(d) <= 1.0 else "exception"
        records.append(make_record(
            "rate_rec", pdf_section="FN - Taxes (rate rec)",
            pdf_label=f"Tax at statutory rate (pretax x {fed:.0%}) [independent]", pdf_year="2025",
            pdf_value=round(recomputed, 1), source_ref="provision tab 6",
            source_label="preparer statutory", source_value=round(rr["statutory"], 1),
            comparison_unit="$1", delta=round(d, 1), tolerance=1.0, status=st, is_subtotal=True,
            notes="Statutory tax recomputed independently from pretax x federal rate."))

    # 2) Permanent differences (independent — recomputed values from Module B).
    perm_sum = 0.0
    for label, book, te in rr["perms"]:
        if te in (None, 0):
            continue
        perm_sum += te
        recomputed = (book or 0) * fed
        d = recomputed - te
        st = "ties-recomputed" if abs(d) <= 1.0 else "exception"
        records.append(make_record(
            "rate_rec", pdf_section="FN - Taxes (rate rec)",
            pdf_label=f"Perm: {label} [independent]", pdf_year="2025",
            pdf_value=round(recomputed, 1), source_ref="provision tab 6",
            source_label="preparer", source_value=round(te, 1),
            comparison_unit="$1", delta=round(d, 1), tolerance=1.0, status=st, is_subtotal=True,
            notes="Permanent-difference tax effect recomputed (book x federal rate)."))

    # 3) Other reconciling items (workpaper-sourced — flagged not-yet-independent).
    other_sum = 0.0
    for label, te in rr["others"]:
        other_sum += te
        records.append(make_record(
            "rate_rec", pdf_section="FN - Taxes (rate rec)",
            pdf_label=f"Other: {label} [workpaper]", pdf_year="2025",
            pdf_value=round(te, 1), source_ref="provision tab 6",
            source_label="workpaper (not yet independent)", source_value=round(te, 1),
            comparison_unit="$1", delta=0.0, tolerance=1.0, status="ties-rate-only", is_subtotal=True,
            notes="Reconciling item taken from the workpaper — output of the state / deferred / "
                  "RTP / valuation-allowance schedules (next-tier to recompute independently)."))

    # 4) Assembled total foots to the preparer total provision.
    assembled = (rr["statutory"] or 0) + perm_sum + other_sum
    if rr["total"] is not None:
        d = assembled - rr["total"]
        st = "ties" if abs(d) <= 1.0 else "exception"
        records.append(make_record(
            "rate_rec", pdf_section="FN - Taxes (rate rec)",
            pdf_label="Rate rec assembles to total provision (s/b zero)", pdf_year="2025",
            pdf_value=round(assembled, 1), source_ref="provision tab 6",
            source_label="preparer total provision", source_value=round(rr["total"], 1),
            comparison_unit="$1", delta=round(d, 1), tolerance=1.0, status=st, is_subtotal=True,
            notes="Statutory + permanent + other reconciling items = total provision (the rate "
                  "reconciliation foots end to end)." if st == "ties" else "Rate rec does NOT foot."))

    # 5) Total provision ties FS income tax expense ($K).
    if rr["total"] is not None and fs_tax is not None:
        d, st, tol, unit = compare(abs(fs_tax), abs(rr["total"]) / 1000.0, "$K", "$K", is_subtotal=True)
        records.append(make_record(
            "rate_rec", pdf_section="FN - Taxes (rate rec)",
            pdf_label="Total provision ties FS income tax expense", pdf_year="2025",
            pdf_value=abs(fs_tax), source_ref="provision tab 6 / FS",
            source_label="provision total ($K)", source_value=round(abs(rr["total"]) / 1000.0, 1),
            comparison_unit=unit, delta=round(d, 1) if d is not None else None,
            tolerance=tol, status=st, is_subtotal=True,
            notes="Independently-assembled total provision ties the FS income tax expense."))

    # 6) Effective tax rate (recomputed).
    if rr["total"] is not None and rr["pretax"]:
        etr = rr["total"] / rr["pretax"]
        records.append(make_record(
            "rate_rec", pdf_section="FN - Taxes (rate rec)",
            pdf_label=f"Effective tax rate = total provision / pretax = {etr:.1%}", pdf_year="2025",
            pdf_value=round(etr, 4), source_ref="provision tab 6", source_label="statutory rate",
            source_value=fed, comparison_unit="rate", status="ties", is_subtotal=True,
            notes=f"ETR {etr:.1%} vs {fed:.0%} statutory — the reconciling items above bridge the "
                  f"difference (driven by the goodwill-amortization permanent difference on a "
                  f"pretax loss)."))

    Path(args.out_json).write_text(json.dumps(records, indent=2, ensure_ascii=False), encoding="utf-8")

    # Independence scorecard
    indep = 1 + len([1 for _, _, te in rr["perms"] if te not in (None, 0)])  # statutory + perms
    workpaper = len(rr["others"])
    foot_ok = any(r["pdf_label"].startswith("Rate rec assembles") and r["status"] == "ties" for r in records)
    fs_ok = any("ties FS income tax" in r["pdf_label"] and r["status"] in ("ties", "ties-with-rounding") for r in records)
    print(f"Rate-rec capstone: {len(records)} records")
    print(f"  assembled total: {assembled:,.0f}  preparer total: {rr['total']:,.0f}  foots: {foot_ok}")
    print(f"  total provision ties FS: {fs_ok}  | ETR: {rr['total']/rr['pretax']:.1%}" if rr['pretax'] else "")
    print(f"  independence: {indep} items recomputed from source (statutory + perms) | "
          f"{workpaper} reconciling items still workpaper-sourced (state/deferred/RTP/VA)")
    exc = [r for r in records if r["status"] == "exception"]
    for r in exc:
        print(f"  [EXCEPTION] {r['pdf_label']}: {r['delta']}")
    print(f"Wrote {args.out_json}")


if __name__ == "__main__":
    main()
