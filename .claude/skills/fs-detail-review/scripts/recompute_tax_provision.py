"""Provision recompute engine — Module B: recompute permanent differences.

Consumes the learned tax-treatment map (Module A) and the current-year TB, and
*independently* recomputes each permanent-difference tax effect, then compares to the
preparer's workbook. This is the "recompute, then red-box the drift" half of the engine
(the complement to Lane 9, which ties the preparer to the FS).

Per perm item the status reflects how independent the reproduction is:
  - ties-recomputed : book pulled straight from the GL (tb-exact) and the recomputed
                      tax effect ties the preparer — fully independent, no workpaper trust.
  - ties-rate-only  : the tax-effecting rate is verified, but the book amount could not
                      be GL-sourced (name-divergent / schedule-derived) so it was taken
                      from the preparer's file — independence incomplete, sourcing flagged.
  - exception       : the recomputed tax effect does not reconcile — investigate.

Emits records under lane "tax_recompute" -> the "Provision Recompute" report tab.

CLI:
  python recompute_tax_provision.py <inputs.json> <tax-treatment-map.json> <out.json>
"""

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from tie_out_common import make_record  # noqa: E402

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


def load_tb(inputs):
    return {a["account"]: a.get("value")
            for a in inputs.get("tb_consolidated", {}).get("accounts", [])
            if a.get("value") is not None}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("inputs_json")
    ap.add_argument("map_json")
    ap.add_argument("out_json")
    args = ap.parse_args()

    inputs = json.loads(Path(args.inputs_json).read_text(encoding="utf-8"))
    tmap = json.loads(Path(args.map_json).read_text(encoding="utf-8"))
    tb = load_tb(inputs)
    fed = tmap.get("federal_rate") or 0.21

    records = []
    for e in tmap.get("perm_items", []):
        item = e["item"]
        prep_book = e.get("cur_book")
        prep_te = e.get("cur_tax_effect")
        rate = e.get("rate") or fed
        method = e.get("source_method")

        # Skip items that are zero / absent in the current year (event-driven perms).
        if prep_book in (None, 0) and prep_te in (None, 0):
            continue

        if method == "tb-exact" and e.get("tb_account") in tb:
            # Independent: book straight from the GL, tax-effect at the federal rate.
            gl_book = tb[e["tb_account"]]
            recomputed = gl_book * fed
            delta = (recomputed - prep_te) if prep_te is not None else None
            status = ("ties-recomputed" if delta is not None and abs(delta) <= 1.0
                      else "exception")
            note = (f"Independently recomputed from GL '{e['tb_account']}' "
                    f"({gl_book:,.0f} x {fed:.0%}); reproduces the preparer."
                    if status == "ties-recomputed" else
                    f"GL '{e['tb_account']}' x {fed:.0%} = {recomputed:,.0f} does NOT match "
                    f"the preparer's {prep_te:,.0f} — investigate.")
            src_book = gl_book
        else:
            # Not GL-independent yet: verify the tax-effecting rate off the preparer book,
            # and flag the sourcing gap (what it would take to source from the GL).
            recomputed = (prep_book or 0) * fed
            delta = (recomputed - prep_te) if (prep_te is not None and prep_book is not None) else None
            rate_ok = delta is not None and abs(delta) <= 1.0
            status = "ties-rate-only" if rate_ok else "exception"
            if method == "tb-name-divergent":
                gap = (prep_book - (e.get("tb_value") or 0)) if prep_book is not None else None
                src = (f"Candidate GL '{e.get('tb_account')}' = {e.get('tb_value'):,.0f} but the "
                       f"preparer book is {prep_book:,.0f} (gap {gap:,.0f}) — an aggregation/"
                       f"adjustment rule is needed to source this from the GL.")
            elif method == "schedule-derived":
                src = ("No single GL account supplies this (e.g. SBC comes off the cap-table "
                       "schedule) — needs the supporting sub-schedule to source independently.")
            else:
                src = "Book amount source not yet identified in the GL."
            note = (f"Tax-effecting verified at {fed:.0%} on the preparer's book; "
                    f"book amount NOT yet GL-independent. {src}"
                    if rate_ok else
                    f"Recomputed tax effect does not reconcile to the preparer. {src}")
            src_book = prep_book

        records.append(make_record(
            "tax_recompute",
            pdf_section="FN - Taxes (recompute)", pdf_label=f"{item} — perm tax effect",
            pdf_year="2025",
            pdf_value=round(recomputed, 1) if recomputed is not None else None,
            source_ref="preparer provision workbook",
            source_label="preparer tax effect", source_value=round(prep_te, 1) if prep_te is not None else None,
            comparison_unit="$1", delta=round(delta, 1) if delta is not None else None,
            tolerance=1.0, status=status, is_subtotal=True, notes=note,
        ))

    Path(args.out_json).write_text(json.dumps(records, indent=2, ensure_ascii=False), encoding="utf-8")

    by = {}
    for r in records:
        by[r["status"]] = by.get(r["status"], 0) + 1
    n_indep = by.get("ties-recomputed", 0)
    print(f"Provision recompute (perms): {len(records)} items | {by}")
    print(f"  Independently GL-sourced & tie: {n_indep} | rate-verified (book not yet independent): "
          f"{by.get('ties-rate-only', 0)} | exceptions: {by.get('exception', 0)}")
    for r in records:
        tag = {"ties-recomputed": "GL-INDEP", "ties-rate-only": "rate-only",
               "exception": "EXCEPTION"}.get(r["status"], r["status"])
        print(f"  [{tag:<9}] {r['pdf_label'][:40]:<40} recomputed={r['pdf_value']:>14,.0f} "
              f"preparer={r['source_value']:>14,.0f}")
    print(f"Wrote {args.out_json}")


if __name__ == "__main__":
    main()
