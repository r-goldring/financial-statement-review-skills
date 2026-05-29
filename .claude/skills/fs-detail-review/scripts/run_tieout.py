"""Orchestrator: run the full tie-out pipeline end-to-end.

Steps:
  1. build_inputs.py        → .work/inputs.json
  2. tie_out_pdf_to_bridge  → .work/tie-lane1-pdf-to-bridge.json
  3. tie_out_bridge_to_tb   → .work/tie-lane2-bridge-to-tb.json
  4. tie_out_pdf_prior_year → .work/tie-lane3-prior-year.json
  5. tie_out_pdf_internal   → .work/tie-lane4-internal.json
  6. annotate_tieout_pdf    → Tieout/<pdf-name>_TIEOUT.pdf
  7. build_exceptions_report → Tieout/Tieout Exceptions - <version>.xlsx

Each step is skippable via --skip flag, useful for partial re-runs.

CLI:
  python run_tieout.py                    # full run with defaults
  python run_tieout.py --skip-inputs      # skip rebuild of inputs.json
  python run_tieout.py --skip-ocr-pages   # restrict annotator pages
"""

import argparse
import json
import subprocess
import sys
from pathlib import Path
from datetime import datetime

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


SCRIPT_DIR = Path(__file__).parent
SKILL_DIR = SCRIPT_DIR.parent
WORK_DIR = SKILL_DIR / ".work"

# Path defaults — change here when bridge/PDF version updates
ROOT = Path(r"C:\path\to\financial-statement-review")
FY25_PDF = ROOT / "Tieout" / "vfinal Tieou" / "Acme Holdings, LLC 2025 Financial Statements.pdf"
FY25_VERSION = "vFinal"
TIEOUT_DIR = ROOT / "Tieout" / "vfinal Tieou"

INPUTS_JSON = WORK_DIR / "inputs.json"
TIE_LANE1 = WORK_DIR / "tie-lane1-pdf-to-bridge.json"
TIE_LANE2 = WORK_DIR / "tie-lane2-bridge-to-tb.json"
TIE_LANE3 = WORK_DIR / "tie-lane3-prior-year.json"
TIE_LANE4 = WORK_DIR / "tie-lane4-internal.json"
TIE_LANE5 = WORK_DIR / "tie-lane5-soe-rollforward.json"
TIE_LANE6 = WORK_DIR / "tie-lane6-footing.json"
TIE_LANE7 = WORK_DIR / "tie-lane7-mapping.json"
TIE_LANE8 = WORK_DIR / "tie-lane8-pbc-to-bridge.json"
TIE_LANE9 = WORK_DIR / "tie-lane9-tax-provision.json"
TIE_LANE10 = WORK_DIR / "tie-lane10-flux.json"
TIE_LANE11 = WORK_DIR / "tie-lane11-tax-recompute.json"
TIE_LANE12 = WORK_DIR / "tie-lane12-deferred.json"
TIE_LANE13 = WORK_DIR / "tie-lane13-current-nol.json"
TIE_LANE14 = WORK_DIR / "tie-lane14-rate-rec.json"
TIE_LANE15 = WORK_DIR / "tie-lane15-state.json"
PBC_INDEX = WORK_DIR / "pbc-index.json"
TAX_TREATMENT_MAP = WORK_DIR / "tax-treatment-map.json"
PBC_2025 = ROOT / "PBCs" / "2025 Audit"
TIE_CARRY_FWD = WORK_DIR / "tie-carry-forward.json"
OCR_CACHE = WORK_DIR / "ocr-cache.json"
FN_PAGE_MAP = WORK_DIR / "fn-page-map.json"
FY24_TIEOUT_PDF = ROOT / "Prior Year Examples" / "2024" / "Tieout" / "vYYYY.M.D_Acme Holdings LLC 2024 Financial Statements Tieout (FINAL).pdf"
ANNOTATED_PDF = TIEOUT_DIR / FY25_PDF.name.replace(".pdf", "_TIEOUT.pdf")


def find_carry_forward_source():
    """Prefer a prior-period tieout PDF that lives in the same TIEOUT_DIR (e.g. the
    user dropped a hand-annotated v5.21 tieout into the vfinal folder before running
    against the final FS) — those marks anchor much more cleanly than FY24's against
    the current PDF. Falls back to FY24_TIEOUT_PDF if no local prior tieout exists.

    Match criterion: a .pdf in TIEOUT_DIR that is NOT the FY25 input itself and NOT
    the annotated output, with 'Tieout' in the filename (case-insensitive). The most
    recently modified candidate wins.
    """
    if not TIEOUT_DIR.exists():
        return FY24_TIEOUT_PDF
    cands = []
    excluded = {FY25_PDF.name, ANNOTATED_PDF.name}
    for p in TIEOUT_DIR.glob("*.pdf"):
        if p.name in excluded:
            continue
        if "tieout" in p.name.lower():
            cands.append(p)
    if not cands:
        return FY24_TIEOUT_PDF
    cands.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return cands[0]
EXCEPTIONS_XLSX = TIEOUT_DIR / f"Tieout Exceptions - {FY25_VERSION}.xlsx"
AUDIT_LOG = SKILL_DIR / "audit_log.json"


def run(name, cmd):
    print(f"\n=== {name} ===")
    print(f"  $ {' '.join(str(c) for c in cmd[:4])} ...")
    result = subprocess.run(cmd, cwd=str(SCRIPT_DIR), capture_output=True, text=True, encoding="utf-8", errors="replace")
    if result.stdout:
        for line in result.stdout.rstrip().splitlines()[-30:]:
            print(f"  {line}")
    if result.returncode != 0:
        print(f"  ERROR (exit {result.returncode}):")
        if result.stderr:
            for line in result.stderr.rstrip().splitlines()[-20:]:
                print(f"    {line}")
        sys.exit(result.returncode)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--skip-inputs", action="store_true")
    ap.add_argument("--skip-lanes", action="store_true")
    ap.add_argument("--skip-annotate", action="store_true")
    ap.add_argument("--skip-exceptions", action="store_true")
    ap.add_argument("--skip-pbc", action="store_true", help="skip PBC index/register/Lane 8")
    ap.add_argument("--pages", default=None, help="restrict annotator to these pages")
    ap.add_argument("--keep-unverified-carry-forward", action="store_true",
                    help="keep carry-forward marks whose anchors didn't verify in the "
                         "current PDF (drawn in blue). Default drops them for a cleaner "
                         "annotated PDF when the carry-forward source has drifted.")
    args = ap.parse_args()

    WORK_DIR.mkdir(parents=True, exist_ok=True)
    TIEOUT_DIR.mkdir(parents=True, exist_ok=True)

    PY = sys.executable

    # Step 0: PBC index (Phase 0 foundation) — only if a PBC tree is present.
    pbc_script = SCRIPT_DIR / "build_pbc_index.py"
    if not args.skip_pbc and pbc_script.exists() and PBC_2025.exists():
        run("Step 0: build PBC index", [
            PY, str(pbc_script), "--root", str(PBC_2025), "--out", str(PBC_INDEX),
        ])
    elif not PBC_2025.exists():
        print(f"(no PBC tree at {PBC_2025} — skipping PBC index/register/Lane 8)")

    if not args.skip_inputs:
        run("Step 1: build inputs.json", [PY, str(SCRIPT_DIR / "build_inputs.py"), str(INPUTS_JSON)])
    else:
        print("(skipping build_inputs)")

    # Detect FN pages once (cached via OCR cache)
    if not args.skip_lanes:
        if FN_PAGE_MAP.exists():
            print(f"\n(FN page map exists at {FN_PAGE_MAP} — reusing)")
        else:
            run("Step 1.5: detect FN section -> PDF page", [
                PY, str(SCRIPT_DIR / "detect_fn_pages.py"),
                str(FY25_PDF), str(FN_PAGE_MAP),
                "--ocr-cache", str(OCR_CACHE),
            ])

        run("Step 2: lane 1 (PDF face <-> Bridge)", [
            PY, str(SCRIPT_DIR / "tie_out_pdf_to_bridge.py"),
            str(INPUTS_JSON), str(TIE_LANE1),
            "--fn-page-map", str(FN_PAGE_MAP),
        ])
        run("Step 3: lane 2 (Bridge <-> TB)",       [PY, str(SCRIPT_DIR / "tie_out_bridge_to_tb.py"), str(INPUTS_JSON), str(TIE_LANE2)])
        run("Step 4: lane 3 (FY25 PY col <-> FY24)", [PY, str(SCRIPT_DIR / "tie_out_pdf_prior_year.py"), str(INPUTS_JSON), str(TIE_LANE3)])
        run("Step 5: lane 4 (internal PDF refs)",   [PY, str(SCRIPT_DIR / "tie_out_pdf_internal.py"), str(INPUTS_JSON), str(TIE_LANE4)])
        # Lane 5 (SOE rollforward) is optional — only run if script exists
        soe_script = SCRIPT_DIR / "tie_out_soe.py"
        if soe_script.exists():
            run("Step 5.5: lane 5 (SOE rollforward)", [
                PY, str(soe_script), str(INPUTS_JSON), str(TIE_LANE5),
            ])
        # Lane 6 (footing checks) — F and xF tickmarks
        foot_script = SCRIPT_DIR / "tie_out_footing.py"
        if foot_script.exists():
            run("Step 5.6: lane 6 (footing checks)", [
                PY, str(foot_script), str(INPUTS_JSON), str(TIE_LANE6),
            ])
        # Carry-forward FY24 marks (text + drawings) — RED if verified, BLUE if needs review
        cf_script = SCRIPT_DIR / "carry_forward_fy24_marks.py"
        cf_source = find_carry_forward_source()
        if cf_script.exists() and cf_source.exists():
            label = "carry-forward marks" + (
                " (from local prior tieout)" if cf_source.parent == TIEOUT_DIR else " (from FY24 final)")
            run(f"Step 5.7: {label}", [
                PY, str(cf_script),
                str(cf_source), str(FY25_PDF),
                str(TIE_CARRY_FWD),
                "--ocr-cache", str(OCR_CACHE),
            ])
        # Lane 7 (mapping completeness & reasonableness) — TB-account findings, no PDF marks
        mc_script = SCRIPT_DIR / "tie_out_mapping_completeness.py"
        if mc_script.exists():
            run("Step 5.8: lane 7 (mapping completeness)", [
                PY, str(mc_script), str(INPUTS_JSON), str(TIE_LANE7),
            ])
        # Lane 8 (PBC -> bridge) — source-to-disclosure; needs the PBC index. No PDF marks.
        pbc_tie_script = SCRIPT_DIR / "tie_out_pbc_to_bridge.py"
        if not args.skip_pbc and pbc_tie_script.exists() and PBC_INDEX.exists():
            run("Step 5.9: lane 8 (PBC -> bridge)", [
                PY, str(pbc_tie_script), str(INPUTS_JSON), str(PBC_INDEX), str(TIE_LANE8),
            ])
        # Lane 9 (tax provision) — FN-07 rate rec / deferreds / book-pretax. Needs PBC index.
        tax_script = SCRIPT_DIR / "tie_out_tax_provision.py"
        if not args.skip_pbc and tax_script.exists() and PBC_INDEX.exists():
            run("Step 5.10: lane 9 (tax provision)", [
                PY, str(tax_script), str(INPUTS_JSON), str(PBC_INDEX), str(TIE_LANE9),
            ])
        # Lane 10 (flux) — YoY analytical review re-baselined to the final FS. Needs PBC index.
        flux_script = SCRIPT_DIR / "build_flux_analysis.py"
        if not args.skip_pbc and flux_script.exists() and PBC_INDEX.exists():
            run("Step 5.11: lane 10 (flux review)", [
                PY, str(flux_script), str(INPUTS_JSON), str(PBC_INDEX), str(TIE_LANE10),
            ])
        # Provision recompute engine (Increment 1: perms). Learn the treatment map from the
        # multi-year provision workbooks, then independently recompute + red-box drift.
        map_script = SCRIPT_DIR / "build_tax_treatment_map.py"
        recompute_script = SCRIPT_DIR / "recompute_tax_provision.py"
        if not args.skip_pbc and map_script.exists() and recompute_script.exists():
            run("Step 5.12: provision engine — learn tax-treatment map", [
                PY, str(map_script), str(INPUTS_JSON), str(TAX_TREATMENT_MAP),
            ])
            if TAX_TREATMENT_MAP.exists():
                run("Step 5.13: provision engine — recompute perms", [
                    PY, str(recompute_script), str(INPUTS_JSON), str(TAX_TREATMENT_MAP), str(TIE_LANE11),
                ])
        # Provision engine Module C: deferred taxes (FS tie + cross-year continuity).
        deferred_script = SCRIPT_DIR / "recompute_deferred_tax.py"
        if not args.skip_pbc and deferred_script.exists():
            run("Step 5.14: provision engine — deferred taxes", [
                PY, str(deferred_script), str(INPUTS_JSON), str(TIE_LANE12),
            ])
        # Provision engine Module D: current tax + NOL (taxable-income build + NOL trend).
        current_script = SCRIPT_DIR / "recompute_current_nol.py"
        if not args.skip_pbc and current_script.exists():
            run("Step 5.15: provision engine — current tax + NOL", [
                PY, str(current_script), str(TIE_LANE13),
            ])
        # Provision engine Module E (capstone): assemble the statutory->effective rate rec.
        rate_rec_script = SCRIPT_DIR / "recompute_rate_rec.py"
        if not args.skip_pbc and rate_rec_script.exists():
            run("Step 5.16: provision engine — rate-rec assembly (capstone)", [
                PY, str(rate_rec_script), str(INPUTS_JSON), str(TIE_LANE14),
            ])
        # Provision engine Module F: state tax rate derivation (apportionment -> blended rate).
        state_script = SCRIPT_DIR / "recompute_state_tax.py"
        if not args.skip_pbc and state_script.exists():
            run("Step 5.17: provision engine — state tax rate", [
                PY, str(state_script), str(TIE_LANE15),
            ])
    else:
        print("(skipping tie-out lanes)")

    if not args.skip_annotate:
        cmd = [PY, str(SCRIPT_DIR / "annotate_tieout_pdf.py"),
               str(FY25_PDF), str(ANNOTATED_PDF),
               str(TIE_LANE1), str(TIE_LANE3), str(TIE_LANE4)]
        if TIE_LANE5.exists():
            cmd.append(str(TIE_LANE5))
        if TIE_LANE6.exists():
            cmd.append(str(TIE_LANE6))
        if TIE_CARRY_FWD.exists():
            cmd.append(str(TIE_CARRY_FWD))
        cmd += ["--ocr-cache", str(OCR_CACHE)]
        # Default: drop unverified carry-forward marks so the annotated PDF only carries
        # marks whose anchors were verified against the current PDF. Use --keep-unverified
        # to restore the old behavior of drawing them in blue.
        if not args.keep_unverified_carry_forward:
            cmd.append("--drop-unverified-carry-forward")
        if args.pages:
            cmd += ["--pages", args.pages]
        run("Step 6: annotate PDF", cmd)
    else:
        print("(skipping annotation)")

    if not args.skip_exceptions:
        cmd = [
            PY, str(SCRIPT_DIR / "build_exceptions_report.py"),
            str(EXCEPTIONS_XLSX),
            str(TIE_LANE1), str(TIE_LANE2), str(TIE_LANE3), str(TIE_LANE4),
        ]
        if TIE_LANE5.exists():
            cmd.append(str(TIE_LANE5))
        if TIE_LANE6.exists():
            cmd.append(str(TIE_LANE6))
        if TIE_LANE7.exists():
            cmd.append(str(TIE_LANE7))
        if TIE_LANE8.exists():
            cmd.append(str(TIE_LANE8))
        if TIE_LANE9.exists():
            cmd.append(str(TIE_LANE9))
        if TIE_LANE10.exists():
            cmd.append(str(TIE_LANE10))
        if TIE_LANE11.exists():
            cmd.append(str(TIE_LANE11))
        if TIE_LANE12.exists():
            cmd.append(str(TIE_LANE12))
        if TIE_LANE13.exists():
            cmd.append(str(TIE_LANE13))
        if TIE_LANE14.exists():
            cmd.append(str(TIE_LANE14))
        if TIE_LANE15.exists():
            cmd.append(str(TIE_LANE15))
        cmd.append("--include-all")
        if PBC_INDEX.exists():
            cmd += ["--pbc-index", str(PBC_INDEX)]
        run("Step 7: exceptions report", cmd)
    else:
        print("(skipping exceptions report)")

    # Append to audit log
    counts = {}
    for lane_file in [TIE_LANE1, TIE_LANE2, TIE_LANE3, TIE_LANE4]:
        if lane_file.exists():
            records = json.loads(lane_file.read_text(encoding="utf-8"))
            for r in records:
                counts[r["status"]] = counts.get(r["status"], 0) + 1

    log_entry = {
        "timestamp": datetime.utcnow().isoformat(),
        "skill": "fs-detail-review",
        "action": "TIEOUT_RUN",
        "fy25_pdf": str(FY25_PDF),
        "fy25_version": FY25_VERSION,
        "status_counts": counts,
        "outputs": {
            "annotated_pdf": str(ANNOTATED_PDF),
            "exceptions_xlsx": str(EXCEPTIONS_XLSX),
        },
    }
    log = []
    if AUDIT_LOG.exists():
        try:
            log = json.loads(AUDIT_LOG.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            log = []
    log.append(log_entry)
    AUDIT_LOG.write_text(json.dumps(log, indent=2, ensure_ascii=False), encoding="utf-8")

    print(f"\n=== DONE ===")
    print(f"  Annotated PDF: {ANNOTATED_PDF}")
    print(f"  Exceptions:    {EXCEPTIONS_XLSX}")
    print(f"  Status counts: {counts}")
    print(f"  Audit log:     {AUDIT_LOG}")


if __name__ == "__main__":
    main()
