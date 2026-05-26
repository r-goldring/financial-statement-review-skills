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
FY25_PDF = ROOT / "Tieout" / "v2 Tieout 5.22.2026 version" / "9. vYYYY.M.D_Acme Holdings LLC 2025 Financial Statements (DRAFT)_clean.pdf"
FY25_VERSION = "vYYYY.M.D"
TIEOUT_DIR = ROOT / "Tieout" / "v2 Tieout 5.22.2026 version"

INPUTS_JSON = WORK_DIR / "inputs.json"
TIE_LANE1 = WORK_DIR / "tie-lane1-pdf-to-bridge.json"
TIE_LANE2 = WORK_DIR / "tie-lane2-bridge-to-tb.json"
TIE_LANE3 = WORK_DIR / "tie-lane3-prior-year.json"
TIE_LANE4 = WORK_DIR / "tie-lane4-internal.json"
TIE_LANE5 = WORK_DIR / "tie-lane5-soe-rollforward.json"
TIE_LANE6 = WORK_DIR / "tie-lane6-footing.json"
TIE_LANE7 = WORK_DIR / "tie-lane7-mapping.json"
TIE_CARRY_FWD = WORK_DIR / "tie-carry-forward.json"
OCR_CACHE = WORK_DIR / "ocr-cache.json"
FN_PAGE_MAP = WORK_DIR / "fn-page-map.json"
FY24_TIEOUT_PDF = ROOT / "Prior Year Examples" / "2024" / "Tieout" / "vYYYY.M.D_Acme Holdings LLC 2024 Financial Statements Tieout (FINAL).pdf"
ANNOTATED_PDF = TIEOUT_DIR / FY25_PDF.name.replace(".pdf", "_TIEOUT.pdf")
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
    ap.add_argument("--pages", default=None, help="restrict annotator to these pages")
    args = ap.parse_args()

    WORK_DIR.mkdir(parents=True, exist_ok=True)
    TIEOUT_DIR.mkdir(parents=True, exist_ok=True)

    PY = sys.executable

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
        if cf_script.exists() and FY24_TIEOUT_PDF.exists():
            run("Step 5.7: carry-forward FY24 marks", [
                PY, str(cf_script),
                str(FY24_TIEOUT_PDF), str(FY25_PDF),
                str(TIE_CARRY_FWD),
                "--ocr-cache", str(OCR_CACHE),
            ])
        # Lane 7 (mapping completeness & reasonableness) — TB-account findings, no PDF marks
        mc_script = SCRIPT_DIR / "tie_out_mapping_completeness.py"
        if mc_script.exists():
            run("Step 5.8: lane 7 (mapping completeness)", [
                PY, str(mc_script), str(INPUTS_JSON), str(TIE_LANE7),
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
        cmd.append("--include-all")
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
