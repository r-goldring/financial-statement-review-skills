"""Phase 0b — PBC index + classifier.

Walks the canonical 2025 PBC tree, classifies every file, and (for the FS-BUILD
subset that actually feeds the bridge / footnotes) maps it to the FS area / footnote /
bridge tab it supports and records its version / as-of date. The huge audit-EVIDENCE
subset (bank statements, AP invoices, revenue/JE samples, walkthroughs, IT) is
classified but flagged `fs_relevant=false` so it doesn't clutter the tie-out.

Emits `.work/pbc-index.json` — the substrate for the PBC Register (completeness
tracker), Lane 8 (PBC->bridge tie-out), the tax module, and the flux module.

When the structure of a 2025 workpaper is unclear, the matching file under
`PBCs/2024 Audit/` is the reference ("check 2024 first").

CLI:
  python build_pbc_index.py [--root <PBCs/2025 Audit dir>] [--out <.work/pbc-index.json>]
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

ROOT = Path(r"C:\path\to\financial-statement-review")
PBC_2025 = ROOT / "PBCs" / "2025 Audit"
DEFAULT_OUT = Path(__file__).parent.parent / ".work" / "pbc-index.json"

# The actual audit area is the path component AFTER the "Year End"/"Interim" prefix.
def get_area(rel_path):
    parts = rel_path.split("/")
    if parts and parts[0] in ("Year End", "Interim") and len(parts) > 1:
        return parts[1]
    return parts[0] if parts else ""

# Areas whose contents feed the FS/bridge build (map these). Everything else is evidence.
FS_BUILD_AREAS = {
    "fs compilation requests (connor group) (old)", "tax provision requests (tanner)",
    "tax (ye)", "goodwill (ye)", "leases (ye)", "leases (int)", "debt (ye)", "debt (int)",
    "financial statements (ye)", "financial statements (int)", "general (ye)",
}
# Areas that are purely audit testing evidence (never FS-build), regardless of keywords.
EVIDENCE_AREAS = {
    "cash (ye)", "cash (int)", "accounts payable and accrued expenses (ye)",
    "revenue (ye)", "revenue (int)", "expenses (ye)", "expenses (int)",
    "it (ye)", "it (int)", "walkthroughs and controls (ye)",
    "walkthroughs and controls (int)", "administrative", "general (int)",
}

# Filename/-path keyword -> (category, fs_area, footnote, bridge_tab). First match wins.
# fs_area/footnote/bridge_tab are the connective tissue mirrored in references/pbc-mapping.md.
FS_BUILD_RULES = [
    (r"tax provision|provision memo|provision_v|deferred tax|rate rec|163j|\bnol\b",
        ("tax", "Income taxes", "FN-07", "FN- Taxes")),
    (r"goodwill|intangible",
        ("goodwill-intangibles", "Goodwill & intangibles", "FN-03/FN-04", "FN - Intangibles / FN - PPE")),
    (r"capitalized software|cap software|internal-use software",
        ("cap-software", "Property & equipment", "FN-03", "FN - PPE")),
    (r"capitalized commission|cap commission",
        ("cap-commissions", "Deferred commissions", "BS", "BS / SCF")),
    (r"asc 842|lease schedule|lease expense|\bleases?\b",
        ("lease", "Leases", "FN-11", "FN - Leases")),
    (r"mufg|loan amort|term loan|credit agreement|effective interest|debt schedule|\bdebt\b",
        ("debt", "Term loans & line of credit", "FN-06", "FN - Debt")),
    (r"401\(?k\)?|401k",
        ("401k", "Employee benefit plan", "FN-13", "FN - 401(k)")),
    (r"ebitda",
        ("ebitda", "Non-GAAP / EBITDA", None, "EBITDA Adj Summary")),
    (r"flux|account reconciliation and flux",
        ("flux", "Analytical / flux review", None, None)),
    (r"going concern",
        ("going-concern", "Going concern", "FN-01", None)),
    (r"general ledger detail|\bgl detail\b|2025 general ledger|full gl|je detail|journal entry detail",
        ("gl-detail", "General ledger (source)", None, "2025 TB")),
    (r"trial balance|\btb by\b|consolidated tb|consolidated is by|consolidating balance sheet|consolidating income statement|detailed tb",
        ("trial-balance", "Trial balance (source)", None, "TB Mapping")),
    (r"department mapping",
        ("dept-mapping", "Department mapping (source)", None, None)),
    (r"fixed asset|\bppe\b|\bfar\b|asset rf",
        ("ppe", "Property & equipment", "FN-03", "FN - PPE")),
]

# Audit-evidence areas (NOT FS-build) — classified for the register but fs_relevant=false.
AUDIT_EVIDENCE_KEYWORDS = (
    "bank statement", "bank reconciliation", "_recon_", "sample support", "selections",
    "invoice", "bill_", "samples", "walkthrough", "process narrative", "rep letter",
    "confirmation", "je sample", "je selection", "journal entry sample", "disbursements",
    "soc1", "audit committee", "management rep",
)

VERSION_RE = re.compile(r"v?20\d{2}[._-]\d{1,2}[._-]\d{1,2}", re.I)
RECON_RE = re.compile(r"_recon_(20\d{2}-\d{2})", re.I)
ASOF_RE = re.compile(r"\b(12[._-]31[._-]\d{2,4}|3[._-]31[._-]\d{2,4})\b")


def extract_version(name):
    for rx in (VERSION_RE, RECON_RE, ASOF_RE):
        m = rx.search(name)
        if m:
            return m.group(0)
    return None


def classify(rel_path):
    """Return dict: category, fs_relevant, fs_area, footnote, bridge_tab.

    Evidence-first precedence. The FS-build subset is small and concentrated in a
    handful of areas (FS Compilation Requests (FS Compilation Partner) (old), Tax Provision
    Requests (Tax Provision Preparer), the YE FS/goodwill/leases/debt/general areas). Everything else
    — the 600+ bank statements, AP invoices, JE/revenue samples, walkthroughs, IT — is
    audit evidence and must NOT be pulled into the tie-out, even when a filename happens
    to contain an FS keyword (e.g. "Chase" is both the term-loan lender AND the company's
    bank, so Chase bank statements would otherwise mis-classify as debt).
    """
    p = rel_path.lower()
    area = get_area(rel_path).lower()

    # 1) Explicit audit-evidence filename/path markers always win.
    if any(k in p for k in AUDIT_EVIDENCE_KEYWORDS):
        return {"category": "audit-evidence", "fs_relevant": False,
                "fs_area": area or None, "footnote": None, "bridge_tab": None}

    # 2) Areas that are purely testing evidence — never FS-build, regardless of keywords.
    if area in EVIDENCE_AREAS:
        return {"category": "audit-evidence", "fs_relevant": False,
                "fs_area": area or None, "footnote": None, "bridge_tab": None}

    # 3) Inside an FS-build area: apply keyword rules to map footnote/bridge tab.
    if area in FS_BUILD_AREAS:
        for pattern, (cat, fs_area, fn, tab) in FS_BUILD_RULES:
            if re.search(pattern, p):
                return {
                    "category": cat, "fs_relevant": True,
                    "fs_area": fs_area, "footnote": fn, "bridge_tab": tab,
                }
        # In an FS-build area but no keyword matched — flag for mapping review.
        return {"category": "fs-build-unmapped", "fs_relevant": True,
                "fs_area": area, "footnote": None, "bridge_tab": None}

    # 4) Anything else (unknown/admin areas) → evidence by default.
    return {"category": "audit-evidence", "fs_relevant": False,
            "fs_area": area or None, "footnote": None, "bridge_tab": None}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", default=str(PBC_2025))
    ap.add_argument("--out", default=str(DEFAULT_OUT))
    args = ap.parse_args()

    root = Path(args.root)
    if not root.exists():
        print(f"ERROR: PBC root not found: {root}")
        sys.exit(1)

    records = []
    for f in sorted(root.rglob("*")):
        if not f.is_file():
            continue
        if f.name.startswith("~$"):  # Excel lock files
            continue
        rel = f.relative_to(root).as_posix()
        c = classify(rel)
        records.append({
            "path": str(f),
            "rel": rel,
            "area": rel.split("/")[0],
            "filename": f.name,
            "ext": f.suffix.lower(),
            "size": f.stat().st_size,
            "version": extract_version(f.name),
            **c,
            "extracted": None,  # Phase 1/2/3 populate structured values here
        })

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(records, indent=2, ensure_ascii=False), encoding="utf-8")

    # Summary
    total = len(records)
    fs_build = [r for r in records if r["fs_relevant"]]
    by_cat = {}
    for r in fs_build:
        by_cat[r["category"]] = by_cat.get(r["category"], 0) + 1
    print(f"PBC index: {total} files | FS-build: {len(fs_build)} | audit-evidence: {total - len(fs_build)}")
    print(f"FS-build by category: {by_cat}")
    print("\nFS-build PBCs (the tie-out targets):")
    for r in fs_build:
        tab = r["bridge_tab"] or "-"
        ver = r["version"] or "-"
        print(f"  [{r['category']:<18}] {r['footnote'] or '--':<8} {ver:<14} {r['rel'][:70]}")
    unmapped = [r for r in fs_build if r["category"] == "fs-build-unmapped"]
    if unmapped:
        print(f"\n  ⚠ {len(unmapped)} FS-area files not keyword-mapped (review pbc-mapping.md):")
        for r in unmapped[:15]:
            print(f"    {r['rel'][:80]}")
    print(f"\nWrote {out}")


if __name__ == "__main__":
    main()
