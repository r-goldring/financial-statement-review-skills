"""Shared helpers for the four tie-out lane scripts.

Provides:
  - parse_value(cell): turn a string/number cell into a float (handles $K formats,
    parens-for-negative, dashes-for-zero, etc.)
  - tolerance_for(comparison_kind, is_subtotal): tolerance per the plan's table
  - compare(left_value, right_value, comparison_kind, is_subtotal): returns
    (delta_in_comparison_unit, status, tolerance_used)
  - normalize_label(s): for fuzzy row-label matching
  - find_table_value(table_rows, label, year_col_idx): given .docx-shaped rows,
    find the row matching a label and return the value at year_col_idx
  - LABEL_ALIASES: known label variants between PDF/bridge/TB
"""

import re
from difflib import SequenceMatcher


# ---------- Parsing values ----------

_NUMBER_RE = re.compile(r"^\s*\(?\$?\s*-?\s*([\d,]+(?:\.\d+)?)\s*\)?\s*$")
_LITERAL_ZERO_DASHES = {"-", "—", "–"}  # literal "no activity" → 0.0


def parse_value(cell):
    """Return float value of a cell, or None if cell has no numeric content.

    Handles:
      - Numbers: 12017, 12017.5, "12017"
      - Formatted: "$X,XXX.XX", "NN,NNN", "$ NN,NNN "
      - Negatives: "($X,XXX.XX)", "-12017", "$($X,XXX.XX)", "$ ($X,XXX.XX)"
      - Literal zero dashes: "-", "—" → 0.0 (used in FS for "no activity")
      - Empty / None / "$" alone / label-only → None (no value present)
    """
    if cell is None:
        return None
    if isinstance(cell, (int, float)):
        return float(cell)
    s = str(cell).strip()
    if not s:
        return None
    if s in _LITERAL_ZERO_DASHES:
        return 0.0
    if s in ("$", "$ ", "($)"):
        return None
    # Detect parens-negative — both "(N)" and "$(N)" / "$ (N)" forms
    stripped_for_paren = s.replace("$", "").replace(" ", "")
    is_neg = stripped_for_paren.startswith("(") and stripped_for_paren.endswith(")")
    # Strip currency, parens, spaces
    cleaned = s.replace("$", "").replace("(", "").replace(")", "").replace(",", "").strip()
    if cleaned in ("-", "—", "–", ""):
        return 0.0 if s in _LITERAL_ZERO_DASHES else None
    try:
        v = float(cleaned)
        if is_neg:
            v = -abs(v)
        return v
    except ValueError:
        return None


# ---------- Tolerance matrix ----------
# All tolerances are EXPRESSED IN THE COMPARISON UNIT.
# After converting both sides to the same unit, abs(delta) is compared to tolerance.

TOLERANCE_MATRIX = {
    # (pdf_unit_norm, source_unit_norm) → (simple, subtotal) tolerance in COMPARISON unit
    # PDF is always $K (display).
    ("$K", "$K"):   (1.0, 5.0),   # both display
    ("$K", "$1"):   (1.0, 5.0),   # convert source $1 -> $K, allow rounding loss
    ("$1", "$K"):   (0.0, 0.0),   # never; we never compare in $1 if source is $K
    ("$1", "$1"):   (0.0, 1.0),   # both raw, should tie to penny
}


def normalize_unit(unit):
    """Reduce a unit label to one of: '$K', '$1', or 'ambiguous'."""
    if not unit:
        return "ambiguous"
    if unit == "$1-with-$K-column":
        return "$1"  # primary column is raw; consumer should opt-in to the $K column
    if unit in ("$K", "$1"):
        return unit
    return "ambiguous"


def compare(left_value, right_value, left_unit, right_unit, is_subtotal=False,
            kind=None):
    """Compare two values that may be in different units.

    Returns (delta, status, tolerance, comparison_unit) where:
      - comparison_unit is the unit both were converted to ($K or $1)
      - delta is left - right in comparison_unit
      - tolerance is the threshold used
      - status in {"ties", "ties-with-rounding", "exception", "missing"}
    """
    if left_value is None or right_value is None:
        return None, "missing", None, None

    left_u = normalize_unit(left_unit)
    right_u = normalize_unit(right_unit)

    # Decide comparison unit: if one side is $K (PDF face), compare in $K
    if "$K" in (left_u, right_u):
        comparison_unit = "$K"
        l = left_value if left_u == "$K" else left_value / 1000.0
        r = right_value if right_u == "$K" else right_value / 1000.0
    elif left_u == "$1" and right_u == "$1":
        comparison_unit = "$1"
        l = left_value
        r = right_value
    elif "ambiguous" in (left_u, right_u):
        return None, "ambiguous-units", None, None
    else:
        comparison_unit = "$1"
        l = left_value
        r = right_value

    delta = l - r
    abs_delta = abs(delta)

    # Special kinds
    if kind == "prior_year":
        tolerance = 0.0
    elif kind == "internal":
        tolerance = 0.0 if not is_subtotal else 5.0
    else:
        key = (left_u, right_u)
        if key in TOLERANCE_MATRIX:
            simple, subtotal = TOLERANCE_MATRIX[key]
        else:
            simple, subtotal = (1.0, 5.0)
        tolerance = subtotal if is_subtotal else simple

    # Sign-inversion check: if abs values are equal (within tolerance), it's likely
    # an FS-vs-bridge sign-convention difference (FS shows "loss" as positive in
    # "Comprehensive loss" lines, bridge stores it as negative).
    abs_match_delta = abs(abs(l) - abs(r))
    sign_inverted = (l != 0 and r != 0 and (l > 0) != (r > 0))

    if abs_delta == 0:
        status = "ties"
    elif abs_delta <= tolerance:
        status = "ties-with-rounding"
    elif sign_inverted and abs_match_delta <= tolerance:
        status = "ties-with-sign-inversion"
    else:
        status = "exception"

    return delta, status, tolerance, comparison_unit


# ---------- Label normalization & matching ----------

# Punctuation/whitespace stripped, lowercased
_LABEL_STRIP = re.compile(r"[^a-z0-9]+")

# Hand-tuned aliases — known label variants. Direction: (PDF label) → (canonical key)
LABEL_ALIASES = {
    # BS items
    "cash and cash equivalents": "cash",
    "cash equivalents": "cash",
    "accounts receivable, net": "accounts_receivable",
    "accounts receivable": "accounts_receivable",
    "prepaid expenses and other current assets": "prepaid",
    "prepaid expenses": "prepaid",
    "deferred commissions, current": "deferred_commissions_current",
    "current deferred commissions": "deferred_commissions_current",
    "total current assets": "total_current_assets",
    "property and equipment, net": "ppe_net",
    "property and equipment": "ppe_net",
    "right-of-use assets": "rou_assets",
    "deferred commissions, noncurrent": "deferred_commissions_noncurrent",
    "noncurrent deferred commissions": "deferred_commissions_noncurrent",
    "intangible assets, net": "intangibles_net",
    "intangible assets": "intangibles_net",
    "goodwill, net": "goodwill",
    "goodwill": "goodwill",
    "other noncurrent assets": "other_noncurrent_assets",
    "other assets": "other_noncurrent_assets",
    "total assets": "total_assets",
    "accounts payable": "accounts_payable",
    "accrued expenses and other current liabilities": "accrued_expenses",
    "accrued expenses and other current liabilites": "accrued_expenses",  # known typo
    "lease liabilities, current portion": "lease_liab_current",
    "current lease liabilities": "lease_liab_current",
    "deferred revenue": "deferred_revenue",
    "convertible notes payable - current": "convertible_notes_current",
    "convertible notes payable, current": "convertible_notes_current",
    "convertible notes": "convertible_notes_current",
    "notes payable": "notes_payable",
    "total current liabilities": "total_current_liabilities",
    "notes payable, net of current portion": "notes_payable_noncurrent",
    "lease liabilities, net of current portion": "lease_liab_noncurrent",
    "noncurrent lease liabilities": "lease_liab_noncurrent",
    "other non-current liabilites": "other_noncurrent_liabilities",  # known typo
    "other noncurrent liabilities": "other_noncurrent_liabilities",
    "total liabilities": "total_liabilities",
    "total members' equity": "total_equity",
    "total members equity": "total_equity",
    "total liabilities and members' equity": "total_liab_and_equity",
    "accumulated deficit": "accumulated_deficit",
    "accumulated other comprehensive income (loss)": "aoci",
    "accumulated other comprehensive loss": "aoci",
    "additional paid-in capital": "apic",
    # IS items
    "revenue": "revenue",
    "cost of revenue (exclusive of depreciation and amortization)": "cogs",
    "cost of revenue": "cogs",
    "gross profit": "gross_profit",
    "sales and marketing": "sm",
    "research and development": "rd",
    "general and administrative": "ga",
    "depreciation and amortization": "da",
    "change in fair value of contingent consideration": "fv_cc",
    "total operating expenses": "total_opex",
    "loss from operations": "loss_from_operations",
    "income from operations": "loss_from_operations",
    "change in fair value of related party convertible notes": "fv_convertible",
    "interest expense": "interest_expense",
    "other income (expense), net": "other_income_net",
    "other income, net": "other_income_net",
    "total other expense": "total_other_expense",
    "total other income (expense)": "total_other_expense",
    "loss before income taxes": "loss_before_tax",
    "income before income taxes": "loss_before_tax",
    "income tax expense": "tax_expense",
    "income tax (expense) benefit": "tax_expense",
    "net loss": "net_loss",
    "net income": "net_loss",
    "foreign currency translation adjustment": "fx_translation",
    "comprehensive loss": "comprehensive_loss",
    "comprehensive income": "comprehensive_loss",
    # SCF items
    "net loss": "net_loss",
    # SOE items
}


def normalize_label(label):
    """Return canonical key for a label, or fallback to stripped-lower form."""
    if not label:
        return ""
    s = str(label).strip().lower()
    # Strip any leading prefixes like "    " from indentation
    if s in LABEL_ALIASES:
        return LABEL_ALIASES[s]
    # Try fuzzy match against aliases (handle whitespace and minor punct diffs)
    s_norm = _LABEL_STRIP.sub("", s)
    for key, val in LABEL_ALIASES.items():
        if _LABEL_STRIP.sub("", key) == s_norm:
            return val
    return _LABEL_STRIP.sub("_", s).strip("_")


def fuzzy_match(needle, haystack, threshold=0.85):
    """Return haystack item with highest similarity to needle (above threshold), else None."""
    needle_n = _LABEL_STRIP.sub("", needle.lower())
    best = None
    best_score = 0
    for h in haystack:
        h_n = _LABEL_STRIP.sub("", h.lower())
        score = SequenceMatcher(None, needle_n, h_n).ratio()
        if score > best_score:
            best_score = score
            best = h
    if best_score >= threshold:
        return best, best_score
    return None, best_score


# ---------- Subtotal/total detection ----------

SUBTOTAL_KEYWORDS = [
    "total", "subtotal", "gross profit", "loss from", "income from",
    "net loss", "net income", "comprehensive loss", "comprehensive income",
    "loss before", "income before", "ending", "balance as of", "balance at",
]

def is_subtotal_label(label):
    if not label:
        return False
    s = str(label).lower()
    return any(kw in s for kw in SUBTOTAL_KEYWORDS)


# ---------- Common record schema ----------

def make_record(lane, **kwargs):
    """Standard tie record schema."""
    rec = {
        "lane": lane,
        "pdf_page": None,
        "pdf_section": None,
        "pdf_label": None,
        "pdf_year": None,
        "pdf_value": None,
        "source_ref": None,
        "source_label": None,
        "source_value": None,
        "comparison_unit": None,
        "delta": None,
        "tolerance": None,
        "status": None,
        "is_subtotal": False,
        "notes": None,
    }
    rec.update(kwargs)
    return rec
