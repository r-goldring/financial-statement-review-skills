# Tie-out Conventions (mirror from FY24 final)

Source artifacts inspected:
- `Prior Year Examples/2024/Tieout/vYYYY.M.D_Acme Holdings LLC 2024 Financial Statements Tieout (FINAL).pdf` (primary)
- `Prior Year Examples/2024/Tieout/vYYYY.M.D Acme Holdings LLC 2024 Financial Statements Tieout.pdf` (earliest draft, for progression)
- `Prior Year Examples/2024/Tieout/vYYYY.M.D_Acme 2024 FS Bridge (DRAFT).xlsx` (bridge structure)

Renders cached at `.work/py24-final-renders/page-NN.png` and `.work/py24-v522-renders/page-NN.png`.

## What this is

Spec for the marks FS Compilation Partner / Acme Corp use to annotate the rendered FS PDF and the bridge workbook during tie-out. Used by `annotate_tieout_pdf.py` to draw matching marks on the FY25 clean PDF.

## PDF: which pages get marks

| Pages | Content | Annotated? |
|---|---|---|
| 1 | Cover | No |
| 2 | TOC / Index | No |
| 3 | Independent Auditor's Report | No |
| 4 | Notes to FS (blank starter) | No |
| 5 | **Balance Sheet** | Yes — every $ value |
| 6 | **Statements of Operations & Comprehensive Loss** | Yes — every $ value |
| 7 | **Statements of Changes in Members' Equity** | Yes — every $ value |
| 8 | **Statement of Cash Flows** | Yes — every $ value |
| 9+ | Footnotes (narrative + tables) | Tables yes, narrative no |

Rule: a page is annotated iff it contains a dollar-value table or a referenced accounting policy paragraph.

## Mark vocabulary

All marks are drawn in **red** (some sources use a related red-purple for source-tags but everything reads as red). Marks are overlaid on the printed PDF — they do not replace text. Native PDF content remains black.

### Column-level tags (placed at top of a value column, with a red downward arrow ↓ pointing at the column)

| Mark | Meaning | Example placement |
|---|---|---|
| `FS Bridge` | Every value in this column ties to the FS Bridge workbook ("Rounded (000s)" column of the corresponding tab) | Top of current-year column on BS / IS / SCF / SOE |
| `PY` | Every value in this column ties to the prior-year final FS (no restatement) | Top of prior-year column on BS / IS / SCF / SOE |
| `<source-name>` | Whole table sourced from named workpaper. Used on footnote tables. | `Tax Provision`, `Goodwill & Intangibles RF`, `Intangibles RF`, `Lease Schedule`, `Cap Table`, `LLC Agreement`, `Updated 409A Valuation Firm Convertible Notes Valuation Report`, `Updated AcquiredCo A Earnout Valuation Report`, `SBC - <filename>` |
| `Rev Rec Policy`, `Cap Commission Policy` | Accounting policy paragraph identified | Bracket spans the relevant paragraph in Note 1 |

### Cell-level tickmarks (placed right of or below the individual number)

| Mark | Meaning |
|---|---|
| `F` | Footed — subtotal verified by horizontal sum of components |
| `V` | Vertical foot — column total verified by vertical sum |
| `<a>`, `<b>`, `<c>`, `<d>` | Lettered footnote tickmark; text-explanation written at page bottom in red. Reserved for *prior-year restatements / reclassifications* (e.g., AcquiredCo A acquisition valuation adjustments) |
| `FN N` (e.g., `FN 3`, `FN 9`) | Ties to footnote N within the FS PDF |
| `/ <ref>` | Ties to another statement or tab — e.g., `NN,NNN / SoCF` means this number equals the same value on the Statement of Cash Flows |
| `GL NNNNNN` | Ties to a specific GL account in the TB; sometimes prefixed `PY GL NNNNNN` for prior year |
| `PL` | Ties to the Income Statement (PL tab) — used on SCF where IS Net Loss flows in |
| `SoSE` | Ties to Statement of Changes in Members' Equity — used on BS where equity components flow |
| `SoCF` | Ties to Statement of Cash Flows — used on BS Cash and Other items |
| `Cap Table` | Ties to the equity unit-count capitalization table |

### Footnote text (bottom of page, in red)

When a `<a>`, `<b>`, `<c>`, `<d>` mark is placed on a number, an explanatory line is added at the bottom of the same page in red, e.g.:
```
<a>  The Company adjusted three prior year amounts to reflect changes related to valuation
     for the AcquiredCo A acquisition and convertible notes
```

## Mark placement (geometric)

- **Column tags** (`FS Bridge`, `PY`, `<source>`): top of column, slightly above the column header, with a downward arrow (red) along the column's left edge for the height of the table.
- **Cell tickmarks** (`F`, `V`, `FN N`, `/ ref`, `GL N`): immediately to the right of the value, in red. If multiple marks on one number, stack them or separate with spaces.
- **Bracket marks** (around grouped sections): red square bracket [ ] enclosing the section; tag appears at the top-left of the bracket.
- **Footnote text** (`<a>` explanation): page bottom, left-aligned, in red.

## Bridge workbook conventions (in-cell, on the .xlsx itself)

The bridge already encodes its own tie-out logic in column structure:

| Column | Unit | Purpose |
|---|---|---|
| `Per TB` | $1 raw | The unadjusted Trial Balance value pulled directly from NetSuite |
| `Topside Entries` (a, e, f...) | $1 raw | Adjusting entries posted on top of TB; `{a}`, `{b}`, etc. are bracketed-letter footnote refs pointing to entries-detail elsewhere in the workbook |
| `Adjusted Total` | $1 raw | `= Per TB + Topside Entries` |
| `Rounded (000s)` | $K | `= round(Adjusted Total / 1000)` — **this is the column that should match the PDF face statement values** |
| `As of <prior-year date> as filed` | $K | Prior-year filed figure |
| `Misstatement Adjustments` | $K | Prior-year corrections (with `{b}` footnotes) |
| `Adjusted <prior-year date>` | $K | `= As filed + Misstatement Adjustments` |

The literal word `rounding` appears in the workbook when a rounding adjustment is made (e.g., a $X,XXX.XX topside entry labeled `rounding` to make Other Assets round cleanly to $355K).

When mapping PDF → bridge:
- **For current-year**: compare PDF value to bridge `Rounded (000s)` column
- **For prior-year**: compare PDF value to bridge `Adjusted <prior-year date>` column
- **For tie-out audit trail**: also show the bridge `Per TB` + `Topside Entries` so the user can see the build-up

## Color palette (for `annotate_tieout_pdf.py`)

- Red `#D52027` — primary annotation color (all marks, all text, all arrows)
- Red-purple `#7B2C5C` — used sparingly in FY24 final for `FS Bridge` tag, but can be subsumed into the primary red
- Black `#000000` — never used for annotations (that's the printed PDF content)

Font: a thin sans-serif at ~9pt for inline marks, ~10pt for column tags. Handwriting-looking (some FY24 marks look hand-drawn) but a clean sans is acceptable since we're programmatic.

## Density / progression (FY24 v5.22 → FINAL)

Comparing the v5.22 draft IS page to the v5.30 FINAL IS page:
- **Convention** is identical — same marks, same placement
- **Differences are in the underlying numbers** — e.g., tax provision $(230)K in v5.22 → $(205)K in FINAL; net loss $($X,XXX.XX)K → $($X,XXX.XX)K
- **Implication**: FS Compilation Partner annotates the FS PDF early in the tie-out process and the marks travel with each new draft. Our annotator can run on any draft of the FY25 PDF; on a re-run with a revised draft, we reapply the same conventions.

## What we are skipping (intentionally)

- **Hand-written marks vs. typed marks**: the FY24 marks may have been hand-drawn or typed via Bluebeam / Acrobat annotation tools. We will use programmatic PyMuPDF placement, which is closer to typed than hand-drawn but visually equivalent for tie-out purposes.
- **Multi-color convention beyond red/black**: only the standard red is needed.
- **Bridge workbook annotation in this pass**: in-cell bridge marks (cell color, comments) are not being applied by our annotator. The bridge already self-documents through its column structure (`Per TB` → `Topside Entries` → `Rounded (000s)`), which is sufficient for our tie-out logic.
