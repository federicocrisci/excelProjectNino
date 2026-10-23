# Independent mapping analysis

Input: `44,1 26-30ott.xlsx`, sheet `Canvass 44 26-30ott`, 6957 rows × 18 columns.
Reference: `New Grieco N°44.xls`, sheet `CANVAS`, 6619 rows × 13 columns.

## Counts and exact mapping evidence

- Input has 6237 real item rows, marked B=`I`, with C code, E units/carton, F cartons/pallet, G description. Numeric C + G is insufficient: footer PREMI rows6942–6952 repeat eleven codes with no B/E/F and O textual cash rebates.
- Reference has6235 coded item rows,6224 unique codes,355 category headings in A,22 description-only detail rows and7 blank/header rows.
- The model contains6235 matched entries associating target rows, source rows and categories. Every target product occurrence has a unique source occurrence. Duplicate product occurrences are intentionally retained.
- EAN, units/carton and cartons/pallet are identical for all6235 matching item occurrences.
- Target category grouping consolidates649 distinct source A labels into355 headings.19 source categories split among multiple brands (not a simple one-to-one category rename). Match established products by code + source category + EAN + source description; new ambiguous products require review.

## Columns

| Output | Source / behavior |
|---|---|
| A | Consolidated category heading, only on separate heading rows |
| B | Source R EAN, stored as text (8 or13 digits, possibly blank) |
| C | Source C article code, stored as text |
| D | Blank order quantity input |
| E | Source E units per carton |
| F | Source G description, with J/K suffix where present and occasional manual editorial changes |
| G | PROMO / NOVITA reference metadata; cannot infer fully from source |
| H | Net price after source discount |
| I | Gross price after source Q VAT |
| J | Varianza reference metadata; source mostly lacks previous price |
| K | CxS external stock data, not available in source; cannot infer |
| L | Source F cartons per pallet, exact match on all occurrences |
| M | Formula D × E × H |

## Prices and formulas

Numeric source P is fractional discount: net≈O×(1−P), TP means unchanged. VAT≈net×(1+Q/100). H/I are static values in original reference, not formulas. The only reference formulas are in M:6235 line totals plus M2 subtotal.

Reference subtotal M2=`SUM(M3:M6324)` erroneously excludes rows6325–6619. All product totals are `Drow*Erow*Hrow` (shared formulas in BIFF blocks). Correct new total should include final row.

Rounding is not a single straightforward decimal rule. Python binary `round(O*(1-P),2)` matches6044/6230 comparable normal/reference-priced occurrences.185 discrepancies are exact half-cent ties; one is an actual anomaly. Exact decimal HALF_EVEN matches5972/6230; exact HALF_UP5818/6230. Gross discrepancies under Python binary round are all exact half-cent ties. Tested multiple algebraic binary variants, none fully match. There are1009 unique `(source O, numeric discount or0)` pairs and608 unique `(target net,VAT)` pairs, each internally consistent. For exact unchanged-source reproduction, reference values can be reused under a signature guard. New prices should use a stated conventional currency rounding rule and be reviewed.

**Clear price anomaly:** code26794 GILLETTE ESPO RICARICA TRIPACCO CM: source O22.45/TP, reference H224.45 (tenfold), I273.83. Correct source-based net22.45 and VAT27.39 unless user explicitly confirms manual override.

## External/manual metadata

- K CxS has2246 nonzero quantities, including arbitrary inventory-like values (419,722,101 etc.). None are present in source. They must be blank, supplied externally, or explicitly copied as old inventory with a warning.
- Product G flags:582 PROMO and9 NOVITA; additionally9 display-composition rows have PROMO. Source formatting and values do not encode all flags.559/591 total PROMO rows have no source M note. Both discounted/TP rows have mixed blank and PROMO flags.
- J has37 nonempty variance strings; only6 have source EX previous-price notes. Most prior prices are absent; no reliable general rule.
- Source M NOVITA annotations do not directly match target NOVITA: e.g. new GLADE has source NOVITA but target blank. Some section-header novelty notes propagate to7 target products. Known metadata can be reused only for unchanged signatures; new/current flags require review.

## Omaggio / special rows

There are12 source SM item rows; every source O is blank. Ten are retained:

- Three target prices remain blank:29015 SACCA ADIDAS,13637 PANNO QUASAR,18350 CLINIANS.
- Five gifts receive the same net price as a normal code elsewhere:18152 .60;17711 .50;28784/28785/28786 1.50.
- Two target gift prices are wholly external/manual:28814 SANTEX1.00 and8534 PIASTRA10.00.
- Source row684/code27469 `+1CT DI` is omitted as a separate gift and incorporated in preceding18152's `BASE750ML+RIC.` gift description. Its normal product occurrence remains.
- Source row982/code36 `IN OMAGGIO...PHON PROFESSIONALE2200w` is completely omitted. The related `ACQUISTANDO20CT` condition also disappears.

Ten other ACQUISTANDO conditions survive (source G), attached before relevant product subsets.

## Description differences

Only15 product descriptions differ from plain source G: three suffix joins (29065+K200CC,28187+Jpaletta...,20215+KSINGOLO), ten gift rewrites, typo ACTIFF BRILANTANTE→BRILLANTANTE code23891, and SAUGELLA adds INTIMO code27570. Source data should win for changed descriptions; known editorial forms can be reused under guard.

## Display composition rows

The22 no-code detail rows include10 purchasing conditions plus12 composition lines:3 ARIASANA,3 BOROTALCO,2 GILLETTE,4 PARODONTAX. Preserve them for unchanged reference-related products; do not use them as independent orderable items. Input-premium footer11 lines are not exported in reference.

## Suggested behavior for next workbook

Use new uploaded data for prices/EAN/quantities/descriptions, retain catalog category/order/style matches for known unchanged identities, keep new products instead of dropping them, and visibly report unknown categories, missing VAT/price, manual metadata and excluded special rows. Never carry stale stock silently. Recreate output as XLSX preserving reference style; workbook user can input D quantities and totals recalculate.

## Rebuilding the offline reference model

`tools/build_template.py` is a development utility. It requires Python3, openpyxl and xlrd. End users only open the included HTML file; no Python, server, installation or Internet connection is required.

Run with explicit paths to the original source XLSX, worked reference XLS, and a LibreOffice-converted reference XLSX:

```sh
python tools/build_template.py --source /path/original.xlsx --reference /path/worked.xls --converted /path/converted/worked.xlsx
```

The builder recreates occurrence matching rather than relying on scratch data. It validates all matching products, fails on missing source occurrences, and rejects conflicting price lookup pairs. It strips shared strings, stale external workbook references, broken defined names and reference author metadata from the template ZIP. Fonts, fills, borders, alignment, number formats, raw style indices, column widths, row heights, print settings and print headers are retained. Every reference row has13 raw style indices and its original row attributes.

Generated `template.js` assigns `window.CANVAS_MODEL`. The model contains6619 rows,6235 source occurrence associations,2 explicitly excluded source signatures,1008 consistent net rounding pairs,607 consistent gross rounding pairs, generic styles and a record of the code26794 price anomaly. This model is necessary for offline layout and catalog/category reconstruction. It contains data from the provided examples and should be shared only with intended recipients.

Validation performed: all remaining XML and relationships in the template parse successfully; externalLinks/sharedStrings are absent; all6235 source occurrence matches and all13 styles per row are present. Production conversion and browser verification are separate checks documented by the converter's test suite.

The model also retains23 source detail records in `sourceDetails`:11 purchasing conditions and12 display-composition lines. Each contains its18 source values, source row number and parent product source row. All22 reference detail rows link to the matching original through `sourceDetailRow` and to their reference product through `parentR` (alias `detailParentR`); conditions precede their first product, compositions follow their display product. The omitted ELVIVE20-carton condition remains documented in sourceDetails but is absent from reference output. GILLETTE composition quantities are source I5 on each of rows4726–4727; descriptions are source J. Its reference-only unit-price parenthesis should not be carried forward when source composition or price changes.

Column widths are restored to the XLS reference's exact1/256-character units, correcting LibreOffice's two-decimal export rounding. All13 widths were compared to `xlrd.colinfo_map.width/256` and match exactly.
