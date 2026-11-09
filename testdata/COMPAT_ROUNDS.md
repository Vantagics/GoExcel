# Excel compatibility rounds

Differential checks compare goexcel with openpyxl, Apache POI, and LibreOffice.
Run them with `make diffcheck` (tools are optional; missing ones are skipped).
POI jars are not committed. Point `POI_JAR_DIR` at a directory of POI 5.4 jars,
or leave the default `/tmp/poi-jars`.

A row is a **defect** when goexcel disagrees with openpyxl or POI on sheet
names, values, formulas, merges, defined names, sheet state, comments, or
hyperlinks, or when LibreOffice cannot convert the file. LibreOffice value
drift (pivot refresh, German UI labels, `Err:502` versus Excel errors) is
recorded as a reference gap, not a defect. openpyxl and POI cannot read strict
OOXML; goexcel can.

## Round 1 — ClosedXML corpus and harness

Samples added under `testdata/excel/r1/` (ClosedXML test resources, MIT):

data validation, comments, sparklines, external sparkline references, Excel
2010 functions, cell borders, strict OOXML, data tables, pivot tables, charts,
row/column styles, tables, escaped apostrophes, sheet names with `!`, a dialog
sheet, sheets without relationship ids, chartsheet+pivot, ceiling/floor, and
sorting with time spans. 22 files. The original 58 fixtures stay in place.
Total compared this round: 80 workbooks.

### Before

| Check | Result |
| --- | --- |
| openpyxl read | Time spans became datetimes (`0:12:15` read as `1899-12-30`). Shared hyperlink `A3:A5` kept only `A3`. Single-cell merge `A6` dropped. Data tables (`t="dataTable"`) kept only the cached value. Strict files still read. |
| POI read | Same time-span and data-table gaps. POI's 1900 window is one day later than the 1899-12-30 epoch for serials before 1900-03-01. |
| Round-trip | Strict packages rewrote a second default `xmlns` and POI/openpyxl rejected them. Dialog sheets were saved with a chartsheet relationship. Pivot caches stored outside `xl/` were dropped. Preserved `x:tableStyles` lost its prefix. Empty cached strings on formulas were written as numeric, so POI reported `0`. Data-table string caches lost `t="str"`. |

### Fixed

- Time-only formats (`h:mm:ss`, `[h]:mm:ss`, built-ins 18–21 and 45–47) stay numeric.
- Hyperlinks apply to every cell of the reference, capped at 4096 cells.
- Single-cell merges (`ref="A6"`) are kept.
- Data tables round-trip as `t="dataTable"` and expose `TABLE(r1[,r2])`.
- Strict rewrites keep one default namespace and the transitional `r` prefix.
- Package parts whose names contain `pivot` are preserved, including caches outside `xl/`.
- Dialog/chartsheet relationship types are preserved.
- Original `styleSheet` namespace declarations are copied so prefixed `tableStyles` stay bound.
- Empty string formula caches keep `t="str"`.

### After

`python3 scripts/diffharness.py --engines openpyxl,poi --roundtrip testdata/excel`: **80 files, 0 defects**.

LibreOffice Calc 24.2 headless converts the set. Remaining LO differences are recalculated pivots, localized captions (`Gesamtergebnis`), and LO error codes. openpyxl/POI still do not read strict OOXML (goexcel does). A few originals point at a missing `xl/sharedStrings.xml`; that is in the source package.

## Round 2 — Apache POI validation, lookup, and external links

Samples in `testdata/excel/r2/` (Apache-2.0): AverageTaxRates, DataValidationEvaluations, FillWithoutColor, FormatChoiceTests, HeaderFooterComplexFormats, HsGetVal, NewStyleConditionalFormattings, StructuredReferences, Tables, WithTwoCharts, xlookup, atp, GroupTest, link-external-workbook-a, workbookProtection-sheet_password-2013. `protected_passtika.xlsx` was downloaded and rejected: it is an OLE/encrypted package (`zip: not a valid zip file`), not OOXML, so it is not committed.

### Before

Time-of-day cells (`14:00:00`, `08:00:00`) were Excel serials in goexcel and clock strings or `datetime.time` in openpyxl, so the harness reported value mismatches on DataValidationEvaluations.

### Fixed

The harness treats `HH:MM:SS` and `datetime.time` as day-fractions, matching the serial goexcel stores for time-only formats. Encrypted/OLE packages are reference gaps.

### After

openpyxl + POI read and round-trip: **0 defects** on this batch.

## Round 3 — LibreOffice Calc QA workbooks

Samples in `testdata/excel/r3/` from LibreOffice `sc/qa/unit/data/xlsx` (MPL-2.0): external charts and refs, named sheet views, pivot cache sync, sparklines, empty table headers, autofilter named ranges, column styles, and conditional-format listeners. 18 files.

### Before

`lo-MissingPathExternal.xlsx` stores times of day as serials (`06:30:00`, `00:00:18`). The first comparison treated openpyxl's clock text as a different value.

### Fixed

Covered by the clock-fraction comparison from round 2. Re-run after that change.

### After

openpyxl + POI read and round-trip: **0 defects**. External chart parts stay in the package.

## Round 4 — ClosedXML hyperlinks, validation, and rich text

Samples in `testdata/excel/r4/` (MIT): hyperlinks (including `mailto:` and file targets), sheet and workbook protection, date/decimal/text/time/whole-number validation, external links, kanji rich text with newlines, and validation formulas that reference another sheet.

### Before

POI rewrote `D:/Test.xlsx` as `D:\Test.xlsx`. A shared string whose text is `=A*B` was treated as a formula by the harness.

### Fixed

Link comparison treats `\` and `/` as the same path. Only cells whose type is formula (or an array/data-table formula object) are compared as formulas. The package target is preserved as stored.

### After

openpyxl + POI read and round-trip: **0 defects**.

## Round 5 — protection, comments, and data tables

Samples in `testdata/excel/r5/` (ClosedXML, MIT, several produced by LibreOffice): sheet passwords, SHA-512 and structure protection, one- and two-variable data tables, comments with color indexes, comment insets, and a text-box hyperlink.

### Before

`cx-one-variable.xlsx` stores the label `=A*B` as a shared string. The harness treated every string that starts with `=` as a formula, so C1 disagreed.

### Fixed

Formula comparison requires a formula cell type (or an array/data-table formula object). The string `=A*B` stays text, matching the package.

### After

openpyxl + POI read and round-trip: **0 defects**.
## Round 6 — fonts, fills, borders, and theme colors

Samples in `testdata/excel/r6/` (ClosedXML, MIT): alignment, borders, fills, fonts, number formats, row/column styles, phonetics, quote-prefix, text rotation, theme colors, merged-cell borders, and differential fills.

### Before

No value, formula, merge, name, comment, or hyperlink defects against openpyxl or POI. Theme colors and dxf fills are carried in `styles.xml`, which round-trips with the original namespace declarations.

### Fixed

None required for cell content. Preserved style parts from round 1 already cover this batch.

### After

openpyxl + POI read and round-trip: **0 defects**.
## Round 7 — tables, headers, and conditional-format order

Samples in `testdata/excel/r7/` (ClosedXML, MIT): inserting and using tables, table column styles, total-row shared strings, headers with line breaks, custom table themes, empty tables, conditional-formatting order, column styles, and a print area that points at another workbook.

### Before

No content defects. Table XML and print-area defined names are preserved as package parts and defined names.

### Fixed

None required.

### After

openpyxl + POI read and round-trip: **0 defects**.
## Round 8 — larger formula and conditional-format workbooks

Samples in `testdata/excel/r8/` (Apache POI, Apache-2.0): ConditionalFormattingSamples (about 650KB), number-format approximations, text formats, structured table references, formula evaluation, shared formulas set blank, and bug69769.

### Before

POI returns an empty string for `Tabelle1!B2` in `poi-bug69769.xlsx`. The shared-string index in the file is `2 ` with a trailing space. goexcel trims the index and reads `Mustermann`, which matches openpyxl.

### Fixed

Recorded as a POI gap, not a goexcel defect. The reader keeps trimming numeric indexes.

### After

openpyxl + POI read and round-trip: **0 defects**. One POI gap on the padded shared-string index.
## Round 9 — macros, templates, embeddings, and protection

Samples in `testdata/excel/r9/` (Apache POI, Apache-2.0): `SimpleMacro.xlsm`, `test.xltx`, an embedded package, a picture plus shape, comments, workbook and sheet passwords, a clone-sheet workbook, and a header containing `&`.

### Before

No cell-content defects. The macro binary and embeddings have to survive as package parts or Excel/POI repair the file.

### Fixed

None required beyond the part-preservation from round 1. `vbaProject.bin` and embeddings stay in the zip, and the workbook content type stays macro-enabled when that part is present.

### After

openpyxl + POI read and round-trip: **0 defects**.
## Round 10 — format, date, and header regression set

Samples in `testdata/excel/r10/` (Apache POI, Apache-2.0), prefixed `poi-` so they do not collide with the first corpus: general and elapsed formats, `FormatKM`, decimal formats, indexed colors, date-format numbers, a SUMPRODUCT with unary minus, newlines in formulas, booleans, strict-adjacent `SampleSS`, various data, and a complex header/footer.

### Before

No new content defects. These files repeat feature areas from earlier rounds (dates, elapsed time, booleans, headers) on fresh copies.

### Fixed

None required. They stay in the tree as regression files for `make diffcheck`.

### After

openpyxl + POI read and round-trip: **0 defects** across r2–r10 (122 files in that run) after the round 2–5 comparison fixes.

## Remaining known gaps

- Strict OOXML is read by goexcel. openpyxl and POI refuse it (POI bug 57699). Rewritten files are transitional OOXML.
- Encrypted workbooks (OLE compound, for example `protected_passtika.xlsx`) are not OOXML zip packages and are not opened.
- LibreOffice recalculates pivots, localizes captions, and reports `Err:502` for some Excel errors. Conversion succeeds; those drifts are not treated as goexcel defects.
- POI drops a shared-string index that contains trailing whitespace (`<v>2 </v>`). goexcel and openpyxl read the string.
- Slicers, sparklines, and pivot layouts are preserved as package parts, not edited in the cell model.
- Dialog sheets and chartsheets are kept and reattached with their original relationship types. They have no grid cells.
- Array-formula anchors are stored on the anchor cell. POI also repeats the formula on members; openpyxl does not. Both match after the anchor check.
## Later rounds

Filled in as each round lands.
