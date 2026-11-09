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
## Later rounds

Filled in as each round lands.
