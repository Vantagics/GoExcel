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

## Gap closure — encryption, modeled features, array refs, strict write

### Encrypted workbooks

ECMA-376 encrypted packages (OLE compound files) open with `OpenFile(path, WithPassword(pw))` and `XLSXReader.ReadPassword`. `Workbook.SetPassword` saves Agile encryption (AES-256-CBC, SHA-512, 100000 spins), the scheme Excel 2010 and later write. Standard encryption (Office 2007 AES-128 / SHA-1, 50000 spins) is read, and the same writer is covered by `encryptStandardPackage` in tests.

Fixtures in `testdata/encrypted/` (password `s3cret`, synthetic cells only):

| File | Produced by |
|------|-------------|
| `poi-agile.xlsx` | Apache POI 5.4.0 Agile |
| `poi-standard.xlsx` | Apache POI 5.4.0 Standard |
| `msoffcrypto-agile.xlsx` | msoffcrypto-tool 6.1.0 Agile |

goexcel decrypts all three. msoffcrypto-tool and POI both decrypt a workbook goexcel saved with `SetPassword` and read `A1=encrypted-ok`. Opening an encrypted file without a password returns `ErrPasswordRequired`. A wrong password returns `ErrWrongPassword`.

### Sparklines, slicers, pivot layout

These are no longer pass-through only.

- `Worksheet.SparklineGroups` / `SetSparklineGroups` read and rewrite `x14:sparklineGroup` (type, markers, colors, formula, sqref). `testdata/excel/r3/lo-Sparklines.xlsx` round-trips; an edit of the sqref to `Z9` and a group type to `stacked` is what openpyxl, POI, and LibreOffice Calc reopen. Untouched saves keep the original sparkline XML.
- `Worksheet.PivotTables` exposes name, cache id, data caption, location ref, style, row/col/page fields, data fields, and cache field names. Editing name, data caption, and location on `ExcelPivotTableSample.xlsx` is visible to openpyxl (`_pivots` name `RenamedPivot`) and POI (`XSSFPivotTable` name `RenamedPivot`). LibreOffice Calc reopens the file and keeps the name.
- `Worksheet.Slicers` and `Workbook.SlicerCaches` expose name, caption, style, column count, row height, and cache source name. Edits are written back into the slicer parts. openpyxl, POI, and LibreOffice open the rewritten package and still read the sheet cells.

### Array formulas

`Worksheet.SetArrayFormula`, `ArrayFormulas`, and `FormulaAt`, plus `Cell.ArrayRef`, `IsArrayFormula`, and `ReportedFormula`, follow Excel: the package stores the expression on the anchor with `t="array"` and `ref`, and every cell in that ref reports the same formula. Member cells do not grow their own `<f>` element, so the differential harness still agrees with openpyxl (anchor only) and POI (member repetition is still the anchor check in `scripts/diffharness.py`). `SetFormulaArray` is unchanged for a single cell.

### Strict OOXML

`Workbook.SetStrictOOXML(true)` writes strict spreadsheet namespaces (`http://purl.oclc.org/ooxml/spreadsheetml/main`), strict office-document relationship types, strict extended properties, and `conformance="strict"`. Package content-types and package relationship namespaces stay on the transitional OPC URIs, matching Excel's own strict files such as `SampleSS.strict.xlsx`. The default save stays transitional, including when the file that was read was strict (`LoadedStrict` reports that without changing the writer). goexcel reads the strict package back. LibreOffice Calc opens it. openpyxl and POI still refuse strict packages (POI bug 57699); that is their limitation.

## Round 11 — RC4 encryption, new pivots and slicers, Calamine and XlsxWriter

### Encryption

Legacy binary RC4 (EncryptionInfo 1.1, MD5, rekey every 512 bytes) and RC4 CryptoAPI (version 4.2, SHA-1, 40-bit or 128-bit) are read. `testdata/encrypted/` (password `s3cret`):

| File | What it checks |
|------|----------------|
| `poi-binaryrc4.xlsx` | Apache POI 5.4.0 binary RC4. goexcel's plaintext matches POI's decryptor byte for byte (SHA-256 `21d0935c…`). POI's encryptor does not finish the ZIP central directory, so the plaintext is not a readable xlsx; the key schedule still matches. |
| `go-binaryrc4.xlsx` | goexcel binary RC4 of a real xlsx. POI `BinaryRC4Decryptor` reads `A1=encrypted-ok`. |
| `go-cryptoapi.xlsx` | goexcel RC4 CryptoAPI (40-bit). POI `CryptoAPIDecryptor.verifyPassword` succeeds and `getDataStream` reads `A1=encrypted-ok`. POI's CryptoAPI *encryptor* throws `IOException: not supported` for an OOXML package, so there is no POI-produced CryptoAPI xlsx. |

msoffcrypto-tool 6.1.0 rejects version 1.1 and treats version 4.2 without the AES flag as standard AES, so it does not decrypt these two modes.

Extensible encryption (minor version 3, or the `fExternal` flag) is not implemented. MS-OFFCRYPTO says that flag means every other field is zero and the cipher comes from an external provider. POI throws `Unknown encryption` and msoffcrypto has an empty extensible class. There is no portable password-based profile to implement or a fixture either tool can produce.

### New pivot tables and slicers

`Worksheet.AddPivotTable(name, source, dest, rowField, dataField)` snapshots a range, writes a pivot cache (shared strings, numeric min/max, cache records) and a pivot table in the shape Excel 2016 writes (`createdVersion="8"`, location, row items, `Sum of …` data field, `PivotStyleLight16`, the x14 and pivotdefaultlayout extensions). The sheet also receives the summarized values. `Worksheet.AddSlicer(pivotName, field, caption)` writes a slicer cache (`tabular` items plus the pivot table link) and a slicer part, and the worksheet `slicerList` extension.

Verified on a North/South amount table:

- openpyxl reads pivot name `SalesPivot`, `E4=North`, `F4=17`, grand total `22`.
- POI `getPivotTables()` returns `SalesPivot`.
- LibreOffice Calc converts the package and keeps the pivot cache (`Region`, `Amount`, `North`) and the pivot name.

### Samples

`testdata/excel/r11/` adds workbooks from two new sources:

- XlsxWriter (BSD) output written for this round: tables, comments, hyperlinks, freeze panes, conditional formats, data validation, a sparkline, a hidden sheet, unicode sheet names, and repeated formulas (`xw-sales.xlsx`, `xw-format.xlsx`, `xw-unicode.xlsx`).
- Calamine  (MIT) `tests/`: shared formulas, booleans, borders, a chart sheet, 1904 dates, ISO dates, quoted formats, backslash part names, case-insensitive part names, and whole-row/column references. Current EPPlus is Polyform Noncommercial, so those files are not vendored.

### Defects fixed from this corpus

- Shared formulas now shift whole-column (`F:G`) and whole-row (`6:7`) references. `cal-column_row_ranges.xlsx` matches openpyxl and POI after POI's `$A5:$XFD6` expansion is normalized back to `5:6`.
- A boolean cell with an empty `<v>` is stored as false, so the cell is not dropped.
- Words `true` and `TRUE` stay boolean true. POI only treats the token `1` as true; that disagreement is a POI gap. Calamine's own test expects `true`/`TRUE` to be true.
- A 1904-system serial that POI prints as a datetime compares equal to the serial (epoch 1904-01-01). Elapsed formats stay numeric in the model.

`make diffcheck` on r11 and on the full corpus (215 workbooks, including r11): **0 content defects**.

## Round 12 — full pivot builder, and ExcelJS / xlsx-populate / rust_xlsxwriter

### Pivot tables

`Worksheet.AddPivotTable` is unchanged: one row field and a sum. `Worksheet.BuildPivotTable` adds the rest:

- several row fields, column fields, and page/report filters (each filter is written above the table as the field name and `(All)`)
- several data fields, each with its own aggregation: sum, count, countNums, average, max, min, product, stdDev, stdDevp, var, varp
- a number format on a data field (`0`, `0.00`, `#,##0`, `#,##0.00`, `0%`, `0.00%`, `yyyy-mm-dd`)
- date grouping (years, quarters, months, days) and numeric bins (`start` / `end` / `interval`)
- subtotals and row/column grand totals, each optional
- pivot field items, `rowItems`, and `colItems`, including `field x="-2"` when there is more than one data field

Empty intersections are left blank. Group labels are stored in the cache records, so the sheet values are filled even when a reader ignores `fieldGroup`.

Checked on a 5-row Region/Product/Channel table (`/tmp/full-pivot.xlsx` from `TestFullPivotBuilder`):

- openpyxl `_pivots` names: `Full`, `ByMonth`, `Bins`. Grand sum `M18=55`, count `N18=5`, average `O18=11`, Store `C18=25`, Web `H18=30`. Date groups: Nov 8, Dec 12, Jan 15, Feb 20. Bins of Qty: 0–9 = 7, 10–19 = 4, 20–29 = 1.
- POI `getPivotTables()` returns those three names. `Full` has location `A10:Q18`, 5 data fields, and 1 page field.
- LibreOffice Calc 24.2 converts the package. Its compact refresh matches the ungrouped totals (Store 25, Web 30, North 35, grand sum 55, count 5, min 5) and shows the Year filter as `- all -`. It does not apply `fieldGroup`: the date pivot lists the source datetimes and the bin pivot lists the raw amounts, and both grand totals still match (55 and 12). It also drops average and max in that compact layout. Those are LibreOffice refresh limits. Excel itself is not installed here, so “opens without repair” is inferred from this package shape plus POI and openpyxl reading it.

Page fields list every item and do not hide rows. Choosing one report-filter item is not implemented.

### Samples

`testdata/excel/r12/` (29 workbooks) from sources not used in earlier rounds:

- ExcelJS (MIT) `spec/integration/data` and `test/data`: 1904 dates, a chart sheet, comments, formulas with `_xlfn.CONCAT`, images, many columns, shared-string `_x005F_` escapes, and the Fibonacci sheet. `huge.xlsx` (1,400,014 cells, 14MB) is `testdata/perf/ej-huge.xlsx` so LibreOffice is not asked to recalc it.
- xlsx-populate (MIT): rich text and styles.
- rust_xlsxwriter (MIT) `tests/input`: array formulas, form-control buttons, comments, dynamic arrays (`_xlfn.UNIQUE` plus `cm` and `xl/metadata.xml`), an image in a cell (`vm` plus `xl/richData`), classic icon sets, an x14 icon set, VML, and ignore-error flags. The upstream x14 icon-set workbook omitted a space in `<calcPr …"fullCalcOnLoad`; the copy in this tree inserts that space so POI and LibreOffice can open it. The x14 data bar, threaded comment, and custom XML parts are small packages written to the OOXML shape those features use, because the new sources did not ship a permissive fixture for them.
- XlsxWriter (BSD) `xw-large.xlsx`: 4,000 rows with a formula every 100 rows, for a mid-size sheet inside the differential run.

SheetJS `test_files` returned HTTP 403, and the public SheetJS tree that was reachable had no `.xlsx` files. Current EPPlus is Polyform Noncommercial and was not vendored. LibreOffice bug-tracker attachments were skipped where the license was not Apache, MIT, BSD, or MPL.

### Defects fixed from this corpus

- Shared strings decode Excel `_xHHHH_` escapes. The literal text `_x000D_` is stored as `_x005F_x000D_` and read back as `_x000D_`. Characters XML 1.0 forbids are stored the same way instead of being stripped, so `a\x01b\x00c` round-trips.
- `xl/metadata.xml`, `cm`, and `vm` round-trip, so dynamic-array anchors and rich-data cells keep their metadata indexes.
- `xl/richData/` round-trips with the workbook relationships that point at it.
- Package relationships in `_rels/.rels` other than the three standard ones are kept, so `customXml/item1.xml` survives a save.
- Threaded comments and the person list were already copied as package parts; the round-12 fixtures lock that in.

`python3 scripts/diffharness.py --roundtrip` on r12: **29 files, 0 content defects**. The same check on the full corpus (244 workbooks, including r12): **0 content defects**. `go test ./...` passes. Opening `testdata/perf/ej-huge.xlsx` reads 1,400,014 cells in about 7.7s and about 1080MB of heap.

## Round 13 — pivot item filters, streaming, cell images, threaded comments

### Pivot filters

`PivotFilter` on `PivotBuild` selects the visible items of a field. The summary grid is computed from the rows that match.

- A page field with one included item writes `pageField item="n"` and shows that label instead of `(All)`.
- Several included items, or `Multiple: true`, set `multipleItemSelectionAllowed="1"`, hide the other items with `h="1"`, and show `(Multiple Items)`.
- A row or column field hides excluded items with `h="1"` and drops them from `rowItems` / `colItems`.

Checked on a Region/Channel/Amount/Year table (`/tmp/pivot-filter.xlsx`):

- Year `2024` only: North sum 30, grand 30, South omitted. openpyxl, POI (pivot `Only2024` at `A10:B12`), and LibreOffice Calc agree. LibreOffice keeps the page caption `2024`.
- Channel `Web` only: North 10, South 8, grand 18. Store is not a column. LibreOffice's compact refresh matches those totals.

### Streaming

`OpenRows` returns a `RowIterator` that reads one worksheet row at a time. Shared strings are loaded once; the cell map is not. `CreateStream` / `NewStreamWriter` write rows directly into the zip using inline strings, so memory stays on the order of one row plus the compressor.

Measured in `TestStreamMemory` (peak `HeapAlloc` while the call runs):

| Path | Time | Peak heap |
| --- | --- | --- |
| `OpenFile` of `testdata/perf/ej-huge.xlsx` (1,400,014 cells) | 9.1s | 1600MB |
| `OpenRows` over the same file | 2.2s | 29MB |
| `CreateStream` of 2,000,000 numeric cells (250,000 × 8) | 1.5s | 3MB |

openpyxl, POI, and LibreOffice Calc open a streamed workbook. The streamed formula `SUM(B2,1)` is `4.5` after LibreOffice recalculates `3.5+1`. POI reads the 20,000-row streamed sheet through `F20000=200005`.

### Cell images and threaded comments

`Cell.SetImage` / `Cell.Image` read and write a rich-value local image (`vm`, `xl/metadata.xml`, `xl/richData`, and the media part). Excel's `#VALUE!` sentinel is what the cell stores. `testdata/excel/r12/rw-embed-image.xlsx` round-trips the PNG bytes.

`Workbook.AddPerson`, `Worksheet.AddThreadedComment`, `ReplyToThreadedComment`, `SetThreadedCommentText`, and `SetThreadedCommentResolved` read and write the person list and the thread (`parentId`, `done`). Creating a thread also writes a legacy note, so LibreOffice and older Excel still show the text. A thread that was only loaded is not given a new note on save.

### Samples

`testdata/excel/r13/` adds 19 workbooks from three sources not used before:

- PhpSpreadsheet (MIT): pivot, data bar, icon set, sparklines, chart sheet, shared formulas, forms and comments, defined names, a table.
- NPOI (Apache-2.0): comments, conditional formatting, newlines in formulas, a workbook with no shared-string table, matrix formulas.
- excelize (BSD-3): merge, overflow numeric, shared strings, calc chain, Book1.

`python3 scripts/diffharness.py --roundtrip` on r13: **19 files, 0 content defects**. The same check on the full corpus (263 workbooks, including r13): **0 content defects**.

## Round 14 — styled streaming, lower OpenFile memory, formula cache

### Streaming writer and reader

`CreateStream` now writes fonts, fills, borders, and number formats (including dates), an optional shared-string table (`UseSharedStrings`), column widths, merged cells, frozen panes, formulas, and hyperlinks. Style and string tables grow with the number of distinct styles and strings, not with the number of rows. A workbook that contains a formula gets `calcPr fullCalcOnLoad="1"` so Excel recalculates on open.

`OpenRows(..., WithFormattedValues())` loads the style table and the 1904 flag once, then applies number and date formats as each row is read. Date cells come back as `time.Time`.

Checked on `/tmp/styled-stream.xlsx`:

- openpyxl reads bold red `Title`, hyperlink `https://example.com`, number format `0.00`, date format `yyyy-mm-dd`, merge `A1:B1`, freeze `B2`, column width 18.
- POI opens the sheet and reads `A1=Title`.
- LibreOffice Calc keeps the merge and freeze and recalculates `SUM(B2,1)` to 4.5.

`TestStreamMemory` again: streaming read of `ej-huge.xlsx` peaks at about 33MB; writing 2,000,000 cells peaks at about 3MB.

### OpenFile memory

Plain numbers and plain strings are stored in compact maps. Cells are applied as each row is parsed instead of being buffered for the whole sheet. Worksheets larger than 1MB of XML stay unloaded until a cell is touched; smaller sheets are still parsed during `Read`, so a broken sheet fails immediately.

On `testdata/perf/ej-huge.xlsx` (1,400,014 cells):

| | Before this round | After |
| --- | --- | --- |
| Load time | about 9s | 3.0s |
| Peak heap while loading | about 1600MB | 415MB |
| Retained heap after GC | about 1080MB | 198MB |

`OpenRows` over the same file stays near 33MB.

### Formulas

The calculation engine evaluates comparisons (`=`, `<>`, `<`, `>`, `<=`, `>=`), so `IF(A1>1,"yes","no")` returns `yes` along with `SUM`, `AVERAGE`, `ROUND`, and `MAX`. Cached results are written on the formula cell. New workbooks, including streamed ones, set `fullCalcOnLoad="1"` (calcId 191029) so Excel recalculates. LibreOffice matches `SUM(1.5,2.5)=4`. Its converter stored the IF string result as numeric 0; the file we wrote caches `yes`.

### Samples

`testdata/excel/r14/` adds three workbooks from plandem/xlsx (MIT): formatted cells, iteration, and a simple sheet. `python3 scripts/diffharness.py --roundtrip` on the full corpus (266 workbooks, including r14): **0 content defects**. `go test ./...` passes.

## Round 15 — formula calculation engine

### Inventory before

The engine evaluated a single function call or one comparison, then `+ - * /`. It had 25 names: `SUM`, `AVERAGE`, `COUNT`, `COUNTA`, `MAX`, `MIN`, `IF`, `ABS`, `ROUND`, `SQRT`, `POWER`, `MOD`, `INT`, `LEN`, `UPPER`, `LOWER`, `TRIM`, `LEFT`, `RIGHT`, `MID`, `CONCATENATE`, `MEDIAN`, `SUMIF`, `COUNTIF`, plus the four arithmetic operators. A probe of 20 formulas (arithmetic, `SUMIF`, `COUNTIF`, `&`, `AND`, `VLOOKUP`, `A1/0`, `COUNT` of a range that contains text) evaluated 16 without an error. The four misses were `&`, `AND`, `VLOOKUP`, and division by zero (the error string was a Go error, not `#DIV/0!`). `COUNT` of a range that contained text counted the text. There was no `%`, `^`, unary minus, sheet reference, defined name, or absolute-ref handling beyond what the cell parser already stripped.

### After

The parser is precedence-based: comparisons, `&`, `+ -`, `* /`, right-associative `^`, unary `+ -`, postfix `%`, then parentheses, strings, numbers, error literals, refs, and calls. It accepts `$` absolute refs, `'Sheet Name'!A1`, `Sheet1!A1:B2`, and workbook defined names. Unknown names and unknown functions are `#NAME?`. A circular reference is `#REF!`. Division by zero is `#DIV/0!`. A multi-cell range used as a scalar is `#VALUE!`. `CalculateAll` stores the Excel error text (`#DIV/0!`, `#N/A`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`) on the cell. New workbooks still set `fullCalcOnLoad="1"` so Excel recalculates functions this engine does not implement.

Functions added or corrected this round, on top of the original 25:

- Math and aggregate: `SUMIFS`, `SUMPRODUCT`, `COUNTIFS`, `AVERAGEIF`, `AVERAGEIFS`, `ROUNDUP`, `ROUNDDOWN`, `CEILING`, `FLOOR`, `SIGN`, `PRODUCT`. `MOD` uses Excel's floor form (`MOD(-3,2)` is 1). `COUNT` counts only numbers. Range aggregates skip text and blanks; a scalar argument still coerces.
- Logical: `AND`, `OR`, `NOT`, `IFERROR`, `IFNA`, `IFS`, `SWITCH`, `XOR`.
- Lookup: `VLOOKUP`, `HLOOKUP`, `INDEX`, `MATCH`, `XLOOKUP`, `XMATCH`, `CHOOSE`, `OFFSET`, `INDIRECT`.
- Text: `CONCAT`, `TEXTJOIN`, `FIND`, `SEARCH`, `SUBSTITUTE`, `REPLACE`, `TEXT`, `VALUE`.
- Date and time: `DATE`, `TODAY`, `NOW`, `YEAR`, `MONTH`, `DAY`, `EDATE`, `EOMONTH`, `DATEDIF`, `WEEKDAY`, `NETWORKDAYS`, `WORKDAY`.
- Info: `ISBLANK`, `ISNUMBER`, `ISTEXT`, `ISERROR`, `ISNA`.

`_xlfn.` and `_xlws.` prefixes are stripped before the name lookup, so a file that stored `IFS` or `XLOOKUP` for Excel still evaluates here. `DATE` follows Excel's 1900 leap-day serial (`DATE(1900,1,1)` is 1 and `DATE(1900,2,29)` is 60). Stored `time.Time` values still use the writer's 1899-12-30 epoch, which already matches Excel serials from 1900-03-01 on.

### Differential

`TestFormulaEngineCoverage` builds one workbook of 65 formulas (the categories above, plus `&`, `%`, `^`, unary minus, a cross-sheet ref, and a defined name) and compares the engine with itself and with LibreOffice Calc 24.2 headless.

| Check | Before | After |
| --- | --- | --- |
| Engine on the 65-formula workbook | 16/20 on the smaller probe (the rest were unsupported) | 65/65 |
| LibreOffice on the same workbook | not measured (those functions were not written) | 63/65 |
| Excel cached values on corpus formulas the engine evaluates | not measured | 560/653 (85%) |

The corpus sample is `FormulaEvalTestData_Copy.xlsx`, `shared_formulas.xlsx`, `r11/xw-sales.xlsx`, `r12/ej-fibonacci.xlsx`, `r13/ps-sharedformula.xlsx`, and `r13/npoi-matrix.xlsx`. Cached `<v>` values written by Excel are the ground truth. The two LibreOffice misses are LibreOffice 24.2 gaps, not engine results: it returns `#NAME?` for `_xlfn.XLOOKUP` and `#VALUE!` for `SUM` of a defined name that the engine sums to 60. `TODAY` and `NOW` are implemented and left out of the strict fixture because the clock moves.

`go test ./...` passes. `python3 scripts/diffharness.py --roundtrip` on the full corpus (266 workbooks): **0 content defects**.

## Round 16 — array formulas, tables, and closer cache agreement

The round 15 engine matched 560 of 653 cached values on the formula sample (85%) and could evaluate 7,415 of 9,280 formula cells in the full corpus.

### What changed

- Implicit intersection for a range used as a scalar, plus explicit `@` / `SINGLE`. Legacy CSE array formulas (`t="array"`) evaluate element-wise and return the cell's slot. Array constants `{1,2;3,4}` parse. Operators on ranges are element-wise inside an array formula.
- Dynamic arrays: `FILTER`, `SORT`, `SORTBY`, `UNIQUE`, `SEQUENCE`, `TRANSPOSE`, `LET`, `LAMBDA`. The anchor cell returns the first value. `RANDARRAY` is implemented and left out of the cache comparison because it is volatile.
- Structured references read `xl/tables`: `Table[Col]`, `[#Headers]`, `[#Totals]`, `[#This Row]`, `[@Col]`, `[#Data]`, `[#All]`.
- Space is intersection (`#NULL!` when the ranges miss). Comma between references is union.
- `NETWORKDAYS` and `WORKDAY` take a holiday list. `NETWORKDAYS.INTL` and `WORKDAY.INTL` take the weekend code or a 7-character mask.
- `DATEDIF` accepts `Y`, `M`, `D`, `MD`, `YM`, and `YD`.
- Comparisons use Excel's order: number, then text, then boolean. `FALSE>"TRUE"` is true. Equality still does not coerce a number and a non-empty string.
- Error cells inside a range are real errors (`#DIV/0!`, `#N/A`, …) instead of text. `SUM` propagates them. `COUNT` skips them.
- Added the functions that showed up most often as `#NAME?`: `COLUMN`, `ROW`, `SUBTOTAL`, `AGGREGATE`, `CEILING.MATH`, `FLOOR.MATH`, the common trig and statistics functions, `PMT`/`PV`/`FV`/`NPV`/`NPER`, `LARGE`/`SMALL`/`MODE`, `ADDRESS`, `CLEAN`, `REPT`, `LOOKUP`, `COUNTBLANK`, `AVERAGEA`, `DAYS360`, `TIME`/`HOUR`/`MINUTE`/`SECOND`, and `NA`.

### Agreement

| Check | Before this round | After |
| --- | --- | --- |
| Engine, generated workbook | 65/65 | 69/69 |
| LibreOffice Calc 24.2 on that workbook | 63/65 | 66/69 |
| Sample workbooks, cached values among formulas the engine evaluates | 560/653 (85%) | 1051/1087 (96%) |
| Full corpus formula cells the engine evaluates | 7415/9280 | 8410/9280 (90.6%) |
| Full corpus cached values among those | 6353/7402 (86%) | 7324/8393 (87.2%) |

The sample is `FormulaEvalTestData_Copy.xlsx`, `shared_formulas.xlsx`, `r11/xw-sales.xlsx`, `r12/ej-fibonacci.xlsx`, `r13/ps-sharedformula.xlsx`, and `r13/npoi-matrix.xlsx`. `RAND`, `RANDBETWEEN`, `RANDARRAY`, `NOW`, and `TODAY` are not part of the cache comparison.

`go test ./...` passes. `python3 scripts/diffharness.py --roundtrip` on the full corpus (266 workbooks): **0 content defects**.

## Round 17 — number formats, spill, and matrix functions

### Number formats

`TEXT` and formatted-value reads (`GetFormattedValue`, `OpenRows` with formatted values) share one Excel format renderer. It now handles:

- elapsed `[h]`, `[m]`, `[s]` with remainders and fractional seconds
- conditional sections such as `[>10]` and `[<=10]`, falling back to General when the condition misses
- color names stripped from the text (`[Red]`, `[Color 3]`)
- locale and currency brackets (`[$-409]`, `[$¥-804]`)
- fractions (`# #/#`, fixed denominators)
- scientific formats, including a multi-digit mantissa (`####.####E+0` → `12.3457E+4`)
- the text section `@`
- padding `_` (a space) and `*` (omitted in `TEXT`, which has no column width)
- two-digit `y`/`yy`, `aaa`/`aaaa` weekdays, and literal East Asian date text such as `yyyy"年"m"月"d"日"`

### Calculation

- `DATE` truncates each component toward zero, and years 0 through 1899 are added to 1900, so `DATE(1,2,1)` is serial 398. Non-numeric components are `#VALUE!`.
- Approximate `VLOOKUP`/`HLOOKUP` returns the first exact match and otherwise the last key that is still less than the lookup.
- `CalculateAll` and `CalculateCell` memoize each cell, so a dependent formula sees the precedent's fresh result instead of recalculating it again or reading a stale cache.
- Dynamic-array results spill. The anchor keeps `cm="1"` and a new workbook writes `xl/metadata.xml` with an `XLDAPR` record. A blocked spill is `#SPILL!`.
- `MMULT`, `MDETERM`, and `MINVERSE` are implemented, along with `COLUMNS`, `ROWS`, `IRR`, `TIMEVALUE`, `NORM.DIST`, `NORM.S.DIST`, `POISSON.DIST`, `PERCENTILE.INC`, and `RANDBETWEEN`.

### Agreement

| Check | Before this round | After |
| --- | --- | --- |
| Sample workbooks, cached values among evaluated formulas | 1051/1087 (96%) | 1064/1096 (97%) |
| Full corpus formula cells the engine evaluates | 8410/9280 (90.6%) | 8447/9280 (91.0%) |
| Full corpus cached values among those | 7324/8393 (87.2%) | 7695/8423 (91.4%) |

The generated workbook is still 69/69 in the engine and 66/69 in LibreOffice Calc 24.2. `go test ./...` passes. `python3 scripts/diffharness.py --roundtrip` on the full corpus (266 workbooks): **0 content defects**.

## Round 18 — format edges and more evaluable formulas

### Number formats

The remaining POI format-suite gaps from round 17 are closed for the cases Excel's cached text defines:

- `?` padding counts grouping commas, so `?,?????????` of 1234567 is four leading spaces plus `1,234,567`.
- Fraction patterns with literals between digit placeholders keep those literals. `23.75` with `#-#-#:#/#` is `-2-3:3/4`, and `0.75` is `--3/4` (the colon is omitted when the whole-number digits are all insignificant).
- `General` uses a fixed decimal for magnitudes down to `1e-9` (`0.000000001`, not `1e-09`), scientific notation with an explicit exponent sign above `1e11` (`1E+11`), and at most six significant digits in that scientific form (`1.12345E+11`). `TEXT(TRUE,"General")` is `TRUE`.
- Negative elapsed times take the absolute duration and a leading minus, so `[h]:m:s` of −7384 seconds is `-2:3:4.000` rather than `3:3:4.000`.
- A percent that lands on a binary halfway value (`0.00035` with `0.00%`) rounds the same way Excel does (`0.04%`).

`TestFormatMatchesExcelCache` checks every formula on the Tests sheet of `ElapsedFormatTests.xlsx`, `DateFormatTests.xlsx`, `FormatConditionTests.xlsx`, and `GeneralFormatTests.xlsx` against the cached value Excel stored.

### Formulas

Formulas may contain newlines; `B1\n+B2` evaluates. A repeated percent sign (`1000%%`) is applied once per sign. New functions, chosen from the names the corpus still could not evaluate: `RATE`, `BINOM.DIST`, `EXPON.DIST`, `COVARIANCE.P`, `COVARIANCE.S`, `ERF.PRECISE`, `ERFC.PRECISE`, `CONFIDENCE.NORM`, `PERCENTILE.EXC`, `NEGBINOM.DIST`, `HYPGEOM.DIST`, `GEOMEAN`, `HARMEAN`, `ISEVEN`, `ISODD`, `ISOWEEKNUM`, `WEEKNUM`, and `DELTA`.

`SplitsItems` in `lo-XlStartupExternal.xlsx` is a name defined in another workbook, not a function. `TABLE(...)` is an Excel data-table formula. Both stay `#NAME?` for Excel to calculate via `fullCalcOnLoad`. LibreOffice 24.2 still reports `#NAME?` for `_xlfn.XLOOKUP` and does not order booleans above text; those are reference-tool gaps.

### Sample

`testdata/excel/r18/formats-and-matrix.xlsx` is a workbook written for this round: percent, `?` padding, an interleaved fraction, elapsed time, a date, `General` at `1e11`, `TEXT`, `BINOM.DIST`, `ERF.PRECISE`, `EXPON.DIST`, `MMULT`, `MDETERM`, `DATE(1,2,1)`, and `1000%%`.

### Agreement

| Check | Before this round | After |
| --- | --- | --- |
| Sample workbooks, cached values among evaluated formulas | 1064/1096 (97%) | 1066/1098 (97%) |
| Full corpus formula cells the engine evaluates | 8447/9280 (91.0%) | 8507/9280 (91.7%) |
| Full corpus cached values among those | 7695/8423 (91.4%) | 7862/8483 (92.7%) |

The generated engine fixture stays 69/69, and LibreOffice Calc 24.2 stays 66/69 on it. `go test ./...` passes. `python3 scripts/diffharness.py --roundtrip` on the full corpus (267 workbooks, including r18): **0 content defects**.

## Remaining known gaps

- openpyxl and POI do not read strict OOXML (POI bug 57699). goexcel reads it and can write it when asked. LibreOffice Calc opens a strict file written by goexcel.
- LibreOffice recalculates pivots, localizes captions, and reports `Err:502` or `Err:522` for some Excel formulas. Conversion succeeds; those drifts are not goexcel defects. LibreOffice 24.2 also aborts (`ScRangeList`) while importing `r1/cx-SampleSparklines.xlsx`. openpyxl and POI read that file with no content diff.
- POI drops a shared-string index that contains trailing whitespace (`<v>2 </v>`). goexcel and openpyxl read the string.
- POI treats only the boolean token `1` as true. `true` and `TRUE` are booleans in the file and in goexcel.
- openpyxl warns that it does not support sparkline or slicer-list extensions and would drop them if it saved the file. goexcel preserves and writes them. LibreOffice's converter keeps a new pivot cache and drops the slicer parts on export.
- msoffcrypto-tool 6.1.0 does not decrypt binary RC4 or RC4 CryptoAPI. It also cannot decrypt its own encrypt of a package under the 4096-byte mini-stream cutoff. goexcel pads Agile packages so both msoffcrypto and POI decrypt them.
- Extensible encryption is an external cipher provider. It is not implemented.
- LibreOffice's refresh ignores pivot `fieldGroup`, so date and numeric groups stay as the precomputed cache labels in goexcel and become the base values after a LibreOffice recalc. Average and max can disappear in LibreOffice's compact layout.
- `PivotFilter.Include` is an allow-list. A numeric field that is only a data field (not on an axis) has no discrete items to filter.
- `Cell.Image` covers rich-value local images. Other rich-value kinds are still package parts. `OpenRows` applies formats only when `WithFormattedValues` is set. The calculation engine covers the functions in rounds 15 through 18. Dynamic-array spill writes the block next to the anchor and records `XLDAPR` metadata on workbooks created here; a file that already has `xl/metadata.xml` keeps that part. `RANDARRAY` and `RANDBETWEEN` are volatile and are not compared with cached values. The POI number-format suite still has scientific edge cases (a zero mantissa written as `0e0` in some bracket layouts) and some fraction patterns beyond the interleaved `#-#-#` form. `SplitsItems(...)` is an external defined name, and `TABLE(...)` is a data table; both stay on `fullCalcOnLoad`. Aggregates whose own inputs still disagree with Excel keep that disagreement in the total. LibreOffice 24.2 returns `#NAME?` for `_xlfn.XLOOKUP` and `#VALUE!` for `SUM` of a defined name, and it compares `FALSE>"TRUE"` as false. Those are LibreOffice gaps. Unknown functions still rely on `fullCalcOnLoad`.
- LibreOffice 24.2 reports `#NAME?` for `_xlfn.UNIQUE` and recalculates some cached formula results (the 4,000-row sheet's cached `0` becomes the computed total). Its converter stored `IF(A1>1,"yes","no")` as numeric 0. Those are recalc gaps. openpyxl still refuses a workbook whose relationship points at a missing `xl/sharedStrings.xml`.
- `OpenFile` of `testdata/perf/ej-huge.xlsx` retains about 198MB after the sheet is loaded (peak about 415MB while parsing). `OpenRows` stays near 33MB. The file stays outside `make diffcheck`. Worksheets larger than 1MB of XML are not parsed until a cell is read.
- Dialog sheets and chartsheets are kept and reattached with their original relationship types. They have no grid cells. `AddSlicer` still attaches a slicer to a pivot created with `AddPivotTable` or `BuildPivotTable`.

