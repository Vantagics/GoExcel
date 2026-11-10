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

## Remaining known gaps

- openpyxl and POI do not read strict OOXML (POI bug 57699). goexcel reads it and can write it when asked. LibreOffice Calc opens a strict file written by goexcel.
- LibreOffice recalculates pivots, localizes captions, and reports `Err:502` or `Err:522` for some Excel formulas. Conversion succeeds; those drifts are not goexcel defects. LibreOffice 24.2 also aborts (`ScRangeList`) while importing `r1/cx-SampleSparklines.xlsx`. openpyxl and POI read that file with no content diff.
- POI drops a shared-string index that contains trailing whitespace (`<v>2 </v>`). goexcel and openpyxl read the string.
- POI treats only the boolean token `1` as true. `true` and `TRUE` are booleans in the file and in goexcel.
- openpyxl warns that it does not support sparkline or slicer-list extensions and would drop them if it saved the file. goexcel preserves and writes them. LibreOffice's converter keeps a new pivot cache and drops the slicer parts on export.
- msoffcrypto-tool 6.1.0 does not decrypt binary RC4 or RC4 CryptoAPI. It also cannot decrypt its own encrypt of a package under the 4096-byte mini-stream cutoff. goexcel pads Agile packages so both msoffcrypto and POI decrypt them.
- Extensible encryption is an external cipher provider. It is not implemented.
- Dialog sheets and chartsheets are kept and reattached with their original relationship types. They have no grid cells.
- `AddPivotTable` builds one row field and one summed data field from a worksheet range. Column fields, filters, and grouping are not generated. `AddSlicer` attaches a slicer to a pivot created that way.

