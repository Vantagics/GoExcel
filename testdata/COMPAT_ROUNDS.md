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

## Later rounds

Filled in as each round lands.
