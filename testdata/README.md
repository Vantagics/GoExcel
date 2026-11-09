# Excel fixtures

These workbooks drive read and round-trip tests in `xlsx_compat_test.go`.
They are real `.xlsx` packages (many produced by Microsoft Excel) covering
shared strings, rich text, inline strings, shared and array formulas, styles,
dates (including the 1904 date system), defined names, merged cells, comments,
hyperlinks, conditional formatting, data validation, tables, charts, drawings,
pivot caches, and strict OOXML.

## Sources and licenses

Only files whose license allows redistribution are committed.

| Files | Source | License |
| --- | --- | --- |
| All `*.xlsx` except the `excelize-*` files and `Book1.xlsx` | [Apache POI](https://github.com/apache/poi) `test-data/spreadsheet` (trunk) | Apache License 2.0. Copyright the Apache Software Foundation. |
| `Book1.xlsx`, `excelize-SharedStrings.xlsx`, `excelize-MergeCell.xlsx`, `excelize-CalcChain.xlsx` | [excelize](https://github.com/qax-os/excelize) `test/` | BSD 3-Clause. Copyright (c) 2016-2024 The excelize Authors. |
| `r1/cx-*.xlsx` and later ClosedXML `rN/cx-*.xlsx` files | [ClosedXML](https://github.com/ClosedXML/ClosedXML) `ClosedXML.Tests/Resource` | MIT. Copyright (c) ClosedXML contributors. |
| `r2/*.xlsx` and `r8/`–`r10/poi-*.xlsx` / `r9/*.xlsm` | [Apache POI](https://github.com/apache/poi) `test-data/spreadsheet` | Apache License 2.0. |
| `r3/lo-*.xlsx` | [LibreOffice](https://github.com/LibreOffice/core) `sc/qa/unit/data/xlsx` | Mozilla Public License 2.0. |

Apache POI test data is redistributed under the Apache License 2.0.
See https://www.apache.org/licenses/LICENSE-2.0.

excelize's BSD 3-Clause license requires retaining this notice:

> Redistribution and use in source and binary forms, with or without
> modification, are permitted provided that the following conditions are met:
> 1. Redistributions of source code must retain the above copyright notice,
>    this list of conditions and the following disclaimer.
> 2. Redistributions in binary form must reproduce the above copyright notice,
>    this list of conditions and the following disclaimer in the documentation
>    and/or other materials provided with the distribution.
> 3. Neither the name of the copyright holder nor the names of its contributors
>    may be used to endorse or promote products derived from this software
>    without specific prior written permission.
>
> THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
> AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
> IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
> ARE DISCLAIMED.

Differential results are in [COMPAT_ROUNDS.md](COMPAT_ROUNDS.md). `make diffcheck` compares this tree with whatever of openpyxl, Apache POI, and LibreOffice is installed.

## Refreshing the files

`download_fixtures.sh` downloads the same set. Run it from the repository root:

```bash
testdata/download_fixtures.sh
```
