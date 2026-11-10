> [English](API_en.md) | [中文](API.md)

# GoExcel API Documentation

Complete API reference covering all public types, functions, and methods.

---

## Table of Contents

- [Core Types](#core-types)
  - [Workbook](#workbook)
  - [Worksheet](#worksheet)
  - [Cell](#cell)
- [Styling System](#styling-system)
- [Coordinate Utilities](#coordinate-utilities)
- [Formula Calculation Engine](#formula-calculation-engine)
- [File I/O](#file-io)
- [Hyperlinks / Comments / Rich Text](#hyperlinks--comments--rich-text)
- [Conditional Formatting](#conditional-formatting)
- [Data Validation](#data-validation)
- [Auto Filter](#auto-filter)
- [Page Setup](#page-setup)
- [Protection](#protection)

---

## Core Types

### Workbook

`Workbook` is the top-level container for a spreadsheet, containing one or more worksheets.

#### Type Definitions

```go
type DocumentProperties struct {
    Creator        string
    LastModifiedBy string
    Title          string
    Subject        string
    Description    string
    Keywords       string
    Category       string
}
```

#### Constructors

| Function | Description |
|----------|-------------|
| `New() *Workbook` | Creates a new workbook with one default worksheet (Sheet1) |
| `NewEmpty() *Workbook` | Creates an empty workbook with no worksheets |

#### Methods

| Method | Description |
|--------|-------------|
| `AddSheet(title string) (*Worksheet, error)` | Adds a new worksheet; title must be unique |
| `GetSheet(index int) (*Worksheet, error)` | Gets a worksheet by index (0-based) |
| `GetSheetByName(title string) (*Worksheet, error)` | Gets a worksheet by title |
| `SheetCount() int` | Returns the number of worksheets |
| `GetSheetNames() []string` | Returns all worksheet titles |
| `RemoveSheet(index int) error` | Removes the worksheet at the given index (at least one must remain) |
| `GetActiveSheet() *Worksheet` | Returns the currently active worksheet |
| `SetActiveSheet(index int) error` | Sets the active worksheet |
| `AddNamedRange(name, reference string)` | Adds a named range (e.g. `"Sheet1!A1:B2"`) |
| `GetNamedRange(name string) (string, error)` | Gets a named range reference |
| `GetNamedRanges() map[string]string` | Gets all named ranges |
| `RemoveNamedRange(name string) error` | Removes a named range |
| `Sheets() []*Worksheet` | Returns all worksheets |
| `CopySheet(srcIndex int, newTitle string) (*Worksheet, error)` | Copies a worksheet |
| `MoveSheet(fromIndex, toIndex int) error` | Moves a worksheet to a new position |
| `SetWorkbookProtection(wp *WorkbookProtection)` | Sets workbook protection |
| `GetWorkbookProtection() *WorkbookProtection` | Gets workbook protection |
| `ClearProtection()` | Removes all workbook protection |
| `Validate() error` | Checks workbook for common issues |
| `Close() error` | Closes the workbook (no-op for in-memory) |

#### Properties

| Field | Type | Description |
|-------|------|-------------|
| `Properties` | `DocumentProperties` | Document properties (title, author, etc.) |

---

### Worksheet

`Worksheet` represents a single sheet within a workbook.

#### Type Definitions

```go
type MergeCell struct {
    StartRow int
    StartCol int
    EndRow   int
    EndCol   int
}
```

#### Constructors

| Function | Description |
|----------|-------------|
| `NewWorksheet(title string) *Worksheet` | Creates a new worksheet (typically used via `Workbook.AddSheet`) |

#### Basic Methods

| Method | Description |
|--------|-------------|
| `Title() string` | Returns the worksheet title |
| `SetTitle(title string) *Worksheet` | Sets the worksheet title |
| `GetCell(row, col int) *Cell` | Gets a cell by 0-based row/col index (auto-creates if not found) |
| `GetCellByName(ref string) (*Cell, error)` | Gets a cell by reference name (e.g. `"A1"`) |
| `SetCellValue(ref string, value interface{}) error` | Sets a cell value |
| `SetCellFormula(ref string, formula string) error` | Sets a cell formula |
| `SetCellStyle(ref string, style *Style) error` | Sets a cell style |
| `GetCellValue(ref string) (interface{}, error)` | Gets a cell value |

#### Merge and Freeze

| Method | Description |
|--------|-------------|
| `MergeCells(rangeStr string) error` | Merges a cell range (e.g. `"A1:C3"`) |
| `GetMergeCells() []MergeCell` | Gets all merged regions |
| `RemoveMergedCell(rangeStr string) error` | Removes a merged cell region by range |
| `FreezePane(ref string) error` | Freezes panes (e.g. `"A2"` freezes the first row) |
| `GetFreezePane() *CellReference` | Gets the freeze position |

#### Row and Column Dimensions

| Method | Description |
|--------|-------------|
| `SetColumnWidth(col int, width float64) *Worksheet` | Sets column width (0-based) |
| `GetColumnWidth(col int) float64` | Gets column width (default 8.43) |
| `SetRowHeight(row int, height float64) *Worksheet` | Sets row height (0-based) |
| `GetRowHeight(row int) float64` | Gets row height (default 15.0) |

#### Row and Column Visibility

| Method | Description |
|--------|-------------|
| `SetRowHidden(row int, hidden bool) *Worksheet` | Sets row visibility (0-based) |
| `IsRowHidden(row int) bool` | Returns whether a row is hidden |
| `SetColumnHidden(col int, hidden bool) *Worksheet` | Sets column visibility (0-based) |
| `IsColumnHidden(col int) bool` | Returns whether a column is hidden |

#### Row and Column Operations

| Method | Description |
|--------|-------------|
| `InsertRow(rowIdx int)` | Inserts an empty row at the given position; existing rows shift down |
| `DeleteRow(rowIdx int)` | Deletes the given row; rows below shift up |
| `InsertColumn(colIdx int)` | Inserts an empty column at the given position; existing columns shift right |
| `DeleteColumn(colIdx int)` | Deletes the given column; columns to the right shift left |
| `CopyRow(srcRow, dstRow int)` | Copies a row (values, formulas, styles) |

> Row/column operations automatically adjust merged cells and row heights/column widths.

#### Sorting

| Method | Description |
|--------|-------------|
| `Sort(column string, firstRow int, order SortOrder) error` | Sorts rows by column values; firstRow is 0-based (skip headers) |

#### SortOrder Constants

| Constant | Value |
|----------|-------|
| `SortOrderAscending` | `0` |
| `SortOrderDescending` | `1` |

#### Iteration and Statistics

| Method | Description |
|--------|-------------|
| `Dimensions() (minRow, minCol, maxRow, maxCol int, err error)` | Returns the used area range (0-based) |
| `RowIterator() ([][]*Cell, error)` | Returns all cells by row (empty positions are nil) |
| `CellCount() int` | Returns the number of non-empty cells |
| `AllCells() []*Cell` | Returns all non-empty cells (sorted by row then column) |

#### Hyperlinks and Comments

| Method | Description |
|--------|-------------|
| `SetCellHyperlink(ref, url string) error` | Sets a cell hyperlink |
| `SetCellComment(ref, author, text string) error` | Sets a cell comment |

#### Border Range Helper

| Method | Description |
|--------|-------------|
| `SetBorderRange(rangeStr string, border Border) error` | Applies a box border to a cell range |

#### Sheet View

| Method | Description |
|--------|-------------|
| `SetSheetView(sv *SheetView) *Worksheet` | Sets sheet view configuration |
| `GetSheetView() *SheetView` | Gets sheet view (creates default if none) |

#### Validation

| Method | Description |
|--------|-------------|
| `Validate() error` | Checks worksheet for common issues |

#### Conditional Formatting / Data Validation / Filter

| Method | Description |
|--------|-------------|
| `AddConditionalFormatting(cf *ConditionalFormatting) *Worksheet` | Adds conditional formatting |
| `GetConditionalFormattings() []*ConditionalFormatting` | Gets all conditional formatting rules |
| `AddDataValidation(dv *DataValidation) *Worksheet` | Adds data validation |
| `GetDataValidations() []*DataValidation` | Gets all data validations |
| `SetAutoFilter(af *AutoFilter) *Worksheet` | Sets auto filter |
| `GetAutoFilter() *AutoFilter` | Gets auto filter |

#### Page Setup and Protection

| Method | Description |
|--------|-------------|
| `SetPageSetup(ps *PageSetup) *Worksheet` | Sets page configuration |
| `GetPageSetup() *PageSetup` | Gets page configuration (creates default if none exists) |
| `SetSheetProtection(sp *SheetProtection) *Worksheet` | Sets worksheet protection |
| `GetSheetProtection() *SheetProtection` | Gets worksheet protection |
| `SetTabColor(color string) *Worksheet` | Sets tab color (hex, e.g. `"FF0000"`) |
| `GetTabColor() string` | Gets tab color |

---

### Cell

`Cell` represents a single cell in a worksheet.

#### Type Definitions

```go
type CellType int

const (
    CellTypeEmpty   CellType = iota // Empty
    CellTypeString                  // String
    CellTypeNumeric                 // Numeric
    CellTypeBool                    // Boolean
    CellTypeFormula                 // Formula
    CellTypeDate                    // Date
    CellTypeError                   // Error
)
```

#### Constructors

| Function | Description |
|----------|-------------|
| `NewCell(row, col int) *Cell` | Creates an empty cell (typically used via `Worksheet.GetCell`) |

#### Methods

| Method | Description |
|--------|-------------|
| `SetValue(v interface{}) *Cell` | Sets value with auto type detection (supports string/int/float64/bool/time.Time) |
| `SetFormula(formula string) *Cell` | Sets a formula |
| `SetFormulaArray(formula string) *Cell` | Sets an array formula (Ctrl+Shift+Enter) |
| `SetStyle(s *Style) *Cell` | Sets a style |
| `SetInlineString(s string) *Cell` | Sets an inline string (not shared) |
| `SetDateWithStyle(t time.Time) *Cell` | Sets a date with default date format applied |
| `SetNumberWithFormat(v float64, nf NumberFormat) *Cell` | Sets a number with format applied |
| `Clear() *Cell` | Resets cell to empty state |
| `GetStringValue() string` | Gets the string representation |
| `GetNumericValue() (float64, error)` | Gets the numeric value (returns error for non-numeric types) |
| `GetBoolValue() (bool, error)` | Gets the boolean value |
| `GetDateValue() (time.Time, error)` | Gets the date value |
| `GetFormattedValue() string` | Gets the value formatted per number format |
| `HasFormula() bool` | Returns true if cell has a formula |
| `IsEmpty() bool` | Returns true if cell has no value or formula |
| `IsNumber() bool` | Returns true if cell is numeric |
| `IsBool() bool` | Returns true if cell is boolean |
| `IsDate() bool` | Returns true if cell is a date |
| `IsString() bool` | Returns true if cell is a string |
| `IsError() bool` | Returns true if cell is an error |
| `Row() int` | Returns the row index (0-based) |
| `Col() int` | Returns the column index (0-based) |
| `SetHyperlink(h *Hyperlink) *Cell` | Sets a hyperlink |
| `SetComment(comment *Comment) *Cell` | Sets a comment |
| `SetRichText(rt *RichText) *Cell` | Sets rich text |

#### Public Fields

| Field | Type | Description |
|-------|------|-------------|
| `Value` | `interface{}` | Cell value |
| `Type` | `CellType` | Value type |
| `Formula` | `string` | Formula string |
| `Style` | `*Style` | Cell style |
| `Hyperlink` | `*Hyperlink` | Hyperlink |
| `Comment` | `*Comment` | Comment |
| `RichText` | `*RichText` | Rich text |

---

## Styling System

### Style

```go
type Style struct {
    Font         *Font
    Fill         *Fill
    Borders      *Borders
    Alignment    *Alignment
    NumberFormat *NumberFormat
}
```

| Method | Description |
|--------|-------------|
| `NewStyle() *Style` | Creates an empty style |
| `SetFont(f *Font) *Style` | Sets font |
| `SetFill(f *Fill) *Style` | Sets fill |
| `SetBorders(b *Borders) *Style` | Sets borders |
| `SetAlignment(a *Alignment) *Style` | Sets alignment |
| `SetNumberFormat(nf *NumberFormat) *Style` | Sets number format |

### Font

```go
type Font struct {
    Name          string  // Font name (default "Calibri")
    Size          float64 // Font size (default 11)
    Bold          bool
    Italic        bool
    Underline     bool
    Strikethrough bool
    Color         string  // Hex color, e.g. "FF0000"
}
```

| Function | Description |
|----------|-------------|
| `DefaultFont() *Font` | Returns the default font (Calibri, 11pt) |

### Fill

```go
type Fill struct {
    Type    string // "solid", "pattern", "none"
    Color   string // Hex color
    Pattern string // Pattern type
}
```

### Borders

```go
type Borders struct {
    Left     Border
    Right    Border
    Top      Border
    Bottom   Border
    Diagonal Border
    DiagonalUp   bool
    DiagonalDown bool
}

type Border struct {
    Style BorderStyle
    Color string
}
```

#### BorderStyle Constants

| Constant | Value |
|----------|-------|
| `BorderNone` | `"none"` |
| `BorderThin` | `"thin"` |
| `BorderMedium` | `"medium"` |
| `BorderThick` | `"thick"` |
| `BorderDashed` | `"dashed"` |
| `BorderDotted` | `"dotted"` |
| `BorderDouble` | `"double"` |

### Alignment

```go
type Alignment struct {
    Horizontal   HorizontalAlignment
    Vertical     VerticalAlignment
    WrapText     bool
    ShrinkToFit  bool
    TextRotation int // -90 to 90 degrees
    Indent       int
}
```

#### HorizontalAlignment Constants

| Constant | Value |
|----------|-------|
| `AlignLeft` | `"left"` |
| `AlignCenter` | `"center"` |
| `AlignRight` | `"right"` |
| `AlignJustify` | `"justify"` |
| `AlignGeneral` | `"general"` |

#### VerticalAlignment Constants

| Constant | Value |
|----------|-------|
| `AlignTop` | `"top"` |
| `AlignMiddle` | `"center"` |
| `AlignBottom` | `"bottom"` |

### NumberFormat

```go
type NumberFormat struct {
    FormatCode string
}
```

#### Predefined Number Formats

| Variable | FormatCode | Description |
|----------|------------|-------------|
| `FormatGeneral` | `"General"` | General |
| `FormatNumber` | `"0"` | Integer |
| `FormatNumber2Dec` | `"0.00"` | Two decimal places |
| `FormatPercent` | `"0%"` | Percentage |
| `FormatPercent2Dec` | `"0.00%"` | Percentage with two decimals |
| `FormatDate` | `"yyyy-mm-dd"` | Date |
| `FormatDateTime` | `"yyyy-mm-dd hh:mm:ss"` | Date and time |
| `FormatTime` | `"hh:mm:ss"` | Time |
| `FormatCurrency` | `#,##0.00"$"` | Currency |
| `FormatAccounting` | `_("$"* #,##0.00_)` | Accounting |
| `FormatText` | `"@"` | Text |

---

## Coordinate Utilities

`coordinate.go` provides cell reference parsing and conversion utilities.

### CellReference

```go
type CellReference struct {
    Column    string // Column name, e.g. "A"
    ColumnIdx int    // 0-based column index
    Row       int    // 1-based row number
}
```

### Functions

| Function | Description |
|----------|-------------|
| `ColumnIndexToName(index int) (string, error)` | Converts column index to name (0→"A", 25→"Z", 26→"AA") |
| `ColumnNameToIndex(name string) (int, error)` | Converts column name to index ("A"→0, "Z"→25, "AA"→26) |
| `ParseCellReference(ref string) (*CellReference, error)` | Parses a cell reference (e.g. `"A1"`) |
| `CellName(row, col int) (string, error)` | Converts 0-based row/col to reference name (e.g. `(0,0)` → `"A1"`) |
| `ParseRange(rangeStr string) (*CellReference, *CellReference, error)` | Parses a range (e.g. `"A1:C3"`) |

---

## Formula Calculation Engine

### CalculationEngine

| Function/Method | Description |
|-----------------|-------------|
| `NewCalculationEngine(wb *Workbook) *CalculationEngine` | Creates a calculation engine |
| `CalculateCell(ws *Worksheet, ref string) (interface{}, error)` | Calculates a single cell formula |
| `CalculateAll() error` | Calculates all formula cells in the workbook |

### Built-in Functions (24)

#### Math Functions

| Function | Syntax | Description |
|----------|--------|-------------|
| SUM | `SUM(range)` | Sum |
| ABS | `ABS(value)` | Absolute value |
| ROUND | `ROUND(value, digits)` | Round to specified decimal places |
| SQRT | `SQRT(value)` | Square root (returns `#NUM!` for negative values) |
| POWER | `POWER(base, exp)` | Exponentiation |
| MOD | `MOD(num, divisor)` | Modulo (returns `#DIV/0!` if divisor is 0) |
| INT | `INT(value)` | Floor (round down) |

#### Statistical Functions

| Function | Syntax | Description |
|----------|--------|-------------|
| AVERAGE | `AVERAGE(range)` | Average (returns `#DIV/0!` for empty range) |
| COUNT | `COUNT(range)` | Counts numeric cells |
| COUNTA | `COUNTA(range)` | Counts non-empty cells |
| MAX | `MAX(range)` | Maximum value |
| MIN | `MIN(range)` | Minimum value |
| MEDIAN | `MEDIAN(range)` | Median value |

#### Logic Functions

| Function | Syntax | Description |
|----------|--------|-------------|
| IF | `IF(condition, true_val [, false_val])` | Conditional (2 or 3 arguments) |

#### Text Functions

| Function | Syntax | Description |
|----------|--------|-------------|
| LEN | `LEN(text)` | Character count (Unicode-aware) |
| UPPER | `UPPER(text)` | Convert to uppercase |
| LOWER | `LOWER(text)` | Convert to lowercase |
| TRIM | `TRIM(text)` | Remove leading/trailing whitespace |
| LEFT | `LEFT(text, count)` | Left substring |
| RIGHT | `RIGHT(text, count)` | Right substring |
| MID | `MID(text, start, length)` | Middle substring (start is 1-based) |
| CONCATENATE | `CONCATENATE(val1, val2, ...)` | String concatenation |

#### Conditional Functions

| Function | Syntax | Description |
|----------|--------|-------------|
| SUMIF | `SUMIF(criteria_range, criteria [, sum_range])` | Conditional sum |
| COUNTIF | `COUNTIF(range, criteria)` | Conditional count |

Supported criteria: exact match, `">N"`, `"<N"`, `">=N"`, `"<=N"`, `"<>N"`

### Operator Support

Formulas support arithmetic operators: `+`, `-`, `*`, `/`, as well as cell references and nested function calls.

---

## File I/O

### IOFactory (Auto Format Detection)

| Function | Description |
|----------|-------------|
| `OpenFile(filename string, opts ...Option) (*Workbook, error)` | Auto-selects reader by file extension (.xlsx / .csv). `WithPassword` opens an ECMA-376 encrypted workbook. |
| `SaveFile(wb *Workbook, filename string) error` | Auto-selects writer by file extension. `Workbook.SetPassword` writes Agile encryption. |

### XLSXReader / XLSXWriter

| Function/Method | Description |
|-----------------|-------------|
| `NewXLSXReader() *XLSXReader` | Creates an XLSX reader |
| `(*XLSXReader) Open(filename string) (*Workbook, error)` | Reads from file |
| `(*XLSXReader) Read(reader io.ReaderAt, size int64) (*Workbook, error)` | Reads from io.ReaderAt. Encrypted packages return `ErrPasswordRequired`. |
| `(*XLSXReader) OpenPassword / ReadPassword` | Decrypts Agile or Standard ECMA-376 encryption, then reads the package. |
| `DecryptPackage` / `EncryptPackage` | Decrypt or Agile-encrypt a raw OOXML package. |
| `NewXLSXWriter() *XLSXWriter` | Creates an XLSX writer |
| `(*XLSXWriter) Save(wb *Workbook, filename string) error` | Saves to file |
| `(*XLSXWriter) Write(wb *Workbook, writer io.Writer) error` | Writes to io.Writer. A workbook password emits an OLE encrypted package. `SetStrictOOXML(true)` writes strict ECMA-376 namespaces; the default stays transitional. |

### Array formulas, sparklines, slicers, pivots

| API | Description |
|-----|-------------|
| `(*Worksheet) SetArrayFormula(rangeRef, formula string) error` | Stores an array formula on the anchor. Every cell in the ref reports it. |
| `(*Worksheet) FormulaAt(ref string) (formula, arrayRef string, err error)` | Formula Excel would show, plus the array ref when the cell is covered by one. |
| `(*Cell) ArrayRef / IsArrayFormula / ReportedFormula` | Anchor ref and the formula bar text, including for cells that only belong to the ref. |
| `(*Worksheet) SparklineGroups / SetSparklineGroups` | Read or replace sparkline groups (type, markers, colors, formula, cell). |
| `(*Worksheet) AddPivotTable(name, source, dest, rowField, dataField string)` | Builds a pivot cache from a range and a pivot table at `dest`. |
| `(*Worksheet) AddSlicer(pivotName, field, caption string)` | Creates a slicer and slicer cache for a field of a pivot on this sheet. |
| `(*Worksheet) PivotTables` | Pivot name, location, data caption, style, fields. `SetName`, `SetLocation`, `SetDataCaption`, `SetStyle` write those attributes back. |
| `(*Worksheet) Slicers` / `(*Workbook) SlicerCaches` | Caption, style, column count, row height, and cache source name. |

### CSVReader / CSVWriter

| Function/Method | Description |
|-----------------|-------------|
| `NewCSVReader() *CSVReader` | Creates a CSV reader (default comma delimiter) |
| `(*CSVReader) Open(filename string) (*Workbook, error)` | Reads from file |
| `(*CSVReader) Read(reader io.Reader) (*Workbook, error)` | Reads from io.Reader |
| `NewCSVWriter() *CSVWriter` | Creates a CSV writer |
| `(*CSVWriter) Save(wb *Workbook, filename string) error` | Saves to file |
| `(*CSVWriter) Write(wb *Workbook, writer io.Writer) error` | Writes to io.Writer |

#### CSVReader Fields

| Field | Type | Default | Description |
|-------|------|---------|-------------|
| `Delimiter` | `rune` | `','` | Field delimiter |
| `LazyQuotes` | `bool` | `false` | Relaxed quote parsing |

#### CSVWriter Fields

| Field | Type | Default | Description |
|-------|------|---------|-------------|
| `Delimiter` | `rune` | `','` | Field delimiter |
| `SheetIndex` | `int` | `0` | Worksheet index to write |

---

## Hyperlinks / Comments / Rich Text

### Hyperlink

```go
type Hyperlink struct {
    URL     string
    Tooltip string
}
```

| Function/Method | Description |
|-----------------|-------------|
| `NewHyperlink(url string) *Hyperlink` | Creates a hyperlink |
| `SetTooltip(tooltip string) *Hyperlink` | Sets tooltip text |

### Comment

```go
type Comment struct {
    Author string
    Text   string
}
```

| Function | Description |
|----------|-------------|
| `NewComment(author, text string) *Comment` | Creates a comment |

### RichText

```go
type RichTextRun struct {
    Text string
    Font *Font
}

type RichText struct {
    Runs []RichTextRun
}
```

| Function/Method | Description |
|-----------------|-------------|
| `NewRichText() *RichText` | Creates an empty rich text |
| `AddRun(text string, font *Font) *RichText` | Adds a text run (with optional font) |
| `PlainText() string` | Returns plain text content |

---

## Conditional Formatting

### ConditionalFormatting

```go
type ConditionalFormatting struct {
    Range string           // e.g. "A1:A10"
    Rules []ConditionalRule
}

type ConditionalRule struct {
    Type       ConditionalType
    Operator   ConditionalOperator
    Formula    []string
    Style      *Style
    Priority   int
    StopIfTrue bool
}
```

| Function/Method | Description |
|-----------------|-------------|
| `NewConditionalFormatting(rangeStr string) *ConditionalFormatting` | Creates conditional formatting |
| `AddRule(rule ConditionalRule) *ConditionalFormatting` | Adds a rule |
| `CellIsRule(op ConditionalOperator, formula string, style *Style) ConditionalRule` | Creates a cellIs rule |
| `BetweenRule(formula1, formula2 string, style *Style) ConditionalRule` | Creates a between rule |
| `ExpressionRule(formula string, style *Style) ConditionalRule` | Creates an expression rule |

#### ConditionalType Constants

| Constant | Value | Description |
|----------|-------|-------------|
| `ConditionalCellIs` | `"cellIs"` | Cell value comparison |
| `ConditionalExpression` | `"expression"` | Custom expression |
| `ConditionalColorScale` | `"colorScale"` | Color scale |
| `ConditionalDataBar` | `"dataBar"` | Data bar |
| `ConditionalIconSet` | `"iconSet"` | Icon set |
| `ConditionalTop10` | `"top10"` | Top N items |
| `ConditionalAboveAvg` | `"aboveAverage"` | Above average |
| `ConditionalDuplicates` | `"duplicateValues"` | Duplicate values |
| `ConditionalUniqueVals` | `"uniqueValues"` | Unique values |
| `ConditionalContainsText` | `"containsText"` | Contains text |

#### ConditionalOperator Constants

| Constant | Value |
|----------|-------|
| `OperatorEqual` | `"equal"` |
| `OperatorNotEqual` | `"notEqual"` |
| `OperatorGreaterThan` | `"greaterThan"` |
| `OperatorGreaterOrEqual` | `"greaterThanOrEqual"` |
| `OperatorLessThan` | `"lessThan"` |
| `OperatorLessOrEqual` | `"lessThanOrEqual"` |
| `OperatorBetween` | `"between"` |
| `OperatorNotBetween` | `"notBetween"` |

---

## Data Validation

### DataValidation

```go
type DataValidation struct {
    Range         string
    Type          ValidationType
    Operator      ValidationOperator
    Formula1      string
    Formula2      string
    AllowBlank    bool
    ShowInputMsg  bool
    ShowErrorMsg  bool
    ErrorStyle    ValidationErrorStyle
    ErrorTitle    string
    ErrorMessage  string
    PromptTitle   string
    PromptMessage string
}
```

| Function/Method | Description |
|-----------------|-------------|
| `NewDataValidation(rangeStr string) *DataValidation` | Creates data validation (defaults: allow blank, show messages) |
| `SetType(t ValidationType) *DataValidation` | Sets validation type |
| `SetOperator(op ValidationOperator) *DataValidation` | Sets comparison operator |
| `SetFormula1(f string) *DataValidation` | Sets the first formula/value |
| `SetFormula2(f string) *DataValidation` | Sets the second formula/value (for between) |
| `SetErrorMessage(title, message string) *DataValidation` | Sets error message |
| `SetPromptMessage(title, message string) *DataValidation` | Sets input prompt |
| `SetListValues(values []string) *DataValidation` | Sets dropdown list values |

#### ValidationType Constants

| Constant | Value | Description |
|----------|-------|-------------|
| `ValidationNone` | `"none"` | No validation |
| `ValidationWhole` | `"whole"` | Whole number |
| `ValidationDecimal` | `"decimal"` | Decimal |
| `ValidationList` | `"list"` | List |
| `ValidationDate` | `"date"` | Date |
| `ValidationTime` | `"time"` | Time |
| `ValidationTextLength` | `"textLength"` | Text length |
| `ValidationCustom` | `"custom"` | Custom |

#### ValidationErrorStyle Constants

| Constant | Value |
|----------|-------|
| `ErrorStyleStop` | `"stop"` |
| `ErrorStyleWarning` | `"warning"` |
| `ErrorStyleInformation` | `"information"` |

#### ValidationOperator Constants

| Constant | Value |
|----------|-------|
| `ValOperatorBetween` | `"between"` |
| `ValOperatorNotBetween` | `"notBetween"` |
| `ValOperatorEqual` | `"equal"` |
| `ValOperatorNotEqual` | `"notEqual"` |
| `ValOperatorGreaterThan` | `"greaterThan"` |
| `ValOperatorLessThan` | `"lessThan"` |
| `ValOperatorGreaterThanOrEqual` | `"greaterThanOrEqual"` |
| `ValOperatorLessThanOrEqual` | `"lessThanOrEqual"` |

---

## Auto Filter

### AutoFilter

```go
type AutoFilter struct {
    Range   string // e.g. "A1:D100"
    Columns []AutoFilterColumn
}

type AutoFilterColumn struct {
    ColumnIndex int
    FilterType  FilterType
    Conditions  []FilterCondition
    Values      []string
    ShowButton  bool
}

type FilterCondition struct {
    Operator FilterOperator
    Value    string
}
```

| Function/Method | Description |
|-----------------|-------------|
| `NewAutoFilter(rangeStr string) *AutoFilter` | Creates an auto filter |
| `AddColumn(col AutoFilterColumn) *AutoFilter` | Adds a filter column configuration |
| `AddValueFilter(colIndex int, values []string) *AutoFilter` | Adds a value filter |
| `AddCustomFilter(colIndex int, conditions ...FilterCondition) *AutoFilter` | Adds a custom filter |

#### FilterType Constants

| Constant | Value |
|----------|-------|
| `FilterCustom` | `"custom"` |
| `FilterDynamic` | `"dynamic"` |
| `FilterTop10` | `"top10"` |
| `FilterValues` | `"values"` |

#### FilterOperator Constants

| Constant | Value |
|----------|-------|
| `FilterOpEqual` | `"equal"` |
| `FilterOpNotEqual` | `"notEqual"` |
| `FilterOpGreaterThan` | `"greaterThan"` |
| `FilterOpGreaterOrEqual` | `"greaterThanOrEqual"` |
| `FilterOpLessThan` | `"lessThan"` |
| `FilterOpLessOrEqual` | `"lessThanOrEqual"` |

---

## Sheet View

### SheetView

```go
type SheetView struct {
    ZoomScale         uint   // zoom percentage (10-400), default 100
    ShowGridlines     bool
    ShowRowColHeaders bool
    ShowRuler         bool
    TopLeftCell       string // e.g. "A1"
}
```

### SortOrder

```go
type SortOrder int

const (
    SortOrderAscending  SortOrder = 0
    SortOrderDescending SortOrder = 1
)
```

---

## Page Setup

### PageSetup

```go
type PageSetup struct {
    PaperSize          PaperSize
    Orientation        Orientation
    Scale              int         // 10-400%
    FitToWidth         int
    FitToHeight        int
    Margins            PageMargins
    HeaderFooter       *HeaderFooter
    PrintArea          *PrintArea
    PrintGridlines     bool
    PrintHeadings      bool
    CenterHorizontally bool
    CenterVertically   bool
    RepeatRows         string      // e.g. "1:2"
    RepeatColumns      string      // e.g. "A:B"
}
```

| Function/Method | Description |
|-----------------|-------------|
| `NewPageSetup() *PageSetup` | Creates default page setup (A4, portrait, 100%) |
| `SetPaperSize(size PaperSize) *PageSetup` | Sets paper size |
| `SetOrientation(o Orientation) *PageSetup` | Sets orientation |
| `SetScale(scale int) *PageSetup` | Sets scale (10-400) |
| `SetFitToPage(width, height int) *PageSetup` | Sets fit to page |
| `SetMargins(m PageMargins) *PageSetup` | Sets margins |
| `SetPrintArea(rangeStr string) *PageSetup` | Sets print area |
| `SetRepeatRows(rows string) *PageSetup` | Sets repeat rows |
| `SetRepeatColumns(cols string) *PageSetup` | Sets repeat columns |

### PageMargins

```go
type PageMargins struct {
    Top, Bottom, Left, Right, Header, Footer float64 // Unit: inches
}
```

| Function | Description |
|----------|-------------|
| `DefaultPageMargins() PageMargins` | Default margins (top/bottom 0.75, left/right 0.7, header/footer 0.3) |

### HeaderFooter

```go
type HeaderFooter struct {
    OddHeader        string
    OddFooter        string
    EvenHeader       string
    EvenFooter       string
    DifferentOddEven bool
}
```

#### PaperSize Constants

| Constant | Value | Description |
|----------|-------|-------------|
| `PaperLetter` | 1 | Letter (8.5×11") |
| `PaperTabloid` | 3 | Tabloid (11×17") |
| `PaperLegal` | 5 | Legal (8.5×14") |
| `PaperA3` | 8 | A3 (297×420mm) |
| `PaperA4` | 9 | A4 (210×297mm) |
| `PaperA5` | 11 | A5 (148×210mm) |
| `PaperB4` | 12 | B4 (250×353mm) |
| `PaperB5` | 13 | B5 (176×250mm) |

#### Orientation Constants

| Constant | Value |
|----------|-------|
| `OrientationPortrait` | `"portrait"` |
| `OrientationLandscape` | `"landscape"` |

---

## Protection

### SheetProtection

```go
type SheetProtection struct {
    Sheet               bool
    Objects             bool
    Scenarios           bool
    FormatCells         bool
    FormatColumns       bool
    FormatRows          bool
    InsertColumns       bool
    InsertRows          bool
    InsertHyperlinks    bool
    DeleteColumns       bool
    DeleteRows          bool
    SelectLockedCells   bool
    Sort                bool
    AutoFilter          bool
    PivotTables         bool
    SelectUnlockedCells bool
    Password            string // Hashed password
}
```

| Function/Method | Description |
|-----------------|-------------|
| `NewSheetProtection() *SheetProtection` | Creates worksheet protection (defaults: protect Sheet/Objects/Scenarios) |
| `SetPassword(password string) *SheetProtection` | Sets password (SHA-512 hashed) |
| `AllowFormatCells() *SheetProtection` | Allows formatting cells |
| `AllowInsertRows() *SheetProtection` | Allows inserting rows |
| `AllowDeleteRows() *SheetProtection` | Allows deleting rows |
| `AllowSort() *SheetProtection` | Allows sorting |
| `AllowAutoFilter() *SheetProtection` | Allows auto filter |

### WorkbookProtection

```go
type WorkbookProtection struct {
    LockStructure bool
    LockWindows   bool
    Password      string
}
```

| Function/Method | Description |
|-----------------|-------------|
| `NewWorkbookProtection() *WorkbookProtection` | Creates workbook protection (defaults: lock structure) |
| `SetPassword(password string) *WorkbookProtection` | Sets password |
