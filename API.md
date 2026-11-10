> [English](API_en.md) | [中文](API.md)

# GoExcel API 文档

完整的 API 参考手册，涵盖所有公开类型、函数和方法。

---

## 目录

- [核心类型](#核心类型)
  - [Workbook（工作簿）](#workbook工作簿)
  - [Worksheet（工作表）](#worksheet工作表)
  - [Cell（单元格）](#cell单元格)
- [样式系统](#样式系统)
- [坐标工具](#坐标工具)
- [公式计算引擎](#公式计算引擎)
- [文件读写](#文件读写)
- [超链接 / 批注 / 富文本](#超链接--批注--富文本)
- [条件格式](#条件格式)
- [数据验证](#数据验证)
- [自动筛选](#自动筛选)
- [页面设置](#页面设置)
- [保护](#保护)

---

## 核心类型

### Workbook（工作簿）

`Workbook` 是电子表格的顶层容器，包含一个或多个工作表。

#### 类型定义

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

#### 构造函数

| 函数 | 说明 |
|------|------|
| `New() *Workbook` | 创建包含一个默认工作表（Sheet1）的新工作簿 |
| `NewEmpty() *Workbook` | 创建不含工作表的空工作簿 |

#### 方法

| 方法 | 说明 |
|------|------|
| `AddSheet(title string) (*Worksheet, error)` | 添加新工作表，标题不可重复 |
| `GetSheet(index int) (*Worksheet, error)` | 按索引获取工作表（0-based） |
| `GetSheetByName(title string) (*Worksheet, error)` | 按标题获取工作表 |
| `SheetCount() int` | 返回工作表数量 |
| `GetSheetNames() []string` | 返回所有工作表标题 |
| `RemoveSheet(index int) error` | 删除指定索引的工作表（至少保留一个） |
| `GetActiveSheet() *Worksheet` | 返回当前活动工作表 |
| `SetActiveSheet(index int) error` | 设置活动工作表 |
| `AddNamedRange(name, reference string)` | 添加命名范围（如 `"Sheet1!A1:B2"`） |
| `GetNamedRange(name string) (string, error)` | 获取命名范围引用 |
| `GetNamedRanges() map[string]string` | 获取所有命名范围 |

#### 属性

| 字段 | 类型 | 说明 |
|------|------|------|
| `Properties` | `DocumentProperties` | 文档属性（标题、作者等） |

---

### Worksheet（工作表）

`Worksheet` 表示工作簿中的一个工作表。

#### 类型定义

```go
type MergeCell struct {
    StartRow int
    StartCol int
    EndRow   int
    EndCol   int
}
```

#### 构造函数

| 函数 | 说明 |
|------|------|
| `NewWorksheet(title string) *Worksheet` | 创建新工作表（通常通过 `Workbook.AddSheet` 使用） |

#### 基本方法

| 方法 | 说明 |
|------|------|
| `Title() string` | 返回工作表标题 |
| `SetTitle(title string) *Worksheet` | 设置工作表标题 |
| `GetCell(row, col int) *Cell` | 按 0-based 行列索引获取单元格（不存在则自动创建） |
| `GetCellByName(ref string) (*Cell, error)` | 按引用名获取单元格（如 `"A1"`） |
| `SetCellValue(ref string, value interface{}) error` | 设置单元格值 |
| `SetCellFormula(ref string, formula string) error` | 设置单元格公式 |
| `SetCellStyle(ref string, style *Style) error` | 设置单元格样式 |
| `GetCellValue(ref string) (interface{}, error)` | 获取单元格值 |

#### 合并与冻结

| 方法 | 说明 |
|------|------|
| `MergeCells(rangeStr string) error` | 合并单元格范围（如 `"A1:C3"`） |
| `GetMergeCells() []MergeCell` | 获取所有合并区域 |
| `FreezePane(ref string) error` | 冻结窗格（如 `"A2"` 冻结首行） |
| `GetFreezePane() *CellReference` | 获取冻结位置 |

#### 行列尺寸

| 方法 | 说明 |
|------|------|
| `SetColumnWidth(col int, width float64) *Worksheet` | 设置列宽（0-based） |
| `GetColumnWidth(col int) float64` | 获取列宽（默认 8.43） |
| `SetRowHeight(row int, height float64) *Worksheet` | 设置行高（0-based） |
| `GetRowHeight(row int) float64` | 获取行高（默认 15.0） |

#### 行列操作

| 方法 | 说明 |
|------|------|
| `InsertRow(rowIdx int)` | 在指定位置插入空行，已有行下移 |
| `DeleteRow(rowIdx int)` | 删除指定行，下方行上移 |
| `InsertColumn(colIdx int)` | 在指定位置插入空列，已有列右移 |
| `DeleteColumn(colIdx int)` | 删除指定列，右侧列左移 |
| `CopyRow(srcRow, dstRow int)` | 复制行（含值、公式、样式） |

> 行列操作会自动调整合并单元格、行高/列宽。

#### 遍历与统计

| 方法 | 说明 |
|------|------|
| `Dimensions() (minRow, minCol, maxRow, maxCol int, err error)` | 返回已用区域范围（0-based） |
| `RowIterator() ([][]*Cell, error)` | 按行返回所有单元格（空位为 nil） |
| `CellCount() int` | 返回非空单元格数量 |
| `AllCells() []*Cell` | 返回所有非空单元格（按行列排序） |

#### 超链接与批注

| 方法 | 说明 |
|------|------|
| `SetCellHyperlink(ref, url string) error` | 设置单元格超链接 |
| `SetCellComment(ref, author, text string) error` | 设置单元格批注 |

#### 条件格式 / 数据验证 / 筛选

| 方法 | 说明 |
|------|------|
| `AddConditionalFormatting(cf *ConditionalFormatting) *Worksheet` | 添加条件格式 |
| `GetConditionalFormattings() []*ConditionalFormatting` | 获取所有条件格式 |
| `AddDataValidation(dv *DataValidation) *Worksheet` | 添加数据验证 |
| `GetDataValidations() []*DataValidation` | 获取所有数据验证 |
| `SetAutoFilter(af *AutoFilter) *Worksheet` | 设置自动筛选 |
| `GetAutoFilter() *AutoFilter` | 获取自动筛选 |

#### 页面设置与保护

| 方法 | 说明 |
|------|------|
| `SetPageSetup(ps *PageSetup) *Worksheet` | 设置页面配置 |
| `GetPageSetup() *PageSetup` | 获取页面配置（不存在则创建默认值） |
| `SetSheetProtection(sp *SheetProtection) *Worksheet` | 设置工作表保护 |
| `GetSheetProtection() *SheetProtection` | 获取工作表保护 |
| `SetTabColor(color string) *Worksheet` | 设置标签颜色（十六进制，如 `"FF0000"`） |
| `GetTabColor() string` | 获取标签颜色 |

---

### Cell（单元格）

`Cell` 表示工作表中的一个单元格。

#### 类型定义

```go
type CellType int

const (
    CellTypeEmpty   CellType = iota // 空
    CellTypeString                  // 字符串
    CellTypeNumeric                 // 数值
    CellTypeBool                    // 布尔
    CellTypeFormula                 // 公式
    CellTypeDate                    // 日期
    CellTypeError                   // 错误
)
```

#### 构造函数

| 函数 | 说明 |
|------|------|
| `NewCell(row, col int) *Cell` | 创建空单元格（通常通过 `Worksheet.GetCell` 使用） |

#### 方法

| 方法 | 说明 |
|------|------|
| `SetValue(v interface{}) *Cell` | 设置值并自动检测类型（支持 string/int/float64/bool/time.Time） |
| `SetFormula(formula string) *Cell` | 设置公式 |
| `SetStyle(s *Style) *Cell` | 设置样式 |
| `GetStringValue() string` | 获取字符串表示 |
| `GetNumericValue() (float64, error)` | 获取数值（非数值类型返回错误） |
| `GetBoolValue() (bool, error)` | 获取布尔值 |
| `GetDateValue() (time.Time, error)` | 获取日期值 |
| `Row() int` | 返回行索引（0-based） |
| `Col() int` | 返回列索引（0-based） |
| `SetHyperlink(h *Hyperlink) *Cell` | 设置超链接 |
| `SetComment(comment *Comment) *Cell` | 设置批注 |
| `SetRichText(rt *RichText) *Cell` | 设置富文本 |

#### 公开字段

| 字段 | 类型 | 说明 |
|------|------|------|
| `Value` | `interface{}` | 单元格值 |
| `Type` | `CellType` | 值类型 |
| `Formula` | `string` | 公式字符串 |
| `Style` | `*Style` | 单元格样式 |
| `Hyperlink` | `*Hyperlink` | 超链接 |
| `Comment` | `*Comment` | 批注 |
| `RichText` | `*RichText` | 富文本 |

---

## 样式系统

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

| 方法 | 说明 |
|------|------|
| `NewStyle() *Style` | 创建空样式 |
| `SetFont(f *Font) *Style` | 设置字体 |
| `SetFill(f *Fill) *Style` | 设置填充 |
| `SetBorders(b *Borders) *Style` | 设置边框 |
| `SetAlignment(a *Alignment) *Style` | 设置对齐 |
| `SetNumberFormat(nf *NumberFormat) *Style` | 设置数字格式 |

### Font

```go
type Font struct {
    Name          string  // 字体名称（默认 "Calibri"）
    Size          float64 // 字号（默认 11）
    Bold          bool
    Italic        bool
    Underline     bool
    Strikethrough bool
    Color         string  // 十六进制颜色，如 "FF0000"
}
```

| 函数 | 说明 |
|------|------|
| `DefaultFont() *Font` | 返回默认字体（Calibri, 11pt） |

### Fill

```go
type Fill struct {
    Type    string // "solid", "pattern", "none"
    Color   string // 十六进制颜色
    Pattern string // 图案类型
}
```

### Borders

```go
type Borders struct {
    Left   Border
    Right  Border
    Top    Border
    Bottom Border
}

type Border struct {
    Style BorderStyle
    Color string
}
```

#### BorderStyle 常量

| 常量 | 值 |
|------|------|
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
    TextRotation int // -90 到 90 度
    Indent       int
}
```

#### HorizontalAlignment 常量

| 常量 | 值 |
|------|------|
| `AlignLeft` | `"left"` |
| `AlignCenter` | `"center"` |
| `AlignRight` | `"right"` |
| `AlignJustify` | `"justify"` |
| `AlignGeneral` | `"general"` |

#### VerticalAlignment 常量

| 常量 | 值 |
|------|------|
| `AlignTop` | `"top"` |
| `AlignMiddle` | `"center"` |
| `AlignBottom` | `"bottom"` |

### NumberFormat

```go
type NumberFormat struct {
    FormatCode string
}
```

#### 预定义数字格式

| 变量 | FormatCode | 说明 |
|------|------------|------|
| `FormatGeneral` | `"General"` | 常规 |
| `FormatNumber` | `"0"` | 整数 |
| `FormatNumber2Dec` | `"0.00"` | 两位小数 |
| `FormatPercent` | `"0%"` | 百分比 |
| `FormatPercent2Dec` | `"0.00%"` | 两位小数百分比 |
| `FormatDate` | `"yyyy-mm-dd"` | 日期 |
| `FormatDateTime` | `"yyyy-mm-dd hh:mm:ss"` | 日期时间 |
| `FormatTime` | `"hh:mm:ss"` | 时间 |
| `FormatCurrency` | `#,##0.00"$"` | 货币 |
| `FormatAccounting` | `_("$"* #,##0.00_)` | 会计 |
| `FormatText` | `"@"` | 文本 |

---

## 坐标工具

`coordinate.go` 提供单元格引用解析和转换工具。

### CellReference

```go
type CellReference struct {
    Column    string // 列名，如 "A"
    ColumnIdx int    // 0-based 列索引
    Row       int    // 1-based 行号
}
```

### 函数

| 函数 | 说明 |
|------|------|
| `ColumnIndexToName(index int) (string, error)` | 列索引转列名（0→"A", 25→"Z", 26→"AA"） |
| `ColumnNameToIndex(name string) (int, error)` | 列名转列索引（"A"→0, "Z"→25, "AA"→26） |
| `ParseCellReference(ref string) (*CellReference, error)` | 解析单元格引用（如 `"A1"`） |
| `CellName(row, col int) (string, error)` | 0-based 行列转引用名（如 `(0,0)` → `"A1"`） |
| `ParseRange(rangeStr string) (*CellReference, *CellReference, error)` | 解析范围（如 `"A1:C3"`） |

---

## 公式计算引擎

### CalculationEngine

| 函数/方法 | 说明 |
|-----------|------|
| `NewCalculationEngine(wb *Workbook) *CalculationEngine` | 创建计算引擎 |
| `CalculateCell(ws *Worksheet, ref string) (interface{}, error)` | 计算单个单元格公式 |
| `CalculateAll() error` | 计算工作簿中所有公式单元格 |

### 内置函数（24 个）

#### 数学函数

| 函数 | 语法 | 说明 |
|------|------|------|
| SUM | `SUM(range)` | 求和 |
| ABS | `ABS(value)` | 绝对值 |
| ROUND | `ROUND(value, digits)` | 四舍五入到指定小数位 |
| SQRT | `SQRT(value)` | 平方根（负数返回 `#NUM!`） |
| POWER | `POWER(base, exp)` | 幂运算 |
| MOD | `MOD(num, divisor)` | 取模（除数为 0 返回 `#DIV/0!`） |
| INT | `INT(value)` | 向下取整 |

#### 统计函数

| 函数 | 语法 | 说明 |
|------|------|------|
| AVERAGE | `AVERAGE(range)` | 平均值（空范围返回 `#DIV/0!`） |
| COUNT | `COUNT(range)` | 计数数值单元格 |
| COUNTA | `COUNTA(range)` | 计数非空单元格 |
| MAX | `MAX(range)` | 最大值 |
| MIN | `MIN(range)` | 最小值 |
| MEDIAN | `MEDIAN(range)` | 中位数 |

#### 逻辑函数

| 函数 | 语法 | 说明 |
|------|------|------|
| IF | `IF(condition, true_val [, false_val])` | 条件判断（2 或 3 个参数） |

#### 文本函数

| 函数 | 语法 | 说明 |
|------|------|------|
| LEN | `LEN(text)` | 字符数（支持 Unicode） |
| UPPER | `UPPER(text)` | 转大写 |
| LOWER | `LOWER(text)` | 转小写 |
| TRIM | `TRIM(text)` | 去除首尾空白 |
| LEFT | `LEFT(text, count)` | 左截取 |
| RIGHT | `RIGHT(text, count)` | 右截取 |
| MID | `MID(text, start, length)` | 中间截取（start 从 1 开始） |
| CONCATENATE | `CONCATENATE(val1, val2, ...)` | 字符串连接 |

#### 条件函数

| 函数 | 语法 | 说明 |
|------|------|------|
| SUMIF | `SUMIF(criteria_range, criteria [, sum_range])` | 条件求和 |
| COUNTIF | `COUNTIF(range, criteria)` | 条件计数 |

条件支持：精确匹配、`">N"`、`"<N"`、`">=N"`、`"<=N"`、`"<>N"`

### 运算符支持

公式支持四则运算：`+`、`-`、`*`、`/`，以及单元格引用和嵌套函数调用。

---

## 文件读写

### IOFactory（自动格式检测）

| 函数 | 说明 |
|------|------|
| `OpenFile(filename string, opts ...Option) (*Workbook, error)` | 按扩展名自动选择读取器。.xlsx 可用 `WithPassword` 打开 ECMA-376 加密文件 |
| `SaveFile(wb *Workbook, filename string) error` | 按扩展名自动选择写入器。`SetPassword` 会写成 Agile 加密包 |

### XLSXReader / XLSXWriter

| 函数/方法 | 说明 |
|-----------|------|
| `NewXLSXReader() *XLSXReader` | 创建 XLSX 读取器 |
| `(*XLSXReader) Open(filename string) (*Workbook, error)` | 从文件读取 |
| `(*XLSXReader) Read(reader io.ReaderAt, size int64) (*Workbook, error)` | 从 io.ReaderAt 读取。加密包返回 `ErrPasswordRequired` |
| `(*XLSXReader) OpenPassword / ReadPassword` | 解密 Agile 或 Standard 加密后再读取 |
| `DecryptPackage` / `EncryptPackage` | 解密或按 Agile 加密 OOXML 包 |
| `NewXLSXWriter() *XLSXWriter` | 创建 XLSX 写入器 |
| `(*XLSXWriter) Save(wb *Workbook, filename string) error` | 保存到文件 |
| `(*XLSXWriter) Write(wb *Workbook, writer io.Writer) error` | 写入到 io.Writer。设置了密码时写出 OLE 加密包；`SetStrictOOXML(true)` 写出严格 OOXML，默认仍是过渡命名空间。 |

### 数组公式、迷你图、切片器、数据透视表

| API | 说明 |
|-----|------|
| `(*Worksheet) SetArrayFormula(rangeRef, formula string) error` | 在锚点单元格写入数组公式，区域内每个单元格都报告该公式 |
| `(*Worksheet) FormulaAt(ref string) (formula, arrayRef string, err error)` | 返回 Excel 公式栏会显示的公式；若属于数组公式，同时返回 ref |
| `(*Cell) ArrayRef / IsArrayFormula / ReportedFormula` | 数组区域与公式文本，包括只落在 ref 内的单元格 |
| `(*Worksheet) SparklineGroups / SetSparklineGroups` | 读取或替换迷你图（类型、标记、颜色、公式、单元格） |
| `(*Worksheet) AddPivotTable(name, source, dest, rowField, dataField string)` | 从区域建立透视缓存，并在 `dest` 写入透视表 |
| `(*Worksheet) AddSlicer(pivotName, field, caption string)` | 为该表上的透视表字段新建切片器及缓存 |
| `(*Worksheet) PivotTables` | 透视表名称、位置、数据标题、样式和字段。`SetName`、`SetLocation`、`SetDataCaption`、`SetStyle` 写回这些属性 |
| `(*Worksheet) Slicers` / `(*Workbook) SlicerCaches` | 标题、样式、列数、行高，以及缓存的源字段名 |

### CSVReader / CSVWriter

| 函数/方法 | 说明 |
|-----------|------|
| `NewCSVReader() *CSVReader` | 创建 CSV 读取器（默认逗号分隔） |
| `(*CSVReader) Open(filename string) (*Workbook, error)` | 从文件读取 |
| `(*CSVReader) Read(reader io.Reader) (*Workbook, error)` | 从 io.Reader 读取 |
| `NewCSVWriter() *CSVWriter` | 创建 CSV 写入器 |
| `(*CSVWriter) Save(wb *Workbook, filename string) error` | 保存到文件 |
| `(*CSVWriter) Write(wb *Workbook, writer io.Writer) error` | 写入到 io.Writer |

#### CSVReader 字段

| 字段 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `Delimiter` | `rune` | `','` | 字段分隔符 |
| `LazyQuotes` | `bool` | `false` | 宽松引号解析 |

#### CSVWriter 字段

| 字段 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `Delimiter` | `rune` | `','` | 字段分隔符 |
| `SheetIndex` | `int` | `0` | 要写入的工作表索引 |

---

## 超链接 / 批注 / 富文本

### Hyperlink

```go
type Hyperlink struct {
    URL     string
    Tooltip string
}
```

| 函数/方法 | 说明 |
|-----------|------|
| `NewHyperlink(url string) *Hyperlink` | 创建超链接 |
| `SetTooltip(tooltip string) *Hyperlink` | 设置提示文本 |

### Comment

```go
type Comment struct {
    Author string
    Text   string
}
```

| 函数 | 说明 |
|------|------|
| `NewComment(author, text string) *Comment` | 创建批注 |

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

| 函数/方法 | 说明 |
|-----------|------|
| `NewRichText() *RichText` | 创建空富文本 |
| `AddRun(text string, font *Font) *RichText` | 添加文本段（可带字体） |
| `PlainText() string` | 返回纯文本内容 |

---

## 条件格式

### ConditionalFormatting

```go
type ConditionalFormatting struct {
    Range string           // 如 "A1:A10"
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

| 函数/方法 | 说明 |
|-----------|------|
| `NewConditionalFormatting(rangeStr string) *ConditionalFormatting` | 创建条件格式 |
| `AddRule(rule ConditionalRule) *ConditionalFormatting` | 添加规则 |
| `CellIsRule(op ConditionalOperator, formula string, style *Style) ConditionalRule` | 创建 cellIs 规则 |
| `BetweenRule(formula1, formula2 string, style *Style) ConditionalRule` | 创建 between 规则 |
| `ExpressionRule(formula string, style *Style) ConditionalRule` | 创建表达式规则 |

#### ConditionalType 常量

| 常量 | 值 | 说明 |
|------|------|------|
| `ConditionalCellIs` | `"cellIs"` | 单元格值比较 |
| `ConditionalExpression` | `"expression"` | 自定义表达式 |
| `ConditionalColorScale` | `"colorScale"` | 色阶 |
| `ConditionalDataBar` | `"dataBar"` | 数据条 |
| `ConditionalIconSet` | `"iconSet"` | 图标集 |
| `ConditionalTop10` | `"top10"` | 前 N 项 |
| `ConditionalAboveAvg` | `"aboveAverage"` | 高于平均值 |
| `ConditionalDuplicates` | `"duplicateValues"` | 重复值 |
| `ConditionalUniqueVals` | `"uniqueValues"` | 唯一值 |
| `ConditionalContainsText` | `"containsText"` | 包含文本 |

#### ConditionalOperator 常量

| 常量 | 值 |
|------|------|
| `OperatorEqual` | `"equal"` |
| `OperatorNotEqual` | `"notEqual"` |
| `OperatorGreaterThan` | `"greaterThan"` |
| `OperatorGreaterOrEqual` | `"greaterThanOrEqual"` |
| `OperatorLessThan` | `"lessThan"` |
| `OperatorLessOrEqual` | `"lessThanOrEqual"` |
| `OperatorBetween` | `"between"` |
| `OperatorNotBetween` | `"notBetween"` |

---

## 数据验证

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

| 函数/方法 | 说明 |
|-----------|------|
| `NewDataValidation(rangeStr string) *DataValidation` | 创建数据验证（默认允许空白、显示消息） |
| `SetType(t ValidationType) *DataValidation` | 设置验证类型 |
| `SetOperator(op ValidationOperator) *DataValidation` | 设置比较运算符 |
| `SetFormula1(f string) *DataValidation` | 设置第一个公式/值 |
| `SetFormula2(f string) *DataValidation` | 设置第二个公式/值（between 用） |
| `SetErrorMessage(title, message string) *DataValidation` | 设置错误提示 |
| `SetPromptMessage(title, message string) *DataValidation` | 设置输入提示 |
| `SetListValues(values []string) *DataValidation` | 设置下拉列表值 |

#### ValidationType 常量

| 常量 | 值 | 说明 |
|------|------|------|
| `ValidationNone` | `"none"` | 无验证 |
| `ValidationWhole` | `"whole"` | 整数 |
| `ValidationDecimal` | `"decimal"` | 小数 |
| `ValidationList` | `"list"` | 列表 |
| `ValidationDate` | `"date"` | 日期 |
| `ValidationTime` | `"time"` | 时间 |
| `ValidationTextLength` | `"textLength"` | 文本长度 |
| `ValidationCustom` | `"custom"` | 自定义 |

#### ValidationErrorStyle 常量

| 常量 | 值 |
|------|------|
| `ErrorStyleStop` | `"stop"` |
| `ErrorStyleWarning` | `"warning"` |
| `ErrorStyleInformation` | `"information"` |

#### ValidationOperator 常量

| 常量 | 值 |
|------|------|
| `ValOperatorBetween` | `"between"` |
| `ValOperatorNotBetween` | `"notBetween"` |
| `ValOperatorEqual` | `"equal"` |
| `ValOperatorNotEqual` | `"notEqual"` |
| `ValOperatorGreaterThan` | `"greaterThan"` |
| `ValOperatorLessThan` | `"lessThan"` |
| `ValOperatorGreaterThanOrEqual` | `"greaterThanOrEqual"` |
| `ValOperatorLessThanOrEqual` | `"lessThanOrEqual"` |

---

## 自动筛选

### AutoFilter

```go
type AutoFilter struct {
    Range   string // 如 "A1:D100"
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

| 函数/方法 | 说明 |
|-----------|------|
| `NewAutoFilter(rangeStr string) *AutoFilter` | 创建自动筛选 |
| `AddColumn(col AutoFilterColumn) *AutoFilter` | 添加筛选列配置 |
| `AddValueFilter(colIndex int, values []string) *AutoFilter` | 添加值筛选 |
| `AddCustomFilter(colIndex int, conditions ...FilterCondition) *AutoFilter` | 添加自定义筛选 |

#### FilterType 常量

| 常量 | 值 |
|------|------|
| `FilterCustom` | `"custom"` |
| `FilterDynamic` | `"dynamic"` |
| `FilterTop10` | `"top10"` |
| `FilterValues` | `"values"` |

#### FilterOperator 常量

| 常量 | 值 |
|------|------|
| `FilterOpEqual` | `"equal"` |
| `FilterOpNotEqual` | `"notEqual"` |
| `FilterOpGreaterThan` | `"greaterThan"` |
| `FilterOpGreaterOrEqual` | `"greaterThanOrEqual"` |
| `FilterOpLessThan` | `"lessThan"` |
| `FilterOpLessOrEqual` | `"lessThanOrEqual"` |

---

## 页面设置

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
    RepeatRows         string      // 如 "1:2"
    RepeatColumns      string      // 如 "A:B"
}
```

| 函数/方法 | 说明 |
|-----------|------|
| `NewPageSetup() *PageSetup` | 创建默认页面设置（A4、纵向、100%） |
| `SetPaperSize(size PaperSize) *PageSetup` | 设置纸张大小 |
| `SetOrientation(o Orientation) *PageSetup` | 设置方向 |
| `SetScale(scale int) *PageSetup` | 设置缩放比例（10-400） |
| `SetFitToPage(width, height int) *PageSetup` | 设置适合页面 |
| `SetMargins(m PageMargins) *PageSetup` | 设置页边距 |
| `SetPrintArea(rangeStr string) *PageSetup` | 设置打印区域 |
| `SetRepeatRows(rows string) *PageSetup` | 设置重复行 |
| `SetRepeatColumns(cols string) *PageSetup` | 设置重复列 |

### PageMargins

```go
type PageMargins struct {
    Top, Bottom, Left, Right, Header, Footer float64 // 单位：英寸
}
```

| 函数 | 说明 |
|------|------|
| `DefaultPageMargins() PageMargins` | 默认页边距（上下 0.75, 左右 0.7, 页眉页脚 0.3） |

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

#### PaperSize 常量

| 常量 | 值 | 说明 |
|------|------|------|
| `PaperLetter` | 1 | Letter (8.5×11") |
| `PaperTabloid` | 3 | Tabloid (11×17") |
| `PaperLegal` | 5 | Legal (8.5×14") |
| `PaperA3` | 8 | A3 (297×420mm) |
| `PaperA4` | 9 | A4 (210×297mm) |
| `PaperA5` | 11 | A5 (148×210mm) |
| `PaperB4` | 12 | B4 (250×353mm) |
| `PaperB5` | 13 | B5 (176×250mm) |

#### Orientation 常量

| 常量 | 值 |
|------|------|
| `OrientationPortrait` | `"portrait"` |
| `OrientationLandscape` | `"landscape"` |

---

## 保护

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
    Password            string // 哈希后的密码
}
```

| 函数/方法 | 说明 |
|-----------|------|
| `NewSheetProtection() *SheetProtection` | 创建工作表保护（默认保护 Sheet/Objects/Scenarios） |
| `SetPassword(password string) *SheetProtection` | 设置密码（SHA-512 哈希） |
| `AllowFormatCells() *SheetProtection` | 允许格式化单元格 |
| `AllowInsertRows() *SheetProtection` | 允许插入行 |
| `AllowDeleteRows() *SheetProtection` | 允许删除行 |
| `AllowSort() *SheetProtection` | 允许排序 |
| `AllowAutoFilter() *SheetProtection` | 允许自动筛选 |

### WorkbookProtection

```go
type WorkbookProtection struct {
    LockStructure bool
    LockWindows   bool
    Password      string
}
```

| 函数/方法 | 说明 |
|-----------|------|
| `NewWorkbookProtection() *WorkbookProtection` | 创建工作簿保护（默认锁定结构） |
| `SetPassword(password string) *WorkbookProtection` | 设置密码 |
