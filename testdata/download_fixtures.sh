#!/bin/sh
# Download the Excel workbooks used by xlsx_compat_test.go.
# See testdata/README.md for sources and licenses.
set -eu
cd "$(dirname "$0")/excel"
base="https://raw.githubusercontent.com/apache/poi/trunk/test-data/spreadsheet"
while IFS= read -r f; do
  [ -z "$f" ] && continue
  echo "GET $f"
  enc=$(python3 -c 'import urllib.parse,sys; print(urllib.parse.quote(sys.argv[1]))' "$f")
  curl -fsSL "$base/$enc" -o "$f"
done <<'FILES'
SampleSS.xlsx
SampleSS.strict.xlsx
SimpleStrict.xlsx
sample.xlsx
sample.strict.xlsx
InlineString.xlsx
InlineStrings.xlsx
shared_formulas.xlsx
comments.xlsx
SimpleWithComments.xlsx
commentTest.xlsx
WithConditionalFormatting.xlsx
conditional_formatting_cell_is.xlsx
55406_Conditional_formatting_sample.xlsx
DataValidations-49244.xlsx
DateFormatTests.xlsx
NumberFormatTests.xlsx
styles.xlsx
Formatting.xlsx
WidthsAndHeights.xlsx
SheetTabColors.xlsx
TwoSheetsOneHidden.xlsx
sharedhyperlink.xlsx
ExcelTables.xlsx
WithChart.xlsx
WithThreeCharts.xlsx
ExcelPivotTableSample.xlsx
simple-monthly-budget.xlsx
WithVariousData.xlsx
WithMoreVariousData.xlsx
simple-table-named-range.xlsx
RepeatingRowsCols.xlsx
Booleans.xlsx
unicodeSheetName.xlsx
reordered_sheets.xlsx
headerFooterTest.xlsx
Themes.xlsx
54084 - Greek - beyond BMP.xlsx
NewlineInFormulas.xlsx
MatrixFormulaEvalTestData.xlsx
ShrinkToFit.xlsx
GeneralFormatTests.xlsx
ElapsedFormatTests.xlsx
WithDrawing.xlsx
table-sample.xlsx
FormulaEvalTestData_Copy.xlsx
workbookProtection_worksheet_protected.xlsx
noSharedStringTable.xlsx
testSharedFormulasRangeSetBlankBug.xlsx
DataTableCities.xlsx
FormatConditionTests.xlsx
customIndexedColors.xlsx
right-to-left.xlsx
WithTextBox.xlsx
FILES
curl -fsSL "https://raw.githubusercontent.com/qax-os/excelize/master/test/Book1.xlsx" -o Book1.xlsx
curl -fsSL "https://raw.githubusercontent.com/qax-os/excelize/master/test/SharedStrings.xlsx" -o excelize-SharedStrings.xlsx
curl -fsSL "https://raw.githubusercontent.com/qax-os/excelize/master/test/MergeCell.xlsx" -o excelize-MergeCell.xlsx
curl -fsSL "https://raw.githubusercontent.com/qax-os/excelize/master/test/CalcChain.xlsx" -o excelize-CalcChain.xlsx
echo "downloaded $(ls -1 | wc -l) files"
