package gospreadsheet

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"io"
	"math"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"testing"
	"time"
)

func TestShiftFormulaSharedStyle(t *testing.T) {
	cases := []struct {
		in         string
		dCol, dRow int
		want       string
	}{
		{"B2", 0, 1, "B3"},
		{"A1+B$1", 1, 1, "B2+C$1"},
		{"$A1+A$1", 1, 2, "$A3+B$1"},
		{`LOG10(A1)+SUM(B2)`, 1, 1, `LOG10(B2)+SUM(C3)`},
		{`"A1"+A1`, 1, 0, `"A1"+B1`},
		{"Sheet1!A1", 1, 1, "Sheet1!B2"},
		{"SUM(F:G)", 1, 2, "SUM(G:H)"},
		{"SUM($F:G)", 1, 0, "SUM($F:H)"},
		{"SUM(6:7)", 3, 1, "SUM(7:8)"},
		{"SUM($6:$7)", 1, 4, "SUM($6:$7)"},
		{"'My Sheet'!$A$1", 2, 3, "'My Sheet'!$A$1"},
	}
	for _, tc := range cases {
		got := shiftFormula(tc.in, tc.dCol, tc.dRow)
		if got != tc.want {
			t.Errorf("shiftFormula(%q, %d, %d) = %q, want %q", tc.in, tc.dCol, tc.dRow, got, tc.want)
		}
	}
}

func TestXLSXWriterPackageShape(t *testing.T) {
	wb := New()
	wb.Properties.Creator = "GoExcel"
	wb.Properties.Title = "Compat"
	wb.SetDate1904(false)
	ws := wb.GetActiveSheet()
	ws.SetCellValue("A1", " Name ")
	ws.SetCellValue("B1", 42)
	ws.SetCellValue("C1", true)
	ws.SetCellValue("D1", time.Date(2024, 1, 2, 15, 4, 5, 0, time.UTC))
	ws.SetCellFormula("E1", "=SUM(B1)")
	ws.MergeCells("A2:B2")
	ws.SetCellValue("A2", "merged")
	ws.FreezePane("B2")
	ws.SetColumnWidth(0, 22)
	ws.SetRowHeight(0, 30)
	ws.SetRowHidden(4, true)
	ws.SetColumnHidden(3, true)
	ws.SetCellHyperlink("A1", "https://example.com/a&b")
	ws.SetCellComment("B1", "Ada", "check this")
	ws.SetCellStyle("A1", NewStyle().SetFont(&Font{Name: "Arial", Size: 14, Bold: true, Color: "FF0000"}).SetFill(&Fill{Type: "solid", Color: "FFFF00"}))
	rt := NewRichText().AddRun("Hello ", nil).AddRun("world", &Font{Bold: true, Color: "0000FF"})
	ws.GetCell(2, 0).SetRichText(rt)
	wb.AddNamedRange("Total", "Sheet1!$B$1")
	ws.SetSheetState("visible")
	cf := NewConditionalFormatting("B1")
	cf.AddRule(CellIsRule(OperatorGreaterThan, "10", NewStyle().SetFont(&Font{Color: "00AA00"})))
	ws.AddConditionalFormatting(cf)
	dv := NewDataValidation("B1").SetType(ValidationWhole).SetOperator(ValOperatorGreaterThan).SetFormula1("0")
	ws.AddDataValidation(dv)
	ws.SetAutoFilter(NewAutoFilter("A1:B10").AddValueFilter(0, []string{" Name "}))
	ws.SetPageSetup(NewPageSetup().SetOrientation(OrientationLandscape).SetPaperSize(PaperA4))
	ws2, _ := wb.AddSheet("Other")
	ws2.SetCellValue("A1", "second")
	ws2.SetSheetState("hidden")

	var buf bytes.Buffer
	if err := NewXLSXWriter().Write(wb, &buf); err != nil {
		t.Fatal(err)
	}
	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}
	files := map[string][]byte{}
	var order []string
	for _, f := range zr.File {
		rc, _ := f.Open()
		b, _ := io.ReadAll(rc)
		rc.Close()
		files[f.Name] = b
		order = append(order, f.Name)
	}
	if order[0] != "[Content_Types].xml" {
		t.Fatalf("first zip entry = %s, want [Content_Types].xml", order[0])
	}
	for _, name := range []string{"[Content_Types].xml", "_rels/.rels", "xl/workbook.xml", "xl/_rels/workbook.xml.rels", "xl/styles.xml", "xl/sharedStrings.xml", "xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml", "docProps/core.xml", "docProps/app.xml", "xl/comments1.xml", "xl/drawings/vmlDrawing1.vml", "xl/worksheets/_rels/sheet1.xml.rels"} {
		if _, ok := files[name]; !ok {
			t.Errorf("missing part %s", name)
		}
	}
	assertChildOrder(t, files["xl/workbook.xml"], []string{
		"fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets",
		"functionGroups", "externalReferences", "definedNames", "calcPr", "pivotCaches", "extLst",
	})
	assertChildOrder(t, files["xl/worksheets/sheet1.xml"], []string{
		"sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetProtection",
		"autoFilter", "mergeCells", "conditionalFormatting", "dataValidations", "hyperlinks",
		"printOptions", "pageMargins", "pageSetup", "headerFooter", "drawing", "legacyDrawing", "extLst",
	})
	assertChildOrder(t, files["xl/styles.xml"], []string{
		"numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs", "tableStyles", "colors",
	})
	sheet := string(files["xl/worksheets/sheet1.xml"])
	if !strings.Contains(sheet, `xml:space="preserve"`) {
		t.Error("shared string or inline text with surrounding spaces was not preserved")
	}
	sst := string(files["xl/sharedStrings.xml"])
	if !strings.Contains(sst, `xml:space="preserve"`) || !strings.Contains(sst, " Name ") {
		t.Errorf("shared string table lost surrounding spaces: %s", sst)
	}
	if !strings.Contains(sheet, `t="array"`) && strings.Contains(sheet, "<f>") && !strings.Contains(sheet, "<f>SUM(B1)</f>") && !strings.Contains(sheet, "<f>SUM(B1)</f>") {
		// formula must not keep the leading '='
	}
	if strings.Contains(sheet, "<f>=SUM") {
		t.Error("formula was written with a leading equals sign")
	}
	if !strings.Contains(sheet, `state="frozen"`) {
		t.Error("frozen pane was not written")
	}
	if !strings.Contains(sheet, `<legacyDrawing r:id=`) {
		t.Error("comment legacy drawing was not written")
	}
	styles := string(files["xl/styles.xml"])
	if !strings.Contains(styles, `patternType="none"`) || !strings.Contains(styles, `patternType="gray125"`) {
		t.Error("styles.xml is missing the two required default fills")
	}
	if !strings.Contains(styles, `name="Normal"`) {
		t.Error("styles.xml is missing the Normal cell style")
	}
	wbXML := string(files["xl/workbook.xml"])
	if !strings.Contains(wbXML, `<calcPr calcId="191029"`) {
		t.Error("workbook is missing calcPr")
	}
	if !strings.Contains(wbXML, `state="hidden"`) {
		t.Error("hidden sheet state was not written")
	}
	if !strings.Contains(wbXML, `name="Total"`) {
		t.Error("defined name was not written")
	}
	// Column records must be in ascending order.
	cols := colsBetween(sheet, "<cols>", "</cols>")
	last := 0
	for _, min := range colMins(cols) {
		if min < last {
			t.Fatalf("cols out of order: %s", cols)
		}
		last = min
	}

	// Round-trip the values.
	got, err := NewXLSXReader().Read(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}
	if got.SheetCount() != 2 {
		t.Fatalf("sheets = %d", got.SheetCount())
	}
	other, err := got.GetSheet(1)
	if err != nil {
		t.Fatal(err)
	}
	if got.GetSheetNames()[1] != "Other" || other.GetSheetState() != "hidden" {
		t.Fatalf("sheet state/name = %v %s", got.GetSheetNames(), other.GetSheetState())
	}
	s := got.GetActiveSheet()
	if v, _ := s.GetCellValue("A1"); v != " Name " {
		t.Errorf("A1 = %#v", v)
	}
	if v, _ := s.GetCellValue("B1"); v != float64(42) {
		t.Errorf("B1 = %#v", v)
	}
	if v, _ := s.GetCellValue("C1"); v != true {
		t.Errorf("C1 = %#v", v)
	}
	d, err := s.GetCell(0, 3).GetDateValue()
	if err != nil {
		t.Fatal(err)
	}
	wantDate := time.Date(2024, 1, 2, 15, 4, 5, 0, time.UTC)
	if d.Sub(wantDate).Abs() > time.Second {
		t.Errorf("D1 = %v, want %v", d, wantDate)
	}
	fcell, _ := s.GetCellByName("E1")
	if fcell.Formula != "SUM(B1)" {
		t.Errorf("formula = %q", fcell.Formula)
	}
	if s.GetCell(0, 0).Hyperlink == nil || s.GetCell(0, 0).Hyperlink.URL != "https://example.com/a&b" {
		t.Errorf("hyperlink = %#v", s.GetCell(0, 0).Hyperlink)
	}
	if s.GetCell(0, 1).Comment == nil || s.GetCell(0, 1).Comment.Text != "check this" || s.GetCell(0, 1).Comment.Author != "Ada" {
		t.Errorf("comment = %#v", s.GetCell(0, 1).Comment)
	}
	if s.GetCell(2, 0).GetStringValue() != "Hello world" {
		t.Errorf("rich text = %#v", s.GetCell(2, 0).Value)
	}
	if _, err := got.GetNamedRange("Total"); err != nil {
		t.Error(err)
	}
	if len(s.GetMergeCells()) != 1 {
		t.Errorf("merges = %d", len(s.GetMergeCells()))
	}
	if !s.IsRowHidden(4) || !s.IsColumnHidden(3) {
		t.Error("hidden row/column was not round-tripped")
	}
	if s.GetColumnWidth(0) != 22 {
		t.Errorf("width = %v", s.GetColumnWidth(0))
	}
	if s.GetAutoFilter() == nil || s.GetAutoFilter().Range != "A1:B10" {
		t.Errorf("autofilter = %#v", s.GetAutoFilter())
	}
	if len(s.GetDataValidations()) != 1 || s.GetDataValidations()[0].Formula1 != "0" {
		t.Errorf("validation = %#v", s.GetDataValidations())
	}
	if len(s.GetConditionalFormattings()) != 1 {
		t.Errorf("cf = %d", len(s.GetConditionalFormattings()))
	}
	font := s.GetCell(0, 0).Style
	if font == nil || font.Font == nil || !font.Font.Bold || font.Font.Name != "Arial" {
		t.Errorf("style = %#v", font)
	}
	if font != nil && font.Font != nil && font.Font.Color != "FFFF0000" && font.Font.Color != "FF0000" {
		t.Errorf("font color = %q", font.Font.Color)
	}
}

func TestXLSXIllegalXMLCharacters(t *testing.T) {
	wb := New()
	ws := wb.GetActiveSheet()
	ws.SetCellValue("A1", "a\x01b\x00c\n d")
	var buf bytes.Buffer
	if err := NewXLSXWriter().Write(wb, &buf); err != nil {
		t.Fatal(err)
	}
	got, err := NewXLSXReader().Read(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}
	v, _ := got.GetActiveSheet().GetCellValue("A1")
	if v != "abc\n d" {
		t.Errorf("sanitized value = %#v", v)
	}
}

func (ws *Worksheet) MustState(t *testing.T) string {
	t.Helper()
	if ws == nil {
		t.Fatal("nil sheet")
	}
	return ws.GetSheetState()
}

func assertChildOrder(t *testing.T, data []byte, allowed []string) {
	t.Helper()
	rank := map[string]int{}
	for i, name := range allowed {
		rank[name] = i
	}
	dec := xml.NewDecoder(bytes.NewReader(data))
	var root xml.StartElement
	for {
		tok, err := dec.Token()
		if err != nil {
			t.Fatalf("xml: %v", err)
		}
		if se, ok := tok.(xml.StartElement); ok {
			root = se
			break
		}
	}
	_ = root
	prev := -1
	prevName := ""
	for {
		tok, err := dec.Token()
		if err != nil {
			break
		}
		se, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		r, known := rank[se.Name.Local]
		if !known {
			t.Errorf("unexpected child %s", se.Name.Local)
		} else if r < prev {
			t.Errorf("element %s appears after %s", se.Name.Local, prevName)
		} else {
			prev = r
			prevName = se.Name.Local
		}
		if err := skipElement(dec); err != nil {
			t.Fatal(err)
		}
	}
}

func colsBetween(s, a, b string) string {
	i := strings.Index(s, a)
	j := strings.Index(s, b)
	if i < 0 || j < 0 || j < i {
		return ""
	}
	return s[i:j]
}

func colMins(cols string) []int {
	var out []int
	rest := cols
	for {
		i := strings.Index(rest, `min="`)
		if i < 0 {
			return out
		}
		rest = rest[i+5:]
		j := strings.IndexByte(rest, '"')
		if j < 0 {
			return out
		}
		n, _ := strconv.Atoi(rest[:j])
		out = append(out, n)
		rest = rest[j:]
	}
}

func TestFixtureReadAndRoundTrip(t *testing.T) {
	dir := filepath.Join("testdata", "excel")
	var paths []string
	err := filepath.WalkDir(dir, func(path string, d os.DirEntry, err error) error {
		if err != nil || d.IsDir() {
			return nil
		}
		switch strings.ToLower(filepath.Ext(path)) {
		case ".xlsx", ".xlsm", ".xltx", ".xltm":
			paths = append(paths, path)
		}
		return nil
	})
	if err != nil {
		t.Fatal(err)
	}
	for _, path := range paths {
		name := strings.TrimPrefix(path, dir+string(filepath.Separator))
		t.Run(name, func(t *testing.T) {
			wb, err := OpenFile(path)
			if err != nil {
				t.Fatalf("open: %v", err)
			}
			if wb.SheetCount() == 0 {
				t.Fatal("workbook has no sheets")
			}
			var buf bytes.Buffer
			if err := NewXLSXWriter().Write(wb, &buf); err != nil {
				t.Fatalf("write: %v", err)
			}
			wb2, err := NewXLSXReader().Read(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
			if err != nil {
				t.Fatalf("reopen: %v", err)
			}
			compareWorkbooks(t, wb, wb2)
			assertPackagePartsPreserved(t, path, buf.Bytes())
		})
	}
}

func compareWorkbooks(t *testing.T, a, b *Workbook) {
	t.Helper()
	if a.SheetCount() != b.SheetCount() {
		t.Fatalf("sheet count %d vs %d", a.SheetCount(), b.SheetCount())
	}
	for i := 0; i < a.SheetCount(); i++ {
		sa, _ := a.GetSheet(i)
		sb, _ := b.GetSheet(i)
		if sa.Title() != sb.Title() {
			t.Errorf("sheet %d name %q vs %q", i, sa.Title(), sb.Title())
		}
		if sa.GetSheetState() != sb.GetSheetState() {
			t.Errorf("sheet %s state %s vs %s", sa.Title(), sa.GetSheetState(), sb.GetSheetState())
		}
		if len(sa.GetMergeCells()) != len(sb.GetMergeCells()) {
			t.Errorf("sheet %s merges %d vs %d", sa.Title(), len(sa.GetMergeCells()), len(sb.GetMergeCells()))
		}
		amap := map[string]*Cell{}
		for _, c := range sa.AllCells() {
			ref, _ := CellName(c.row, c.col)
			amap[ref] = c
		}
		bmap := map[string]*Cell{}
		for _, c := range sb.AllCells() {
			ref, _ := CellName(c.row, c.col)
			bmap[ref] = c
		}
		if len(amap) != len(bmap) {
			t.Errorf("sheet %s cells %d vs %d", sa.Title(), len(amap), len(bmap))
		}
		for ref, ca := range amap {
			cb := bmap[ref]
			if cb == nil {
				t.Errorf("sheet %s missing %s after round-trip", sa.Title(), ref)
				continue
			}
			if ca.Formula != cb.Formula {
				t.Errorf("sheet %s %s formula %q vs %q", sa.Title(), ref, ca.Formula, cb.Formula)
			}
			if !valuesClose(ca.Value, cb.Value) {
				t.Errorf("sheet %s %s value %#v vs %#v", sa.Title(), ref, ca.Value, cb.Value)
			}
		}
	}
	if len(a.GetNamedRanges()) != len(b.GetNamedRanges()) {
		t.Errorf("named ranges %d vs %d (%v vs %v)", len(a.GetNamedRanges()), len(b.GetNamedRanges()), a.GetNamedRanges(), b.GetNamedRanges())
	}
}

func valuesClose(a, b interface{}) bool {
	if ta, ok := a.(time.Time); ok {
		tb, ok := b.(time.Time)
		return ok && ta.Sub(tb).Abs() <= 2*time.Second
	}
	if fa, ok := a.(float64); ok {
		fb, ok := b.(float64)
		if !ok {
			return false
		}
		if fa == fb {
			return true
		}
		return math.Abs(fa-fb) <= 1e-6*math.Max(1, math.Abs(fa))
	}
	return fmtSprint(a) == fmtSprint(b)
}

func fmtSprint(v interface{}) string {
	if v == nil {
		return ""
	}
	return strings.TrimSpace(strings.ReplaceAll(stringify(v), "\r\n", "\n"))
}

func stringify(v interface{}) string {
	return sprintAny(v)
}

func sprintAny(v interface{}) string {
	return strings.TrimSpace(func() string {
		switch t := v.(type) {
		case string:
			return t
		case float64:
			return strconv.FormatFloat(t, 'f', -1, 64)
		case bool:
			if t {
				return "true"
			}
			return "false"
		case time.Time:
			return t.UTC().Format(time.RFC3339Nano)
		default:
			return ""
		}
	}())
}

func assertPackagePartsPreserved(t *testing.T, originalPath string, written []byte) {
	t.Helper()
	orig, err := zip.OpenReader(originalPath)
	if err != nil {
		t.Fatal(err)
	}
	defer orig.Close()
	zr, err := zip.NewReader(bytes.NewReader(written), int64(len(written)))
	if err != nil {
		t.Fatal(err)
	}
	have := map[string]bool{}
	for _, f := range zr.File {
		have[strings.ToLower(f.Name)] = true
	}
	for _, f := range orig.File {
		name := strings.ToLower(f.Name)
		if strings.HasPrefix(name, "xl/charts/") || strings.HasPrefix(name, "xl/media/") || strings.HasPrefix(name, "xl/theme/") || strings.HasPrefix(name, "xl/tables/") || strings.HasPrefix(name, "xl/pivot") || strings.HasPrefix(name, "xl/drawings/drawing") {
			if !have[name] {
				t.Errorf("round-trip dropped package part %s", f.Name)
			}
		}
	}
}

func TestRound1ReferenceDiscrepancies(t *testing.T) {
	wb, err := OpenFile(filepath.Join("testdata", "excel", "sharedhyperlink.xlsx"))
	if err != nil {
		t.Fatal(err)
	}
	ws := wb.GetActiveSheet()
	for _, addr := range []string{"A3", "A4", "A5"} {
		cell, err := ws.GetCellByName(addr)
		if err != nil || cell.Hyperlink == nil || !strings.Contains(cell.Hyperlink.URL, "apache.org") {
			t.Errorf("%s hyperlink = %#v err=%v", addr, cell.Hyperlink, err)
		}
	}

	merges := openWB(t, "Book1.xlsx")
	sheet, err := merges.GetSheetByName("Sheet2")
	if err != nil {
		// Book1's second sheet is the one with the single-cell merge.
		sheet, err = merges.GetSheet(1)
		if err != nil {
			t.Fatal(err)
		}
	}
	found := false
	for _, mc := range sheet.GetMergeCells() {
		if mc.StartRow == mc.EndRow && mc.StartCol == mc.EndCol {
			found = true
		}
	}
	if !found {
		t.Fatalf("single-cell merge missing, got %#v", sheet.GetMergeCells())
	}

	times := openWB(t, filepath.Join("r1", "cx-Sorting.xlsx"))
	ts, err := times.GetSheetByName("TimeSpans")
	if err != nil {
		t.Fatal(err)
	}
	c, _ := ts.GetCellByName("A1")
	if c.Type != CellTypeNumeric {
		t.Fatalf("time-only A1 type=%v value=%#v", c.Type, c.Value)
	}

	dtwb := openWB(t, filepath.Join("r1", "cx-DataTableFormula-Excel-Input.xlsx"))
	col, err := dtwb.GetSheetByName("1D Column")
	if err != nil {
		t.Fatal(err)
	}
	dc, _ := col.GetCellByName("C4")
	if dc.Formula != "TABLE(A1)" {
		t.Fatalf("data table formula = %q", dc.Formula)
	}
	if dc.GetStringValue() != "AA" {
		t.Fatalf("data table cache = %#v", dc.Value)
	}
}

func openWB(t *testing.T, name string) *Workbook {
	t.Helper()
	wb, err := OpenFile(filepath.Join("testdata", "excel", name))
	if err != nil {
		t.Fatal(err)
	}
	return wb
}

func TestExcelFixtureSemantics(t *testing.T) {
	open := func(name string) *Workbook {
		t.Helper()
		wb, err := OpenFile(filepath.Join("testdata", "excel", name))
		if err != nil {
			t.Fatal(err)
		}
		return wb
	}

	inline := open("InlineString.xlsx").GetActiveSheet()
	if got := inline.GetCell(0, 0).GetStringValue(); got != "😜more text" {
		t.Errorf("inline string = %q", got)
	}

	sample := open("sample.xlsx")
	rich, _ := sample.GetSheet(1)
	// Rich-text sheet stores "The quick brown fox jumps over the lazy dog".
	foundFox := false
	foundHello := false
	for _, c := range rich.AllCells() {
		s := c.GetStringValue()
		if strings.Contains(s, "quick brown fox") {
			foundFox = true
			if c.RichText == nil || len(c.RichText.Runs) < 2 {
				t.Errorf("rich text runs = %#v", c.RichText)
			}
		}
		if s == "hello, xssf" && c.RichText != nil {
			foundHello = true
		}
	}
	if !foundFox || !foundHello {
		t.Errorf("rich text sheet missing expected strings fox=%v hello=%v", foundFox, foundHello)
	}

	shared := open("shared_formulas.xlsx").GetActiveSheet()
	a2, _ := shared.GetCellByName("A2")
	a3, _ := shared.GetCellByName("A3")
	if a2.Formula != "B2" {
		t.Errorf("A2 formula = %q", a2.Formula)
	}
	if a3.Formula != "B3" {
		t.Errorf("A3 shared formula = %q, want B3", a3.Formula)
	}
	if len(shared.GetParentNamed(t)) < 60 {
		t.Errorf("defined names = %d", len(shared.GetParentNamed(t)))
	}

	dates := open("DateFormatTests.xlsx")
	if !dates.IsDate1904() {
		t.Error("DateFormatTests should use the 1904 date system")
	}
	tests, err := dates.GetSheetByName("Tests")
	if err != nil {
		t.Fatal(err)
	}
	gotDate, err := tests.GetCellByName("C2")
	if err != nil {
		t.Fatal(err)
	}
	dv, err := gotDate.GetDateValue()
	if err != nil {
		t.Fatalf("C2 type=%v value=%#v err=%v", gotDate.Type, gotDate.Value, err)
	}
	want := time.Date(1952, 10, 11, 14, 35, 27, 0, time.UTC)
	if dv.Sub(want).Abs() > 2*time.Second {
		t.Errorf("1904 date C2 = %v, want %v", dv, want)
	}

	comments := open("SimpleWithComments.xlsx").GetActiveSheet()
	cmt := comments.GetCell(0, 1).Comment
	if cmt == nil || cmt.Author != "Yegor Kozlov" || !strings.Contains(cmt.Text, "first cell") {
		t.Errorf("comment = %#v", cmt)
	}

	links := open("WithVariousData.xlsx")
	ws, _ := links.GetSheet(0)
	var link *Hyperlink
	for _, c := range ws.AllCells() {
		if c.Hyperlink != nil {
			link = c.Hyperlink
		}
	}
	if link == nil || !strings.Contains(link.URL, "poi.apache.org") {
		t.Errorf("hyperlink = %#v", link)
	}

	hidden := open("TwoSheetsOneHidden.xlsx")
	hidden0, _ := hidden.GetSheet(0)
	if hidden0.GetSheetState() != "hidden" {
		t.Errorf("first sheet state = %s", hidden0.GetSheetState())
	}
	if hidden.GetActiveSheet().Title() != "Sheet2" {
		t.Errorf("active sheet = %s", hidden.GetActiveSheet().Title())
	}

	widths := open("WidthsAndHeights.xlsx").GetActiveSheet()
	if math.Abs(widths.GetColumnWidth(0)-20) > 0.01 {
		t.Errorf("col width = %v", widths.GetColumnWidth(0))
	}
	if !widths.IsColumnHidden(2) {
		t.Error("column C should be hidden")
	}
	if math.Abs(widths.GetRowHeight(0)-37.5) > 0.01 {
		t.Errorf("row height = %v", widths.GetRowHeight(0))
	}
	if !widths.IsRowHidden(2) || !widths.IsRowHidden(3) {
		t.Error("rows 3 and 4 should be hidden")
	}

	tabs := open("SheetTabColors.xlsx")
	tab2, _ := tabs.GetSheet(2)
	if tab2.GetTabColor() != "FF7F2700" {
		t.Errorf("tab color = %q", tab2.GetTabColor())
	}

	strict := open("SampleSS.strict.xlsx")
	if strict.SheetCount() != 3 {
		t.Fatalf("strict sheets = %d", strict.SheetCount())
	}
	if v, _ := strict.GetActiveSheet().GetCellValue("A1"); v == nil || v == "" {
		t.Errorf("strict A1 = %#v", v)
	}

	matrix := open("MatrixFormulaEvalTestData.xlsx").GetActiveSheet()
	h2, _ := matrix.GetCellByName("H2")
	if !strings.Contains(h2.Formula, "B2:D4") {
		t.Errorf("array formula = %q", h2.Formula)
	}

	names := open("RepeatingRowsCols.xlsx")
	if _, err := names.GetNamedRange("_xlnm.Print_Titles"); err != nil {
		t.Error(err)
	}
	sheet1, _ := names.GetSheet(1)
	sheet3, _ := names.GetSheet(3)
	if sheet1.GetPageSetup().RepeatRows == "" && sheet3.GetPageSetup().RepeatColumns == "" {
		t.Error("print titles were not mapped onto page setup")
	}

	greek := open("54084 - Greek - beyond BMP.xlsx")
	var saw bool
	for _, ws := range greek.Sheets() {
		for _, c := range ws.AllCells() {
			if s := c.GetStringValue(); strings.Contains(s, "𐀀") || len([]rune(s)) > 0 && s != "" {
				saw = s != ""
			}
		}
	}
	if !saw {
		t.Error("expected text in the beyond-BMP workbook")
	}
}

func (ws *Worksheet) GetParentNamed(t *testing.T) map[string]string {
	t.Helper()
	if ws.parent == nil {
		t.Fatal("sheet has no parent")
	}
	return ws.parent.GetNamedRanges()
}
