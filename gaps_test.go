package gospreadsheet

import (
	"archive/zip"
	"bytes"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestEncryptedRoundTripAndExternalSamples(t *testing.T) {
	wb := New()
	ws := wb.GetActiveSheet()
	ws.SetCellValue("A1", "encrypted-ok")
	ws.SetCellValue("B2", 42.0)
	wb.SetPassword("s3cret")
	var buf bytes.Buffer
	if err := NewXLSXWriter().Write(wb, &buf); err != nil {
		t.Fatal(err)
	}
	if !isOLE(buf.Bytes()) {
		t.Fatal("password save did not produce an OLE package")
	}
	if _, err := NewXLSXReader().Read(bytes.NewReader(buf.Bytes()), int64(buf.Len())); err != ErrPasswordRequired {
		t.Fatalf("open without password: %v", err)
	}
	got, err := NewXLSXReader().ReadPassword(bytes.NewReader(buf.Bytes()), int64(buf.Len()), "s3cret")
	if err != nil {
		t.Fatal(err)
	}
	v, _ := got.GetActiveSheet().GetCellValue("A1")
	if v != "encrypted-ok" {
		t.Fatalf("A1 = %#v", v)
	}
	n, _ := got.GetActiveSheet().GetCellValue("B2")
	if n != 42.0 {
		t.Fatalf("B2 = %#v", n)
	}
	if _, err := DecryptPackage(buf.Bytes(), "nope"); err != ErrWrongPassword {
		t.Fatalf("wrong password: %v", err)
	}

	// Standard encryption is what Office 2007 wrote. Decrypt our own standard
	// package and the POI/msoffcrypto fixtures checked in under testdata.
	std, err := encryptStandardPackage(mustZip(t, wbWithoutPassword(t)), "s3cret")
	if err != nil {
		t.Fatal(err)
	}
	plain, err := DecryptPackage(std, "s3cret")
	if err != nil {
		t.Fatal(err)
	}
	opened, err := NewXLSXReader().Read(bytes.NewReader(plain), int64(len(plain)))
	if err != nil {
		t.Fatal(err)
	}
	if v, _ := opened.GetActiveSheet().GetCellValue("A1"); v != "encrypted-ok" {
		t.Fatalf("standard A1 = %#v", v)
	}

	samples := []struct {
		name string
		cell string
		want interface{}
	}{
		{"poi-agile.xlsx", "A1", "encrypted-ok"},
		{"poi-standard.xlsx", "A1", "encrypted-ok"},
		{"msoffcrypto-agile.xlsx", "A1", "row-0"},
	}
	for _, s := range samples {
		path := filepath.Join("testdata", "encrypted", s.name)
		wb, err := OpenFile(path, WithPassword("s3cret"))
		if err != nil {
			t.Errorf("%s: %v", s.name, err)
			continue
		}
		got, _ := wb.GetActiveSheet().GetCellValue(s.cell)
		if got != s.want {
			t.Errorf("%s %s = %#v, want %#v", s.name, s.cell, got, s.want)
		}
		if _, err := OpenFile(path); err != ErrPasswordRequired {
			t.Errorf("%s without password: %v", s.name, err)
		}
	}
}

func wbWithoutPassword(t *testing.T) *Workbook {
	t.Helper()
	wb := New()
	wb.GetActiveSheet().SetCellValue("A1", "encrypted-ok")
	return wb
}

func mustZip(t *testing.T, wb *Workbook) []byte {
	t.Helper()
	var buf bytes.Buffer
	if err := NewXLSXWriter().Write(wb, &buf); err != nil {
		t.Fatal(err)
	}
	return buf.Bytes()
}

func TestStrictOOXMLOptional(t *testing.T) {
	loaded, err := OpenFile(filepath.Join("testdata", "excel", "SampleSS.strict.xlsx"))
	if err != nil {
		t.Fatal(err)
	}
	if !loaded.LoadedStrict() {
		t.Fatal("strict sample was not detected")
	}
	if loaded.StrictOOXML() {
		t.Fatal("reading strict must not switch the writer to strict")
	}

	wb := New()
	wb.GetActiveSheet().SetCellValue("A1", "strict-cell")
	wb.SetStrictOOXML(true)
	var buf bytes.Buffer
	if err := NewXLSXWriter().Write(wb, &buf); err != nil {
		t.Fatal(err)
	}
	files := zipFiles(t, buf.Bytes())
	wbXML := string(files["xl/workbook.xml"])
	if !strings.Contains(wbXML, `conformance="strict"`) {
		t.Fatalf("workbook missing conformance: %s", wbXML[:min(200, len(wbXML))])
	}
	if !strings.Contains(wbXML, "http://purl.oclc.org/ooxml/spreadsheetml/main") {
		t.Fatal("workbook is not strict spreadsheet markup")
	}
	if strings.Contains(string(files["[Content_Types].xml"]), "http://purl.oclc.org/ooxml/package") {
		t.Fatal("content types namespace should stay the OPC transitional URI")
	}
	rels := string(files["xl/_rels/workbook.xml.rels"])
	if !strings.Contains(rels, "http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet") {
		t.Fatalf("worksheet relationship was not rewritten: %s", rels)
	}
	if !strings.Contains(rels, "http://schemas.openxmlformats.org/package/2006/relationships") {
		t.Fatal("package relationship namespace changed")
	}
	got, err := NewXLSXReader().Read(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}
	if !got.LoadedStrict() {
		t.Fatal("rewritten strict file was not detected on read")
	}
	v, _ := got.GetActiveSheet().GetCellValue("A1")
	if v != "strict-cell" {
		t.Fatalf("strict round trip A1 = %#v", v)
	}

	wb.SetStrictOOXML(false)
	buf.Reset()
	if err := NewXLSXWriter().Write(wb, &buf); err != nil {
		t.Fatal(err)
	}
	plain := string(zipFiles(t, buf.Bytes())["xl/workbook.xml"])
	if strings.Contains(plain, `conformance="strict"`) || strings.Contains(plain, "purl.oclc.org/ooxml/spreadsheetml/main") {
		t.Fatal("default save became strict")
	}
}

func TestArrayFormulaRange(t *testing.T) {
	wb := New()
	ws := wb.GetActiveSheet()
	if err := ws.SetArrayFormula("B2:C3", "A1:B2"); err != nil {
		t.Fatal(err)
	}
	anchor, _ := ws.GetCellByName("B2")
	if !anchor.IsArrayFormula() || anchor.ArrayRef() != "B2:C3" {
		t.Fatalf("anchor ref=%q array=%v formula=%q", anchor.ArrayRef(), anchor.IsArrayFormula(), anchor.Formula)
	}
	if anchor.Formula != "{A1:B2}" {
		t.Fatalf("anchor formula = %q", anchor.Formula)
	}
	// The member has no stored formula of its own.
	member := ws.GetCell(2, 2) // C3
	if member.Formula != "" {
		t.Fatalf("member stored formula = %q", member.Formula)
	}
	if !member.IsArrayFormula() || member.ReportedFormula() != "{A1:B2}" || member.ArrayRef() != "B2:C3" {
		t.Fatalf("member reports formula %q ref %q", member.ReportedFormula(), member.ArrayRef())
	}
	formula, ref, err := ws.FormulaAt("C3")
	if err != nil || formula != "{A1:B2}" || ref != "B2:C3" {
		t.Fatalf("FormulaAt = %q %q %v", formula, ref, err)
	}
	var buf bytes.Buffer
	if err := NewXLSXWriter().Write(wb, &buf); err != nil {
		t.Fatal(err)
	}
	sheet := string(zipFiles(t, buf.Bytes())["xl/worksheets/sheet1.xml"])
	if strings.Count(sheet, `t="array"`) != 1 || !strings.Contains(sheet, `ref="B2:C3"`) {
		t.Fatalf("array markup = %s", sheet)
	}
	got, err := NewXLSXReader().Read(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}
	f, r, err := got.GetActiveSheet().FormulaAt("C3")
	if err != nil || f != "{A1:B2}" || r != "B2:C3" {
		t.Fatalf("reopen FormulaAt = %q %q %v", f, r, err)
	}

	matrix, err := OpenFile(filepath.Join("testdata", "excel", "MatrixFormulaEvalTestData.xlsx"))
	if err != nil {
		t.Fatal(err)
	}
	var saw bool
	for _, sh := range matrix.Sheets() {
		for _, af := range sh.ArrayFormulas() {
			if !strings.Contains(af.Ref, ":") {
				continue
			}
			saw = true
			parts := strings.Split(af.Ref, ":")
			// The bottom-right cell reports the anchor formula.
			f, r, err := sh.FormulaAt(parts[1])
			if err != nil || f == "" || r != af.Ref {
				t.Fatalf("%s FormulaAt(%s) = %q %q %v", sh.Title(), parts[1], f, r, err)
			}
		}
	}
	if !saw {
		t.Fatal("matrix workbook has no multi-cell array formula")
	}
}

func TestSparklineReadEditRoundTrip(t *testing.T) {
	wb, err := OpenFile(filepath.Join("testdata", "excel", "r3", "lo-Sparklines.xlsx"))
	if err != nil {
		t.Fatal(err)
	}
	var ws *Worksheet
	var groups []SparklineGroup
	for _, sh := range wb.Sheets() {
		g := sh.SparklineGroups()
		if len(g) > 0 {
			ws = sh
			groups = g
			break
		}
	}
	if ws == nil {
		t.Fatal("no sparklines parsed")
	}
	if groups[0].Type != "line" || !groups[0].Markers || groups[0].ColorSeries == "" {
		t.Fatalf("group0 = %+v", groups[0])
	}
	if len(groups[0].Sparklines) == 0 || groups[0].Sparklines[0].Ref != "A2" || !strings.Contains(groups[0].Sparklines[0].Formula, "B1:M1") {
		t.Fatalf("sparkline = %+v", groups[0].Sparklines)
	}
	if len(groups) < 2 || groups[1].Type != "column" {
		t.Fatalf("groups = %+v", groups)
	}
	groups[0].Sparklines[0].Ref = "Z9"
	groups[1].Type = "stacked"
	ws.SetSparklineGroups(groups)
	var buf bytes.Buffer
	if err := NewXLSXWriter().Write(wb, &buf); err != nil {
		t.Fatal(err)
	}
	sheet := string(zipFiles(t, buf.Bytes())["xl/worksheets/sheet1.xml"])
	if !strings.Contains(sheet, "<xm:sqref>Z9</xm:sqref>") || !strings.Contains(sheet, `type="stacked"`) {
		t.Fatalf("edited sparkline markup missing: %s", sheet[strings.LastIndex(sheet, "sparkline"):])
	}
	got, err := NewXLSXReader().Read(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}
	var again []SparklineGroup
	for _, sh := range got.Sheets() {
		for _, g := range sh.SparklineGroups() {
			for _, s := range g.Sparklines {
				if s.Ref == "Z9" {
					again = sh.SparklineGroups()
				}
			}
		}
	}
	if len(again) < 2 || again[0].Sparklines[0].Ref != "Z9" || again[1].Type != "stacked" {
		t.Fatalf("reopen sparklines = %+v", again)
	}
	if again[0].Markers != true || again[0].ColorSeries == "" || !strings.Contains(again[0].Sparklines[0].Formula, "B1:M1") {
		t.Fatalf("sparkline details dropped: %+v", again[0])
	}
	// Untouched read/write keeps the original sparkline ref.
	fresh, err := OpenFile(filepath.Join("testdata", "excel", "r3", "lo-Sparklines.xlsx"))
	if err != nil {
		t.Fatal(err)
	}
	buf.Reset()
	if err := NewXLSXWriter().Write(fresh, &buf); err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(zipFiles(t, buf.Bytes())["xl/worksheets/sheet1.xml"]), "<xm:sqref>A2</xm:sqref>") {
		t.Fatal("untouched sparkline round-trip rewrote the location")
	}
}

func TestPivotReadEditRoundTrip(t *testing.T) {
	wb, err := OpenFile(filepath.Join("testdata", "excel", "ExcelPivotTableSample.xlsx"))
	if err != nil {
		t.Fatal(err)
	}
	tables := wb.PivotTables()
	if len(tables) < 2 {
		t.Fatalf("pivots = %d", len(tables))
	}
	var pt *PivotTable
	for _, p := range tables {
		if p.Name == "PivotTable1" {
			pt = p
		}
	}
	if pt == nil {
		t.Fatalf("missing PivotTable1 (%d tables)", len(tables))
	}
	if pt.Location == "" || pt.DataCaption == "" || len(pt.CacheFields) == 0 || len(pt.RowFields) == 0 || len(pt.DataFields) == 0 {
		t.Fatalf("pivot = %+v", pt)
	}
	pt.SetName("RenamedPivot")
	pt.SetDataCaption("Values")
	pt.SetLocation("A3:D12")
	var buf bytes.Buffer
	if err := NewXLSXWriter().Write(wb, &buf); err != nil {
		t.Fatal(err)
	}
	files := zipFiles(t, buf.Bytes())
	raw := string(files[pt.part])
	if !strings.Contains(raw, `name="RenamedPivot"`) || !strings.Contains(raw, `dataCaption="Values"`) || !strings.Contains(raw, `ref="A3:D12"`) {
		t.Fatalf("pivot part %s xml head %s", pt.part, raw[:min(220, len(raw))])
	}
	got, err := NewXLSXReader().Read(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}
	var found *PivotTable
	for _, p := range got.PivotTables() {
		if p.Name == "RenamedPivot" {
			found = p
		}
	}
	if found == nil || found.DataCaption != "Values" || found.Location != "A3:D12" || len(found.CacheFields) == 0 {
		t.Fatalf("reopen pivot = %+v", found)
	}
}

func TestSlicerReadEditRoundTrip(t *testing.T) {
	raw := workbookWithSlicer(t)
	wb, err := NewXLSXReader().Read(bytes.NewReader(raw), int64(len(raw)))
	if err != nil {
		t.Fatal(err)
	}
	var slicer *Slicer
	for _, sh := range wb.Sheets() {
		if ss := sh.Slicers(); len(ss) > 0 {
			slicer = ss[0]
		}
	}
	if slicer == nil || slicer.Name != "Region" || slicer.Caption != "Region" || slicer.Cache != "Slicer_Region" {
		t.Fatalf("slicer = %+v", slicer)
	}
	caches := wb.SlicerCaches()
	if len(caches) != 1 || caches[0].SourceName != "Region" {
		t.Fatalf("caches = %+v", caches)
	}
	slicer.SetCaption("Sales region")
	slicer.SetStyle("SlicerStyleDark1")
	slicer.SetColumnCount(2)
	caches[0].SetSourceName("Area")
	var buf bytes.Buffer
	if err := NewXLSXWriter().Write(wb, &buf); err != nil {
		t.Fatal(err)
	}
	files := zipFiles(t, buf.Bytes())
	slicerXML := string(files["xl/slicers/slicer1.xml"])
	cacheXML := string(files["xl/slicerCaches/slicerCache1.xml"])
	if !strings.Contains(slicerXML, `caption="Sales region"`) || !strings.Contains(slicerXML, `style="SlicerStyleDark1"`) || !strings.Contains(slicerXML, `columnCount="2"`) {
		t.Fatalf("slicer xml = %s", slicerXML)
	}
	if !strings.Contains(cacheXML, `sourceName="Area"`) {
		t.Fatalf("cache xml = %s", cacheXML)
	}
	got, err := NewXLSXReader().Read(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}
	s := got.Sheets()[0].Slicers()
	if len(s) != 1 || s[0].Caption != "Sales region" || s[0].Style != "SlicerStyleDark1" || s[0].ColumnCount != 2 {
		t.Fatalf("reopen slicer = %+v", s)
	}
	if got.SlicerCaches()[0].SourceName != "Area" {
		t.Fatalf("reopen cache = %+v", got.SlicerCaches()[0])
	}
	if v, _ := got.GetActiveSheet().GetCellValue("A1"); v != "Region" {
		t.Fatalf("cell lost = %#v", v)
	}
	if err := os.WriteFile(filepath.Join(os.TempDir(), "slicer-edited.xlsx"), buf.Bytes(), 0644); err != nil {
		t.Fatal(err)
	}
}

func workbookWithSlicer(t *testing.T) []byte {
	t.Helper()
	wb := New()
	wb.GetActiveSheet().SetCellValue("A1", "Region")
	wb.GetActiveSheet().SetCellValue("A2", "North")
	var buf bytes.Buffer
	if err := NewXLSXWriter().Write(wb, &buf); err != nil {
		t.Fatal(err)
	}
	files := zipFiles(t, buf.Bytes())
	files["xl/slicers/slicer1.xml"] = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">` +
		`<slicer name="Region" cache="Slicer_Region" caption="Region" columnCount="1" rowHeight="20" style="SlicerStyleLight1"/>` +
		`</slicers>`)
	files["xl/slicerCaches/slicerCache1.xml"] = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="Slicer_Region" sourceName="Region"/>`)
	ct := string(files["[Content_Types].xml"])
	ct = strings.Replace(ct, "</Types>",
		`<Override PartName="/xl/slicers/slicer1.xml" ContentType="application/vnd.ms-excel.slicer+xml"/>`+
			`<Override PartName="/xl/slicerCaches/slicerCache1.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"/>`+
			`</Types>`, 1)
	files["[Content_Types].xml"] = []byte(ct)
	rels := string(files["xl/_rels/workbook.xml.rels"])
	rels = strings.Replace(rels, "</Relationships>",
		`<Relationship Id="rIdSlicerCache" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="slicerCaches/slicerCache1.xml"/></Relationships>`, 1)
	files["xl/_rels/workbook.xml.rels"] = []byte(rels)
	files["xl/worksheets/_rels/sheet1.xml.rels"] = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
		`<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="../slicers/slicer1.xml"/>` +
		`</Relationships>`)
	var out bytes.Buffer
	zw := zip.NewWriter(&out)
	for name, body := range files {
		f, err := zw.Create(name)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := f.Write(body); err != nil {
			t.Fatal(err)
		}
	}
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	return out.Bytes()
}

func zipFiles(t *testing.T, data []byte) map[string][]byte {
	t.Helper()
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatal(err)
	}
	out := map[string][]byte{}
	for _, f := range zr.File {
		rc, err := f.Open()
		if err != nil {
			t.Fatal(err)
		}
		b, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			t.Fatal(err)
		}
		out[f.Name] = b
	}
	return out
}
