package gospreadsheet

import (
	"bytes"
	"strings"
	"testing"
)

func TestOOXMLTextEscape(t *testing.T) {
	if got := unescapeOOXMLText("_x005F_x000D_"); got != "_x000D_" {
		t.Fatalf("unescape escape-of-escape = %q", got)
	}
	if got := unescapeOOXMLText("_x000D_"); got != "\r" {
		t.Fatalf("unescape CR = %q", got)
	}
	if got := escapeOOXMLText("_x000D_"); got != "_x005F_x000D_" {
		t.Fatalf("escape literal = %q", got)
	}
	if got := escapeOOXMLText("\x01"); got != "_x0001_" {
		t.Fatalf("escape control = %q", got)
	}
	if got := unescapeOOXMLText(escapeOOXMLText("keep _x000D_ and \x01")); got != "keep _x000D_ and \x01" {
		t.Fatalf("round trip = %q", got)
	}
	wb, err := OpenFile("testdata/excel/r12/ej-shared-escape.xlsx")
	if err != nil {
		t.Fatal(err)
	}
	v, _ := wb.GetActiveSheet().GetCellValue("A1")
	if v != "_x000D_" {
		t.Fatalf("A1 = %#v", v)
	}
	var buf bytes.Buffer
	if err := NewXLSXWriter().Write(wb, &buf); err != nil {
		t.Fatal(err)
	}
	files := zipFiles(t, buf.Bytes())
	sst := string(files["xl/sharedStrings.xml"])
	if !strings.Contains(sst, "_x005F_x000D_") {
		t.Fatalf("shared string was not re-escaped: %s", sst)
	}
	again, err := NewXLSXReader().Read(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}
	v, _ = again.GetActiveSheet().GetCellValue("A1")
	if v != "_x000D_" {
		t.Fatalf("reopened A1 = %#v", v)
	}
}

func TestRound12PackageFeatures(t *testing.T) {
	dyn, err := OpenFile("testdata/excel/r12/rw-dynamic-array02.xlsx")
	if err != nil {
		t.Fatal(err)
	}
	cell, err := dyn.GetActiveSheet().GetCellByName("B1")
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(cell.Formula, "_xlfn.UNIQUE") || cell.cellMeta != "1" {
		t.Fatalf("dynamic array formula=%q cm=%q", cell.Formula, cell.cellMeta)
	}
	var buf bytes.Buffer
	if err := NewXLSXWriter().Write(dyn, &buf); err != nil {
		t.Fatal(err)
	}
	files := zipFiles(t, buf.Bytes())
	if _, ok := files["xl/metadata.xml"]; !ok {
		t.Fatal("dynamic array rewrite dropped xl/metadata.xml")
	}
	sheet := sheetXML(files)
	if !strings.Contains(sheet, `cm="1"`) || !strings.Contains(sheet, "_xlfn.UNIQUE") {
		t.Fatalf("rewritten dynamic array sheet missing cm or formula: %s", sheet)
	}

	img, err := OpenFile("testdata/excel/r12/rw-embed-image.xlsx")
	if err != nil {
		t.Fatal(err)
	}
	ic, _ := img.GetActiveSheet().GetCellByName("A1")
	if ic.valueMeta != "1" || ic.Type != CellTypeError {
		t.Fatalf("embed image cell type=%v vm=%q value=%#v", ic.Type, ic.valueMeta, ic.Value)
	}
	buf.Reset()
	if err := NewXLSXWriter().Write(img, &buf); err != nil {
		t.Fatal(err)
	}
	files = zipFiles(t, buf.Bytes())
	for _, name := range []string{
		"xl/metadata.xml",
		"xl/richData/rdrichvalue.xml",
		"xl/richData/rdRichValueTypes.xml",
		"xl/media/image1.png",
	} {
		if _, ok := files[name]; !ok {
			t.Errorf("embed image rewrite dropped %s", name)
		}
	}
	if !strings.Contains(sheetXML(files), `vm="1"`) {
		t.Fatalf("embed image cell lost vm: %s", sheetXML(files))
	}
	root := string(files["_rels/.rels"])
	wbrels := string(files["xl/_rels/workbook.xml.rels"])
	if !strings.Contains(wbrels, "richValueRel") && !strings.Contains(wbrels, "rdRichValue") {
		t.Fatalf("rich data relationships dropped: %s", wbrels)
	}
	_ = root

	icon, err := OpenFile("testdata/excel/r12/rw-iconset.xlsx")
	if err != nil {
		t.Fatal(err)
	}
	buf.Reset()
	if err := NewXLSXWriter().Write(icon, &buf); err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(sheetXML(zipFiles(t, buf.Bytes())), `type="iconSet"`) {
		t.Fatal("icon set conditional formatting was dropped")
	}

	bar, err := OpenFile("testdata/excel/r12/rw-x14-databar.xlsx")
	if err != nil {
		t.Fatal(err)
	}
	buf.Reset()
	if err := NewXLSXWriter().Write(bar, &buf); err != nil {
		t.Fatal(err)
	}
	barSheet := sheetXML(zipFiles(t, buf.Bytes()))
	if !strings.Contains(barSheet, `type="dataBar"`) || !strings.Contains(barSheet, "x14:dataBar") {
		t.Fatalf("x14 data bar was dropped: %s", barSheet)
	}
	x14, err := OpenFile("testdata/excel/r12/rw-x14-iconset.xlsx")
	if err != nil {
		t.Fatal(err)
	}
	buf.Reset()
	if err := NewXLSXWriter().Write(x14, &buf); err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(sheetXML(zipFiles(t, buf.Bytes())), "x14:iconSet") {
		t.Fatal("x14 icon set extension was dropped")
	}

	th, err := OpenFile("testdata/excel/r12/rw-threaded.xlsx")
	if err != nil {
		t.Fatal(err)
	}
	buf.Reset()
	if err := NewXLSXWriter().Write(th, &buf); err != nil {
		t.Fatal(err)
	}
	files = zipFiles(t, buf.Bytes())
	body := string(files["xl/threadedComments/threadedComment1.xml"])
	if !strings.Contains(body, "thread note") {
		t.Fatalf("threaded comment dropped: %s", body)
	}
	if _, ok := files["xl/persons/person.xml"]; !ok {
		t.Fatal("person part dropped")
	}

	cx, err := OpenFile("testdata/excel/r12/rw-customxml.xlsx")
	if err != nil {
		t.Fatal(err)
	}
	buf.Reset()
	if err := NewXLSXWriter().Write(cx, &buf); err != nil {
		t.Fatal(err)
	}
	files = zipFiles(t, buf.Bytes())
	item := string(files["customXml/item1.xml"])
	if !strings.Contains(item, "kept") {
		t.Fatalf("custom xml dropped: %s", item)
	}
	if !strings.Contains(string(files["_rels/.rels"]), "customXml/item1.xml") {
		t.Fatalf("custom xml relationship dropped: %s", files["_rels/.rels"])
	}

	btn, err := OpenFile("testdata/excel/r12/rw-form-button.xlsx")
	if err != nil {
		t.Fatal(err)
	}
	buf.Reset()
	if err := NewXLSXWriter().Write(btn, &buf); err != nil {
		t.Fatal(err)
	}
	files = zipFiles(t, buf.Bytes())
	var vml string
	for name, b := range files {
		if strings.Contains(name, "vmlDrawing") {
			vml = string(b)
		}
	}
	if !strings.Contains(vml, `ObjectType="Button"`) {
		t.Fatalf("form control button was dropped: %s", vml)
	}
}

func sheetXML(files map[string][]byte) string {
	for name, b := range files {
		if strings.HasPrefix(name, "xl/worksheets/sheet") && strings.HasSuffix(name, ".xml") {
			return string(b)
		}
	}
	return ""
}
