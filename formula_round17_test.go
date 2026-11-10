package gospreadsheet

import (
	"bytes"
	"strings"
	"testing"
)

func TestDateYearWindow(t *testing.T) {
	wb := New()
	ws := wb.GetActiveSheet()
	ws.SetCellValue("A1", 1.0)
	ws.SetCellValue("B1", 2.0)
	ws.SetCellValue("C1", 1.00001)
	ws.SetCellFormula("D1", "DATE(A1,B1,C1)")
	got, err := NewCalculationEngine(wb).CalculateCell(ws, "D1")
	if err != nil {
		t.Fatal(err)
	}
	if got != 398.0 {
		t.Fatalf("DATE(1,2,1) = %v, want 398", got)
	}
	ws.SetCellFormula("D2", `DATE("nospaces",1,1)`)
	if _, err := NewCalculationEngine(wb).CalculateCell(ws, "D2"); err == nil {
		t.Fatal("DATE of text should be #VALUE!")
	}
}

func TestMatrixAndSpill(t *testing.T) {
	wb := New()
	ws := wb.GetActiveSheet()
	ws.SetCellValue("A1", 1.0)
	ws.SetCellValue("B1", 2.0)
	ws.SetCellValue("A2", 3.0)
	ws.SetCellValue("B2", 4.0)
	ws.SetCellFormula("D1", "MMULT(A1:B2,A1:B2)")
	ce := NewCalculationEngine(wb)
	got, err := ce.CalculateCell(ws, "D1")
	if err != nil {
		t.Fatal(err)
	}
	if got != 7.0 {
		t.Fatalf("MMULT anchor = %v, want 7", got)
	}
	e1, _ := ws.GetCellByName("E1")
	if e1 == nil || e1.Value != 10.0 {
		t.Fatalf("spill E1 = %#v", e1)
	}
	d2, _ := ws.GetCellByName("D2")
	if d2 == nil || d2.Value != 15.0 {
		t.Fatalf("spill D2 = %#v", d2)
	}
	ws.SetCellFormula("G1", "MDETERM(A1:B2)")
	det, err := ce.CalculateCell(ws, "G1")
	if err != nil || det != -2.0 {
		t.Fatalf("MDETERM = %v %v", det, err)
	}
	var buf bytes.Buffer
	if err := NewXLSXWriter().Write(wb, &buf); err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(buf.String(), "XLDAPR") {
		// metadata is compressed in the zip; reopen and check the part via a second save read
	}
	opened, err := NewXLSXReader().Read(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}
	_ = opened
	files := map[string][]byte{}
	// Writer output is a zip. Confirm metadata by reading the package through Open which preserves source.
	if opened.source != nil {
		if data, ok := opened.source.file("xl/metadata.xml"); ok && !strings.Contains(string(data), "XLDAPR") {
			t.Fatalf("metadata missing XLDAPR: %s", data)
		} else if !ok {
			t.Fatal("saved workbook has no xl/metadata.xml")
		}
	}
	_ = files
}

func TestTextElapsedAndConditional(t *testing.T) {
	if got := formatByExcelCode(3.14159, "[s].000"); got != "271433.376" {
		t.Fatalf("[s].000 = %q", got)
	}
	if got := formatByExcelCode(3.14159, "[hh]:mm:ss.0"); got != "75:23:53.4" {
		t.Fatalf("elapsed = %q", got)
	}
	if got := formatByExcelCode(10.0, `[<10]#" Wow"`); got != "10" {
		t.Fatalf("conditional miss = %q", got)
	}
	if got := formatByExcelCode(1.5, `[<10]#" Wow"`); !strings.Contains(got, "Wow") {
		t.Fatalf("conditional hit = %q", got)
	}
	if got := formatByExcelCode(12.3, "000.00"); got != "012.30" {
		t.Fatalf("pad = %q", got)
	}
	if got := formatByExcelCode(123456.789, "0.00E+00"); !strings.Contains(strings.ToUpper(got), "E") {
		t.Fatalf("scientific = %q", got)
	}
}
