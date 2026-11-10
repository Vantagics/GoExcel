package gospreadsheet

import (
	"math"
	"os"
	"os/exec"
	"path/filepath"
	"strconv"
	"strings"
	"testing"
	"time"
)

func TestFormulaEngineCoverage(t *testing.T) {
	wb := New()
	ws := wb.GetActiveSheet()
	ws.SetTitle("Data")
	mustSet(t, ws, "A1", 10.0)
	mustSet(t, ws, "A2", 20.0)
	mustSet(t, ws, "A3", 30.0)
	mustSet(t, ws, "A4", "Hello")
	mustSet(t, ws, "A5", "hello world")
	mustSet(t, ws, "A9", -4.0)
	mustSet(t, ws, "A10", 2.5)
	mustSet(t, ws, "B1", "apple")
	mustSet(t, ws, "B2", "pear")
	mustSet(t, ws, "B3", "apple")
	mustSet(t, ws, "C1", 5.0)
	mustSet(t, ws, "C2", 7.0)
	mustSet(t, ws, "C3", 9.0)
	mustSet(t, ws, "D1", time.Date(2024, 3, 5, 0, 0, 0, 0, time.UTC))
	other, err := wb.AddSheet("Other")
	if err != nil {
		t.Fatal(err)
	}
	mustSet(t, other, "A1", 100.0)
	wb.AddNamedRange("Total", "Data!A1:A3")

	cases := []struct {
		cell string
		f    string
		want interface{}
	}{
		{"E1", "SUM(A1:A3)", 60.0},
		{"E2", "SUM(A1)+1", 11.0},
		{"E3", "AVERAGE(A1:A3)", 20.0},
		{"E4", "COUNT(A1:A5)", 3.0},
		{"E5", "COUNTA(A1:A5)", 5.0},
		{"E6", "MIN(A1:A3)", 10.0},
		{"E7", "MAX(A1:A3)", 30.0},
		{"E8", "ABS(A9)", 4.0},
		{"E9", "ROUND(A10,0)", 3.0},
		{"E10", "ROUNDUP(2.2,0)", 3.0},
		{"E11", "ROUNDDOWN(2.8,0)", 2.0},
		{"E12", "INT(2.9)", 2.0},
		{"E13", "MOD(10,3)", 1.0},
		{"E14", "MOD(-3,2)", 1.0},
		{"E15", "POWER(2,3)", 8.0},
		{"E16", "2^3", 8.0},
		{"E17", "A10*10%", 0.25},
		{"E18", "A1&A4", "10Hello"},
		{"E19", "-A9", 4.0},
		{"E20", "IF(A1>5,\"yes\",\"no\")", "yes"},
		{"E21", "AND(A1>1,A2>1)", true},
		{"E22", "OR(A1>100,A2>1)", true},
		{"E23", "NOT(A1>100)", true},
		{"E24", "_xlfn.XOR(TRUE,FALSE)", true},
		{"E25", "IFERROR(A1/0,9)", 9.0},
		{"E26", "_xlfn.IFNA(VLOOKUP(\"nope\",B1:C3,2,FALSE),7)", 7.0},
		{"E27", "_xlfn.IFS(A1>100,1,A1>5,2)", 2.0},
		{"E28", "_xlfn.SWITCH(B1,\"pear\",1,\"apple\",2)", 2.0},
		{"E29", "SUMIF(B1:B3,\"apple\",C1:C3)", 14.0},
		{"E30", "SUMIFS(C1:C3,B1:B3,\"apple\",C1:C3,\">6\")", 9.0},
		{"E31", "COUNTIF(B1:B3,\"apple\")", 2.0},
		{"E32", "COUNTIFS(B1:B3,\"apple\",C1:C3,\">6\")", 1.0},
		{"E33", "AVERAGEIF(B1:B3,\"apple\",C1:C3)", 7.0},
		{"E34", "SUMPRODUCT(C1:C2,C1:C2)", 74.0},
		{"E35", "VLOOKUP(\"pear\",B1:C3,2,FALSE)", 7.0},
		{"E36", "INDEX(C1:C3,2)", 7.0},
		{"E37", "MATCH(\"pear\",B1:B3,0)", 2.0},
		{"E38", "_xlfn.XLOOKUP(\"apple\",B1:B3,C1:C3)", 5.0},
		{"E39", "CHOOSE(2,\"a\",\"b\",\"c\")", "b"},
		{"E40", "OFFSET(A1,1,0)", 20.0},
		{"E41", "INDIRECT(\"A2\")", 20.0},
		{"E42", "Other!A1", 100.0},
		{"E43", "SUM(Total)", 60.0},
		{"E44", "LEFT(A4,2)", "He"},
		{"E45", "RIGHT(A4,2)", "lo"},
		{"E46", "MID(A5,7,5)", "world"},
		{"E47", "LEN(A4)", 5.0},
		{"E48", "FIND(\"l\",A4)", 3.0},
		{"E49", "SEARCH(\"L\",A4)", 3.0},
		{"E50", "SUBSTITUTE(A4,\"l\",\"L\")", "HeLLo"},
		{"E51", "UPPER(A4)", "HELLO"},
		{"E52", "LOWER(A4)", "hello"},
		{"E53", "TRIM(\"  hi  \")", "hi"},
		{"E54", "_xlfn.CONCAT(A4,\"!\")", "Hello!"},
		{"E55", "_xlfn.TEXTJOIN(\"-\",TRUE,A4,\"x\")", "Hello-x"},
		{"E56", "VALUE(\"12.5\")", 12.5},
		{"E57", "YEAR(DATE(2024,3,5))", 2024.0},
		{"E58", "MONTH(D1)", 3.0},
		{"E59", "DAY(D1)", 5.0},
		{"E60", "DATEDIF(DATE(2024,1,1),DATE(2024,3,5),\"D\")", 64.0},
		{"E61", "ISBLANK(A6)", true},
		{"E62", "ISNUMBER(A1)", true},
		{"E63", "ISTEXT(A4)", true},
		{"E64", "ISERROR(A1/0)", true},
		{"E65", "A1/0", "#DIV/0!"},
		{"E66", "FALSE>\"TRUE\"", true},
		{"E67", "TIME(27,0,0)", 0.125},
		{"E68", "ROW(A1)", 1.0},
		{"E69", "DATEDIF(DATE(2020,1,15),DATE(2020,3,20),\"MD\")", 5.0},
	}
	ce := NewCalculationEngine(wb)
	matched := 0
	for _, tc := range cases {
		if err := ws.SetCellFormula(tc.cell, tc.f); err != nil {
			t.Fatal(err)
		}
		got, err := ce.CalculateCell(ws, tc.cell)
		if err != nil {
			got = err.Error()
		}
		if !calcEqual(got, tc.want) {
			t.Errorf("%s %s = %#v, want %#v", tc.cell, tc.f, got, tc.want)
			continue
		}
		matched++
	}
	t.Logf("engine self-check %d/%d", matched, len(cases))
	if matched != len(cases) {
		t.Fatalf("engine mismatches %d", len(cases)-matched)
	}

	dir := t.TempDir()
	path := filepath.Join(dir, "formulas.xlsx")
	f, err := os.Create(path)
	if err != nil {
		t.Fatal(err)
	}
	if err := NewXLSXWriter().Write(wb, f); err != nil {
		f.Close()
		t.Fatal(err)
	}
	f.Close()
	if _, err := exec.LookPath("soffice"); err != nil {
		t.Skip("libreoffice not installed")
	}
	out := t.TempDir()
	cmd := exec.Command("soffice", "--headless", "--norestore", "--convert-to", "xlsx", "--outdir", out, path)
	if outb, err := cmd.CombinedOutput(); err != nil {
		t.Fatalf("soffice: %v %s", err, outb)
	}
	lo, err := OpenFile(filepath.Join(out, "formulas.xlsx"))
	if err != nil {
		t.Fatal(err)
	}
	sheet, err := lo.GetSheetByName("Data")
	if err != nil {
		t.Fatal(err)
	}
	loOK := 0
	for _, tc := range cases {
		got, _ := sheet.GetCellValue(tc.cell)
		if calcEqual(got, tc.want) || loClose(got, tc.want) {
			loOK++
			continue
		}
		t.Logf("libreoffice mismatch %s %s = %#v, want %#v", tc.cell, tc.f, got, tc.want)
	}
	t.Logf("libreoffice match %d/%d", loOK, len(cases))
	if loOK*100/len(cases) < 90 {
		t.Fatalf("libreoffice match rate %d/%d", loOK, len(cases))
	}
}

func mustSet(t *testing.T, ws *Worksheet, ref string, v interface{}) {
	t.Helper()
	if err := ws.SetCellValue(ref, v); err != nil {
		t.Fatal(err)
	}
}

func calcEqual(got, want interface{}) bool {
	if gs, ok := got.(string); ok {
		if ws, ok := want.(string); ok {
			if isExcelError(gs) && isExcelError(ws) {
				return strings.EqualFold(gs, ws)
			}
			return gs == ws
		}
	}
	if gb, ok := got.(bool); ok {
		if wb, ok := want.(bool); ok {
			return gb == wb
		}
		if wf, ok := want.(float64); ok {
			return (gb && wf != 0) || (!gb && wf == 0)
		}
	}
	gf, gok := asNumber(got, true)
	wf, wok := asNumber(want, true)
	if gok && wok {
		return math.Abs(gf-wf) < 1e-6
	}
	return false
}

func loClose(got, want interface{}) bool {
	// LibreOffice sometimes returns 1/0 for booleans.
	s := ""
	switch v := got.(type) {
	case string:
		s = v
	case float64:
		s = strconv.FormatFloat(v, 'f', -1, 64)
	}
	if b, ok := want.(bool); ok {
		if b && (s == "1" || s == "TRUE") {
			return true
		}
		if !b && (s == "0" || s == "FALSE") {
			return true
		}
	}
	return false
}

func TestCorpusCachedFormulas(t *testing.T) {
	dir := filepath.Join("testdata", "excel")
	entries, err := os.ReadDir(dir)
	if err != nil {
		t.Fatal(err)
	}
	// A bounded sample of real workbooks. Cached Excel values are the ground truth.
	files := []string{
		"FormulaEvalTestData_Copy.xlsx",
		"shared_formulas.xlsx",
		filepath.Join("r11", "xw-sales.xlsx"),
		filepath.Join("r12", "ej-fibonacci.xlsx"),
		filepath.Join("r13", "ps-sharedformula.xlsx"),
		filepath.Join("r13", "npoi-matrix.xlsx"),
	}
	_ = entries
	checked, matched := 0, 0
	for _, name := range files {
		path := filepath.Join(dir, name)
		if _, err := os.Stat(path); err != nil {
			continue
		}
		wb, err := OpenFile(path)
		if err != nil {
			t.Logf("skip %s: %v", name, err)
			continue
		}
		ce := NewCalculationEngine(wb)
		for _, ws := range wb.Sheets() {
			for _, cell := range ws.AllCells() {
				if cell.Type != CellTypeFormula || cell.Value == nil {
					continue
				}
				ref, _ := CellName(cell.row, cell.col)
				got, err := ce.CalculateCell(ws, ref)
				if err != nil {
					continue
				}
				checked++
				if calcEqual(got, cell.Value) {
					matched++
				} else if checked-matched <= 12 {
					t.Logf("mismatch %s!%s %s got %#v cached %#v", ws.Title(), ref, cell.Formula, got, cell.Value)
				}
			}
		}
	}
	if checked == 0 {
		t.Fatal("no cached formulas were evaluated")
	}
	rate := matched * 100 / checked
	t.Logf("corpus cached match %d/%d (%d%%)", matched, checked, rate)
	if rate < 94 {
		t.Fatalf("corpus match rate %d%%", rate)
	}
	// Full corpus: how many formula cells evaluate, and how many cached
	// values they reproduce. Volatile functions are not compared.
	total, evaluated, fullChecked, fullMatched := 0, 0, 0, 0
	filepath.Walk(dir, func(path string, info os.FileInfo, err error) error {
		if err != nil || info.IsDir() || !strings.HasSuffix(strings.ToLower(path), ".xlsx") {
			return nil
		}
		wb, err := OpenFile(path)
		if err != nil {
			return nil
		}
		ce := NewCalculationEngine(wb)
		for _, ws := range wb.Sheets() {
			for _, cell := range ws.AllCells() {
				if cell.Type != CellTypeFormula || strings.TrimSpace(cell.Formula) == "" {
					continue
				}
				total++
				ref, _ := CellName(cell.row, cell.col)
				got, err := ce.CalculateCell(ws, ref)
				if err != nil {
					continue
				}
				evaluated++
				if cell.Value == nil || volatileFormula(cell.Formula) {
					continue
				}
				fullChecked++
				if calcEqual(got, cell.Value) {
					fullMatched++
				}
			}
		}
		return nil
	})
	t.Logf("corpus formulas %d evaluated %d cached match %d/%d", total, evaluated, fullMatched, fullChecked)
	if total == 0 || evaluated*100/total < 90 {
		t.Fatalf("evaluable formulas %d/%d", evaluated, total)
	}
	if fullChecked == 0 || fullMatched*100/fullChecked < 91 {
		t.Fatalf("full corpus match %d/%d", fullMatched, fullChecked)
	}
}

func volatileFormula(f string) bool {
	f = strings.ToUpper(f)
	for _, name := range []string{"RANDARRAY(", "RANDBETWEEN(", "RAND(", "NOW(", "TODAY("} {
		if strings.Contains(f, name) {
			return true
		}
	}
	return false
}
