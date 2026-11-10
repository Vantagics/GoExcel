package gospreadsheet

import (
	"bytes"
	"os"
	"path/filepath"
	"testing"
)

func TestCreatePivotAndSlicer(t *testing.T) {
	wb := New()
	ws := wb.GetActiveSheet()
	ws.SetCellValue("A1", "Region")
	ws.SetCellValue("B1", "Amount")
	ws.SetCellValue("A2", "North")
	ws.SetCellValue("B2", 10.0)
	ws.SetCellValue("A3", "South")
	ws.SetCellValue("B3", 5.0)
	ws.SetCellValue("A4", "North")
	ws.SetCellValue("B4", 7.0)
	pt, err := ws.AddPivotTable("SalesPivot", "A1:B4", "E3", "Region", "Amount")
	if err != nil {
		t.Fatal(err)
	}
	if pt.Location == "" || pt.DataCaption != "Values" || len(pt.CacheFields) != 2 {
		t.Fatalf("pivot = %+v", pt)
	}
	if v, _ := ws.GetCellValue("E4"); v != "North" {
		t.Fatalf("E4 = %#v", v)
	}
	if v, _ := ws.GetCellValue("F4"); v != 17.0 {
		t.Fatalf("North sum = %#v", v)
	}
	if v, _ := ws.GetCellValue("F6"); v != 22.0 {
		t.Fatalf("grand = %#v", v)
	}
	sl, err := ws.AddSlicer("SalesPivot", "Region", "Region")
	if err != nil {
		t.Fatal(err)
	}
	if sl.Caption != "Region" || len(sl.author.items) != 2 {
		t.Fatalf("slicer = %+v items=%v", sl, sl.author.items)
	}
	var buf bytes.Buffer
	if err := NewXLSXWriter().Write(wb, &buf); err != nil {
		t.Fatal(err)
	}
	out := filepath.Join(os.TempDir(), "created-pivot.xlsx")
	if err := os.WriteFile(out, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	got, err := NewXLSXReader().Read(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}
	tables := got.PivotTables()
	if len(tables) != 1 || tables[0].Name != "SalesPivot" || tables[0].Location == "" {
		t.Fatalf("reloaded pivots = %+v", tables)
	}
	if len(tables[0].RowFields) == 0 || tables[0].RowFields[0] != "Region" {
		t.Fatalf("row fields = %v cache=%v", tables[0].RowFields, tables[0].CacheFields)
	}
	if len(tables[0].DataFields) != 1 || tables[0].DataFields[0].FieldName != "Amount" {
		t.Fatalf("data fields = %+v", tables[0].DataFields)
	}
	slicers := got.Sheets()[0].Slicers()
	if len(slicers) != 1 || slicers[0].Caption != "Region" || slicers[0].Cache == "" {
		t.Fatalf("slicers = %+v", slicers)
	}
	if len(got.SlicerCaches()) != 1 || got.SlicerCaches()[0].SourceName != "Region" {
		t.Fatalf("caches = %+v", got.SlicerCaches())
	}
}
