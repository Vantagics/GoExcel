package gospreadsheet

import (
	"bytes"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"
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

func TestFullPivotBuilder(t *testing.T) {
	wb := New()
	ws := wb.GetActiveSheet()
	headers := []string{"Region", "Product", "Channel", "Amount", "Qty", "Year", "When"}
	for i, h := range headers {
		cell, _ := CellName(0, i)
		ws.SetCellValue(cell, h)
	}
	type rec struct {
		region, product, channel string
		amount, qty, year        float64
		when                     time.Time
	}
	rows := []rec{
		{"North", "Widget", "Web", 10, 2, 2024, time.Date(2024, 1, 15, 0, 0, 0, 0, time.UTC)},
		{"North", "Gadget", "Store", 20, 1, 2024, time.Date(2024, 2, 2, 0, 0, 0, 0, time.UTC)},
		{"North", "Widget", "Store", 5, 3, 2024, time.Date(2024, 1, 20, 0, 0, 0, 0, time.UTC)},
		{"South", "Widget", "Web", 8, 4, 2023, time.Date(2023, 11, 1, 0, 0, 0, 0, time.UTC)},
		{"South", "Gadget", "Web", 12, 2, 2023, time.Date(2023, 12, 15, 0, 0, 0, 0, time.UTC)},
	}
	for i, r := range rows {
		ws.SetCellValue(cellAt(i+1, 0), r.region)
		ws.SetCellValue(cellAt(i+1, 1), r.product)
		ws.SetCellValue(cellAt(i+1, 2), r.channel)
		ws.SetCellValue(cellAt(i+1, 3), r.amount)
		ws.SetCellValue(cellAt(i+1, 4), r.qty)
		ws.SetCellValue(cellAt(i+1, 5), r.year)
		ws.SetCellValue(cellAt(i+1, 6), r.when)
	}
	on := true
	off := false
	pt, err := ws.BuildPivotTable(PivotBuild{
		Name: "Full", Source: "A1:G6", Dest: "A10",
		RowFields:    []string{"Region", "Product"},
		ColumnFields: []string{"Channel"},
		PageFields:   []string{"Year"},
		Values: []PivotValueField{
			{Field: "Amount", Function: "sum", NumberFormat: "0.00"},
			{Field: "Qty", Function: "count"},
			{Field: "Amount", Function: "average"},
			{Field: "Amount", Function: "max"},
			{Field: "Amount", Function: "min"},
		},
		RowGrandTotals: &on, ColumnGrandTotals: &on, Subtotals: &on,
	})
	if err != nil {
		t.Fatal(err)
	}
	if len(pt.RowFields) != 2 || len(pt.ColFields) != 1 || len(pt.PageFields) != 1 || len(pt.DataFields) != 5 {
		t.Fatalf("axes row=%v col=%v page=%v data=%d", pt.RowFields, pt.ColFields, pt.PageFields, len(pt.DataFields))
	}
	if pt.DataFields[0].Subtotal != "sum" || pt.DataFields[2].Subtotal != "average" {
		t.Fatalf("data = %+v", pt.DataFields)
	}
	// Report filter sits above the table. Empty intersections stay blank.
	if v, _ := ws.GetCellValue("A8"); v != "Year" {
		t.Fatalf("page field label = %#v", v)
	}
	if v, _ := ws.GetCellValue("B8"); v != "(All)" {
		t.Fatalf("page field value = %#v", v)
	}
	checks := map[string]interface{}{
		"C12": 20.0, "H13": 10.0, "M14": 35.0, "M18": 55.0, "N18": 5.0,
		"O18": 11.0, "P18": 20.0, "Q18": 5.0, "C18": 25.0, "H18": 30.0,
	}
	for addr, want := range checks {
		got, _ := ws.GetCellValue(addr)
		if got != want {
			t.Errorf("%s = %#v, want %#v", addr, got, want)
		}
	}
	if c, _ := ws.GetCellByName("H12"); c != nil && c.Value != nil {
		t.Fatalf("empty intersection H12 = %#v", c.Value)
	}
	grouped, err := ws.BuildPivotTable(PivotBuild{
		Name: "ByMonth", Source: "A1:G6", Dest: "A30",
		RowFields:      []string{"When"},
		Values:         []PivotValueField{{Field: "Amount", Function: "sum"}},
		Groupings:      []PivotGrouping{{Field: "When", Dates: []string{"years", "months"}}},
		RowGrandTotals: &on, ColumnGrandTotals: &off, Subtotals: &off,
	})
	if err != nil {
		t.Fatal(err)
	}
	if !containsStr(grouped.RowFields, "Years") || !containsStr(grouped.RowFields, "When") {
		t.Fatalf("grouped rows = %v", grouped.RowFields)
	}
	bins, err := ws.BuildPivotTable(PivotBuild{
		Name: "Bins", Source: "A1:G6", Dest: "A40",
		RowFields:      []string{"Amount"},
		Values:         []PivotValueField{{Field: "Qty", Function: "sum"}},
		Groupings:      []PivotGrouping{{Field: "Amount", Start: 0, End: 30, Interval: 10}},
		RowGrandTotals: &on, ColumnGrandTotals: &off, Subtotals: &off,
	})
	if err != nil {
		t.Fatal(err)
	}
	if bins.Location == "" {
		t.Fatal("bins location empty")
	}
	groupedChecks := map[string]interface{}{
		"C31": 12.0, "C32": 8.0, "C33": 20.0, "C34": 15.0, "C35": 55.0,
		"B41": 7.0, "B42": 4.0, "B43": 1.0, "B44": 12.0,
	}
	for addr, want := range groupedChecks {
		got, _ := ws.GetCellValue(addr)
		if got != want {
			t.Errorf("%s = %#v, want %#v", addr, got, want)
		}
	}
	var buf bytes.Buffer
	if err := NewXLSXWriter().Write(wb, &buf); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(os.TempDir(), "full-pivot.xlsx")
	if err := os.WriteFile(path, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}
	files := zipFiles(t, buf.Bytes())
	var table string
	for name, b := range files {
		if strings.HasSuffix(name, "pivotTable1.xml") {
			table = string(b)
		}
	}
	for _, needle := range []string{`subtotal="sum"`, `subtotal="count"`, `subtotal="average"`, `subtotal="max"`, `subtotal="min"`, `numFmtId="2"`, `<pageField`, `field x="-2"`, `<rowItems`, `<colItems`, `axis="axisPage"`} {
		if !strings.Contains(table, needle) {
			t.Errorf("pivot table XML missing %s", needle)
		}
	}
	got, err := NewXLSXReader().Read(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}
	if len(got.PivotTables()) != 3 {
		t.Fatalf("reloaded %d pivots", len(got.PivotTables()))
	}
}

func cellAt(row, col int) string {
	s, _ := CellName(row, col)
	return s
}

func containsStr(ss []string, want string) bool {
	for _, s := range ss {
		if s == want {
			return true
		}
	}
	return false
}
