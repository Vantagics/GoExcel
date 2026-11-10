package gospreadsheet

import (
	"bytes"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"testing"
	"time"
)

func TestCellImageRoundTrip(t *testing.T) {
	wb, err := OpenFile("testdata/excel/r12/rw-embed-image.xlsx")
	if err != nil {
		t.Fatal(err)
	}
	img, ok := wb.GetActiveSheet().GetCell(0, 0).Image()
	if !ok || len(img.Data) < 8 || img.Extension != "png" {
		t.Fatalf("loaded image ok=%v len=%d ext=%s", ok, len(img.Data), img.Extension)
	}
	var buf bytes.Buffer
	if err := NewXLSXWriter().Write(wb, &buf); err != nil {
		t.Fatal(err)
	}
	got, err := NewXLSXReader().Read(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}
	again, ok := got.GetActiveSheet().GetCell(0, 0).Image()
	if !ok || !bytes.Equal(again.Data, img.Data) {
		t.Fatalf("rewritten image len=%d", len(again.Data))
	}

	fresh := New()
	fresh.GetActiveSheet().GetCell(1, 1).SetImage([]byte{0x89, 'P', 'N', 'G', 9, 8, 7}, "png")
	buf.Reset()
	if err := NewXLSXWriter().Write(fresh, &buf); err != nil {
		t.Fatal(err)
	}
	files := zipFiles(t, buf.Bytes())
	for _, name := range []string{"xl/metadata.xml", "xl/richData/rdrichvalue.xml", "xl/richData/richValueRel.xml"} {
		if _, ok := files[name]; !ok {
			t.Errorf("created workbook missing %s", name)
		}
	}
	if !strings.Contains(string(files["xl/worksheets/sheet1.xml"]), `vm="1"`) {
		t.Fatal("created image cell has no vm")
	}
	re, err := NewXLSXReader().Read(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}
	created, ok := re.GetActiveSheet().GetCell(1, 1).Image()
	if !ok || !bytes.Equal(created.Data, []byte{0x89, 'P', 'N', 'G', 9, 8, 7}) {
		t.Fatalf("created image = %v", created.Data)
	}
}

func TestThreadedCommentAPI(t *testing.T) {
	wb, err := OpenFile("testdata/excel/r12/rw-threaded.xlsx")
	if err != nil {
		t.Fatal(err)
	}
	threads := wb.GetActiveSheet().ThreadedComments()
	if len(threads) != 1 || threads[0].Text != "thread note" || threads[0].Cell != "A1" {
		t.Fatalf("loaded threads = %+v", threads)
	}
	if len(wb.Persons()) != 1 || wb.Persons()[0].DisplayName != "Ada" {
		t.Fatalf("persons = %+v", wb.Persons())
	}

	fresh := New()
	ws := fresh.GetActiveSheet()
	ada := fresh.AddPerson("Ada", "ada@example.com")
	root, err := ws.AddThreadedComment("B2", ada.ID, "hello")
	if err != nil {
		t.Fatal(err)
	}
	if _, err := ws.ReplyToThreadedComment(root.ID, ada.ID, "reply"); err != nil {
		t.Fatal(err)
	}
	if err := ws.SetThreadedCommentResolved(root.ID, true); err != nil {
		t.Fatal(err)
	}
	note, err := ws.GetCellByName("B2")
	if err != nil || note.Comment == nil || !strings.Contains(note.Comment.Text, "hello") || !strings.Contains(note.Comment.Text, "reply") {
		t.Fatalf("legacy note = %#v err=%v", note.Comment, err)
	}
	var buf bytes.Buffer
	if err := NewXLSXWriter().Write(fresh, &buf); err != nil {
		t.Fatal(err)
	}
	files := zipFiles(t, buf.Bytes())
	body := string(files["xl/threadedComments/threadedComment1.xml"])
	if !strings.Contains(body, "hello") || !strings.Contains(body, "reply") || !strings.Contains(body, `done="1"`) || !strings.Contains(body, "parentId") {
		t.Fatalf("thread xml = %s", body)
	}
	if !strings.Contains(string(files["xl/comments1.xml"]), "hello") {
		t.Fatal("legacy comment part missing the thread text")
	}
	got, err := NewXLSXReader().Read(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}
	list := got.GetActiveSheet().ThreadedComments()
	if len(list) != 2 || !list[0].Resolved || list[1].ParentID == "" {
		t.Fatalf("reloaded threads = %+v", list)
	}
	legacy, _ := got.GetActiveSheet().GetCellByName("B2")
	if legacy.Comment == nil || !strings.Contains(legacy.Comment.Text, "reply") {
		t.Fatalf("reloaded note = %#v", legacy.Comment)
	}
}

func TestStreamRoundTrip(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "stream.xlsx")
	w, err := CreateStream(path)
	if err != nil {
		t.Fatal(err)
	}
	sh, err := w.AddSheet("Data")
	if err != nil {
		t.Fatal(err)
	}
	if err := sh.WriteRow(1, []StreamCell{{Column: 0, Value: "name"}, {Column: 1, Value: "n"}}); err != nil {
		t.Fatal(err)
	}
	if err := sh.WriteRow(2, []StreamCell{{Column: 0, Value: "a"}, {Column: 1, Value: 1.5}, {Column: 2, Value: true}}); err != nil {
		t.Fatal(err)
	}
	if err := sh.WriteRow(4, []StreamCell{{Column: 0, Formula: "SUM(B2)", Value: 1.5}}); err != nil {
		t.Fatal(err)
	}
	if err := w.Close(); err != nil {
		t.Fatal(err)
	}
	it, err := OpenRows(path, "Data")
	if err != nil {
		t.Fatal(err)
	}
	defer it.Close()
	var rows []StreamRow
	for it.Next() {
		r := it.Row()
		r.Cells = append([]StreamCell(nil), r.Cells...)
		rows = append(rows, r)
	}
	if err := it.Err(); err != nil {
		t.Fatal(err)
	}
	if len(rows) != 3 || rows[0].Cells[0].Value != "name" || rows[1].Cells[1].Value != 1.5 || rows[1].Cells[2].Value != true {
		t.Fatalf("rows = %#v", rows)
	}
	if rows[2].Number != 4 || rows[2].Cells[0].Formula != "SUM(B2)" || rows[2].Cells[0].Value != 1.5 {
		t.Fatalf("formula row = %#v", rows[2])
	}
}

func TestStreamStylesAndFormulas(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "styled.xlsx")
	w, err := CreateStream(path)
	if err != nil {
		t.Fatal(err)
	}
	w.UseSharedStrings()
	sh, err := w.AddSheet("Styled")
	if err != nil {
		t.Fatal(err)
	}
	if err := sh.SetColWidth(0, 18); err != nil {
		t.Fatal(err)
	}
	if err := sh.Freeze(1, 1); err != nil {
		t.Fatal(err)
	}
	if err := sh.MergeCells("A1:B1"); err != nil {
		t.Fatal(err)
	}
	bold := NewStyle().SetFont(&Font{Name: "Calibri", Size: 12, Bold: true, Color: "FF0000"})
	fill := NewStyle().SetFill(&Fill{Type: "solid", Color: "FFFF00"})
	border := NewStyle().SetBorders(&Borders{Left: Border{Style: BorderThin, Color: "000000"}, Right: Border{Style: BorderThin}, Top: Border{Style: BorderThin}, Bottom: Border{Style: BorderThin}})
	num := NewStyle()
	num.NumberFormat = &FormatNumber2Dec
	when := time.Date(2024, 3, 5, 0, 0, 0, 0, time.UTC)
	if err := sh.WriteRow(1, []StreamCell{{Column: 0, Value: "Title", Style: bold}, {Column: 1, Value: "Title"}}); err != nil {
		t.Fatal(err)
	}
	if err := sh.WriteRow(2, []StreamCell{
		{Column: 0, Value: "bolt", Style: fill, Link: "https://example.com"},
		{Column: 1, Value: 3.5, Style: num},
		{Column: 2, Value: when},
		{Column: 3, Formula: "SUM(B2,1)", Value: 4.5, Style: border},
	}); err != nil {
		t.Fatal(err)
	}
	if err := w.Close(); err != nil {
		t.Fatal(err)
	}
	files := zipFiles(t, mustRead(t, path))
	sst := string(files["xl/sharedStrings.xml"])
	if !strings.Contains(sst, "bolt") || strings.Count(sst, "<si>") != 2 {
		t.Fatalf("shared strings = %s", sst)
	}
	if !strings.Contains(string(files["xl/workbook.xml"]), `fullCalcOnLoad="1"`) {
		t.Fatal("workbook missing fullCalcOnLoad")
	}
	sheet := string(files["xl/worksheets/sheet1.xml"])
	for _, needle := range []string{`customWidth="1"`, `state="frozen"`, `mergeCell ref="A1:B1"`, `hyperlink`, `<f>SUM(B2,1)</f>`} {
		if !strings.Contains(sheet, needle) {
			t.Errorf("sheet missing %s", needle)
		}
	}
	wb, err := OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	ws := wb.GetActiveSheet()
	if ws.GetFreezePane() == nil {
		t.Fatal("freeze pane was not read")
	}
	if len(ws.GetMergeCells()) != 1 {
		t.Fatalf("merges = %#v", ws.GetMergeCells())
	}
	link, _ := ws.GetCellByName("A2")
	if link.Hyperlink == nil || link.Hyperlink.URL != "https://example.com" {
		t.Fatalf("link = %#v", link.Hyperlink)
	}
	if link.Style == nil || link.Style.Fill == nil {
		t.Fatal("fill style missing")
	}
	n, _ := ws.GetCellByName("B2")
	if n.GetFormattedValue() != "3.50" {
		t.Fatalf("formatted number = %q", n.GetFormattedValue())
	}
	it, err := OpenRows(path, "Styled", WithFormattedValues())
	if err != nil {
		t.Fatal(err)
	}
	defer it.Close()
	var sawDate bool
	for it.Next() {
		for _, c := range it.Row().Cells {
			if c.Column == 2 && it.Row().Number == 2 {
				sawDate = true
				if c.Formatted == "" {
					t.Fatal("date cell has no formatted value")
				}
				if _, ok := c.Value.(time.Time); !ok {
					t.Fatalf("date value = %#v formatted %q", c.Value, c.Formatted)
				}
			}
			if c.Column == 1 && it.Row().Number == 2 && c.Formatted != "3.50" {
				t.Fatalf("stream formatted = %q", c.Formatted)
			}
		}
	}
	if !sawDate {
		t.Fatal("date cell was not streamed")
	}
}

func mustRead(t *testing.T, path string) []byte {
	t.Helper()
	b, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	return b
}

func TestCalcCommonFunctions(t *testing.T) {
	wb := New()
	ws := wb.GetActiveSheet()
	_ = ws.SetCellValue("A1", 1.5)
	_ = ws.SetCellValue("A2", 2.5)
	_ = ws.SetCellFormula("A3", "SUM(A1:A2)")
	_ = ws.SetCellFormula("A4", "AVERAGE(A1:A2)")
	_ = ws.SetCellFormula("A5", "IF(A1>1,\"yes\",\"no\")")
	_ = ws.SetCellFormula("A6", "ROUND(A1,0)")
	_ = ws.SetCellFormula("A7", "MAX(A1,A2)")
	if err := NewCalculationEngine(wb).CalculateAll(); err != nil {
		t.Fatal(err)
	}
	if v, _ := ws.GetCellValue("A3"); v != 4.0 {
		t.Fatalf("SUM = %#v", v)
	}
	if v, _ := ws.GetCellValue("A4"); v != 2.0 {
		t.Fatalf("AVERAGE = %#v", v)
	}
	if v, _ := ws.GetCellValue("A5"); v != "yes" {
		t.Fatalf("IF = %#v", v)
	}
	if v, _ := ws.GetCellValue("A6"); v != 2.0 {
		t.Fatalf("ROUND = %#v", v)
	}
	if v, _ := ws.GetCellValue("A7"); v != 2.5 {
		t.Fatalf("MAX = %#v", v)
	}
}

func TestStreamMemory(t *testing.T) {
	if _, err := os.Stat("testdata/perf/ej-huge.xlsx"); err != nil {
		t.Skip(err)
	}
	lazyTime, lazyHeap := peakHeap(func() {
		wb, err := OpenFile("testdata/perf/ej-huge.xlsx")
		if err != nil {
			t.Fatal(err)
		}
		if wb.SheetCount() < 1 {
			t.Fatal("no sheets")
		}
		runtime.KeepAlive(wb)
	})
	fullTime, fullHeap := peakHeap(func() {
		wb, err := OpenFile("testdata/perf/ej-huge.xlsx")
		if err != nil {
			t.Fatal(err)
		}
		if wb.GetActiveSheet().CellCount() < 1000000 {
			t.Fatalf("cells %d", wb.GetActiveSheet().CellCount())
		}
		runtime.KeepAlive(wb)
	})
	streamTime, streamHeap := peakHeap(func() {
		it, err := OpenRows("testdata/perf/ej-huge.xlsx", "")
		if err != nil {
			t.Fatal(err)
		}
		n := 0
		var last interface{}
		for it.Next() {
			n += len(it.Row().Cells)
			if it.Row().Number == 2 && len(it.Row().Cells) > 0 {
				last = it.Row().Cells[0].Value
			}
		}
		if err := it.Err(); err != nil {
			t.Fatal(err)
		}
		it.Close()
		if n < 1000000 {
			t.Fatalf("streamed cells %d last %v", n, last)
		}
		runtime.KeepAlive(last)
	})
	dir := t.TempDir()
	out := filepath.Join(dir, "million.xlsx")
	writeTime, writeHeap := peakHeap(func() {
		w, err := CreateStream(out)
		if err != nil {
			t.Fatal(err)
		}
		sh, err := w.AddSheet("Big")
		if err != nil {
			t.Fatal(err)
		}
		const rows = 250000
		const cols = 8
		cells := make([]StreamCell, cols)
		for r := 1; r <= rows; r++ {
			for c := 0; c < cols; c++ {
				cells[c] = StreamCell{Column: c, Value: float64(r*cols + c)}
			}
			if err := sh.WriteRow(r, cells); err != nil {
				t.Fatal(err)
			}
		}
		if err := w.Close(); err != nil {
			t.Fatal(err)
		}
	})
	t.Logf("open lazy %s heap=%dMB", lazyTime, lazyHeap/1024/1024)
	t.Logf("open loaded %s heap=%dMB", fullTime, fullHeap/1024/1024)
	t.Logf("stream read %s heap=%dMB", streamTime, streamHeap/1024/1024)
	t.Logf("stream write 2000000 cells %s heap=%dMB", writeTime, writeHeap/1024/1024)
	if streamHeap > fullHeap/2 && streamHeap > 256*1024*1024 {
		t.Fatalf("stream read heap %dMB is not bounded vs full %dMB", streamHeap/1024/1024, fullHeap/1024/1024)
	}
	if writeHeap > 128*1024*1024 {
		t.Fatalf("stream write heap %dMB", writeHeap/1024/1024)
	}
	it, err := OpenRows(out, "Big")
	if err != nil {
		t.Fatal(err)
	}
	defer it.Close()
	if !it.Next() {
		t.Fatal(it.Err())
	}
	if it.Row().Cells[0].Value != float64(8) && it.Row().Cells[0].Value != float64(1*8+0) {
		// r=1, c=0 -> 1*8+0 = 8
		t.Fatalf("first cell %#v", it.Row().Cells[0].Value)
	}
}

func peakHeap(fn func()) (time.Duration, uint64) {
	runtime.GC()
	var peak uint64
	stop := make(chan struct{})
	done := make(chan struct{})
	go func() {
		defer close(done)
		var m runtime.MemStats
		for {
			runtime.ReadMemStats(&m)
			if m.HeapAlloc > peak {
				peak = m.HeapAlloc
			}
			select {
			case <-stop:
				return
			case <-time.After(15 * time.Millisecond):
			}
		}
	}()
	start := time.Now()
	fn()
	close(stop)
	<-done
	return time.Since(start), peak
}
