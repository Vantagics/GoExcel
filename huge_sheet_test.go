package gospreadsheet

import (
	"runtime"
	"testing"
	"time"
)

func TestHugeSheet(t *testing.T) {
	runtime.GC()
	var before runtime.MemStats
	runtime.ReadMemStats(&before)
	start := time.Now()
	wb, err := OpenFile("testdata/perf/ej-huge.xlsx")
	if err != nil {
		t.Fatal(err)
	}
	ws := wb.GetActiveSheet()
	n := ws.CellCount()
	d := time.Since(start)
	runtime.GC()
	var after runtime.MemStats
	runtime.ReadMemStats(&after)
	heap := after.Alloc - before.Alloc
	t.Logf("open+load %s sheets=%d cells=%d retained=%dMB", d, wb.SheetCount(), n, heap/1024/1024)
	if d > 30*time.Second {
		t.Fatalf("open too slow: %s", d)
	}
	if ws.CellCount() < 1000000 {
		t.Fatalf("cells = %d", ws.CellCount())
	}
}
