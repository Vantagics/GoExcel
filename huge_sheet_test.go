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
	d := time.Since(start)
	var after runtime.MemStats
	runtime.ReadMemStats(&after)
	ws := wb.GetActiveSheet()
	heap := after.Alloc - before.Alloc
	t.Logf("open %s sheets=%d cells=%d heap=%dMB", d, wb.SheetCount(), ws.CellCount(), heap/1024/1024)
	if d > 30*time.Second {
		t.Fatalf("open too slow: %s", d)
	}
	if ws.CellCount() < 1000000 {
		t.Fatalf("cells = %d", ws.CellCount())
	}
}
