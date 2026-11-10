package gospreadsheet

import (
	"archive/zip"
	"fmt"
	"math"
	"math/rand"
	"strings"
)

func (ce *CalculationEngine) writeSpill(ws *Worksheet, anchor *Cell, rg *formulaRange) error {
	if ws == nil || anchor == nil || rg == nil || rg.cellCount() <= 1 {
		return nil
	}
	for r := 0; r < rg.rows; r++ {
		for c := 0; c < rg.cols; c++ {
			if r == 0 && c == 0 {
				continue
			}
			row, col := anchor.row+r, anchor.col+c
			if row < 0 || col < 0 {
				return fmt.Errorf("#SPILL!")
			}
			if existing := ws.GetCellIfExists(row, col); existing != nil && existing.Formula != "" {
				return fmt.Errorf("#SPILL!")
			}
		}
	}
	for r := 0; r < rg.rows; r++ {
		for c := 0; c < rg.cols; c++ {
			if r == 0 && c == 0 {
				continue
			}
			cell := ws.GetCell(anchor.row+r, anchor.col+c)
			val := rg.at(r, c)
			if err := excelError(val); err != nil {
				cell.Value = err.Error()
				cell.Type = CellTypeError
				continue
			}
			if val == nil {
				cell.SetValue("")
				continue
			}
			cell.SetValue(val)
		}
	}
	if anchor.cellMeta == "" {
		anchor.cellMeta = "1"
	}
	if ws.parent != nil {
		ws.parent.dynamicArray = true
	}
	return nil
}

func dynamicArrayMetadataXML() string {
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	b.WriteString(`<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">`)
	b.WriteString(`<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1"/></metadataTypes>`)
	b.WriteString(`<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{B025F08A-C4E3-4398-8D55-6B4B62B1B4E1}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>`)
	b.WriteString(`<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>`)
	b.WriteString(`</metadata>`)
	return b.String()
}

func (w *XLSXWriter) writeDynamicArrayMeta(zw *zip.Writer, wb *Workbook) error {
	if wb == nil || !wb.dynamicArray || wb.hasCellImages() {
		return nil
	}
	if wb.source != nil {
		if _, ok := wb.source.file("xl/metadata.xml"); ok {
			return nil
		}
	}
	return writeZip(zw, "xl/metadata.xml", dynamicArrayMetadataXML())
}

func fnMMult(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) != 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	a, err := ce.evalRange(ws, parts[0])
	if err != nil {
		return nil, err
	}
	b, err := ce.evalRange(ws, parts[1])
	if err != nil {
		return nil, err
	}
	if err := rangeError(a); err != nil {
		return nil, err
	}
	if err := rangeError(b); err != nil {
		return nil, err
	}
	if a.cols != b.rows {
		return nil, fmt.Errorf("#VALUE!")
	}
	vals := make([]interface{}, a.rows*b.cols)
	for i := 0; i < a.rows; i++ {
		for j := 0; j < b.cols; j++ {
			sum := 0.0
			for k := 0; k < a.cols; k++ {
				av, aok := asNumber(a.at(i, k), false)
				bv, bok := asNumber(b.at(k, j), false)
				if !aok || !bok {
					return nil, fmt.Errorf("#VALUE!")
				}
				sum += av * bv
			}
			vals[i*b.cols+j] = sum
		}
	}
	return &formulaRange{values: vals, rows: a.rows, cols: b.cols, virtual: true, row0: ce.row, col0: ce.col}, nil
}

func fnMDeterm(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	rg, err := ce.evalRange(ws, args)
	if err != nil {
		return nil, err
	}
	m, err := squareMatrix(rg)
	if err != nil {
		return nil, err
	}
	return determinant(m), nil
}

func fnMInverse(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	rg, err := ce.evalRange(ws, args)
	if err != nil {
		return nil, err
	}
	m, err := squareMatrix(rg)
	if err != nil {
		return nil, err
	}
	inv, err := invertMatrix(m)
	if err != nil {
		return nil, err
	}
	n := len(inv)
	vals := make([]interface{}, n*n)
	for i := 0; i < n; i++ {
		for j := 0; j < n; j++ {
			vals[i*n+j] = inv[i][j]
		}
	}
	return &formulaRange{values: vals, rows: n, cols: n, virtual: true, row0: ce.row, col0: ce.col}, nil
}

func squareMatrix(rg *formulaRange) ([][]float64, error) {
	if rg.rows != rg.cols || rg.rows == 0 {
		return nil, fmt.Errorf("#VALUE!")
	}
	if err := rangeError(rg); err != nil {
		return nil, err
	}
	n := rg.rows
	m := make([][]float64, n)
	for i := 0; i < n; i++ {
		m[i] = make([]float64, n)
		for j := 0; j < n; j++ {
			v, ok := asNumber(rg.at(i, j), false)
			if !ok {
				return nil, fmt.Errorf("#VALUE!")
			}
			m[i][j] = v
		}
	}
	return m, nil
}

func determinant(m [][]float64) float64 {
	n := len(m)
	a := copyMatrix(m)
	det := 1.0
	for k := 0; k < n; k++ {
		pivot := k
		max := math.Abs(a[k][k])
		for i := k + 1; i < n; i++ {
			if math.Abs(a[i][k]) > max {
				max = math.Abs(a[i][k])
				pivot = i
			}
		}
		if max < 1e-12 {
			return 0
		}
		if pivot != k {
			a[k], a[pivot] = a[pivot], a[k]
			det = -det
		}
		det *= a[k][k]
		for i := k + 1; i < n; i++ {
			f := a[i][k] / a[k][k]
			for j := k; j < n; j++ {
				a[i][j] -= f * a[k][j]
			}
		}
	}
	return det
}

func invertMatrix(m [][]float64) ([][]float64, error) {
	n := len(m)
	a := make([][]float64, n)
	for i := 0; i < n; i++ {
		a[i] = make([]float64, 2*n)
		copy(a[i], m[i])
		a[i][n+i] = 1
	}
	for k := 0; k < n; k++ {
		pivot := k
		max := math.Abs(a[k][k])
		for i := k + 1; i < n; i++ {
			if math.Abs(a[i][k]) > max {
				max = math.Abs(a[i][k])
				pivot = i
			}
		}
		if max < 1e-12 {
			return nil, fmt.Errorf("#NUM!")
		}
		a[k], a[pivot] = a[pivot], a[k]
		div := a[k][k]
		for j := 0; j < 2*n; j++ {
			a[k][j] /= div
		}
		for i := 0; i < n; i++ {
			if i == k {
				continue
			}
			f := a[i][k]
			for j := 0; j < 2*n; j++ {
				a[i][j] -= f * a[k][j]
			}
		}
	}
	out := make([][]float64, n)
	for i := 0; i < n; i++ {
		out[i] = append([]float64(nil), a[i][n:]...)
	}
	return out, nil
}

func copyMatrix(m [][]float64) [][]float64 {
	out := make([][]float64, len(m))
	for i := range m {
		out[i] = append([]float64(nil), m[i]...)
	}
	return out
}

func fnColumns(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	rg, err := refGeometry(ce, ws, args)
	if err != nil {
		return nil, err
	}
	return float64(rg.cols), nil
}

func fnRows(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	rg, err := refGeometry(ce, ws, args)
	if err != nil {
		return nil, err
	}
	return float64(rg.rows), nil
}

func fnRandBetween(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) != 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	lo, err := evalMust(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	hi, err := evalMust(ce, ws, parts[1])
	if err != nil {
		return nil, err
	}
	a, b := int(math.Ceil(lo)), int(math.Floor(hi))
	if a > b {
		return nil, fmt.Errorf("#NUM!")
	}
	return float64(a + rand.Intn(b-a+1)), nil
}

func fnIrr(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 1 {
		return nil, fmt.Errorf("#VALUE!")
	}
	cash, err := collectList(ce, ws, parts[0], false)
	if err != nil {
		return nil, err
	}
	guess := 0.1
	if len(parts) > 1 && strings.TrimSpace(parts[1]) != "" {
		guess, err = evalMust(ce, ws, parts[1])
		if err != nil {
			return nil, err
		}
	}
	rate := guess
	for i := 0; i < 40; i++ {
		npv, deriv := 0.0, 0.0
		for t, c := range cash {
			den := math.Pow(1+rate, float64(t))
			npv += c / den
			if t > 0 {
				deriv -= float64(t) * c / (den * (1 + rate))
			}
		}
		if math.Abs(npv) < 1e-8 {
			return rate, nil
		}
		if deriv == 0 || math.IsNaN(deriv) {
			break
		}
		rate -= npv / deriv
		if rate <= -1 {
			rate = -0.9
		}
	}
	return nil, fmt.Errorf("#NUM!")
}

func fnTimeValue(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	v, err := ce.evaluate(ws, args)
	if err != nil {
		return nil, err
	}
	s, err := ce.textArg(v)
	if err != nil {
		if n, ok := asNumber(v, true); ok {
			frac := n - math.Floor(n)
			if frac < 0 {
				frac += 1
			}
			return frac, nil
		}
		return nil, err
	}
	var h, m, sec float64
	if _, err := fmt.Sscanf(s, "%f:%f:%f", &h, &m, &sec); err != nil {
		if _, err2 := fmt.Sscanf(s, "%f:%f", &h, &m); err2 != nil {
			return nil, fmt.Errorf("#VALUE!")
		}
	}
	day := (h*3600 + m*60 + sec) / 86400
	day -= math.Floor(day)
	return day, nil
}

func fnNormDist(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 4 {
		return nil, fmt.Errorf("#VALUE!")
	}
	x, err := evalMust(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	mean, err := evalMust(ce, ws, parts[1])
	if err != nil {
		return nil, err
	}
	sd, err := evalMust(ce, ws, parts[2])
	if err != nil {
		return nil, err
	}
	if sd <= 0 {
		return nil, fmt.Errorf("#NUM!")
	}
	cumV, err := ce.evaluate(ws, parts[3])
	if err != nil {
		return nil, err
	}
	cum, _ := isTruthy(cumV)
	z := (x - mean) / sd
	if cum {
		return 0.5 * (1 + math.Erf(z/math.Sqrt2)), nil
	}
	return math.Exp(-z*z/2) / (sd * math.Sqrt(2*math.Pi)), nil
}

func fnNormSDist(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 1 {
		return nil, fmt.Errorf("#VALUE!")
	}
	z, err := evalMust(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	cum := true
	if len(parts) > 1 {
		v, err := ce.evaluate(ws, parts[1])
		if err != nil {
			return nil, err
		}
		cum, _ = isTruthy(v)
	}
	if cum {
		return 0.5 * (1 + math.Erf(z/math.Sqrt2)), nil
	}
	return math.Exp(-z*z/2) / math.Sqrt(2*math.Pi), nil
}

func fnPoisson(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	x, err := evalMust(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	mean, err := evalMust(ce, ws, parts[1])
	if err != nil {
		return nil, err
	}
	if mean < 0 || x < 0 {
		return nil, fmt.Errorf("#NUM!")
	}
	cum := false
	if len(parts) > 2 {
		v, err := ce.evaluate(ws, parts[2])
		if err != nil {
			return nil, err
		}
		cum, _ = isTruthy(v)
	}
	k := int(math.Floor(x))
	if !cum {
		return poissonPMF(k, mean), nil
	}
	sum := 0.0
	for i := 0; i <= k; i++ {
		sum += poissonPMF(i, mean)
	}
	return sum, nil
}

func poissonPMF(k int, mean float64) float64 {
	if k < 0 {
		return 0
	}
	// exp(-mean) * mean^k / k!
	logp := -mean + float64(k)*math.Log(mean) - logFact(k)
	return math.Exp(logp)
}

func logFact(n int) float64 {
	s := 0.0
	for i := 2; i <= n; i++ {
		s += math.Log(float64(i))
	}
	return s
}

func fnPercentile(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) != 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	nums, err := collectList(ce, ws, parts[0], false)
	if err != nil {
		return nil, err
	}
	if len(nums) == 0 {
		return nil, fmt.Errorf("#NUM!")
	}
	k, err := evalMust(ce, ws, parts[1])
	if err != nil {
		return nil, err
	}
	if k < 0 || k > 1 {
		return nil, fmt.Errorf("#NUM!")
	}
	for i := 0; i < len(nums); i++ {
		for j := i + 1; j < len(nums); j++ {
			if nums[j] < nums[i] {
				nums[i], nums[j] = nums[j], nums[i]
			}
		}
	}
	pos := k * float64(len(nums)-1)
	lo := int(math.Floor(pos))
	hi := int(math.Ceil(pos))
	if lo == hi {
		return nums[lo], nil
	}
	w := pos - float64(lo)
	return nums[lo]*(1-w) + nums[hi]*w, nil
}
