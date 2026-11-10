package gospreadsheet

import (
	"fmt"
	"math"
	"strings"
)

func powExcel(ln, rn float64) (interface{}, error) {
	pow := math.Pow(ln, rn)
	if math.IsNaN(pow) || math.IsInf(pow, 0) {
		return nil, fmt.Errorf("#NUM!")
	}
	return pow, nil
}

// excelFault is an Excel error stored in a range so aggregates can see it.
type excelFault string

func (ce *CalculationEngine) evalFormulaCell(ws *Worksheet, cell *Cell) (interface{}, error) {
	key := ""
	if ce.memo != nil && ws != nil {
		ref, _ := CellName(cell.row, cell.col)
		key = ws.Title() + "!" + ref
		if hit, ok := ce.memo[key]; ok {
			return hit.v, hit.err
		}
	}
	finish := func(v interface{}, err error) (interface{}, error) {
		if key != "" && ce.memo != nil {
			ce.memo[key] = calcMemo{v: v, err: err}
		}
		return v, err
	}
	prevR, prevC, prevA := ce.row, ce.col, ce.arrayMode
	ce.row, ce.col = cell.row, cell.col
	ce.arrayMode = cell.arrayAnchor || cell.IsArrayFormula()
	v, err := ce.evaluate(ws, cell.Formula)
	if err != nil {
		ce.row, ce.col, ce.arrayMode = prevR, prevC, prevA
		return finish(nil, err)
	}
	var out interface{}
	if ce.arrayMode {
		out, err = ce.pickArrayCell(v, cell)
	} else if rg, ok := v.(*formulaRange); ok {
		if rg.virtual {
			if rg.cellCount() > 1 {
				if spillErr := ce.writeSpill(ws, cell, rg); spillErr != nil {
					err = spillErr
				}
			}
			if err == nil {
				if len(rg.values) == 0 {
					err = fmt.Errorf("#N/A")
				} else if err = excelError(rg.values[0]); err == nil {
					out = rg.values[0]
				}
			}
		} else {
			out, err = ce.asScalar(rg)
		}
	} else {
		out = v
	}
	ce.row, ce.col, ce.arrayMode = prevR, prevC, prevA
	if err != nil {
		return finish(nil, err)
	}
	if f, ok := out.(excelFault); ok {
		return finish(nil, fmt.Errorf("%s", string(f)))
	}
	if out == nil {
		return finish(0.0, nil)
	}
	return finish(out, nil)
}

func (ce *CalculationEngine) pickArrayCell(v interface{}, cell *Cell) (interface{}, error) {
	rg, ok := v.(*formulaRange)
	if !ok {
		return v, nil
	}
	if rg.rows == 1 && rg.cols == 1 {
		item := rg.values[0]
		if err := excelError(item); err != nil {
			return nil, err
		}
		return item, nil
	}
	offR, offC := 0, 0
	if ref := cell.ArrayRef(); ref != "" {
		part := stripAbs(ref)
		if i := strings.IndexByte(part, ':'); i >= 0 {
			part = part[:i]
		}
		if start, err := ParseCellReference(part); err == nil {
			offR = cell.row - (start.Row - 1)
			offC = cell.col - start.ColumnIdx
		}
	}
	if offR < 0 || offC < 0 || offR >= rg.rows || offC >= rg.cols {
		return nil, fmt.Errorf("#N/A")
	}
	item := rg.at(offR, offC)
	if err := excelError(item); err != nil {
		return nil, err
	}
	return item, nil
}

func (ce *CalculationEngine) asScalar(v interface{}) (interface{}, error) {
	rg, ok := v.(*formulaRange)
	if !ok {
		return v, nil
	}
	if len(rg.areas) > 0 {
		var hit interface{}
		n := 0
		for _, a := range rg.areas {
			if ce.row >= 0 && !areaContains(a, ce.row, ce.col) {
				continue
			}
			s, err := ce.asScalar(a)
			if err != nil {
				continue
			}
			hit = s
			n++
		}
		if n == 1 {
			return hit, nil
		}
		return nil, fmt.Errorf("#VALUE!")
	}
	if rg.rows <= 1 && rg.cols <= 1 {
		if len(rg.values) == 0 {
			return nil, nil
		}
		return rg.values[0], nil
	}
	if rg.virtual || ce.row < 0 {
		return nil, fmt.Errorf("#VALUE!")
	}
	if rg.rows == 1 {
		c := ce.col - rg.col0
		if c < 0 || c >= rg.cols {
			return nil, fmt.Errorf("#VALUE!")
		}
		return rg.at(0, c), nil
	}
	if rg.cols == 1 {
		r := ce.row - rg.row0
		if r < 0 || r >= rg.rows {
			return nil, fmt.Errorf("#VALUE!")
		}
		return rg.at(r, 0), nil
	}
	r := ce.row - rg.row0
	c := ce.col - rg.col0
	if r < 0 || c < 0 || r >= rg.rows || c >= rg.cols {
		return nil, fmt.Errorf("#VALUE!")
	}
	return rg.at(r, c), nil
}

func areaContains(rg *formulaRange, row, col int) bool {
	if rg == nil {
		return false
	}
	return row >= rg.row0 && row < rg.row0+rg.rows && col >= rg.col0 && col < rg.col0+rg.cols
}

func (ce *CalculationEngine) textArg(v interface{}) (string, error) {
	sv, err := ce.asScalar(v)
	if err != nil {
		return "", err
	}
	if err := excelError(sv); err != nil {
		return "", err
	}
	switch t := sv.(type) {
	case nil:
		return "", nil
	case string:
		return t, nil
	case bool:
		return excelText(t), nil
	default:
		if _, ok := asNumber(t, false); ok {
			return "", fmt.Errorf("#VALUE!")
		}
		return excelText(t), nil
	}
}

func (ce *CalculationEngine) mustNum(v interface{}) (float64, error) {
	sv, err := ce.asScalar(v)
	if err != nil {
		return 0, err
	}
	if err := excelError(sv); err != nil {
		return 0, err
	}
	return arithNumber(sv)
}

func flattenAreas(rg *formulaRange) []interface{} {
	if rg == nil {
		return nil
	}
	if len(rg.areas) == 0 {
		return rg.values
	}
	var out []interface{}
	for _, a := range rg.areas {
		out = append(out, flattenAreas(a)...)
	}
	return out
}

func (p *formulaParser) bin(left, right interface{}, op string) (interface{}, error) {
	if err := excelError(left); err != nil {
		return nil, err
	}
	if err := excelError(right); err != nil {
		return nil, err
	}
	if p.ce.arrayMode && (isMultiRange(left) || isMultiRange(right)) {
		return broadcast(left, right, func(a, b interface{}) (interface{}, error) {
			return p.scalarBin(a, b, op)
		})
	}
	ls, err := p.ce.asScalar(left)
	if err != nil {
		return nil, err
	}
	rs, err := p.ce.asScalar(right)
	if err != nil {
		return nil, err
	}
	return p.scalarBin(ls, rs, op)
}

func (p *formulaParser) scalarBin(a, b interface{}, op string) (interface{}, error) {
	if err := excelError(a); err != nil {
		return nil, err
	}
	if err := excelError(b); err != nil {
		return nil, err
	}
	if op == "&" {
		return excelText(a) + excelText(b), nil
	}
	switch op {
	case "=", "<>", ">", "<", ">=", "<=":
		return compareValues(a, b, op), nil
	}
	ln, err := arithNumber(a)
	if err != nil {
		return nil, err
	}
	rn, err := arithNumber(b)
	if err != nil {
		return nil, err
	}
	if op == "^" {
		return powExcel(ln, rn)
	}
	return applyOp(ln, rn, op)
}

func isMultiRange(v interface{}) bool {
	rg, ok := v.(*formulaRange)
	return ok && rg.cellCount() > 1
}

func broadcast(left, right interface{}, fn func(a, b interface{}) (interface{}, error)) (interface{}, error) {
	lg, _ := left.(*formulaRange)
	rg, _ := right.(*formulaRange)
	rows, cols := 1, 1
	if lg != nil && lg.cellCount() > 1 {
		rows, cols = lg.rows, lg.cols
	}
	if rg != nil && rg.cellCount() > 1 {
		if rows == 1 && cols == 1 {
			rows, cols = rg.rows, rg.cols
		} else if rows != rg.rows || cols != rg.cols {
			return nil, fmt.Errorf("#N/A")
		}
	}
	out := make([]interface{}, rows*cols)
	for r := 0; r < rows; r++ {
		for c := 0; c < cols; c++ {
			a, b := left, right
			if lg != nil && lg.cellCount() > 1 {
				a = lg.at(r, c)
			} else if lg != nil {
				a = lg.at(0, 0)
			}
			if rg != nil && rg.cellCount() > 1 {
				b = rg.at(r, c)
			} else if rg != nil {
				b = rg.at(0, 0)
			}
			v, err := fn(a, b)
			if err != nil {
				return nil, err
			}
			out[r*cols+c] = v
		}
	}
	return &formulaRange{values: out, rows: rows, cols: cols, virtual: true}, nil
}

func intersectRanges(a, b interface{}) (interface{}, error) {
	ag, aok := a.(*formulaRange)
	bg, bok := b.(*formulaRange)
	if !aok || !bok || len(ag.areas) > 0 || len(bg.areas) > 0 {
		return nil, fmt.Errorf("#NULL!")
	}
	if ag.sheet != nil && bg.sheet != nil && ag.sheet != bg.sheet {
		return nil, fmt.Errorf("#NULL!")
	}
	r1 := ag.row0
	if bg.row0 > r1 {
		r1 = bg.row0
	}
	c1 := ag.col0
	if bg.col0 > c1 {
		c1 = bg.col0
	}
	r2 := ag.row0 + ag.rows
	if bg.row0+bg.rows < r2 {
		r2 = bg.row0 + bg.rows
	}
	c2 := ag.col0 + ag.cols
	if bg.col0+bg.cols < c2 {
		c2 = bg.col0 + bg.cols
	}
	if r1 >= r2 || c1 >= c2 {
		return nil, fmt.Errorf("#NULL!")
	}
	rows, cols := r2-r1, c2-c1
	vals := make([]interface{}, 0, rows*cols)
	for r := r1; r < r2; r++ {
		for c := c1; c < c2; c++ {
			vals = append(vals, ag.at(r-ag.row0, c-ag.col0))
		}
	}
	sheet := ag.sheet
	if sheet == nil {
		sheet = bg.sheet
	}
	return &formulaRange{values: vals, rows: rows, cols: cols, row0: r1, col0: c1, sheet: sheet}, nil
}

func unionRanges(a, b interface{}) (interface{}, error) {
	ag := asArea(a)
	bg := asArea(b)
	if ag == nil || bg == nil {
		return nil, fmt.Errorf("#VALUE!")
	}
	areas := append(append([]*formulaRange{}, expandAreas(ag)...), expandAreas(bg)...)
	var vals []interface{}
	for _, area := range areas {
		vals = append(vals, flattenAreas(area)...)
	}
	return &formulaRange{values: vals, rows: len(vals), cols: 1, virtual: true, areas: areas}, nil
}

func asArea(v interface{}) *formulaRange {
	if rg, ok := v.(*formulaRange); ok {
		return rg
	}
	return &formulaRange{values: []interface{}{v}, rows: 1, cols: 1, virtual: true}
}

func expandAreas(rg *formulaRange) []*formulaRange {
	if len(rg.areas) == 0 {
		return []*formulaRange{rg}
	}
	return rg.areas
}

func (p *formulaParser) parseRefOps(v interface{}) (interface{}, error) {
	for {
		j := p.i
		for j < len(p.s) && (p.s[j] == ' ' || p.s[j] == '\t') {
			j++
		}
		if j >= len(p.s) {
			return v, nil
		}
		if !p.inArray && p.s[j] == ',' && j+1 < len(p.s) && isRefStart(p.s[j+1]) {
			p.i = j + 1
			right, err := p.parseAtom()
			if err != nil {
				return nil, err
			}
			v, err = unionRanges(v, right)
			if err != nil {
				return nil, err
			}
			continue
		}
		if j > p.i && isRefStart(p.s[j]) {
			p.i = j
			right, err := p.parseAtom()
			if err != nil {
				return nil, err
			}
			v, err = intersectRanges(v, right)
			if err != nil {
				return nil, err
			}
			continue
		}
		return v, nil
	}
}

func isRefStart(c byte) bool {
	return isIdentStart(c) || c == '$' || c == '\'' || c == '@'
}

func (p *formulaParser) parseArrayConst() (interface{}, error) {
	p.i++ // {
	prev := p.inArray
	p.inArray = true
	defer func() { p.inArray = prev }()
	var rows [][]interface{}
	var row []interface{}
	for {
		p.skip()
		if p.i >= len(p.s) {
			return nil, fmt.Errorf("#VALUE!")
		}
		if p.s[p.i] == '}' {
			p.i++
			rows = append(rows, row)
			break
		}
		v, err := p.parseCompare()
		if err != nil {
			return nil, err
		}
		sv, err := p.ce.asScalar(v)
		if err != nil {
			return nil, err
		}
		row = append(row, sv)
		p.skip()
		if p.i >= len(p.s) {
			return nil, fmt.Errorf("#VALUE!")
		}
		switch p.s[p.i] {
		case ',':
			p.i++
		case ';':
			p.i++
			rows = append(rows, row)
			row = nil
		case '}':
			// loop continues
		default:
			return nil, fmt.Errorf("#VALUE!")
		}
	}
	if len(rows) == 0 {
		return nil, fmt.Errorf("#VALUE!")
	}
	cols := len(rows[0])
	for _, r := range rows {
		if len(r) != cols {
			return nil, fmt.Errorf("#VALUE!")
		}
	}
	vals := make([]interface{}, 0, len(rows)*cols)
	for _, r := range rows {
		vals = append(vals, r...)
	}
	return &formulaRange{values: vals, rows: len(rows), cols: cols, virtual: true}, nil
}
