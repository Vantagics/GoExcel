package gospreadsheet

import (
	"fmt"
	"math"
	"strings"
	"time"
)

func extraFunctions() map[string]formulaFunc {
	return map[string]formulaFunc{
		"SUMIFS": fnSumIfs, "SUMPRODUCT": fnSumProduct, "COUNTIFS": fnCountIfs,
		"AVERAGEIF": fnAverageIf, "AVERAGEIFS": fnAverageIfs,
		"ROUNDUP": fnRoundUp, "ROUNDDOWN": fnRoundDown, "CEILING": fnCeiling, "FLOOR": fnFloor,
		"SIGN": fnSign, "PRODUCT": fnProduct,
		"AND": fnAnd, "OR": fnOr, "NOT": fnNot, "IFERROR": fnIfError, "IFNA": fnIfNA,
		"IFS": fnIfs, "SWITCH": fnSwitch, "XOR": fnXor,
		"VLOOKUP": fnVLookup, "HLOOKUP": fnHLookup, "INDEX": fnIndex, "MATCH": fnMatch,
		"XLOOKUP": fnXLookup, "XMATCH": fnXMatch, "CHOOSE": fnChoose, "OFFSET": fnOffset, "INDIRECT": fnIndirect,
		"TEXTJOIN": fnTextJoin, "FIND": fnFind, "SEARCH": fnSearch, "SUBSTITUTE": fnSubstitute,
		"REPLACE": fnReplace, "TEXT": fnText, "VALUE": fnValue,
		"DATE": fnDate, "TODAY": fnToday, "NOW": fnNow, "YEAR": fnYear, "MONTH": fnMonth, "DAY": fnDay,
		"EDATE": fnEDate, "EOMONTH": fnEOMonth, "DATEDIF": fnDateDif, "WEEKDAY": fnWeekday,
		"NETWORKDAYS": fnNetWorkDays, "WORKDAY": fnWorkday,
		"ISBLANK": fnIsBlank, "ISNUMBER": fnIsNumber, "ISTEXT": fnIsText, "ISERROR": fnIsError, "ISNA": fnIsNA,
	}
}

func fnSumIfs(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 3 || len(parts)%2 == 0 {
		return nil, fmt.Errorf("#VALUE!")
	}
	sumVals, err := ce.resolveRange(ws, parts[0])
	if err != nil {
		return nil, err
	}
	masks := make([][]bool, 0)
	for i := 1; i+1 < len(parts); i += 2 {
		critVals, err := ce.resolveRange(ws, parts[i])
		if err != nil {
			return nil, err
		}
		crit, err := criteriaText(ce, ws, parts[i+1])
		if err != nil {
			return nil, err
		}
		ok := make([]bool, len(critVals))
		for j, v := range critVals {
			ok[j] = matchesCriteria(v, crit)
		}
		masks = append(masks, ok)
	}
	sum := 0.0
	for i, v := range sumVals {
		keep := true
		for _, m := range masks {
			if i >= len(m) || !m[i] {
				keep = false
				break
			}
		}
		if keep {
			if n, ok := asNumber(v, false); ok {
				sum += n
			}
		}
	}
	return sum, nil
}

func fnCountIfs(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 2 || len(parts)%2 != 0 {
		return nil, fmt.Errorf("#VALUE!")
	}
	base, err := ce.resolveRange(ws, parts[0])
	if err != nil {
		return nil, err
	}
	count := 0.0
	for i := range base {
		keep := true
		for j := 0; j+1 < len(parts); j += 2 {
			vals, err := ce.resolveRange(ws, parts[j])
			if err != nil {
				return nil, err
			}
			crit, err := criteriaText(ce, ws, parts[j+1])
			if err != nil {
				return nil, err
			}
			if i >= len(vals) || !matchesCriteria(vals[i], crit) {
				keep = false
				break
			}
		}
		if keep {
			count++
		}
	}
	return count, nil
}

func fnAverageIf(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 2 || len(parts) > 3 {
		return nil, fmt.Errorf("#VALUE!")
	}
	sumRange := parts[0]
	if len(parts) == 3 {
		sumRange = parts[2]
	}
	sum, err := fnSumIf(ce, ws, parts[0]+","+parts[1]+","+sumRange)
	if err != nil {
		return nil, err
	}
	n, err := fnCountIf(ce, ws, parts[0]+","+parts[1])
	if err != nil {
		return nil, err
	}
	if toFloat(n) == 0 {
		return nil, fmt.Errorf("#DIV/0!")
	}
	return toFloat(sum) / toFloat(n), nil
}

func fnAverageIfs(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 3 {
		return nil, fmt.Errorf("#VALUE!")
	}
	sum, err := fnSumIfs(ce, ws, args)
	if err != nil {
		return nil, err
	}
	n, err := fnCountIfs(ce, ws, strings.Join(parts[1:], ","))
	if err != nil {
		return nil, err
	}
	if toFloat(n) == 0 {
		return nil, fmt.Errorf("#DIV/0!")
	}
	return toFloat(sum) / toFloat(n), nil
}

func fnSumProduct(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) == 0 {
		return 0.0, nil
	}
	cols := make([][]float64, len(parts))
	n := -1
	scalar := false
	for i, part := range parts {
		rg, err := ce.evalRange(ws, part)
		if err != nil {
			return nil, err
		}
		if err := rangeError(rg); err != nil {
			return nil, err
		}
		if rg.cellCount() > 1 {
			if n < 0 {
				n = len(rg.values)
			} else if len(rg.values) != n {
				return nil, fmt.Errorf("#VALUE!")
			}
		} else {
			scalar = true
		}
		nums := make([]float64, len(rg.values))
		for j, v := range rg.values {
			if err := excelError(v); err != nil {
				return nil, err
			}
			if x, ok := asNumber(v, false); ok {
				nums[j] = x
			} else if v != nil && rg.cellCount() == 1 {
				return nil, fmt.Errorf("#VALUE!")
			}
		}
		cols[i] = nums
		if n < 0 {
			n = len(nums)
		}
	}
	if scalar && n > 1 {
		return nil, fmt.Errorf("#VALUE!")
	}
	sum := 0.0
	for i := 0; i < n; i++ {
		prod := 1.0
		for _, col := range cols {
			prod *= col[i]
		}
		sum += prod
	}
	return sum, nil
}

func rangeError(rg *formulaRange) error {
	if rg == nil {
		return nil
	}
	for _, v := range flattenAreas(rg) {
		if err := excelError(v); err != nil {
			return err
		}
	}
	return nil
}

func fnRoundUp(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return roundDir(ce, ws, args, true)
}

func fnRoundDown(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return roundDir(ce, ws, args, false)
}

func roundDir(ce *CalculationEngine, ws *Worksheet, args string, up bool) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) != 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	val, err := ce.evaluate(ws, parts[0])
	if err != nil {
		return nil, err
	}
	dig, err := ce.evaluate(ws, parts[1])
	if err != nil {
		return nil, err
	}
	d, err := ce.mustNum(dig)
	if err != nil {
		return nil, err
	}
	n, err := ce.mustNum(val)
	if err != nil {
		return nil, err
	}
	pow := math.Pow(10, d)
	x := n * pow
	if up {
		if x < 0 {
			return math.Floor(x) / pow, nil
		}
		return math.Ceil(x) / pow, nil
	}
	if x < 0 {
		return math.Ceil(x) / pow, nil
	}
	return math.Floor(x) / pow, nil
}

func fnCeiling(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return ceilingFloor(ce, ws, args, true)
}

func fnFloor(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return ceilingFloor(ce, ws, args, false)
}

func ceilingFloor(ce *CalculationEngine, ws *Worksheet, args string, ceil bool) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 1 || len(parts) > 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	val, err := ce.evaluate(ws, parts[0])
	if err != nil {
		return nil, err
	}
	sig := 1.0
	if len(parts) == 2 {
		s, err := ce.evaluate(ws, parts[1])
		if err != nil {
			return nil, err
		}
		sig, err = ce.mustNum(s)
		if err != nil {
			return nil, err
		}
	}
	n, err := ce.mustNum(val)
	if err != nil {
		return nil, err
	}
	if sig == 0 {
		if n == 0 {
			return 0.0, nil
		}
		return nil, fmt.Errorf("#DIV/0!")
	}
	x := n / sig
	if ceil {
		return math.Ceil(x) * sig, nil
	}
	return math.Floor(x) * sig, nil
}

func fnSign(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	v, err := ce.evaluate(ws, args)
	if err != nil {
		return nil, err
	}
	n, err := ce.mustNum(v)
	if err != nil {
		return nil, err
	}
	if n > 0 {
		return 1.0, nil
	}
	if n < 0 {
		return -1.0, nil
	}
	return 0.0, nil
}

func fnProduct(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	values, err := collectNumericValues(ce, ws, args)
	if err != nil {
		return nil, err
	}
	if len(values) == 0 {
		return 0.0, nil
	}
	p := 1.0
	for _, v := range values {
		p *= v
	}
	return p, nil
}

func fnAnd(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return foldLogic(ce, ws, args, true)
}

func fnOr(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return foldLogic(ce, ws, args, false)
}

func foldLogic(ce *CalculationEngine, ws *Worksheet, args string, and bool) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) == 0 {
		return nil, fmt.Errorf("#VALUE!")
	}
	acc := and
	seen := false
	for _, part := range parts {
		v, err := ce.evaluate(ws, part)
		if err != nil {
			return nil, err
		}
		scalar := true
		items := []interface{}{v}
		if rg, ok := v.(*formulaRange); ok {
			items = rg.values
			scalar = false
		}
		for _, item := range items {
			if item == nil {
				continue
			}
			if _, isNum := asNumber(item, false); !isNum {
				if _, isBool := item.(bool); !isBool {
					if scalar {
						if _, isStr := item.(string); isStr {
							return nil, fmt.Errorf("#VALUE!")
						}
					}
					continue
				}
			}
			ok, err := isTruthy(item)
			if err != nil {
				return nil, err
			}
			seen = true
			if and {
				acc = acc && ok
			} else {
				acc = acc || ok
			}
		}
	}
	if !seen {
		return nil, fmt.Errorf("#VALUE!")
	}
	return acc, nil
}

func fnNot(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	v, err := ce.evaluate(ws, args)
	if err != nil {
		return nil, err
	}
	ok, err := isTruthy(v)
	if err != nil {
		return nil, err
	}
	return !ok, nil
}

func fnIfError(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) != 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	v, err := ce.evaluate(ws, parts[0])
	if err != nil || excelError(v) != nil {
		return ce.evaluate(ws, parts[1])
	}
	return v, nil
}

func fnIfNA(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) != 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	v, err := ce.evaluate(ws, parts[0])
	if err != nil && strings.Contains(strings.ToUpper(err.Error()), "#N/A") {
		return ce.evaluate(ws, parts[1])
	}
	if err != nil {
		return nil, err
	}
	if s, ok := v.(string); ok && strings.EqualFold(s, "#N/A") {
		return ce.evaluate(ws, parts[1])
	}
	return v, nil
}

func fnIfs(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 2 || len(parts)%2 != 0 {
		return nil, fmt.Errorf("#VALUE!")
	}
	for i := 0; i+1 < len(parts); i += 2 {
		cond, err := ce.evaluate(ws, parts[i])
		if err != nil {
			return nil, err
		}
		ok, err := isTruthy(cond)
		if err != nil {
			return nil, err
		}
		if ok {
			return ce.evaluate(ws, parts[i+1])
		}
	}
	return nil, fmt.Errorf("#N/A")
}

func fnSwitch(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 3 {
		return nil, fmt.Errorf("#VALUE!")
	}
	expr, err := ce.evaluate(ws, parts[0])
	if err != nil {
		return nil, err
	}
	i := 1
	for i+1 < len(parts) {
		val, err := ce.evaluate(ws, parts[i])
		if err != nil {
			return nil, err
		}
		if compareValues(expr, val, "=") {
			return ce.evaluate(ws, parts[i+1])
		}
		i += 2
	}
	if i < len(parts) {
		return ce.evaluate(ws, parts[i])
	}
	return nil, fmt.Errorf("#N/A")
}

func fnXor(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	n := 0
	for _, part := range parts {
		v, err := ce.evaluate(ws, part)
		if err != nil {
			return nil, err
		}
		ok, err := isTruthy(v)
		if err != nil {
			return nil, err
		}
		if ok {
			n++
		}
	}
	return n%2 == 1, nil
}

func fnVLookup(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return lookupHV(ce, ws, args, true)
}

func fnHLookup(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return lookupHV(ce, ws, args, false)
}

func lookupHV(ce *CalculationEngine, ws *Worksheet, args string, vertical bool) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 3 || len(parts) > 4 {
		return nil, fmt.Errorf("#VALUE!")
	}
	key, err := ce.evaluate(ws, parts[0])
	if err != nil {
		return nil, err
	}
	key, err = ce.asScalar(key)
	if err != nil {
		return nil, err
	}
	rg, err := ce.evalRange(ws, parts[1])
	if err != nil {
		return nil, err
	}
	idxf, err := ce.evaluate(ws, parts[2])
	if err != nil {
		return nil, err
	}
	idxN, err := ce.mustNum(idxf)
	if err != nil {
		return nil, err
	}
	idx := int(idxN)
	approx := false
	if len(parts) == 4 {
		m, err := ce.evaluate(ws, parts[3])
		if err != nil {
			return nil, err
		}
		ok, _ := isTruthy(m)
		approx = ok
	}
	if vertical {
		if idx < 1 || idx > rg.cols {
			return nil, fmt.Errorf("#REF!")
		}
		found := -1
		for r := 0; r < rg.rows; r++ {
			cell := rg.values[r*rg.cols]
			if approx {
				if compareValues(cell, key, "<=") {
					found = r
				} else {
					break
				}
			} else if compareValues(cell, key, "=") {
				found = r
				break
			}
		}
		if found < 0 {
			return nil, fmt.Errorf("#N/A")
		}
		return rg.values[found*rg.cols+idx-1], nil
	}
	if idx < 1 || idx > rg.rows {
		return nil, fmt.Errorf("#REF!")
	}
	found := -1
	for c := 0; c < rg.cols; c++ {
		cell := rg.values[c]
		if approx {
			if compareValues(cell, key, "<=") {
				found = c
			} else {
				break
			}
		} else if compareValues(cell, key, "=") {
			found = c
			break
		}
	}
	if found < 0 {
		return nil, fmt.Errorf("#N/A")
	}
	return rg.values[(idx-1)*rg.cols+found], nil
}

func fnIndex(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 2 || len(parts) > 3 {
		return nil, fmt.Errorf("#VALUE!")
	}
	rg, err := ce.evalRange(ws, parts[0])
	if err != nil {
		return nil, err
	}
	rowV, err := ce.evaluate(ws, parts[1])
	if err != nil {
		return nil, err
	}
	rowN, err := ce.mustNum(rowV)
	if err != nil {
		return nil, err
	}
	row := int(rowN)
	col := 1
	if len(parts) == 3 {
		colV, err := ce.evaluate(ws, parts[2])
		if err != nil {
			return nil, err
		}
		colN, err := ce.mustNum(colV)
		if err != nil {
			return nil, err
		}
		col = int(colN)
	}
	if row == 0 || (len(parts) == 3 && col == 0) {
		r1, r2 := 0, rg.rows-1
		c1, c2 := 0, rg.cols-1
		if row > 0 {
			if row > rg.rows {
				return nil, fmt.Errorf("#REF!")
			}
			r1, r2 = row-1, row-1
		}
		if col > 0 {
			if col > rg.cols {
				return nil, fmt.Errorf("#REF!")
			}
			c1, c2 = col-1, col-1
		}
		vals := make([]interface{}, 0, (r2-r1+1)*(c2-c1+1))
		for r := r1; r <= r2; r++ {
			for c := c1; c <= c2; c++ {
				vals = append(vals, rg.at(r, c))
			}
		}
		return &formulaRange{values: vals, rows: r2 - r1 + 1, cols: c2 - c1 + 1, row0: rg.row0 + r1, col0: rg.col0 + c1, sheet: rg.sheet}, nil
	}
	if row < 1 || col < 1 || row > rg.rows || col > rg.cols {
		return nil, fmt.Errorf("#REF!")
	}
	return rg.values[(row-1)*rg.cols+col-1], nil
}

func fnMatch(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 2 || len(parts) > 3 {
		return nil, fmt.Errorf("#VALUE!")
	}
	key, err := ce.evaluate(ws, parts[0])
	if err != nil {
		return nil, err
	}
	rg, err := ce.evalRange(ws, parts[1])
	if err != nil {
		return nil, err
	}
	typ := 1.0
	if len(parts) == 3 {
		t, err := ce.evaluate(ws, parts[2])
		if err != nil {
			return nil, err
		}
		typ = toFloat(t)
	}
	if typ == 0 {
		for i, v := range rg.values {
			if compareValues(v, key, "=") {
				return float64(i + 1), nil
			}
		}
		return nil, fmt.Errorf("#N/A")
	}
	found := -1
	for i, v := range rg.values {
		if typ > 0 {
			if compareValues(v, key, "<=") {
				found = i
			} else {
				break
			}
		} else if compareValues(v, key, ">=") {
			found = i
		} else {
			break
		}
	}
	if found < 0 {
		return nil, fmt.Errorf("#N/A")
	}
	return float64(found + 1), nil
}

func fnXLookup(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 3 || len(parts) > 5 {
		return nil, fmt.Errorf("#VALUE!")
	}
	key, err := ce.evaluate(ws, parts[0])
	if err != nil {
		return nil, err
	}
	lookup, err := ce.evalRange(ws, parts[1])
	if err != nil {
		return nil, err
	}
	ret, err := ce.evalRange(ws, parts[2])
	if err != nil {
		return nil, err
	}
	for i, v := range lookup.values {
		if compareValues(v, key, "=") {
			if i < len(ret.values) {
				return ret.values[i], nil
			}
			return nil, fmt.Errorf("#REF!")
		}
	}
	if len(parts) >= 4 && strings.TrimSpace(parts[3]) != "" {
		return ce.evaluate(ws, parts[3])
	}
	return nil, fmt.Errorf("#N/A")
}

func fnXMatch(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	return fnMatch(ce, ws, parts[0]+","+parts[1]+",0")
}

func fnChoose(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	idxV, err := ce.evaluate(ws, parts[0])
	if err != nil {
		return nil, err
	}
	idx := int(toFloat(idxV))
	if idx < 1 || idx >= len(parts) {
		return nil, fmt.Errorf("#VALUE!")
	}
	return ce.evaluate(ws, parts[idx])
}

func fnOffset(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 3 || len(parts) > 5 {
		return nil, fmt.Errorf("#VALUE!")
	}
	sheet, ref, err := ce.resolveRef(ws, stripAbs(parts[0]))
	if err != nil {
		return nil, err
	}
	base, err := ParseCellReference(ref)
	if err != nil {
		return nil, fmt.Errorf("#REF!")
	}
	rowV, err := ce.evaluate(ws, parts[1])
	if err != nil {
		return nil, err
	}
	colV, err := ce.evaluate(ws, parts[2])
	if err != nil {
		return nil, err
	}
	rowN, err := ce.mustNum(rowV)
	if err != nil {
		return nil, err
	}
	colN, err := ce.mustNum(colV)
	if err != nil {
		return nil, err
	}
	h, w := 1, 1
	if len(parts) >= 4 && strings.TrimSpace(parts[3]) != "" {
		hv, err := ce.evaluate(ws, parts[3])
		if err != nil {
			return nil, err
		}
		hn, err := ce.mustNum(hv)
		if err != nil {
			return nil, err
		}
		h = int(hn)
	}
	if len(parts) == 5 && strings.TrimSpace(parts[4]) != "" {
		wv, err := ce.evaluate(ws, parts[4])
		if err != nil {
			return nil, err
		}
		wn, err := ce.mustNum(wv)
		if err != nil {
			return nil, err
		}
		w = int(wn)
	}
	r := base.Row + int(rowN)
	c := base.ColumnIdx + int(colN)
	if r < 1 || c < 0 || h < 1 || w < 1 {
		return nil, fmt.Errorf("#REF!")
	}
	a1, err := CellName(r-1, c)
	if err != nil {
		return nil, fmt.Errorf("#REF!")
	}
	a2, err := CellName(r-1+h-1, c+w-1)
	if err != nil {
		return nil, fmt.Errorf("#REF!")
	}
	if h == 1 && w == 1 {
		return ce.resolveCellValue(sheet, a1)
	}
	return ce.readArea(sheet, a1+":"+a2)
}

func fnIndirect(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 1 {
		return nil, fmt.Errorf("#VALUE!")
	}
	v, err := ce.evaluate(ws, parts[0])
	if err != nil {
		return nil, err
	}
	ref := excelText(v)
	return ce.evaluate(ws, ref)
}

func fnTextJoin(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 3 {
		return nil, fmt.Errorf("#VALUE!")
	}
	delimV, err := ce.evaluate(ws, parts[0])
	if err != nil {
		return nil, err
	}
	ignV, err := ce.evaluate(ws, parts[1])
	if err != nil {
		return nil, err
	}
	ignore, _ := isTruthy(ignV)
	delim := excelText(delimV)
	var bits []string
	for _, part := range parts[2:] {
		v, err := ce.evaluate(ws, part)
		if err != nil {
			return nil, err
		}
		items := []interface{}{v}
		if rg, ok := v.(*formulaRange); ok {
			items = rg.values
		}
		for _, item := range items {
			if item == nil && ignore {
				continue
			}
			text := excelText(item)
			if text == "" && ignore {
				continue
			}
			bits = append(bits, text)
		}
	}
	return strings.Join(bits, delim), nil
}

func fnFind(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return findSearch(ce, ws, args, true)
}

func fnSearch(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return findSearch(ce, ws, args, false)
}

func findSearch(ce *CalculationEngine, ws *Worksheet, args string, sensitive bool) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 2 || len(parts) > 3 {
		return nil, fmt.Errorf("#VALUE!")
	}
	findV, err := ce.evaluate(ws, parts[0])
	if err != nil {
		return nil, err
	}
	withinV, err := ce.evaluate(ws, parts[1])
	if err != nil {
		return nil, err
	}
	start := 1
	if len(parts) == 3 {
		s, err := ce.evaluate(ws, parts[2])
		if err != nil {
			return nil, err
		}
		start = int(toFloat(s))
	}
	needle := excelText(findV)
	hay := []rune(excelText(withinV))
	if start < 1 {
		start = 1
	}
	if start > len(hay)+1 {
		return nil, fmt.Errorf("#VALUE!")
	}
	n := []rune(needle)
	h := hay[start-1:]
	if !sensitive {
		n = []rune(strings.ToLower(string(n)))
		h = []rune(strings.ToLower(string(h)))
	}
	idx := strings.Index(string(h), string(n))
	if idx < 0 {
		return nil, fmt.Errorf("#VALUE!")
	}
	return float64(start + len([]rune(string(h)[:idx]))), nil
}

func fnSubstitute(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 3 || len(parts) > 4 {
		return nil, fmt.Errorf("#VALUE!")
	}
	textV, err := ce.evaluate(ws, parts[0])
	if err != nil {
		return nil, err
	}
	oldV, err := ce.evaluate(ws, parts[1])
	if err != nil {
		return nil, err
	}
	newV, err := ce.evaluate(ws, parts[2])
	if err != nil {
		return nil, err
	}
	text, old, repl := excelText(textV), excelText(oldV), excelText(newV)
	if len(parts) == 3 {
		return strings.ReplaceAll(text, old, repl), nil
	}
	instV, err := ce.evaluate(ws, parts[3])
	if err != nil {
		return nil, err
	}
	inst := int(toFloat(instV))
	if inst < 1 || old == "" {
		return nil, fmt.Errorf("#VALUE!")
	}
	seen := 0
	var b strings.Builder
	for {
		i := strings.Index(text, old)
		if i < 0 {
			b.WriteString(text)
			break
		}
		seen++
		b.WriteString(text[:i])
		if seen == inst {
			b.WriteString(repl)
		} else {
			b.WriteString(old)
		}
		text = text[i+len(old):]
	}
	return b.String(), nil
}

func fnReplace(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) != 4 {
		return nil, fmt.Errorf("#VALUE!")
	}
	textV, err := ce.evaluate(ws, parts[0])
	if err != nil {
		return nil, err
	}
	startV, err := ce.evaluate(ws, parts[1])
	if err != nil {
		return nil, err
	}
	lenV, err := ce.evaluate(ws, parts[2])
	if err != nil {
		return nil, err
	}
	newV, err := ce.evaluate(ws, parts[3])
	if err != nil {
		return nil, err
	}
	text, err := ce.textArg(textV)
	if err != nil {
		return nil, err
	}
	startN, err := ce.mustNum(startV)
	if err != nil {
		return nil, err
	}
	nlen, err := ce.mustNum(lenV)
	if err != nil {
		return nil, err
	}
	if startN < 1 || nlen < 0 {
		return nil, fmt.Errorf("#VALUE!")
	}
	runes := []rune(text)
	start := int(startN) - 1
	n := int(nlen)
	if start > len(runes) {
		start = len(runes)
	}
	end := start + n
	if end > len(runes) {
		end = len(runes)
	}
	return string(runes[:start]) + excelText(newV) + string(runes[end:]), nil
}

func fnText(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) != 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	val, err := ce.evaluate(ws, parts[0])
	if err != nil {
		return nil, err
	}
	fmtV, err := ce.evaluate(ws, parts[1])
	if err != nil {
		return nil, err
	}
	sv, err := ce.asScalar(val)
	if err != nil {
		return nil, err
	}
	if err := excelError(sv); err != nil {
		return nil, err
	}
	code, err := ce.textArg(fmtV)
	if err != nil {
		return nil, err
	}
	return formatByExcelCode(sv, code), nil
}

func fnValue(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	v, err := ce.evaluate(ws, args)
	if err != nil {
		return nil, err
	}
	if n, ok := asNumber(v, true); ok {
		return n, nil
	}
	return nil, fmt.Errorf("#VALUE!")
}

func fnDate(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) != 3 {
		return nil, fmt.Errorf("#VALUE!")
	}
	y, err := ce.evaluate(ws, parts[0])
	if err != nil {
		return nil, err
	}
	m, err := ce.evaluate(ws, parts[1])
	if err != nil {
		return nil, err
	}
	d, err := ce.evaluate(ws, parts[2])
	if err != nil {
		return nil, err
	}
	yi := int(toFloat(y))
	mi := int(toFloat(m))
	di := int(toFloat(d))
	if yi < 0 || yi > 9999 {
		return nil, fmt.Errorf("#NUM!")
	}
	return excelDateSerial(yi, mi, di), nil
}

func fnToday(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	now := time.Now().UTC()
	t := time.Date(now.Year(), now.Month(), now.Day(), 0, 0, 0, 0, time.UTC)
	return dateToSerial(t), nil
}

func fnNow(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return dateToSerial(time.Now().UTC()), nil
}

func dateParts(ce *CalculationEngine, ws *Worksheet, expr string) (year, month, day int, err error) {
	v, err := ce.evaluate(ws, expr)
	if err != nil {
		return 0, 0, 0, err
	}
	sv, err := ce.asScalar(v)
	if err != nil {
		return 0, 0, 0, err
	}
	if n, ok := asNumber(sv, true); ok {
		if n < 0 {
			return 0, 0, 0, fmt.Errorf("#NUM!")
		}
		if n == 0 {
			return 1900, 1, 0, nil
		}
		t := excelSerialToTime(n)
		return t.Year(), int(t.Month()), t.Day(), nil
	}
	if t, ok := sv.(time.Time); ok {
		return t.Year(), int(t.Month()), t.Day(), nil
	}
	return 0, 0, 0, fmt.Errorf("#VALUE!")
}

func fnYear(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	y, _, _, err := dateParts(ce, ws, args)
	if err != nil {
		return nil, err
	}
	return float64(y), nil
}

func fnMonth(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	_, m, _, err := dateParts(ce, ws, args)
	if err != nil {
		return nil, err
	}
	return float64(m), nil
}

func fnDay(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	_, _, d, err := dateParts(ce, ws, args)
	if err != nil {
		return nil, err
	}
	return float64(d), nil
}

func fnEDate(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) != 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	t, err := evalDate(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	m, err := ce.evaluate(ws, parts[1])
	if err != nil {
		return nil, err
	}
	return dateToSerial(t.AddDate(0, int(toFloat(m)), 0)), nil
}

func fnEOMonth(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) != 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	t, err := evalDate(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	m, err := ce.evaluate(ws, parts[1])
	if err != nil {
		return nil, err
	}
	shifted := t.AddDate(0, int(toFloat(m))+1, 0)
	end := time.Date(shifted.Year(), shifted.Month(), 1, 0, 0, 0, 0, time.UTC).AddDate(0, 0, -1)
	return dateToSerial(end), nil
}

func fnDateDif(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) != 3 {
		return nil, fmt.Errorf("#VALUE!")
	}
	start, err := evalDate(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	end, err := evalDate(ce, ws, parts[1])
	if err != nil {
		return nil, err
	}
	unitV, err := ce.evaluate(ws, parts[2])
	if err != nil {
		return nil, err
	}
	if end.Before(start) {
		return nil, fmt.Errorf("#NUM!")
	}
	switch strings.ToUpper(excelText(unitV)) {
	case "Y":
		y := end.Year() - start.Year()
		if end.Month() < start.Month() || (end.Month() == start.Month() && end.Day() < start.Day()) {
			y--
		}
		return float64(y), nil
	case "M":
		m := (end.Year()-start.Year())*12 + int(end.Month()) - int(start.Month())
		if end.Day() < start.Day() {
			m--
		}
		return float64(m), nil
	case "D":
		return dateDiffDays(start, end), nil
	case "MD":
		d := end.Day() - start.Day()
		if d < 0 {
			prev := time.Date(end.Year(), end.Month(), 0, 0, 0, 0, 0, time.UTC)
			d += prev.Day()
		}
		return float64(d), nil
	case "YM":
		m := int(end.Month()) - int(start.Month())
		if end.Day() < start.Day() {
			m--
		}
		if m < 0 {
			m += 12
		}
		return float64(m), nil
	case "YD":
		s := time.Date(end.Year(), start.Month(), start.Day(), 0, 0, 0, 0, time.UTC)
		if s.After(end) {
			s = time.Date(end.Year()-1, start.Month(), start.Day(), 0, 0, 0, 0, time.UTC)
		}
		return dateDiffDays(s, end), nil
	default:
		return nil, fmt.Errorf("#NUM!")
	}
}

func dateDiffDays(start, end time.Time) float64 {
	return math.Floor(end.Sub(start).Hours()/24 + 1e-9)
}

func fnWeekday(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 1 {
		return nil, fmt.Errorf("#VALUE!")
	}
	t, err := evalDate(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	// Excel default return_type 1: Sunday = 1.
	wd := int(t.Weekday()) + 1
	return float64(wd), nil
}

func fnNetWorkDays(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	start, err := evalDate(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	end, err := evalDate(ce, ws, parts[1])
	if err != nil {
		return nil, err
	}
	weekend := [7]bool{time.Sunday: true, time.Saturday: true}
	var holidays []time.Time
	if len(parts) >= 3 {
		holidays, err = holidayList(ce, ws, parts[2])
		if err != nil {
			return nil, err
		}
	}
	return float64(networkDayCount(start, end, weekend, holidays)), nil
}

func fnWorkday(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	start, err := evalDate(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	daysV, err := ce.evaluate(ws, parts[1])
	if err != nil {
		return nil, err
	}
	daysN, err := ce.mustNum(daysV)
	if err != nil {
		return nil, err
	}
	weekend := [7]bool{time.Sunday: true, time.Saturday: true}
	var holidays []time.Time
	if len(parts) >= 3 {
		holidays, err = holidayList(ce, ws, parts[2])
		if err != nil {
			return nil, err
		}
	}
	return dateToSerial(workdayShift(start, int(daysN), weekend, holidays)), nil
}

func fnIsBlank(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	v, err := ce.evaluate(ws, args)
	if err != nil {
		return false, nil
	}
	sv, err := ce.asScalar(v)
	if err != nil {
		return false, nil
	}
	return sv == nil, nil
}

func fnIsNumber(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	v, err := ce.evaluate(ws, args)
	if err != nil {
		return false, nil
	}
	sv, err := ce.asScalar(v)
	if err != nil {
		return false, nil
	}
	_, ok := asNumber(sv, false)
	return ok, nil
}

func fnIsText(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	v, err := ce.evaluate(ws, args)
	if err != nil {
		return false, nil
	}
	sv, err := ce.asScalar(v)
	if err != nil {
		return false, nil
	}
	_, ok := sv.(string)
	if ok && excelError(sv) != nil {
		return false, nil
	}
	return ok, nil
}

func fnIsError(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	v, err := ce.evaluate(ws, args)
	if err != nil {
		return true, nil
	}
	sv, err := ce.asScalar(v)
	if err != nil {
		return true, nil
	}
	return excelError(sv) != nil, nil
}

func fnIsNA(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	v, err := ce.evaluate(ws, args)
	if err != nil {
		return strings.Contains(strings.ToUpper(err.Error()), "#N/A"), nil
	}
	sv, err := ce.asScalar(v)
	if err != nil {
		return false, nil
	}
	if f, ok := sv.(excelFault); ok {
		return strings.EqualFold(string(f), "#N/A"), nil
	}
	s, ok := sv.(string)
	return ok && strings.EqualFold(s, "#N/A"), nil
}

func criteriaText(ce *CalculationEngine, ws *Worksheet, expr string) (string, error) {
	expr = strings.TrimSpace(expr)
	if strings.HasPrefix(expr, `"`) && strings.HasSuffix(expr, `"`) && len(expr) >= 2 {
		return expr[1 : len(expr)-1], nil
	}
	v, err := ce.evaluate(ws, expr)
	if err != nil {
		return "", err
	}
	return excelText(v), nil
}

// excelDateSerial is Excel's DATE(y, m, d), including the 1900-02-29 leap bug.
// Go's time.Date skips that day, so a day count that crosses it is one too high.
func excelDateSerial(y, m, d int) float64 {
	t := time.Date(y, time.Month(m), d, 0, 0, 0, 0, time.UTC)
	serial := dateToSerial(t)
	leap := time.Date(1900, 3, 1, 0, 0, 0, 0, time.UTC)
	if t.Before(leap) {
		return serial - 1
	}
	base := time.Date(y, time.Month(m), 1, 0, 0, 0, 0, time.UTC)
	nominal := base.AddDate(0, 0, d-1)
	if base.Before(leap) && !nominal.Before(leap) {
		return serial - 1
	}
	return serial
}

// excelSerialToTime interprets a formula serial the way Excel does.
// dateToSerial's 1899-12-30 epoch is one day high before the fake 1900-02-29.
func excelSerialToTime(serial float64) time.Time {
	t := serialToDate(serial)
	if serial > 0 && serial < 60 {
		t = t.AddDate(0, 0, 1)
	}
	return t
}

func evalDate(ce *CalculationEngine, ws *Worksheet, expr string) (time.Time, error) {
	v, err := ce.evaluate(ws, expr)
	if err != nil {
		return time.Time{}, err
	}
	switch t := v.(type) {
	case time.Time:
		return t, nil
	default:
		if n, ok := asNumber(v, true); ok {
			if n < 0 {
				return time.Time{}, fmt.Errorf("#NUM!")
			}
			return excelSerialToTime(n), nil
		}
		return time.Time{}, fmt.Errorf("#VALUE!")
	}
}

func looksDateCode(code string) bool {
	c := strings.ToLower(code)
	return strings.Contains(c, "yy") || strings.Contains(c, "dd") || strings.Contains(c, "mmm")
}
