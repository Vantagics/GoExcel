package gospreadsheet

import (
	"fmt"
	"math"
	"math/rand"
	"sort"
	"strings"
)

type lambdaVal struct {
	params []string
	body   string
	ws     *Worksheet
}

func (p *formulaParser) callLambda(lam *lambdaVal) (interface{}, error) {
	p.i++ // (
	depth := 1
	inStr := false
	start := p.i
	for p.i < len(p.s) && depth > 0 {
		ch := p.s[p.i]
		if ch == '"' {
			inStr = !inStr
		} else if !inStr {
			if ch == '(' {
				depth++
			} else if ch == ')' {
				depth--
			}
		}
		p.i++
	}
	if depth != 0 {
		return nil, fmt.Errorf("#VALUE!")
	}
	args := parseArgs(p.s[start : p.i-1])
	if len(args) != len(lam.params) {
		return nil, fmt.Errorf("#VALUE!")
	}
	saved := p.ce.locals
	locals := map[string]interface{}{}
	for k, v := range saved {
		locals[k] = v
	}
	for i, name := range lam.params {
		val, err := p.ce.evaluate(lam.ws, args[i])
		if err != nil {
			return nil, err
		}
		locals[strings.ToUpper(name)] = val
	}
	p.ce.locals = locals
	out, err := p.ce.evaluate(lam.ws, lam.body)
	p.ce.locals = saved
	return out, err
}

func fnLet(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 3 || len(parts)%2 == 0 {
		return nil, fmt.Errorf("#VALUE!")
	}
	saved := ce.locals
	locals := map[string]interface{}{}
	for k, v := range saved {
		locals[k] = v
	}
	ce.locals = locals
	defer func() { ce.locals = saved }()
	for i := 0; i+1 < len(parts); i += 2 {
		name := strings.TrimSpace(parts[i])
		if !isNameToken(name) {
			return nil, fmt.Errorf("#NAME?")
		}
		val, err := ce.evaluate(ws, parts[i+1])
		if err != nil {
			return nil, err
		}
		locals[strings.ToUpper(name)] = val
	}
	return ce.evaluate(ws, parts[len(parts)-1])
}

func isNameToken(s string) bool {
	if s == "" || !isIdentStart(s[0]) {
		return false
	}
	for i := 1; i < len(s); i++ {
		if !isIdent(s[i]) && s[i] != '.' {
			return false
		}
	}
	return true
}

func fnLambda(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	params := make([]string, 0, len(parts)-1)
	for _, p := range parts[:len(parts)-1] {
		name := strings.TrimSpace(p)
		if !isNameToken(name) {
			return nil, fmt.Errorf("#NAME?")
		}
		params = append(params, name)
	}
	return &lambdaVal{params: params, body: parts[len(parts)-1], ws: ws}, nil
}

func fnSingle(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	v, err := ce.evaluate(ws, args)
	if err != nil {
		return nil, err
	}
	return ce.asScalar(v)
}

func fnTranspose(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	rg, err := ce.evalRange(ws, args)
	if err != nil {
		return nil, err
	}
	if err := rangeError(rg); err != nil {
		return nil, err
	}
	vals := make([]interface{}, 0, rg.rows*rg.cols)
	for c := 0; c < rg.cols; c++ {
		for r := 0; r < rg.rows; r++ {
			vals = append(vals, rg.at(r, c))
		}
	}
	return &formulaRange{values: vals, rows: rg.cols, cols: rg.rows, virtual: true, row0: ce.row, col0: ce.col}, nil
}

func fnSequence(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 1 || len(parts) > 4 {
		return nil, fmt.Errorf("#VALUE!")
	}
	rowsN, err := evalMust(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	colsN := 1.0
	start := 1.0
	step := 1.0
	if len(parts) > 1 && strings.TrimSpace(parts[1]) != "" {
		colsN, err = evalMust(ce, ws, parts[1])
		if err != nil {
			return nil, err
		}
	}
	if len(parts) > 2 && strings.TrimSpace(parts[2]) != "" {
		start, err = evalMust(ce, ws, parts[2])
		if err != nil {
			return nil, err
		}
	}
	if len(parts) > 3 && strings.TrimSpace(parts[3]) != "" {
		step, err = evalMust(ce, ws, parts[3])
		if err != nil {
			return nil, err
		}
	}
	rows, cols := int(rowsN), int(colsN)
	if rows < 1 || cols < 1 || rows*cols > 200000 {
		return nil, fmt.Errorf("#VALUE!")
	}
	vals := make([]interface{}, rows*cols)
	n := start
	for i := range vals {
		vals[i] = n
		n += step
	}
	return &formulaRange{values: vals, rows: rows, cols: cols, virtual: true, row0: ce.row, col0: ce.col}, nil
}

func evalMust(ce *CalculationEngine, ws *Worksheet, expr string) (float64, error) {
	v, err := ce.evaluate(ws, expr)
	if err != nil {
		return 0, err
	}
	return ce.mustNum(v)
}

func fnUnique(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 1 {
		return nil, fmt.Errorf("#VALUE!")
	}
	rg, err := ce.evalRange(ws, parts[0])
	if err != nil {
		return nil, err
	}
	byCol := false
	if len(parts) >= 2 {
		v, err := ce.evaluate(ws, parts[1])
		if err != nil {
			return nil, err
		}
		byCol, _ = isTruthy(v)
	}
	exact := false
	if len(parts) >= 3 {
		v, err := ce.evaluate(ws, parts[2])
		if err != nil {
			return nil, err
		}
		exact, _ = isTruthy(v)
	}
	if byCol {
		return nil, fmt.Errorf("#VALUE!")
	}
	seen := map[string]int{}
	var keep []int
	for r := 0; r < rg.rows; r++ {
		key := rowKey(rg, r)
		seen[key]++
		if seen[key] == 1 {
			keep = append(keep, r)
		}
	}
	if exact {
		var once []int
		for _, r := range keep {
			if seen[rowKey(rg, r)] == 1 {
				once = append(once, r)
			}
		}
		keep = once
	}
	vals := make([]interface{}, 0, len(keep)*rg.cols)
	for _, r := range keep {
		for c := 0; c < rg.cols; c++ {
			vals = append(vals, rg.at(r, c))
		}
	}
	if len(keep) == 0 {
		return nil, fmt.Errorf("#N/A")
	}
	return &formulaRange{values: vals, rows: len(keep), cols: rg.cols, virtual: true, row0: ce.row, col0: ce.col}, nil
}

func rowKey(rg *formulaRange, r int) string {
	var b strings.Builder
	for c := 0; c < rg.cols; c++ {
		if c > 0 {
			b.WriteByte('|')
		}
		b.WriteString(fmt.Sprintf("%#v", rg.at(r, c)))
	}
	return b.String()
}

func fnFilter(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	rg, err := ce.evalRange(ws, parts[0])
	if err != nil {
		return nil, err
	}
	inc, err := ce.evalRange(ws, parts[1])
	if err != nil {
		return nil, err
	}
	if inc.rows != rg.rows || (inc.cols != 1 && inc.cols != rg.cols) {
		return nil, fmt.Errorf("#VALUE!")
	}
	var rows []int
	for r := 0; r < rg.rows; r++ {
		ok, err := isTruthy(inc.at(r, 0))
		if err != nil {
			return nil, err
		}
		if ok {
			rows = append(rows, r)
		}
	}
	if len(rows) == 0 {
		if len(parts) >= 3 {
			return ce.evaluate(ws, parts[2])
		}
		return nil, fmt.Errorf("#N/A")
	}
	vals := make([]interface{}, 0, len(rows)*rg.cols)
	for _, r := range rows {
		for c := 0; c < rg.cols; c++ {
			vals = append(vals, rg.at(r, c))
		}
	}
	return &formulaRange{values: vals, rows: len(rows), cols: rg.cols, virtual: true, row0: ce.row, col0: ce.col}, nil
}

func fnSort(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 1 {
		return nil, fmt.Errorf("#VALUE!")
	}
	rg, err := ce.evalRange(ws, parts[0])
	if err != nil {
		return nil, err
	}
	index := 1
	order := 1
	if len(parts) >= 2 && strings.TrimSpace(parts[1]) != "" {
		n, err := evalMust(ce, ws, parts[1])
		if err != nil {
			return nil, err
		}
		index = int(n)
	}
	if len(parts) >= 3 && strings.TrimSpace(parts[2]) != "" {
		n, err := evalMust(ce, ws, parts[2])
		if err != nil {
			return nil, err
		}
		order = int(n)
	}
	if index < 1 || index > rg.cols {
		return nil, fmt.Errorf("#VALUE!")
	}
	return sortRows(rg, index-1, order), nil
}

func fnSortBy(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	rg, err := ce.evalRange(ws, parts[0])
	if err != nil {
		return nil, err
	}
	by, err := ce.evalRange(ws, parts[1])
	if err != nil {
		return nil, err
	}
	order := 1
	if len(parts) >= 3 && strings.TrimSpace(parts[2]) != "" {
		n, err := evalMust(ce, ws, parts[2])
		if err != nil {
			return nil, err
		}
		order = int(n)
	}
	if by.rows != rg.rows {
		return nil, fmt.Errorf("#VALUE!")
	}
	idx := make([]int, rg.rows)
	for i := range idx {
		idx[i] = i
	}
	sort.SliceStable(idx, func(i, j int) bool {
		cmp := rankCompare(by.at(idx[i], 0), by.at(idx[j], 0))
		if order < 0 {
			return cmp > 0
		}
		return cmp < 0
	})
	return takeRows(rg, idx), nil
}

func sortRows(rg *formulaRange, col, order int) *formulaRange {
	idx := make([]int, rg.rows)
	for i := range idx {
		idx[i] = i
	}
	sort.SliceStable(idx, func(i, j int) bool {
		cmp := rankCompare(rg.at(idx[i], col), rg.at(idx[j], col))
		if order < 0 {
			return cmp > 0
		}
		return cmp < 0
	})
	return takeRows(rg, idx)
}

func takeRows(rg *formulaRange, idx []int) *formulaRange {
	vals := make([]interface{}, 0, len(idx)*rg.cols)
	for _, r := range idx {
		for c := 0; c < rg.cols; c++ {
			vals = append(vals, rg.at(r, c))
		}
	}
	return &formulaRange{values: vals, rows: len(idx), cols: rg.cols, virtual: true}
}

func rankCompare(a, b interface{}) int {
	if compareValues(a, b, "<") {
		return -1
	}
	if compareValues(a, b, ">") {
		return 1
	}
	return 0
}

func fnRandArray(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	rows, cols := 1, 1
	var err error
	if len(parts) > 0 && strings.TrimSpace(parts[0]) != "" {
		n, e := evalMust(ce, ws, parts[0])
		if e != nil {
			return nil, e
		}
		rows = int(n)
	}
	if len(parts) > 1 && strings.TrimSpace(parts[1]) != "" {
		n, e := evalMust(ce, ws, parts[1])
		if e != nil {
			return nil, e
		}
		cols = int(n)
	}
	min, max := 0.0, 1.0
	if len(parts) > 2 && strings.TrimSpace(parts[2]) != "" {
		min, err = evalMust(ce, ws, parts[2])
		if err != nil {
			return nil, err
		}
	}
	if len(parts) > 3 && strings.TrimSpace(parts[3]) != "" {
		max, err = evalMust(ce, ws, parts[3])
		if err != nil {
			return nil, err
		}
	}
	whole := false
	if len(parts) > 4 {
		v, err := ce.evaluate(ws, parts[4])
		if err != nil {
			return nil, err
		}
		whole, _ = isTruthy(v)
	}
	if rows < 1 || cols < 1 {
		return nil, fmt.Errorf("#VALUE!")
	}
	vals := make([]interface{}, rows*cols)
	for i := range vals {
		x := min + rand.Float64()*(max-min)
		if whole {
			x = math.Floor(x)
		}
		vals[i] = x
	}
	return &formulaRange{values: vals, rows: rows, cols: cols, virtual: true, row0: ce.row, col0: ce.col}, nil
}
