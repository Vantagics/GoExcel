package gospreadsheet

import (
	"fmt"
	"math"
	"strconv"
	"strings"
)

// formulaRange is a row-major list of cell values, including blanks (nil).
type formulaRange struct {
	values []interface{}
	rows   int
	cols   int
}

type formulaParser struct {
	ce *CalculationEngine
	ws *Worksheet
	s  string
	i  int
}

func (p *formulaParser) skip() {
	for p.i < len(p.s) && (p.s[p.i] == ' ' || p.s[p.i] == '\t') {
		p.i++
	}
}

func (p *formulaParser) parseCompare() (interface{}, error) {
	left, err := p.parseConcat()
	if err != nil {
		return nil, err
	}
	for {
		p.skip()
		if p.i >= len(p.s) {
			return left, nil
		}
		op := ""
		switch {
		case strings.HasPrefix(p.s[p.i:], "<>"):
			op = "<>"
		case strings.HasPrefix(p.s[p.i:], ">="):
			op = ">="
		case strings.HasPrefix(p.s[p.i:], "<="):
			op = "<="
		case p.i < len(p.s) && (p.s[p.i] == '>' || p.s[p.i] == '<' || p.s[p.i] == '='):
			op = p.s[p.i : p.i+1]
		}
		if op == "" {
			return left, nil
		}
		p.i += len(op)
		right, err := p.parseConcat()
		if err != nil {
			return nil, err
		}
		if err := excelError(left); err != nil {
			return nil, err
		}
		if err := excelError(right); err != nil {
			return nil, err
		}
		ls, err := scalarValue(left)
		if err != nil {
			return nil, err
		}
		rs, err := scalarValue(right)
		if err != nil {
			return nil, err
		}
		left = compareValues(ls, rs, op)
	}
}

func (p *formulaParser) parseConcat() (interface{}, error) {
	left, err := p.parseAdd()
	if err != nil {
		return nil, err
	}
	for {
		p.skip()
		if p.i >= len(p.s) || p.s[p.i] != '&' {
			return left, nil
		}
		p.i++
		right, err := p.parseAdd()
		if err != nil {
			return nil, err
		}
		if err := excelError(left); err != nil {
			return nil, err
		}
		if err := excelError(right); err != nil {
			return nil, err
		}
		ls, err := scalarValue(left)
		if err != nil {
			return nil, err
		}
		rs, err := scalarValue(right)
		if err != nil {
			return nil, err
		}
		left = excelText(ls) + excelText(rs)
	}
}

func (p *formulaParser) parseAdd() (interface{}, error) {
	left, err := p.parseMul()
	if err != nil {
		return nil, err
	}
	for {
		p.skip()
		if p.i >= len(p.s) || (p.s[p.i] != '+' && p.s[p.i] != '-') {
			return left, nil
		}
		op := string(p.s[p.i])
		p.i++
		right, err := p.parseMul()
		if err != nil {
			return nil, err
		}
		ls, err := scalarValue(left)
		if err != nil {
			return nil, err
		}
		rs, err := scalarValue(right)
		if err != nil {
			return nil, err
		}
		ln, err := arithNumber(ls)
		if err != nil {
			return nil, err
		}
		rn, err := arithNumber(rs)
		if err != nil {
			return nil, err
		}
		n, err := applyOp(ln, rn, op)
		if err != nil {
			return nil, err
		}
		left = n
	}
}

func (p *formulaParser) parseMul() (interface{}, error) {
	left, err := p.parsePow()
	if err != nil {
		return nil, err
	}
	for {
		p.skip()
		if p.i >= len(p.s) || (p.s[p.i] != '*' && p.s[p.i] != '/') {
			return left, nil
		}
		op := string(p.s[p.i])
		p.i++
		right, err := p.parsePow()
		if err != nil {
			return nil, err
		}
		ls, err := scalarValue(left)
		if err != nil {
			return nil, err
		}
		rs, err := scalarValue(right)
		if err != nil {
			return nil, err
		}
		ln, err := arithNumber(ls)
		if err != nil {
			return nil, err
		}
		rn, err := arithNumber(rs)
		if err != nil {
			return nil, err
		}
		n, err := applyOp(ln, rn, op)
		if err != nil {
			return nil, err
		}
		left = n
	}
}

func (p *formulaParser) parsePow() (interface{}, error) {
	left, err := p.parseUnary()
	if err != nil {
		return nil, err
	}
	p.skip()
	if p.i < len(p.s) && p.s[p.i] == '^' {
		p.i++
		right, err := p.parsePow()
		if err != nil {
			return nil, err
		}
		ls, err := scalarValue(left)
		if err != nil {
			return nil, err
		}
		rs, err := scalarValue(right)
		if err != nil {
			return nil, err
		}
		ln, err := arithNumber(ls)
		if err != nil {
			return nil, err
		}
		rn, err := arithNumber(rs)
		if err != nil {
			return nil, err
		}
		pow := math.Pow(ln, rn)
		if math.IsNaN(pow) || math.IsInf(pow, 0) {
			return nil, fmt.Errorf("#NUM!")
		}
		return pow, nil
	}
	return left, nil
}

func (p *formulaParser) parseUnary() (interface{}, error) {
	p.skip()
	if p.i < len(p.s) && (p.s[p.i] == '+' || p.s[p.i] == '-') {
		op := p.s[p.i]
		p.i++
		v, err := p.parseUnary()
		if err != nil {
			return nil, err
		}
		sv, err := scalarValue(v)
		if err != nil {
			return nil, err
		}
		n, err := arithNumber(sv)
		if err != nil {
			return nil, err
		}
		if op == '-' {
			return -n, nil
		}
		return n, nil
	}
	v, err := p.parsePrimary()
	if err != nil {
		return nil, err
	}
	p.skip()
	if p.i < len(p.s) && p.s[p.i] == '%' {
		p.i++
		sv, err := scalarValue(v)
		if err != nil {
			return nil, err
		}
		n, err := arithNumber(sv)
		if err != nil {
			return nil, err
		}
		return n / 100, nil
	}
	return v, nil
}

func (p *formulaParser) parsePrimary() (interface{}, error) {
	p.skip()
	if p.i >= len(p.s) {
		return nil, fmt.Errorf("unexpected end of formula")
	}
	if p.s[p.i] == '(' {
		p.i++
		v, err := p.parseCompare()
		if err != nil {
			return nil, err
		}
		p.skip()
		if p.i >= len(p.s) || p.s[p.i] != ')' {
			return nil, fmt.Errorf("#VALUE!")
		}
		p.i++
		return v, nil
	}
	if p.s[p.i] == '"' {
		s, err := p.parseString()
		if err != nil {
			return nil, err
		}
		return s, nil
	}
	if isNumStart(p.s, p.i) {
		return p.parseNumber(), nil
	}
	if p.s[p.i] == '#' {
		return p.parseErrorLit()
	}
	if isIdentStart(p.s[p.i]) || p.s[p.i] == '$' || p.s[p.i] == '\'' {
		return p.parseRefOrCall()
	}
	return nil, fmt.Errorf("cannot evaluate: %s", p.s[p.i:])
}

func (p *formulaParser) parseString() (string, error) {
	p.i++ // opening quote
	var b strings.Builder
	for p.i < len(p.s) {
		if p.s[p.i] == '"' {
			if p.i+1 < len(p.s) && p.s[p.i+1] == '"' {
				b.WriteByte('"')
				p.i += 2
				continue
			}
			p.i++
			return b.String(), nil
		}
		b.WriteByte(p.s[p.i])
		p.i++
	}
	return "", fmt.Errorf("#VALUE!")
}

func (p *formulaParser) parseNumber() float64 {
	start := p.i
	for p.i < len(p.s) && (isDigit(p.s[p.i]) || p.s[p.i] == '.') {
		p.i++
	}
	if p.i < len(p.s) && (p.s[p.i] == 'e' || p.s[p.i] == 'E') {
		p.i++
		if p.i < len(p.s) && (p.s[p.i] == '+' || p.s[p.i] == '-') {
			p.i++
		}
		for p.i < len(p.s) && isDigit(p.s[p.i]) {
			p.i++
		}
	}
	n, _ := strconv.ParseFloat(p.s[start:p.i], 64)
	return n
}

func (p *formulaParser) parseErrorLit() (interface{}, error) {
	rest := p.s[p.i:]
	for _, e := range []string{"#DIV/0!", "#N/A", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#NULL!"} {
		if strings.HasPrefix(strings.ToUpper(rest), e) {
			p.i += len(e)
			return nil, fmt.Errorf("%s", e)
		}
	}
	return nil, fmt.Errorf("#NAME?")
}

func (p *formulaParser) parseRefOrCall() (interface{}, error) {
	start := p.i
	sheet, rest, hasSheet, end := scanSheetRef(p.s, p.i)
	if hasSheet {
		p.i = end
		return p.finishRef(sheet, rest)
	}
	// identifier
	if p.s[p.i] == '$' {
		return p.finishRef(p.ws.Title(), "")
	}
	for p.i < len(p.s) && (isIdent(p.s[p.i]) || p.s[p.i] == '.' || p.s[p.i] == '_') {
		p.i++
	}
	name := p.s[start:p.i]
	p.skip()
	if p.i < len(p.s) && p.s[p.i] == '(' {
		return p.parseCall(name)
	}
	upper := strings.ToUpper(name)
	if upper == "TRUE" {
		return true, nil
	}
	if upper == "FALSE" {
		return false, nil
	}
	p.i = start
	return p.finishRef(p.ws.Title(), "")
}

func (p *formulaParser) parseCall(name string) (interface{}, error) {
	p.i++ // (
	depth := 1
	inStr := false
	argStart := p.i
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
	args := p.s[argStart : p.i-1]
	key := strings.ToUpper(strings.TrimSpace(name))
	key = strings.TrimPrefix(key, "_XLFN.")
	key = strings.TrimPrefix(key, "_XLWS.")
	fn, ok := getBuiltinFunctions()[key]
	if !ok {
		return nil, fmt.Errorf("#NAME?")
	}
	return fn(p.ce, p.ws, args)
}

func (p *formulaParser) finishRef(sheetName, already string) (interface{}, error) {
	sheet, err := p.sheetByName(sheetName)
	if err != nil {
		return nil, err
	}
	body := already
	if body == "" {
		body = p.scanRefBody()
	}
	body = strings.TrimSpace(body)
	if body == "" {
		return nil, fmt.Errorf("#NAME?")
	}
	if strings.Contains(body, ":") {
		return p.ce.readArea(sheet, body)
	}
	if isA1(stripAbs(body)) {
		return p.ce.resolveCellValue(sheet, stripAbs(body))
	}
	if p.ce.workbook != nil {
		if ref, err := p.ce.workbook.GetNamedRange(body); err == nil {
			return p.ce.evaluate(sheet, ref)
		}
	}
	return nil, fmt.Errorf("#NAME?")
}

func (p *formulaParser) scanRefBody() string {
	start := p.i
	for p.i < len(p.s) {
		ch := p.s[p.i]
		if isIdent(ch) || isDigit(ch) || ch == '$' || ch == ':' {
			p.i++
			continue
		}
		break
	}
	return p.s[start:p.i]
}

func (p *formulaParser) sheetByName(name string) (*Worksheet, error) {
	if p.ws != nil && (name == "" || strings.EqualFold(name, p.ws.Title())) {
		return p.ws, nil
	}
	if p.ce.workbook == nil {
		return nil, fmt.Errorf("#REF!")
	}
	ws, err := p.ce.workbook.GetSheetByName(name)
	if err != nil {
		return nil, fmt.Errorf("#REF!")
	}
	return ws, nil
}

func scanSheetRef(s string, i int) (sheet, rest string, ok bool, end int) {
	if i >= len(s) {
		return "", "", false, i
	}
	if s[i] == '\'' {
		j := i + 1
		for j < len(s) {
			if s[j] == '\'' {
				if j+1 < len(s) && s[j+1] == '\'' {
					j += 2
					continue
				}
				break
			}
			j++
		}
		if j >= len(s) || s[j] != '\'' || j+1 >= len(s) || s[j+1] != '!' {
			return "", "", false, i
		}
		return strings.ReplaceAll(s[i+1:j], "''", "'"), "", true, j + 2
	}
	j := i
	for j < len(s) && (isIdent(s[j]) || isDigit(s[j]) || s[j] == ' ') {
		j++
	}
	if j > i && j < len(s) && s[j] == '!' && !isA1(s[i:j]) {
		return s[i:j], "", true, j + 1
	}
	return "", "", false, i
}

func (ce *CalculationEngine) evalRange(ws *Worksheet, expr string) (*formulaRange, error) {
	v, err := ce.evaluate(ws, expr)
	if err != nil {
		return nil, err
	}
	if rg, ok := v.(*formulaRange); ok {
		return rg, nil
	}
	return &formulaRange{values: []interface{}{v}, rows: 1, cols: 1}, nil
}

func (ce *CalculationEngine) readArea(ws *Worksheet, body string) (*formulaRange, error) {
	parts := strings.SplitN(stripAbs(body), ":", 2)
	if len(parts) != 2 {
		return nil, fmt.Errorf("#REF!")
	}
	start, err := ParseCellReference(parts[0])
	if err != nil {
		return nil, fmt.Errorf("#REF!")
	}
	end, err := ParseCellReference(parts[1])
	if err != nil {
		return nil, fmt.Errorf("#REF!")
	}
	r1, r2 := start.Row, end.Row
	if r1 > r2 {
		r1, r2 = r2, r1
	}
	c1, c2 := start.ColumnIdx, end.ColumnIdx
	if c1 > c2 {
		c1, c2 = c2, c1
	}
	rows := r2 - r1 + 1
	cols := c2 - c1 + 1
	if rows < 1 || cols < 1 || rows*cols > 200000 {
		return nil, fmt.Errorf("#VALUE!")
	}
	vals := make([]interface{}, 0, rows*cols)
	for r := r1; r <= r2; r++ {
		for c := c1; c <= c2; c++ {
			ref, _ := CellName(r-1, c)
			v, err := ce.resolveCellValue(ws, ref)
			if err != nil {
				return nil, err
			}
			vals = append(vals, v)
		}
	}
	return &formulaRange{values: vals, rows: rows, cols: cols}, nil
}

func (ce *CalculationEngine) resolveRef(ws *Worksheet, ref string) (*Worksheet, string, error) {
	ref = strings.TrimSpace(ref)
	sheet := ws
	body := ref
	if sh, rest, ok, _ := scanSheetRef(ref, 0); ok {
		var err error
		if ce.workbook != nil {
			sheet, err = ce.workbook.GetSheetByName(sh)
			if err != nil {
				return nil, "", fmt.Errorf("#REF!")
			}
		}
		body = rest
		if body == "" {
			// scanSheetRef for quoted names returns end after ! and rest empty.
			// The caller already consumed the sheet. For resolveRef the input
			// still contains the cell.
			if i := strings.LastIndex(ref, "!"); i >= 0 {
				body = ref[i+1:]
			}
		}
	}
	body = stripAbs(body)
	if !isA1(body) {
		return nil, "", fmt.Errorf("#NAME?")
	}
	return sheet, body, nil
}

func scalar(v interface{}) interface{} {
	sv, err := scalarValue(v)
	if err != nil {
		return v
	}
	return sv
}

func scalarValue(v interface{}) (interface{}, error) {
	if rg, ok := v.(*formulaRange); ok {
		if len(rg.values) == 1 {
			return rg.values[0], nil
		}
		return nil, fmt.Errorf("#VALUE!")
	}
	return v, nil
}

func excelText(v interface{}) string {
	switch t := v.(type) {
	case nil:
		return ""
	case string:
		return t
	case bool:
		if t {
			return "TRUE"
		}
		return "FALSE"
	case float64:
		return formatNumber(t)
	case int:
		return fmt.Sprintf("%d", t)
	default:
		if n, ok := asNumber(v, false); ok {
			return formatNumber(n)
		}
		return fmt.Sprint(v)
	}
}

func isTruthy(v interface{}) (bool, error) {
	if err := excelError(v); err != nil {
		return false, err
	}
	switch t := v.(type) {
	case bool:
		return t, nil
	case float64:
		return t != 0, nil
	case int:
		return t != 0, nil
	case string:
		return false, fmt.Errorf("#VALUE!")
	case nil:
		return false, nil
	default:
		if n, ok := asNumber(v, false); ok {
			return n != 0, nil
		}
		return false, fmt.Errorf("#VALUE!")
	}
}

func isExcelError(s string) bool {
	switch strings.ToUpper(strings.TrimSpace(s)) {
	case "#DIV/0!", "#N/A", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#NULL!":
		return true
	default:
		return false
	}
}

func stripAbs(s string) string {
	return strings.ReplaceAll(s, "$", "")
}

func isA1(s string) bool {
	s = stripAbs(strings.TrimSpace(s))
	if s == "" {
		return false
	}
	i := 0
	for i < len(s) && isLetter(s[i]) {
		i++
	}
	if i == 0 || i > 3 || i == len(s) {
		return false
	}
	for _, c := range s[i:] {
		if c < '0' || c > '9' {
			return false
		}
	}
	return true
}

func isNumStart(s string, i int) bool {
	if i >= len(s) {
		return false
	}
	if s[i] >= '0' && s[i] <= '9' {
		return true
	}
	return s[i] == '.' && i+1 < len(s) && isDigit(s[i+1])
}

func isDigit(c byte) bool { return c >= '0' && c <= '9' }
func isLetter(c byte) bool {
	return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z')
}
func isIdentStart(c byte) bool { return isLetter(c) || c == '_' }
func isIdent(c byte) bool      { return isLetter(c) || isDigit(c) || c == '_' }
