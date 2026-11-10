package gospreadsheet

import (
	"fmt"
	"strconv"
	"strings"
	"time"
	"unicode"
)

// CalculationEngine evaluates cell formulas.
type CalculationEngine struct {
	workbook  *Workbook
	visiting  map[string]bool
	row, col  int
	arrayMode bool
	locals    map[string]interface{}
	memo      map[string]calcMemo
}

type calcMemo struct {
	v   interface{}
	err error
}

// NewCalculationEngine creates a new calculation engine for a workbook.
func NewCalculationEngine(wb *Workbook) *CalculationEngine {
	return &CalculationEngine{workbook: wb, row: -1, col: -1}
}

// CalculateCell evaluates the formula in a cell and returns the result.
func (ce *CalculationEngine) CalculateCell(ws *Worksheet, ref string) (interface{}, error) {
	cell, err := ws.GetCellByName(ref)
	if err != nil {
		return nil, err
	}
	if cell.Type != CellTypeFormula {
		return cell.Value, nil
	}
	freshMemo := ce.memo == nil
	if freshMemo {
		ce.memo = map[string]calcMemo{}
		defer func() { ce.memo = nil }()
	}
	v, err := ce.evalFormulaCell(ws, cell)
	if err != nil {
		return nil, err
	}
	if v == nil {
		return 0.0, nil
	}
	return v, nil
}

// CalculateAll evaluates all formula cells in the workbook.
func (ce *CalculationEngine) CalculateAll() error {
	if ce.memo == nil {
		ce.memo = map[string]calcMemo{}
		defer func() { ce.memo = nil }()
	}
	for i := 0; i < ce.workbook.SheetCount(); i++ {
		ws, err := ce.workbook.GetSheet(i)
		if err != nil {
			return fmt.Errorf("accessing sheet %d: %w", i, err)
		}
		for _, cell := range ws.AllCells() {
			if cell.Type == CellTypeFormula {
				val, err := ce.evalFormulaCell(ws, cell)
				if err != nil {
					msg := err.Error()
					if strings.HasPrefix(msg, "#") {
						cell.Value = msg
					} else {
						cell.Value = "#ERROR!"
					}
				} else {
					cell.Value = val
				}
			}
		}
	}
	return nil
}

// evaluate parses and evaluates a formula string.
func (ce *CalculationEngine) evaluate(ws *Worksheet, formula string) (interface{}, error) {
	formula = strings.TrimSpace(formula)
	if strings.HasPrefix(formula, "=") {
		formula = strings.TrimSpace(formula[1:])
	}
	// Array constants and CSE formulas both use braces. parseAtom reads
	// `{...}` itself, so the braces stay part of the expression.
	if formula == "" {
		return nil, fmt.Errorf("empty formula")
	}
	if isExcelError(formula) {
		return nil, fmt.Errorf("%s", formula)
	}
	p := &formulaParser{ce: ce, ws: ws, s: formula}
	v, err := p.parseCompare()
	if err != nil {
		return nil, err
	}
	p.skip()
	if p.i < len(p.s) {
		return nil, fmt.Errorf("cannot evaluate: %s", formula)
	}
	if rg, ok := v.(*formulaRange); ok && len(rg.values) == 1 {
		return rg.values[0], nil
	}
	return v, nil
}

func (ce *CalculationEngine) evalCompare(ws *Worksheet, formula string) (interface{}, bool, error) {
	for _, op := range []string{"<>", ">=", "<=", "=", ">", "<"} {
		idx := findCompare(formula, op)
		if idx <= 0 {
			continue
		}
		left, err := ce.evaluate(ws, formula[:idx])
		if err != nil {
			return nil, true, err
		}
		right, err := ce.evaluate(ws, formula[idx+len(op):])
		if err != nil {
			return nil, true, err
		}
		return compareValues(left, right, op), true, nil
	}
	return nil, false, nil
}

func findCompare(formula, op string) int {
	depth := 0
	inStr := false
	for i := 0; i+len(op) <= len(formula); i++ {
		ch := formula[i]
		if ch == '"' {
			inStr = !inStr
			continue
		}
		if inStr {
			continue
		}
		if ch == '(' {
			depth++
			continue
		}
		if ch == ')' {
			if depth > 0 {
				depth--
			}
			continue
		}
		if depth == 0 && formula[i:i+len(op)] == op {
			if i == 0 {
				continue
			}
			return i
		}
	}
	return -1
}

func compareValues(left, right interface{}, op string) bool {
	if op == "=" && ((left == nil && blankLike(right)) || (right == nil && blankLike(left))) {
		return true
	}
	if op == "<>" && ((left == nil && blankLike(right)) || (right == nil && blankLike(left))) {
		return false
	}
	lk, ln, lt := valueRank(left)
	rk, rn, rt := valueRank(right)
	var cmp int
	switch {
	case lk == 1 && rk == 1:
		if ln < rn {
			cmp = -1
		} else if ln > rn {
			cmp = 1
		}
	case lk == rk && lk == 2:
		cmp = strings.Compare(lt, rt)
	case lk == rk && lk == 3:
		if ln < rn {
			cmp = -1
		} else if ln > rn {
			cmp = 1
		}
	case lk != rk:
		// Excel orders number < text < boolean for inequality.
		// Equality does not coerce across those kinds.
		if op == "=" {
			return false
		}
		if op == "<>" {
			return true
		}
		if lk < rk {
			cmp = -1
		} else {
			cmp = 1
		}
	}
	switch op {
	case "=":
		return cmp == 0
	case "<>":
		return cmp != 0
	case ">":
		return cmp > 0
	case "<":
		return cmp < 0
	case ">=":
		return cmp >= 0
	case "<=":
		return cmp <= 0
	default:
		return false
	}
}

// valueRank is Excel's comparison kind: 1 number (blank is 0), 2 text, 3 boolean.
func valueRank(v interface{}) (kind int, num float64, text string) {
	if v == nil {
		return 1, 0, ""
	}
	switch t := v.(type) {
	case bool:
		if t {
			return 3, 1, ""
		}
		return 3, 0, ""
	case string:
		return 2, 0, strings.ToLower(t)
	case time.Time:
		return 1, dateToSerial(t), ""
	default:
		if n, ok := asNumber(v, false); ok {
			return 1, n, ""
		}
		if s, ok := v.(fmt.Stringer); ok {
			return 2, 0, strings.ToLower(s.String())
		}
		return 2, 0, strings.ToLower(fmt.Sprint(v))
	}
}

func findOperator(formula string, op string) int {
	depth := 0
	// For correct operator precedence, always search right-to-left.
	// The caller (evaluate) tries +/- before */÷, so the rightmost
	// lowest-precedence operator is split first, which is correct.
	for i := len(formula) - 1; i > 0; i-- {
		ch := formula[i]
		if ch == ')' {
			depth++
		} else if ch == '(' {
			depth--
		} else if depth == 0 && string(ch) == op {
			// Avoid matching a negative sign right after another operator
			// e.g. "3*-2" should not split on the '-'
			if (op == "-" || op == "+") && i > 0 {
				prev := formula[i-1]
				if prev == '+' || prev == '-' || prev == '*' || prev == '/' || prev == '(' {
					continue
				}
			}
			return i
		}
	}
	return -1
}

func applyOp(left, right float64, op string) (float64, error) {
	switch op {
	case "+":
		return left + right, nil
	case "-":
		return left - right, nil
	case "*":
		return left * right, nil
	case "/":
		if right == 0 {
			return 0, fmt.Errorf("#DIV/0!")
		}
		return left / right, nil
	}
	return 0, fmt.Errorf("unknown operator: %s", op)
}

func blankLike(v interface{}) bool {
	if v == nil || v == false {
		return true
	}
	if s, ok := v.(string); ok && s == "" {
		return true
	}
	if n, ok := asNumber(v, false); ok && n == 0 {
		return true
	}
	return false
}

func arithNumber(v interface{}) (float64, error) {
	if v == nil {
		return 0, nil
	}
	if n, ok := asNumber(v, false); ok {
		return n, nil
	}
	if b, ok := v.(bool); ok {
		if b {
			return 1, nil
		}
		return 0, nil
	}
	if s, ok := v.(string); ok {
		if f, err := strconv.ParseFloat(strings.TrimSpace(s), 64); err == nil {
			return f, nil
		}
		return 0, fmt.Errorf("#VALUE!")
	}
	return 0, fmt.Errorf("#VALUE!")
}

func requireNumber(v interface{}) (float64, error) {
	sv, err := scalarValue(v)
	if err != nil {
		return 0, err
	}
	return arithNumber(sv)
}

func toFloat(v interface{}) float64 {
	switch val := v.(type) {
	case float64:
		return val
	case int:
		return float64(val)
	case bool:
		if val {
			return 1
		}
		return 0
	case time.Time:
		return dateToSerial(val)
	case string:
		if f, err := strconv.ParseFloat(val, 64); err == nil {
			return f
		}
	}
	return 0
}

func isCellRef(s string) bool {
	s = strings.TrimSpace(s)
	if s == "" {
		return false
	}
	hasLetter := false
	hasDigit := false
	for i, ch := range s {
		if unicode.IsLetter(ch) {
			if hasDigit {
				return false
			}
			hasLetter = true
		} else if unicode.IsDigit(ch) {
			if i == 0 {
				return false
			}
			hasDigit = true
		} else {
			return false
		}
	}
	return hasLetter && hasDigit
}

func (ce *CalculationEngine) resolveCellValue(ws *Worksheet, ref string) (interface{}, error) {
	sheet, cellRef, err := ce.resolveRef(ws, ref)
	if err != nil {
		return nil, err
	}
	key := sheet.Title() + "!" + cellRef
	if ce.visiting == nil {
		ce.visiting = map[string]bool{}
	}
	if ce.visiting[key] {
		return nil, fmt.Errorf("#REF!")
	}
	ce.visiting[key] = true
	defer delete(ce.visiting, key)
	cell := sheet.GetCellIfExists(mustRef(cellRef).Row-1, mustRef(cellRef).ColumnIdx)
	if cell == nil {
		return nil, nil
	}
	if cell.Type == CellTypeFormula {
		return ce.evalFormulaCell(sheet, cell)
	}
	if cell.Type == CellTypeError {
		s := strings.ToUpper(strings.TrimSpace(fmt.Sprint(cell.Value)))
		if !isExcelError(s) {
			s = "#VALUE!"
		}
		return excelFault(s), nil
	}
	if cell.Type == CellTypeEmpty && cell.Value == nil {
		return nil, nil
	}
	return cell.Value, nil
}

func mustRef(ref string) *CellReference {
	cr, err := ParseCellReference(stripAbs(ref))
	if err != nil {
		return &CellReference{}
	}
	return cr
}

// resolveRange resolves a range like "A1:A10" or "Sheet2!A1:B2".
// Every cell is returned, including blanks, so criteria ranges stay aligned.
func (ce *CalculationEngine) resolveRange(ws *Worksheet, rangeStr string) ([]interface{}, error) {
	rg, err := ce.evalRange(ws, rangeStr)
	if err != nil {
		return nil, err
	}
	return rg.values, nil
}

// parseArgs splits function arguments, respecting nested parentheses.
func parseArgs(args string) []string {
	var result []string
	depth := 0
	inStr := false
	var current strings.Builder
	for _, ch := range args {
		if ch == '"' {
			inStr = !inStr
			current.WriteRune(ch)
			continue
		}
		if inStr {
			current.WriteRune(ch)
			continue
		}
		if ch == '(' {
			depth++
			current.WriteRune(ch)
		} else if ch == ')' {
			depth--
			current.WriteRune(ch)
		} else if ch == ',' && depth == 0 {
			result = append(result, strings.TrimSpace(current.String()))
			current.Reset()
		} else {
			current.WriteRune(ch)
		}
	}
	if current.Len() > 0 {
		result = append(result, strings.TrimSpace(current.String()))
	}
	return result
}
