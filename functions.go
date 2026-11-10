package gospreadsheet

import (
	"fmt"
	"math"
	"sort"
	"strconv"
	"strings"
	"time"
)

type formulaFunc func(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error)

// builtinFunctions is the cached map of builtin functions, initialized lazily.
var builtinFunctions map[string]formulaFunc

// getBuiltinFunctions returns the map of builtin functions.
// Uses lazy initialization to avoid initialization cycle.
func getBuiltinFunctions() map[string]formulaFunc {
	if builtinFunctions == nil {
		builtinFunctions = map[string]formulaFunc{
			"SUM": fnSum, "AVERAGE": fnAverage, "COUNT": fnCount, "COUNTA": fnCountA,
			"MAX": fnMax, "MIN": fnMin, "IF": fnIf, "ABS": fnAbs, "ROUND": fnRound,
			"SQRT": fnSqrt, "POWER": fnPower, "MOD": fnMod, "INT": fnInt,
			"LEN": fnLen, "UPPER": fnUpper, "LOWER": fnLower, "TRIM": fnTrim,
			"LEFT": fnLeft, "RIGHT": fnRight, "MID": fnMid, "CONCATENATE": fnConcatenate,
			"CONCAT": fnConcatenate, "MEDIAN": fnMedian, "SUMIF": fnSumIf, "COUNTIF": fnCountIf,
		}
		for name, fn := range extraFunctions() {
			builtinFunctions[name] = fn
		}
		for name, fn := range moreFunctions() {
			builtinFunctions[name] = fn
		}
	}
	return builtinFunctions
}

// collectNumericValues resolves args (ranges or single values) into float64 slice.
func collectNumericValues(ce *CalculationEngine, ws *Worksheet, args string) ([]float64, error) {
	parts := parseArgs(args)
	var values []float64
	for _, part := range parts {
		val, err := ce.evaluate(ws, part)
		if err != nil {
			return nil, err
		}
		nums, err := numericList(val)
		if err != nil {
			return nil, err
		}
		values = append(values, nums...)
	}
	return values, nil
}

// numericList returns numbers Excel would aggregate. Range text and blanks are
// skipped. A scalar is coerced (TRUE -> 1, numeric text -> number).
func numericList(v interface{}) ([]float64, error) {
	if rg, ok := v.(*formulaRange); ok {
		items := rg.values
		if len(rg.areas) > 0 {
			items = flattenAreas(rg)
		}
		var out []float64
		for _, item := range items {
			if err := excelError(item); err != nil {
				return nil, err
			}
			if n, ok := asNumber(item, false); ok {
				out = append(out, n)
			}
		}
		return out, nil
	}
	if err := excelError(v); err != nil {
		return nil, err
	}
	if n, ok := asNumber(v, true); ok {
		return []float64{n}, nil
	}
	return nil, nil
}

func asNumber(v interface{}, coerce bool) (float64, bool) {
	switch t := v.(type) {
	case float64:
		return t, true
	case int:
		return float64(t), true
	case int64:
		return float64(t), true
	case bool:
		if !coerce {
			return 0, false
		}
		if t {
			return 1, true
		}
		return 0, true
	case time.Time:
		return dateToSerial(t), true
	case string:
		if !coerce {
			return 0, false
		}
		if f, err := strconv.ParseFloat(strings.TrimSpace(t), 64); err == nil {
			return f, true
		}
		return 0, false
	default:
		return 0, false
	}
}

func excelError(v interface{}) error {
	switch t := v.(type) {
	case excelFault:
		return fmt.Errorf("%s", string(t))
	case string:
		if isExcelError(t) {
			return fmt.Errorf("%s", strings.ToUpper(t))
		}
	}
	return nil
}

func fnSum(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	values, err := collectNumericValues(ce, ws, args)
	if err != nil {
		return nil, err
	}
	sum := 0.0
	for _, v := range values {
		sum += v
	}
	return sum, nil
}

func fnAverage(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	values, err := collectNumericValues(ce, ws, args)
	if err != nil {
		return nil, err
	}
	if len(values) == 0 {
		return nil, fmt.Errorf("#DIV/0!")
	}
	sum := 0.0
	for _, v := range values {
		sum += v
	}
	return sum / float64(len(values)), nil
}

func fnCount(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	n := 0
	for _, part := range parts {
		val, err := ce.evaluate(ws, part)
		if err != nil {
			return nil, err
		}
		items := []interface{}{val}
		if rg, ok := val.(*formulaRange); ok {
			items = flattenAreas(rg)
		}
		for _, item := range items {
			if excelError(item) != nil {
				continue
			}
			if _, ok := asNumber(item, false); ok {
				n++
			}
		}
	}
	return float64(n), nil
}

func fnCountA(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	count := 0
	for _, part := range parts {
		val, err := ce.evaluate(ws, part)
		if err != nil {
			return nil, err
		}
		items := []interface{}{val}
		if rg, ok := val.(*formulaRange); ok {
			items = rg.values
		}
		for _, item := range items {
			if item != nil {
				count++
			}
		}
	}
	return float64(count), nil
}

func fnMax(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	values, err := collectNumericValues(ce, ws, args)
	if err != nil {
		return nil, err
	}
	if len(values) == 0 {
		return 0.0, nil
	}
	max := values[0]
	for _, v := range values[1:] {
		if v > max {
			max = v
		}
	}
	return max, nil
}

func fnMin(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	values, err := collectNumericValues(ce, ws, args)
	if err != nil {
		return nil, err
	}
	if len(values) == 0 {
		return 0.0, nil
	}
	min := values[0]
	for _, v := range values[1:] {
		if v < min {
			min = v
		}
	}
	return min, nil
}

func fnIf(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 2 || len(parts) > 3 {
		return nil, fmt.Errorf("IF requires 2 or 3 arguments")
	}
	cond, err := ce.evaluate(ws, parts[0])
	if err != nil {
		return nil, err
	}
	condBool, err := isTruthy(cond)
	if err != nil {
		return nil, err
	}
	if condBool {
		return ce.evaluate(ws, parts[1])
	}
	if len(parts) == 3 {
		return ce.evaluate(ws, parts[2])
	}
	return false, nil
}

func fnAbs(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	val, err := ce.evaluate(ws, args)
	if err != nil {
		return nil, err
	}
	return absNum(ce, val)
}

func absNum(ce *CalculationEngine, val interface{}) (interface{}, error) {
	n, err := ce.mustNum(val)
	if err != nil {
		return nil, err
	}
	return math.Abs(n), nil
}

func fnRound(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) != 2 {
		return nil, fmt.Errorf("ROUND requires 2 arguments")
	}
	val, err := ce.evaluate(ws, parts[0])
	if err != nil {
		return nil, err
	}
	digits, err := ce.evaluate(ws, parts[1])
	if err != nil {
		return nil, err
	}
	n, err := ce.mustNum(val)
	if err != nil {
		return nil, err
	}
	d, err := ce.mustNum(digits)
	if err != nil {
		return nil, err
	}
	pow := math.Pow(10, d)
	return math.Round(n*pow) / pow, nil
}

func fnSqrt(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	val, err := ce.evaluate(ws, args)
	if err != nil {
		return nil, err
	}
	f, err := ce.mustNum(val)
	if err != nil {
		return nil, err
	}
	if f < 0 {
		return nil, fmt.Errorf("#NUM!")
	}
	return math.Sqrt(f), nil
}

func fnPower(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) != 2 {
		return nil, fmt.Errorf("POWER requires 2 arguments")
	}
	base, err := ce.evaluate(ws, parts[0])
	if err != nil {
		return nil, err
	}
	exp, err := ce.evaluate(ws, parts[1])
	if err != nil {
		return nil, err
	}
	b, err := ce.mustNum(base)
	if err != nil {
		return nil, err
	}
	e, err := ce.mustNum(exp)
	if err != nil {
		return nil, err
	}
	return powExcel(b, e)
}

func fnMod(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) != 2 {
		return nil, fmt.Errorf("MOD requires 2 arguments")
	}
	num, err := ce.evaluate(ws, parts[0])
	if err != nil {
		return nil, err
	}
	div, err := ce.evaluate(ws, parts[1])
	if err != nil {
		return nil, err
	}
	n, err := ce.mustNum(num)
	if err != nil {
		return nil, err
	}
	d, err := ce.mustNum(div)
	if err != nil {
		return nil, err
	}
	if d == 0 {
		return nil, fmt.Errorf("#DIV/0!")
	}
	return n - d*math.Floor(n/d), nil
}

func fnInt(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	val, err := ce.evaluate(ws, args)
	if err != nil {
		return nil, err
	}
	n, err := ce.mustNum(val)
	if err != nil {
		return nil, err
	}
	return math.Floor(n), nil
}

func fnLen(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	val, err := ce.evaluate(ws, args)
	if err != nil {
		return nil, err
	}
	s, err := ce.textArg(val)
	if err != nil {
		return nil, err
	}
	return float64(len([]rune(s))), nil
}

func fnUpper(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	val, err := ce.evaluate(ws, args)
	if err != nil {
		return nil, err
	}
	s, err := ce.textArg(val)
	if err != nil {
		return nil, err
	}
	return strings.ToUpper(s), nil
}

func fnLower(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	val, err := ce.evaluate(ws, args)
	if err != nil {
		return nil, err
	}
	s, err := ce.textArg(val)
	if err != nil {
		return nil, err
	}
	return strings.ToLower(s), nil
}

func fnTrim(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	val, err := ce.evaluate(ws, args)
	if err != nil {
		return nil, err
	}
	s, err := ce.textArg(val)
	if err != nil {
		return nil, err
	}
	return strings.TrimSpace(s), nil
}

func fnLeft(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) != 2 {
		return nil, fmt.Errorf("LEFT requires 2 arguments")
	}
	val, err := ce.evaluate(ws, parts[0])
	if err != nil {
		return nil, err
	}
	n, err := ce.evaluate(ws, parts[1])
	if err != nil {
		return nil, err
	}
	s, err := ce.textArg(val)
	if err != nil {
		return nil, err
	}
	count, err := ce.mustNum(n)
	if err != nil {
		return nil, err
	}
	if count < 0 {
		return nil, fmt.Errorf("#VALUE!")
	}
	runes := []rune(s)
	k := int(count)
	if k > len(runes) {
		k = len(runes)
	}
	return string(runes[:k]), nil
}

func fnRight(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) != 2 {
		return nil, fmt.Errorf("RIGHT requires 2 arguments")
	}
	val, err := ce.evaluate(ws, parts[0])
	if err != nil {
		return nil, err
	}
	n, err := ce.evaluate(ws, parts[1])
	if err != nil {
		return nil, err
	}
	s, err := ce.textArg(val)
	if err != nil {
		return nil, err
	}
	count, err := ce.mustNum(n)
	if err != nil {
		return nil, err
	}
	if count < 0 {
		return nil, fmt.Errorf("#VALUE!")
	}
	runes := []rune(s)
	k := int(count)
	if k > len(runes) {
		k = len(runes)
	}
	return string(runes[len(runes)-k:]), nil
}

func fnMid(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) != 3 {
		return nil, fmt.Errorf("MID requires 3 arguments")
	}
	val, err := ce.evaluate(ws, parts[0])
	if err != nil {
		return nil, err
	}
	startVal, err := ce.evaluate(ws, parts[1])
	if err != nil {
		return nil, err
	}
	lenVal, err := ce.evaluate(ws, parts[2])
	if err != nil {
		return nil, err
	}
	s, err := ce.textArg(val)
	if err != nil {
		return nil, err
	}
	startN, err := ce.mustNum(startVal)
	if err != nil {
		return nil, err
	}
	lengthN, err := ce.mustNum(lenVal)
	if err != nil {
		return nil, err
	}
	if startN < 1 || lengthN < 0 {
		return nil, fmt.Errorf("#VALUE!")
	}
	runes := []rune(s)
	start := int(startN) - 1
	length := int(lengthN)
	if start >= len(runes) {
		return "", nil
	}
	end := start + length
	if end > len(runes) {
		end = len(runes)
	}
	return string(runes[start:end]), nil
}

func fnConcatenate(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	result := ""
	for _, part := range parts {
		val, err := ce.evaluate(ws, part)
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
		result += excelText(sv)
	}
	return result, nil
}

func fnMedian(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	values, err := collectNumericValues(ce, ws, args)
	if err != nil {
		return nil, err
	}
	if len(values) == 0 {
		return nil, fmt.Errorf("#NUM!")
	}
	sort.Float64s(values)
	n := len(values)
	if n%2 == 0 {
		return (values[n/2-1] + values[n/2]) / 2, nil
	}
	return values[n/2], nil
}

func fnSumIf(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 2 || len(parts) > 3 {
		return nil, fmt.Errorf("SUMIF requires 2 or 3 arguments")
	}

	criteriaRange := parts[0]
	criteria, err := criterionOf(ce, ws, parts[1])
	if err != nil {
		return nil, err
	}
	sumRange := criteriaRange
	if len(parts) == 3 {
		sumRange = parts[2]
	}

	criteriaVals, err := ce.resolveRange(ws, criteriaRange)
	if err != nil {
		return nil, err
	}
	sumVals, err := ce.resolveRange(ws, sumRange)
	if err != nil {
		return nil, err
	}

	sum := 0.0
	for i, cv := range criteriaVals {
		if matchesCriterion(cv, criteria) && i < len(sumVals) {
			sum += toFloat(sumVals[i])
		}
	}
	return sum, nil
}

func fnCountIf(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) != 2 {
		return nil, fmt.Errorf("COUNTIF requires 2 arguments")
	}

	rangeStr := parts[0]
	criteria, err := criterionOf(ce, ws, parts[1])
	if err != nil {
		return nil, err
	}

	values, err := ce.resolveRange(ws, rangeStr)
	if err != nil {
		return nil, err
	}

	count := 0.0
	for _, v := range values {
		if matchesCriterion(v, criteria) {
			count++
		}
	}
	return count, nil
}

// matchesCriteria checks if a value matches a criteria string.
// Supports: exact match, ">N", "<N", ">=N", "<=N", "<>N"
func criterionOf(ce *CalculationEngine, ws *Worksheet, expr string) (interface{}, error) {
	v, err := ce.evaluate(ws, expr)
	if err != nil {
		msg := strings.ToUpper(strings.TrimSpace(err.Error()))
		if isExcelError(msg) {
			return excelFault(msg), nil
		}
		return nil, err
	}
	return ce.asScalar(v)
}

func matchesCriterion(value, crit interface{}) bool {
	if f, ok := crit.(excelFault); ok {
		g, ok := value.(excelFault)
		return ok && strings.EqualFold(string(f), string(g))
	}
	if s, ok := crit.(string); ok {
		return matchesCriteria(value, s)
	}
	if b, ok := crit.(bool); ok {
		vb, vok := value.(bool)
		return vok && vb == b
	}
	if n, ok := asNumber(crit, false); ok {
		vn, vok := asNumber(value, false)
		return vok && vn == n
	}
	return matchesCriteria(value, excelText(crit))
}

func matchesCriteria(value interface{}, criteria string) bool {
	if strings.HasPrefix(criteria, ">=") {
		threshold, err := strconv.ParseFloat(criteria[2:], 64)
		if err != nil {
			return false
		}
		return toFloat(value) >= threshold
	}
	if strings.HasPrefix(criteria, "<=") {
		threshold, err := strconv.ParseFloat(criteria[2:], 64)
		if err != nil {
			return false
		}
		return toFloat(value) <= threshold
	}
	if strings.HasPrefix(criteria, "<>") {
		return fmt.Sprintf("%v", value) != criteria[2:]
	}
	if strings.HasPrefix(criteria, ">") {
		threshold, err := strconv.ParseFloat(criteria[1:], 64)
		if err != nil {
			return false
		}
		return toFloat(value) > threshold
	}
	if strings.HasPrefix(criteria, "<") {
		threshold, err := strconv.ParseFloat(criteria[1:], 64)
		if err != nil {
			return false
		}
		return toFloat(value) < threshold
	}
	// Exact match (string or numeric)
	if v, err := strconv.ParseFloat(criteria, 64); err == nil {
		return toFloat(value) == v
	}
	return strings.EqualFold(fmt.Sprintf("%v", value), criteria)
}
