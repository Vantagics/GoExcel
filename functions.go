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
		if err := excelError(val); err != nil {
			return nil, err
		}
		values = append(values, numericList(val)...)
	}
	return values, nil
}

// numericList returns numbers Excel would aggregate. Range text and blanks are
// skipped. A scalar is coerced (TRUE -> 1, numeric text -> number).
func numericList(v interface{}) []float64 {
	if rg, ok := v.(*formulaRange); ok {
		var out []float64
		for _, item := range rg.values {
			if n, ok := asNumber(item, false); ok {
				out = append(out, n)
			}
		}
		return out
	}
	if n, ok := asNumber(v, true); ok {
		return []float64{n}
	}
	return nil
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
	s, ok := v.(string)
	if !ok || !isExcelError(s) {
		return nil
	}
	return fmt.Errorf("%s", s)
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
	values, err := collectNumericValues(ce, ws, args)
	if err != nil {
		return nil, err
	}
	return float64(len(values)), nil
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
	sv, err := scalarValue(val)
	if err != nil {
		return nil, err
	}
	n, err := arithNumber(sv)
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
	d := toFloat(digits)
	pow := math.Pow(10, d)
	return math.Round(toFloat(val)*pow) / pow, nil
}

func fnSqrt(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	val, err := ce.evaluate(ws, args)
	if err != nil {
		return nil, err
	}
	f := toFloat(val)
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
	return math.Pow(toFloat(base), toFloat(exp)), nil
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
	d := toFloat(div)
	if d == 0 {
		return nil, fmt.Errorf("#DIV/0!")
	}
	n := toFloat(num)
	return n - d*math.Floor(n/d), nil
}

func fnInt(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	val, err := ce.evaluate(ws, args)
	if err != nil {
		return nil, err
	}
	return math.Floor(toFloat(val)), nil
}

func fnLen(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	val, err := ce.evaluate(ws, args)
	if err != nil {
		return nil, err
	}
	s := fmt.Sprintf("%v", val)
	if str, ok := val.(string); ok {
		s = str
	}
	return float64(len([]rune(s))), nil
}

func fnUpper(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	val, err := ce.evaluate(ws, args)
	if err != nil {
		return nil, err
	}
	return strings.ToUpper(fmt.Sprintf("%v", val)), nil
}

func fnLower(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	val, err := ce.evaluate(ws, args)
	if err != nil {
		return nil, err
	}
	return strings.ToLower(fmt.Sprintf("%v", val)), nil
}

func fnTrim(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	val, err := ce.evaluate(ws, args)
	if err != nil {
		return nil, err
	}
	return strings.TrimSpace(fmt.Sprintf("%v", val)), nil
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
	s := fmt.Sprintf("%v", val)
	if str, ok := val.(string); ok {
		s = str
	}
	runes := []rune(s)
	count := int(toFloat(n))
	if count > len(runes) {
		count = len(runes)
	}
	if count < 0 {
		count = 0
	}
	return string(runes[:count]), nil
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
	s := fmt.Sprintf("%v", val)
	if str, ok := val.(string); ok {
		s = str
	}
	runes := []rune(s)
	count := int(toFloat(n))
	if count > len(runes) {
		count = len(runes)
	}
	if count < 0 {
		count = 0
	}
	return string(runes[len(runes)-count:]), nil
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
	s := fmt.Sprintf("%v", val)
	if str, ok := val.(string); ok {
		s = str
	}
	runes := []rune(s)
	start := int(toFloat(startVal)) - 1 // MID is 1-based
	length := int(toFloat(lenVal))
	if start < 0 {
		start = 0
	}
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
		if rg, ok := val.(*formulaRange); ok && len(rg.values) != 1 {
			return nil, fmt.Errorf("#VALUE!")
		}
		sv, err := scalarValue(val)
		if err != nil {
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
	criteria := strings.Trim(parts[1], `"`)
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
		if matchesCriteria(cv, criteria) && i < len(sumVals) {
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
	criteria := strings.Trim(parts[1], `"`)

	values, err := ce.resolveRange(ws, rangeStr)
	if err != nil {
		return nil, err
	}

	count := 0.0
	for _, v := range values {
		if matchesCriteria(v, criteria) {
			count++
		}
	}
	return count, nil
}

// matchesCriteria checks if a value matches a criteria string.
// Supports: exact match, ">N", "<N", ">=N", "<=N", "<>N"
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
