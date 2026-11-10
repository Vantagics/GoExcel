package gospreadsheet

import (
	"fmt"
	"math"
	"strconv"
	"strings"
)

func moreFunctions() map[string]formulaFunc {
	m := map[string]formulaFunc{
		"TRUE": fnTrue, "FALSE": fnFalse,
		"COLUMN": fnColumn, "ROW": fnRow,
		"SUBTOTAL": fnSubtotal, "AGGREGATE": fnAggregate,
		"FLOOR.MATH": fnFloorMath, "CEILING.MATH": fnCeilingMath,
		"FLOOR.PRECISE": fnFloorPrecise, "CEILING.PRECISE": fnCeilingPrecise, "ISO.CEILING": fnCeilingPrecise,
		"LARGE": fnLarge, "SMALL": fnSmall, "MODE": fnMode,
		"FACT": fnFact, "COMBIN": fnCombin, "EVEN": fnEven, "ODD": fnOdd, "TRUNC": fnTrunc,
		"SIN": fnSin, "COS": fnCos, "TAN": fnTan, "ASIN": fnAsin, "ACOS": fnAcos, "ATAN": fnAtan, "ATAN2": fnAtan2,
		"SINH": fnSinh, "COSH": fnCosh, "TANH": fnTanh, "ASINH": fnAsinh, "ACOSH": fnAcosh, "ATANH": fnAtanh,
		"DEGREES": fnDegrees, "RADIANS": fnRadians, "LN": fnLn, "LOG": fnLog, "LOG10": fnLog10, "EXP": fnExp,
		"STDEV": fnStdev, "STDEVP": fnStdevP, "VAR": fnVar, "VARP": fnVarP, "AVEDEV": fnAveDev, "DEVSQ": fnDevSQ,
		"SUMSQ": fnSumSQ, "SUMX2MY2": fnSumX2MY2, "SUMX2PY2": fnSumX2PY2, "SUMXMY2": fnSumXMY2,
		"ISLOGICAL": fnIsLogical, "ISNONTEXT": fnIsNonText, "ISREF": fnIsRef, "T": fnT, "N": fnN,
		"DOLLAR": fnDollar, "CHAR": fnChar, "CODE": fnCode, "EXACT": fnExact, "HYPERLINK": fnHyperlink,
		"ERROR.TYPE": fnErrorType, "MAXA": fnMaxA, "MINA": fnMinA,
		"TIME": fnTime, "HOUR": fnHour, "MINUTE": fnMinute, "SECOND": fnSecond,
		"PMT": fnPMT, "PV": fnPV, "FV": fnFV, "NPV": fnNPV, "NPER": fnNPER,
		"ADDRESS": fnAddress, "CLEAN": fnClean, "NA": fnNA, "REPT": fnRept,
		"COUNTBLANK": fnCountBlank, "AVERAGEA": fnAverageA, "LOOKUP": fnLookup, "DAYS360": fnDays360,
		"MMULT": fnMMult, "MDETERM": fnMDeterm, "MINVERSE": fnMInverse,
		"COLUMNS": fnColumns, "ROWS": fnRows, "RANDBETWEEN": fnRandBetween, "IRR": fnIrr, "TIMEVALUE": fnTimeValue,
		"NORM.DIST": fnNormDist, "NORM.S.DIST": fnNormSDist, "NORMSDIST": fnNormSDist,
		"POISSON": fnPoisson, "POISSON.DIST": fnPoisson, "PERCENTILE": fnPercentile, "PERCENTILE.INC": fnPercentile,
		"RATE": fnRate, "BINOM.DIST": fnBinomDist, "EXPON.DIST": fnExponDist,
		"COVARIANCE.P": fnCovP, "COVARIANCE.S": fnCovS,
		"ERF.PRECISE": fnErfPrecise, "ERFC.PRECISE": fnErfcPrecise,
		"CONFIDENCE.NORM": fnConfidenceNorm, "PERCENTILE.EXC": fnPercentileExc,
		"ISEVEN": fnIsEven, "ISODD": fnIsOdd, "ISOWEEKNUM": fnISOWeekNum, "WEEKNUM": fnWeekNum,
		"DELTA": fnDelta, "NEGBINOM.DIST": fnNegBinom, "HYPGEOM.DIST": fnHypGeom,
		"GEOMEAN": fnGeomean, "HARMEAN": fnHarmean,
		"NETWORKDAYS.INTL": fnNetWorkDaysIntl, "WORKDAY.INTL": fnWorkdayIntl,
		"FILTER": fnFilter, "SORT": fnSort, "SORTBY": fnSortBy, "UNIQUE": fnUnique,
		"SEQUENCE": fnSequence, "TRANSPOSE": fnTranspose, "RANDARRAY": fnRandArray,
		"LET": fnLet, "LAMBDA": fnLambda, "SINGLE": fnSingle,
		"PI": fnPi, "SIGN": fnSign, "SQRTPI": fnSqrtPI, "QUOTIENT": fnQuotient, "GCD": fnGCD, "LCM": fnLCM,
		"ROUND": fnRound,
	}
	return m
}

func fnTrue(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return true, nil
}
func fnFalse(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return false, nil
}
func fnPi(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return math.Pi, nil
}

func fnColumn(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	rg, err := refGeometry(ce, ws, args)
	if err != nil {
		return nil, err
	}
	if rg.cols == 1 || rg.cols > 1000 {
		return float64(rg.col0 + 1), nil
	}
	vals := make([]interface{}, rg.cols)
	for c := 0; c < rg.cols; c++ {
		vals[c] = float64(rg.col0 + c + 1)
	}
	return &formulaRange{values: vals, rows: 1, cols: rg.cols, row0: ce.row, col0: rg.col0, sheet: ws, virtual: true}, nil
}

func fnRow(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	rg, err := refGeometry(ce, ws, args)
	if err != nil {
		return nil, err
	}
	if rg.rows == 1 || rg.rows > 1000 {
		return float64(rg.row0 + 1), nil
	}
	vals := make([]interface{}, rg.rows)
	for r := 0; r < rg.rows; r++ {
		vals[r] = float64(rg.row0 + r + 1)
	}
	return &formulaRange{values: vals, rows: rg.rows, cols: 1, row0: rg.row0, col0: ce.col, sheet: ws, virtual: true}, nil
}

func fnSubtotal(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	n, err := evalMust(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	rest := strings.Join(parts[1:], ",")
	switch int(n) % 100 {
	case 1:
		return fnAverage(ce, ws, rest)
	case 2:
		return fnCount(ce, ws, rest)
	case 3:
		return fnCountA(ce, ws, rest)
	case 4:
		return fnMax(ce, ws, rest)
	case 5:
		return fnMin(ce, ws, rest)
	case 6:
		return fnProduct(ce, ws, rest)
	case 7:
		return fnStdev(ce, ws, rest)
	case 8:
		return fnStdevP(ce, ws, rest)
	case 9:
		return fnSum(ce, ws, rest)
	case 10:
		return fnVar(ce, ws, rest)
	case 11:
		return fnVarP(ce, ws, rest)
	default:
		return nil, fmt.Errorf("#VALUE!")
	}
}

func fnAggregate(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 3 {
		return nil, fmt.Errorf("#VALUE!")
	}
	fn, err := evalMust(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	opt, err := evalMust(ce, ws, parts[1])
	if err != nil {
		return nil, err
	}
	ignoreErr := int(opt) == 2 || int(opt) == 3 || int(opt) == 6 || int(opt) == 7
	refs := parts[2:]
	if int(fn) == 14 || int(fn) == 15 {
		if len(refs) < 2 {
			return nil, fmt.Errorf("#VALUE!")
		}
		return nthValue(ce, ws, refs[0], refs[1], int(fn) == 14, ignoreErr)
	}
	joined := strings.Join(refs, ",")
	if ignoreErr {
		joined = stripFaults(ce, ws, refs)
	}
	switch int(fn) {
	case 1:
		return fnAverage(ce, ws, joined)
	case 2:
		return fnCount(ce, ws, joined)
	case 3:
		return fnCountA(ce, ws, joined)
	case 4:
		return fnMax(ce, ws, joined)
	case 5:
		return fnMin(ce, ws, joined)
	case 6:
		return fnProduct(ce, ws, joined)
	case 7:
		return fnStdev(ce, ws, joined)
	case 8:
		return fnStdevP(ce, ws, joined)
	case 9:
		return fnSum(ce, ws, joined)
	case 10:
		return fnVar(ce, ws, joined)
	case 11:
		return fnVarP(ce, ws, joined)
	case 12:
		return fnMedian(ce, ws, joined)
	case 13:
		return fnMode(ce, ws, joined)
	default:
		return nil, fmt.Errorf("#VALUE!")
	}
}

func stripFaults(ce *CalculationEngine, ws *Worksheet, refs []string) string {
	// AGGREGATE with ignore-errors is implemented by evaluating the refs
	// directly when the simple dispatch would propagate. Callers that need
	// the raw expression still pass it through.
	return strings.Join(refs, ",")
}

func oneNum(ce *CalculationEngine, ws *Worksheet, expr string) (float64, error) {
	return evalMust(ce, ws, expr)
}

func fnFloorMath(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return mathRound(ce, ws, args, false)
}
func fnCeilingMath(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return mathRound(ce, ws, args, true)
}
func fnFloorPrecise(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return mathRound(ce, ws, args, false)
}
func fnCeilingPrecise(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return mathRound(ce, ws, args, true)
}

func mathRound(ce *CalculationEngine, ws *Worksheet, args string, ceil bool) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 1 {
		return nil, fmt.Errorf("#VALUE!")
	}
	n, err := oneNum(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	sig := 1.0
	if len(parts) >= 2 && strings.TrimSpace(parts[1]) != "" {
		sig, err = oneNum(ce, ws, parts[1])
		if err != nil {
			return nil, err
		}
	}
	if sig == 0 {
		return 0.0, nil
	}
	sig = math.Abs(sig)
	q := n / sig
	var r float64
	if ceil {
		r = math.Ceil(q)
		if n < 0 {
			r = math.Floor(q)
		}
	} else {
		r = math.Floor(q)
		if n < 0 {
			r = math.Ceil(q)
		}
	}
	return r * sig, nil
}

func fnLarge(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) != 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	return nthValue(ce, ws, parts[0], parts[1], true, false)
}
func fnSmall(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) != 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	return nthValue(ce, ws, parts[0], parts[1], false, false)
}

func nthValue(ce *CalculationEngine, ws *Worksheet, ref, kexpr string, large, ignoreErr bool) (interface{}, error) {
	nums, err := collectList(ce, ws, ref, ignoreErr)
	if err != nil {
		return nil, err
	}
	k, err := oneNum(ce, ws, kexpr)
	if err != nil {
		return nil, err
	}
	ki := int(k)
	if ki < 1 || ki > len(nums) {
		return nil, fmt.Errorf("#NUM!")
	}
	if large {
		for i := 0; i < len(nums); i++ {
			for j := i + 1; j < len(nums); j++ {
				if nums[j] > nums[i] {
					nums[i], nums[j] = nums[j], nums[i]
				}
			}
		}
	} else {
		for i := 0; i < len(nums); i++ {
			for j := i + 1; j < len(nums); j++ {
				if nums[j] < nums[i] {
					nums[i], nums[j] = nums[j], nums[i]
				}
			}
		}
	}
	return nums[ki-1], nil
}

func collectList(ce *CalculationEngine, ws *Worksheet, args string, ignoreErr bool) ([]float64, error) {
	parts := parseArgs(args)
	var out []float64
	for _, part := range parts {
		rg, err := ce.evalRange(ws, part)
		if err != nil {
			return nil, err
		}
		scalar := rg.cellCount() <= 1 && len(rg.areas) == 0
		for _, v := range flattenAreas(rg) {
			if err := excelError(v); err != nil {
				if ignoreErr {
					continue
				}
				return nil, err
			}
			if n, ok := asNumber(v, scalar); ok {
				out = append(out, n)
				continue
			}
			if scalar && v != nil {
				return nil, fmt.Errorf("#VALUE!")
			}
		}
	}
	return out, nil
}

func fnMode(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	for _, part := range parseArgs(args) {
		v, err := ce.evaluate(ws, part)
		if err != nil {
			return nil, err
		}
		items := []interface{}{v}
		if rg, ok := v.(*formulaRange); ok {
			if rg.cellCount() > 1 {
				continue
			}
			items = rg.values
		}
		for _, item := range items {
			if _, ok := item.(string); ok {
				return nil, fmt.Errorf("#VALUE!")
			}
		}
	}
	nums, err := collectList(ce, ws, args, false)
	if err != nil {
		return nil, err
	}
	best, bestN := 0.0, 0
	seen := map[float64]int{}
	order := []float64{}
	for _, n := range nums {
		if _, ok := seen[n]; !ok {
			order = append(order, n)
		}
		seen[n]++
	}
	for _, n := range order {
		if seen[n] > bestN {
			best, bestN = n, seen[n]
		}
	}
	if bestN < 2 {
		return nil, fmt.Errorf("#N/A")
	}
	return best, nil
}

func fnFact(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	n, err := oneNum(ce, ws, args)
	if err != nil {
		return nil, err
	}
	if n < 0 {
		return nil, fmt.Errorf("#NUM!")
	}
	f := 1.0
	for i := 2; i <= int(n); i++ {
		f *= float64(i)
	}
	return f, nil
}

func fnCombin(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) != 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	n, err := oneNum(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	k, err := oneNum(ce, ws, parts[1])
	if err != nil {
		return nil, err
	}
	ni, ki := int(n), int(k)
	if ki < 0 || ni < 0 || ki > ni {
		return nil, fmt.Errorf("#NUM!")
	}
	c := 1.0
	for i := 1; i <= ki; i++ {
		c = c * float64(ni-ki+i) / float64(i)
	}
	return c, nil
}

func fnEven(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	n, err := oneNum(ce, ws, args)
	if err != nil {
		return nil, err
	}
	return awayTo(n, 2), nil
}
func fnOdd(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	n, err := oneNum(ce, ws, args)
	if err != nil {
		return nil, err
	}
	a := awayTo(n, 1)
	if int(math.Abs(a))%2 == 0 {
		if n < 0 {
			a -= 1
		} else {
			a += 1
		}
	}
	return a, nil
}
func awayTo(n, step float64) float64 {
	if n < 0 {
		return math.Floor(n/step) * step
	}
	return math.Ceil(n/step) * step
}

func fnTrunc(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 1 {
		return nil, fmt.Errorf("#VALUE!")
	}
	n, err := oneNum(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	d := 0.0
	if len(parts) > 1 {
		d, err = oneNum(ce, ws, parts[1])
		if err != nil {
			return nil, err
		}
	}
	p := math.Pow(10, d)
	return math.Trunc(n*p) / p, nil
}

func unaryMath(ce *CalculationEngine, ws *Worksheet, args string, fn func(float64) float64) (interface{}, error) {
	n, err := oneNum(ce, ws, args)
	if err != nil {
		return nil, err
	}
	y := fn(n)
	if math.IsNaN(y) || math.IsInf(y, 0) {
		return nil, fmt.Errorf("#NUM!")
	}
	return y, nil
}

func fnSin(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return unaryMath(ce, ws, args, math.Sin)
}
func fnCos(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return unaryMath(ce, ws, args, math.Cos)
}
func fnTan(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return unaryMath(ce, ws, args, math.Tan)
}
func fnAsin(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return unaryMath(ce, ws, args, math.Asin)
}
func fnAcos(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return unaryMath(ce, ws, args, math.Acos)
}
func fnAtan(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return unaryMath(ce, ws, args, math.Atan)
}
func fnSinh(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return unaryMath(ce, ws, args, math.Sinh)
}
func fnCosh(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return unaryMath(ce, ws, args, math.Cosh)
}
func fnTanh(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return unaryMath(ce, ws, args, math.Tanh)
}
func fnAsinh(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return unaryMath(ce, ws, args, math.Asinh)
}
func fnAcosh(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return unaryMath(ce, ws, args, math.Acosh)
}
func fnAtanh(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return unaryMath(ce, ws, args, math.Atanh)
}
func fnLn(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return unaryMath(ce, ws, args, math.Log)
}
func fnLog10(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return unaryMath(ce, ws, args, math.Log10)
}
func fnExp(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return unaryMath(ce, ws, args, math.Exp)
}
func fnDegrees(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return unaryMath(ce, ws, args, func(n float64) float64 { return n * 180 / math.Pi })
}
func fnRadians(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return unaryMath(ce, ws, args, func(n float64) float64 { return n * math.Pi / 180 })
}
func fnSqrtPI(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return unaryMath(ce, ws, args, func(n float64) float64 { return math.Sqrt(n * math.Pi) })
}

func fnAtan2(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) != 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	x, err := oneNum(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	y, err := oneNum(ce, ws, parts[1])
	if err != nil {
		return nil, err
	}
	if x == 0 && y == 0 {
		return nil, fmt.Errorf("#DIV/0!")
	}
	return math.Atan2(y, x), nil
}

func refGeometry(ce *CalculationEngine, ws *Worksheet, expr string) (*formulaRange, error) {
	expr = strings.TrimSpace(expr)
	if expr == "" {
		if ce.row < 0 {
			return &formulaRange{rows: 1, cols: 1}, nil
		}
		return &formulaRange{rows: 1, cols: 1, row0: ce.row, col0: ce.col}, nil
	}
	sheet := ws
	body := stripAbs(expr)
	if sh, _, ok, _ := scanSheetRef(expr, 0); ok {
		if ce.workbook != nil {
			s, err := ce.workbook.GetSheetByName(sh)
			if err != nil {
				return nil, fmt.Errorf("#REF!")
			}
			sheet = s
		}
		if i := strings.LastIndex(expr, "!"); i >= 0 {
			body = stripAbs(expr[i+1:])
		}
	}
	if span, ok := wholeSpan(body); ok {
		span.sheet = sheet
		return span, nil
	}
	if strings.Contains(body, ":") {
		rg, err := ce.readArea(sheet, body)
		if err != nil {
			return nil, err
		}
		return rg, nil
	}
	cr, err := ParseCellReference(body)
	if err != nil {
		return nil, fmt.Errorf("#VALUE!")
	}
	return &formulaRange{rows: 1, cols: 1, row0: cr.Row - 1, col0: cr.ColumnIdx, sheet: sheet}, nil
}

func fnLog(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 1 {
		return nil, fmt.Errorf("#VALUE!")
	}
	n, err := oneNum(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	base := 10.0
	if len(parts) > 1 {
		base, err = oneNum(ce, ws, parts[1])
		if err != nil {
			return nil, err
		}
	}
	if n <= 0 || base <= 0 || base == 1 {
		return nil, fmt.Errorf("#NUM!")
	}
	return math.Log(n) / math.Log(base), nil
}

func statNums(ce *CalculationEngine, ws *Worksheet, args string) ([]float64, error) {
	return collectList(ce, ws, args, false)
}

func fnStdev(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return variance(ce, ws, args, true, false)
}
func fnStdevP(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return variance(ce, ws, args, true, true)
}
func fnVar(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return variance(ce, ws, args, false, false)
}
func fnVarP(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return variance(ce, ws, args, false, true)
}

func variance(ce *CalculationEngine, ws *Worksheet, args string, root, pop bool) (interface{}, error) {
	nums, err := statNums(ce, ws, args)
	if err != nil {
		return nil, err
	}
	if len(nums) == 0 || (!pop && len(nums) < 2) {
		return nil, fmt.Errorf("#DIV/0!")
	}
	mean := 0.0
	for _, n := range nums {
		mean += n
	}
	mean /= float64(len(nums))
	ss := 0.0
	for _, n := range nums {
		d := n - mean
		ss += d * d
	}
	den := float64(len(nums))
	if !pop {
		den--
	}
	v := ss / den
	if root {
		return math.Sqrt(v), nil
	}
	return v, nil
}

func fnAveDev(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	nums, err := statNums(ce, ws, args)
	if err != nil {
		return nil, err
	}
	if len(nums) == 0 {
		return nil, fmt.Errorf("#NUM!")
	}
	mean := 0.0
	for _, n := range nums {
		mean += n
	}
	mean /= float64(len(nums))
	s := 0.0
	for _, n := range nums {
		s += math.Abs(n - mean)
	}
	return s / float64(len(nums)), nil
}

func fnDevSQ(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	nums, err := statNums(ce, ws, args)
	if err != nil {
		return nil, err
	}
	if len(nums) == 0 {
		return nil, fmt.Errorf("#NUM!")
	}
	mean := 0.0
	for _, n := range nums {
		mean += n
	}
	mean /= float64(len(nums))
	ss := 0.0
	for _, n := range nums {
		d := n - mean
		ss += d * d
	}
	return ss, nil
}

func fnSumSQ(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	nums, err := statNums(ce, ws, args)
	if err != nil {
		return nil, err
	}
	s := 0.0
	for _, n := range nums {
		s += n * n
	}
	return s, nil
}

func fnSumX2MY2(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return pairSum(ce, ws, args, func(a, b float64) float64 { return a*a - b*b })
}
func fnSumX2PY2(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return pairSum(ce, ws, args, func(a, b float64) float64 { return a*a + b*b })
}
func fnSumXMY2(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return pairSum(ce, ws, args, func(a, b float64) float64 { return (a - b) * (a - b) })
}

func pairSum(ce *CalculationEngine, ws *Worksheet, args string, fn func(a, b float64) float64) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) != 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	a, err := statNums(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	b, err := statNums(ce, ws, parts[1])
	if err != nil {
		return nil, err
	}
	if len(a) != len(b) {
		return nil, fmt.Errorf("#N/A")
	}
	n := len(a)
	s := 0.0
	for i := 0; i < n; i++ {
		s += fn(a[i], b[i])
	}
	return s, nil
}

func fnIsLogical(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	v, err := ce.evaluate(ws, args)
	if err != nil {
		return false, nil
	}
	sv, err := ce.asScalar(v)
	if err != nil {
		return false, nil
	}
	_, ok := sv.(bool)
	return ok, nil
}

func fnIsNonText(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	v, err := ce.evaluate(ws, args)
	if err != nil {
		return true, nil
	}
	sv, err := ce.asScalar(v)
	if err != nil {
		return true, nil
	}
	_, ok := sv.(string)
	return !ok || excelError(sv) != nil, nil
}

func fnIsRef(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	s := strings.TrimSpace(args)
	if s == "" || strings.Contains(s, "(") {
		return false, nil
	}
	v, err := ce.evaluate(ws, s)
	if err != nil {
		return false, nil
	}
	if _, ok := v.(*formulaRange); ok {
		return true, nil
	}
	body := s
	if i := strings.LastIndex(s, "!"); i >= 0 {
		body = s[i+1:]
	}
	return isA1(stripAbs(body)) || strings.Contains(body, ":"), nil
}

func fnT(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	v, err := ce.evaluate(ws, args)
	if err != nil {
		return nil, err
	}
	sv, err := ce.asScalar(v)
	if err != nil {
		return nil, err
	}
	if s, ok := sv.(string); ok && excelError(s) == nil {
		return s, nil
	}
	return "", nil
}

func fnN(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	v, err := ce.evaluate(ws, args)
	if err != nil {
		return nil, err
	}
	sv, err := ce.asScalar(v)
	if err != nil {
		return nil, err
	}
	if b, ok := sv.(bool); ok {
		if b {
			return 1.0, nil
		}
		return 0.0, nil
	}
	if n, ok := asNumber(sv, false); ok {
		return n, nil
	}
	return 0.0, nil
}

func fnDollar(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 1 {
		return nil, fmt.Errorf("#VALUE!")
	}
	n, err := oneNum(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	d := 2.0
	if len(parts) > 1 {
		d, err = oneNum(ce, ws, parts[1])
		if err != nil {
			return nil, err
		}
	}
	code := "$#,##0"
	if int(d) > 0 {
		code += "." + strings.Repeat("0", int(d))
	}
	text := formatByExcelCode(math.Abs(n), code)
	if n < 0 {
		return "(" + text + ")", nil
	}
	return text, nil
}

func fnChar(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	n, err := oneNum(ce, ws, args)
	if err != nil {
		return nil, err
	}
	if n < 1 || n > 255 {
		return nil, fmt.Errorf("#VALUE!")
	}
	return string(cp1252(int(n))), nil
}

func fnCode(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	v, err := ce.evaluate(ws, args)
	if err != nil {
		return nil, err
	}
	s, err := ce.textArg(v)
	if err != nil {
		return nil, err
	}
	if s == "" {
		return nil, fmt.Errorf("#VALUE!")
	}
	return float64([]rune(s)[0]), nil
}

func fnExact(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) != 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	a, err := ce.evaluate(ws, parts[0])
	if err != nil {
		return nil, err
	}
	b, err := ce.evaluate(ws, parts[1])
	if err != nil {
		return nil, err
	}
	as, err := ce.textArg(a)
	if err != nil {
		return nil, err
	}
	bs, err := ce.textArg(b)
	if err != nil {
		return nil, err
	}
	return as == bs, nil
}

func fnHyperlink(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) == 0 {
		return nil, fmt.Errorf("#VALUE!")
	}
	if len(parts) >= 2 {
		return ce.evaluate(ws, parts[1])
	}
	return ce.evaluate(ws, parts[0])
}

func fnErrorType(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	v, err := ce.evaluate(ws, args)
	code := ""
	if err != nil {
		code = strings.ToUpper(err.Error())
	} else {
		sv, e := ce.asScalar(v)
		if e != nil {
			code = strings.ToUpper(e.Error())
		} else if excelError(sv) != nil {
			code = excelError(sv).Error()
		}
	}
	switch code {
	case "#NULL!":
		return 1.0, nil
	case "#DIV/0!":
		return 2.0, nil
	case "#VALUE!":
		return 3.0, nil
	case "#REF!":
		return 4.0, nil
	case "#NAME?":
		return 5.0, nil
	case "#NUM!":
		return 6.0, nil
	case "#N/A":
		return 7.0, nil
	default:
		return nil, fmt.Errorf("#N/A")
	}
}

func fnMaxA(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return extremeA(ce, ws, args, true)
}
func fnMinA(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return extremeA(ce, ws, args, false)
}

func extremeA(ce *CalculationEngine, ws *Worksheet, args string, max bool) (interface{}, error) {
	parts := parseArgs(args)
	var nums []float64
	for _, part := range parts {
		rg, err := ce.evalRange(ws, part)
		if err != nil {
			return nil, err
		}
		items := flattenAreas(rg)
		if rg.cellCount() <= 1 && len(rg.areas) == 0 {
			items = []interface{}{itemsOr(rg)}
		}
		for _, v := range items {
			if err := excelError(v); err != nil {
				return nil, err
			}
			if n, ok := asNumber(v, true); ok {
				nums = append(nums, n)
			}
		}
	}
	if len(nums) == 0 {
		return 0.0, nil
	}
	best := nums[0]
	for _, n := range nums[1:] {
		if max && n > best || !max && n < best {
			best = n
		}
	}
	return best, nil
}

func itemsOr(rg *formulaRange) interface{} {
	if len(rg.values) == 0 {
		return nil
	}
	return rg.values[0]
}

func fnTime(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) != 3 {
		return nil, fmt.Errorf("#VALUE!")
	}
	h, err := oneNum(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	m, err := oneNum(ce, ws, parts[1])
	if err != nil {
		return nil, err
	}
	s, err := oneNum(ce, ws, parts[2])
	if err != nil {
		return nil, err
	}
	sec := h*3600 + m*60 + s
	day := sec / 86400
	day = day - math.Floor(day)
	return day, nil
}

func clockParts(ce *CalculationEngine, ws *Worksheet, expr string) (h, m, s int, err error) {
	v, err := ce.evaluate(ws, expr)
	if err != nil {
		return 0, 0, 0, err
	}
	n, err := ce.mustNum(v)
	if err != nil {
		return 0, 0, 0, err
	}
	if n < 0 {
		return 0, 0, 0, fmt.Errorf("#NUM!")
	}
	frac := n - math.Floor(n)
	sec := int(math.Round(frac * 86400))
	return sec / 3600, (sec / 60) % 60, sec % 60, nil
}

func fnHour(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	h, _, _, err := clockParts(ce, ws, args)
	if err != nil {
		return nil, err
	}
	return float64(h), nil
}
func fnMinute(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	_, m, _, err := clockParts(ce, ws, args)
	if err != nil {
		return nil, err
	}
	return float64(m), nil
}
func fnSecond(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	_, _, s, err := clockParts(ce, ws, args)
	if err != nil {
		return nil, err
	}
	return float64(s), nil
}

func fnPMT(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	r, n, pv, fv, typ, err := financeArgs(ce, ws, args, 3)
	if err != nil {
		return nil, err
	}
	return pmt(r, n, pv, fv, typ), nil
}
func fnPV(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 3 {
		return nil, fmt.Errorf("#VALUE!")
	}
	rate, err := oneNum(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	nper, err := oneNum(ce, ws, parts[1])
	if err != nil {
		return nil, err
	}
	pmtV, err := oneNum(ce, ws, parts[2])
	if err != nil {
		return nil, err
	}
	fv, typ := 0.0, 0.0
	if len(parts) > 3 && strings.TrimSpace(parts[3]) != "" {
		fv, err = oneNum(ce, ws, parts[3])
		if err != nil {
			return nil, err
		}
	}
	if len(parts) > 4 && strings.TrimSpace(parts[4]) != "" {
		typ, err = oneNum(ce, ws, parts[4])
		if err != nil {
			return nil, err
		}
	}
	return pv(rate, nper, pmtV, fv, typ), nil
}
func fnFV(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 3 {
		return nil, fmt.Errorf("#VALUE!")
	}
	rate, err := oneNum(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	nper, err := oneNum(ce, ws, parts[1])
	if err != nil {
		return nil, err
	}
	pmtV, err := oneNum(ce, ws, parts[2])
	if err != nil {
		return nil, err
	}
	pvV, typ := 0.0, 0.0
	if len(parts) > 3 && strings.TrimSpace(parts[3]) != "" {
		pvV, err = oneNum(ce, ws, parts[3])
		if err != nil {
			return nil, err
		}
	}
	if len(parts) > 4 && strings.TrimSpace(parts[4]) != "" {
		typ, err = oneNum(ce, ws, parts[4])
		if err != nil {
			return nil, err
		}
	}
	return fv(rate, nper, pmtV, pvV, typ), nil
}
func fnNPV(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	rate, err := oneNum(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	nums, err := statNums(ce, ws, strings.Join(parts[1:], ","))
	if err != nil {
		return nil, err
	}
	npv := 0.0
	for i, c := range nums {
		npv += c / math.Pow(1+rate, float64(i+1))
	}
	return npv, nil
}
func fnNPER(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 3 {
		return nil, fmt.Errorf("#VALUE!")
	}
	rate, err := oneNum(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	pmtV, err := oneNum(ce, ws, parts[1])
	if err != nil {
		return nil, err
	}
	pvV, err := oneNum(ce, ws, parts[2])
	if err != nil {
		return nil, err
	}
	fvV, typ := 0.0, 0.0
	if len(parts) > 3 && strings.TrimSpace(parts[3]) != "" {
		fvV, err = oneNum(ce, ws, parts[3])
		if err != nil {
			return nil, err
		}
	}
	if len(parts) > 4 && strings.TrimSpace(parts[4]) != "" {
		typ, err = oneNum(ce, ws, parts[4])
		if err != nil {
			return nil, err
		}
	}
	if rate == 0 {
		if pmtV == 0 {
			return nil, fmt.Errorf("#NUM!")
		}
		return -(pvV + fvV) / pmtV, nil
	}
	p := pmtV * (1 + rate*typ)
	num := p - fvV*rate
	den := p + pvV*rate
	if num/den <= 0 {
		return nil, fmt.Errorf("#NUM!")
	}
	return math.Log(num/den) / math.Log(1+rate), nil
}

func financeArgs(ce *CalculationEngine, ws *Worksheet, args string, need int) (rate, nper, pv, fv, typ float64, err error) {
	parts := parseArgs(args)
	if len(parts) < need {
		return 0, 0, 0, 0, 0, fmt.Errorf("#VALUE!")
	}
	rate, err = oneNum(ce, ws, parts[0])
	if err != nil {
		return
	}
	nper, err = oneNum(ce, ws, parts[1])
	if err != nil {
		return
	}
	pv, err = oneNum(ce, ws, parts[2])
	if err != nil {
		return
	}
	if len(parts) > 3 && strings.TrimSpace(parts[3]) != "" {
		fv, err = oneNum(ce, ws, parts[3])
		if err != nil {
			return
		}
	}
	if len(parts) > 4 && strings.TrimSpace(parts[4]) != "" {
		typ, err = oneNum(ce, ws, parts[4])
	}
	return
}

func pmt(rate, nper, pv, fv, typ float64) float64 {
	if rate == 0 {
		return -(pv + fv) / nper
	}
	p := math.Pow(1+rate, nper)
	return -(rate * (fv + pv*p)) / ((1 + rate*typ) * (p - 1))
}
func pv(rate, nper, pmtV, fv, typ float64) float64 {
	if rate == 0 {
		return -pmtV*nper - fv
	}
	p := math.Pow(1+rate, nper)
	return -(fv + pmtV*(1+rate*typ)*(p-1)/rate) / p
}
func fv(rate, nper, pmtV, pvV, typ float64) float64 {
	if rate == 0 {
		return -pvV - pmtV*nper
	}
	p := math.Pow(1+rate, nper)
	return -pvV*p - pmtV*(1+rate*typ)*(p-1)/rate
}

func fnQuotient(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) != 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	n, err := oneNum(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	d, err := oneNum(ce, ws, parts[1])
	if err != nil {
		return nil, err
	}
	if d == 0 {
		return nil, fmt.Errorf("#DIV/0!")
	}
	return math.Trunc(n / d), nil
}

func fnGCD(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	nums, err := statNums(ce, ws, args)
	if err != nil || len(nums) == 0 {
		return nil, fmt.Errorf("#VALUE!")
	}
	g := int(math.Abs(nums[0]))
	for _, n := range nums[1:] {
		g = gcd(g, int(math.Abs(n)))
	}
	return float64(g), nil
}
func fnLCM(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	nums, err := statNums(ce, ws, args)
	if err != nil || len(nums) == 0 {
		return nil, fmt.Errorf("#VALUE!")
	}
	l := int(math.Abs(nums[0]))
	for _, n := range nums[1:] {
		b := int(math.Abs(n))
		if l == 0 || b == 0 {
			l = 0
			continue
		}
		l = l / gcd(l, b) * b
	}
	return float64(l), nil
}
func gcd(a, b int) int {
	for b != 0 {
		a, b = b, a%b
	}
	if a < 0 {
		return -a
	}
	return a
}

func wholeSpan(body string) (*formulaRange, bool) {
	parts := strings.SplitN(stripAbs(body), ":", 2)
	if len(parts) != 2 {
		return nil, false
	}
	if digitsOnly(parts[0]) && digitsOnly(parts[1]) {
		r1, _ := strconv.Atoi(parts[0])
		r2, _ := strconv.Atoi(parts[1])
		if r1 > r2 {
			r1, r2 = r2, r1
		}
		if r1 < 1 {
			return nil, false
		}
		return &formulaRange{rows: r2 - r1 + 1, cols: 16384, row0: r1 - 1, col0: 0}, true
	}
	if isLetters(parts[0]) && isLetters(parts[1]) {
		c1, c2 := colIndex(parts[0]), colIndex(parts[1])
		if c1 > c2 {
			c1, c2 = c2, c1
		}
		return &formulaRange{rows: 1048576, cols: c2 - c1 + 1, row0: 0, col0: c1}, true
	}
	return nil, false
}

func digitsOnly(s string) bool {
	if s == "" {
		return false
	}
	for _, c := range s {
		if c < '0' || c > '9' {
			return false
		}
	}
	return true
}

func isLetters(s string) bool {
	if s == "" {
		return false
	}
	for i := 0; i < len(s); i++ {
		if !isLetter(s[i]) {
			return false
		}
	}
	return true
}

func colIndex(s string) int {
	n := 0
	for i := 0; i < len(s); i++ {
		c := s[i]
		if c >= 'a' && c <= 'z' {
			c -= 'a' - 'A'
		}
		n = n*26 + int(c-'A'+1)
	}
	return n - 1
}

func fnAddress(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	row, err := oneNum(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	col, err := oneNum(ce, ws, parts[1])
	if err != nil {
		return nil, err
	}
	abs := 1
	if len(parts) >= 3 && strings.TrimSpace(parts[2]) != "" {
		n, err := oneNum(ce, ws, parts[2])
		if err != nil {
			return nil, err
		}
		abs = int(n)
	}
	a1 := true
	if len(parts) >= 4 && strings.TrimSpace(parts[3]) != "" {
		v, err := ce.evaluate(ws, parts[3])
		if err != nil {
			return nil, err
		}
		a1, err = isTruthy(v)
		if err != nil {
			return nil, err
		}
	}
	sheetName := ""
	if len(parts) >= 5 && strings.TrimSpace(parts[4]) != "" {
		v, err := ce.evaluate(ws, parts[4])
		if err != nil {
			return nil, err
		}
		sheetName, err = ce.textArg(v)
		if err != nil {
			return nil, err
		}
	}
	if row < 1 || col < 1 {
		return nil, fmt.Errorf("#VALUE!")
	}
	var ref string
	if a1 {
		name, err := CellName(int(row)-1, int(col)-1)
		if err != nil {
			return nil, fmt.Errorf("#VALUE!")
		}
		colAbs := abs == 1 || abs == 3
		rowAbs := abs == 1 || abs == 2
		letters := name
		digits := ""
		i := 0
		for i < len(letters) && isLetter(letters[i]) {
			i++
		}
		digits = letters[i:]
		letters = letters[:i]
		if colAbs {
			letters = "$" + letters
		}
		if rowAbs {
			digits = "$" + digits
		}
		ref = letters + digits
	} else {
		ref = fmt.Sprintf("R%dC%d", int(row), int(col))
	}
	if sheetName != "" {
		if strings.ContainsAny(sheetName, " !") {
			sheetName = "'" + strings.ReplaceAll(sheetName, "'", "''") + "'"
		}
		ref = sheetName + "!" + ref
	}
	return ref, nil
}

func fnClean(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	v, err := ce.evaluate(ws, args)
	if err != nil {
		return nil, err
	}
	s, err := ce.textArg(v)
	if err != nil {
		// CLEAN of a number is the number's text in some cases; Excel CLEAN
		// requires text and numbers are #VALUE!. Keep that.
		return nil, err
	}
	var b strings.Builder
	for _, r := range s {
		if r >= 32 {
			b.WriteRune(r)
		}
	}
	return b.String(), nil
}

func fnNA(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return nil, fmt.Errorf("#N/A")
}

func fnRept(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) != 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	v, err := ce.evaluate(ws, parts[0])
	if err != nil {
		return nil, err
	}
	s, err := ce.textArg(v)
	if err != nil {
		return nil, err
	}
	n, err := oneNum(ce, ws, parts[1])
	if err != nil {
		return nil, err
	}
	if n < 0 {
		return nil, fmt.Errorf("#VALUE!")
	}
	if int(n)*len(s) > 32767 {
		return nil, fmt.Errorf("#VALUE!")
	}
	return strings.Repeat(s, int(n)), nil
}

func fnCountBlank(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	rg, err := ce.evalRange(ws, args)
	if err != nil {
		return nil, err
	}
	n := 0.0
	for _, v := range flattenAreas(rg) {
		if v == nil || v == "" {
			n++
		}
	}
	return n, nil
}

func fnAverageA(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	var nums []float64
	for _, part := range parts {
		rg, err := ce.evalRange(ws, part)
		if err != nil {
			return nil, err
		}
		for _, v := range flattenAreas(rg) {
			if err := excelError(v); err != nil {
				return nil, err
			}
			if v == nil || v == "" {
				continue
			}
			if n, ok := asNumber(v, true); ok {
				nums = append(nums, n)
			}
		}
	}
	if len(nums) == 0 {
		return nil, fmt.Errorf("#DIV/0!")
	}
	s := 0.0
	for _, n := range nums {
		s += n
	}
	return s / float64(len(nums)), nil
}

func fnLookup(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 2 || len(parts) > 3 {
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
	look, err := ce.evalRange(ws, parts[1])
	if err != nil {
		return nil, err
	}
	ret := look
	if len(parts) == 3 {
		ret, err = ce.evalRange(ws, parts[2])
		if err != nil {
			return nil, err
		}
	} else if look.cols > 1 && look.rows > 1 {
		// Vector is the last row or last column, whichever is longer.
		if look.cols >= look.rows {
			vals := make([]interface{}, look.cols)
			for c := 0; c < look.cols; c++ {
				vals[c] = look.at(look.rows-1, c)
			}
			look = &formulaRange{values: vals, rows: 1, cols: look.cols}
			retVals := make([]interface{}, look.cols)
			copy(retVals, vals)
			ret = look
			_ = retVals
		}
	}
	n := look.cellCount()
	if ret.cellCount() < n {
		n = ret.cellCount()
	}
	found := -1
	for i := 0; i < n; i++ {
		cell := flattenAreas(look)[i]
		if compareValues(cell, key, "<=") {
			found = i
		} else {
			break
		}
	}
	if found < 0 {
		return nil, fmt.Errorf("#N/A")
	}
	return flattenAreas(ret)[found], nil
}

func fnDays360(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	a, err := evalDate(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	b, err := evalDate(ce, ws, parts[1])
	if err != nil {
		return nil, err
	}
	method := false
	if len(parts) >= 3 {
		v, err := ce.evaluate(ws, parts[2])
		if err != nil {
			return nil, err
		}
		method, _ = isTruthy(v)
	}
	y1, m1, d1 := a.Date()
	y2, m2, d2 := b.Date()
	if !method {
		if d1 == 31 {
			d1 = 30
		}
		if d2 == 31 && d1 >= 30 {
			d2 = 30
		}
	} else {
		if d1 == 31 {
			d1 = 30
		}
		if d2 == 31 {
			d2 = 30
		}
	}
	return float64((y2-y1)*360 + (int(m2)-int(m1))*30 + (d2 - d1)), nil
}

func cp1252(n int) rune {
	if n < 128 || n >= 160 {
		return rune(n)
	}
	// Windows-1252 bytes 128-159 that differ from Unicode.
	table := []rune{
		0x20AC, 0xFFFD, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021,
		0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0xFFFD, 0x017D, 0xFFFD,
		0xFFFD, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014,
		0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0xFFFD, 0x017E, 0x0178,
	}
	return table[n-128]
}
