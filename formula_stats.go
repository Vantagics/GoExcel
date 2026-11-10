package gospreadsheet

import (
	"fmt"
	"math"
	"time"
)

func fnRate(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 3 {
		return nil, fmt.Errorf("#VALUE!")
	}
	nper, err := evalMust(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	pmt, err := evalMust(ce, ws, parts[1])
	if err != nil {
		return nil, err
	}
	pv, err := evalMust(ce, ws, parts[2])
	if err != nil {
		return nil, err
	}
	fv, typ, guess := 0.0, 0.0, 0.1
	if len(parts) > 3 && stringsTrim(parts[3]) != "" {
		fv, err = evalMust(ce, ws, parts[3])
		if err != nil {
			return nil, err
		}
	}
	if len(parts) > 4 && stringsTrim(parts[4]) != "" {
		typ, err = evalMust(ce, ws, parts[4])
		if err != nil {
			return nil, err
		}
	}
	if len(parts) > 5 && stringsTrim(parts[5]) != "" {
		guess, err = evalMust(ce, ws, parts[5])
		if err != nil {
			return nil, err
		}
	}
	rate := guess
	for i := 0; i < 50; i++ {
		f, df := 0.0, 0.0
		pow := 1.0
		for k := 0; k < int(nper); k++ {
			pow *= 1 + rate
		}
		if rate == 0 {
			f = pv + pmt*nper + fv
			df = pmt * nper * (nper - 1) / 2
		} else {
			f = pv*pow + pmt*(1+rate*typ)*(pow-1)/rate + fv
			df = pv*float64(int(nper))*pow/(1+rate) + pmt*(1+rate*typ)*((float64(int(nper))*pow/(1+rate)*rate-(pow-1))/(rate*rate))
		}
		if math.Abs(f) < 1e-7 {
			return rate, nil
		}
		if df == 0 || math.IsNaN(df) {
			break
		}
		rate -= f / df
	}
	if math.IsNaN(rate) || math.IsInf(rate, 0) {
		return nil, fmt.Errorf("#NUM!")
	}
	return rate, nil
}

func stringsTrim(s string) string {
	return trimSpace(s)
}

func trimSpace(s string) string {
	i, j := 0, len(s)
	for i < j && (s[i] == ' ' || s[i] == '\t') {
		i++
	}
	for j > i && (s[j-1] == ' ' || s[j-1] == '\t') {
		j--
	}
	return s[i:j]
}

func fnBinomDist(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 4 {
		return nil, fmt.Errorf("#VALUE!")
	}
	k, err := evalMust(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	n, err := evalMust(ce, ws, parts[1])
	if err != nil {
		return nil, err
	}
	p, err := evalMust(ce, ws, parts[2])
	if err != nil {
		return nil, err
	}
	cumV, err := ce.evaluate(ws, parts[3])
	if err != nil {
		return nil, err
	}
	cum, _ := isTruthy(cumV)
	ki, ni := int(k), int(n)
	if ki < 0 || ni < 0 || ki > ni || p < 0 || p > 1 {
		return nil, fmt.Errorf("#NUM!")
	}
	if !cum {
		return binomPMF(ki, ni, p), nil
	}
	s := 0.0
	for i := 0; i <= ki; i++ {
		s += binomPMF(i, ni, p)
	}
	return s, nil
}

func binomPMF(k, n int, p float64) float64 {
	return math.Exp(logComb(n, k) + float64(k)*math.Log(p) + float64(n-k)*math.Log(1-p))
}

func logComb(n, k int) float64 {
	if k < 0 || k > n {
		return math.Inf(-1)
	}
	s := 0.0
	for i := 1; i <= k; i++ {
		s += math.Log(float64(n-k+i)) - math.Log(float64(i))
	}
	return s
}

func fnExponDist(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 3 {
		return nil, fmt.Errorf("#VALUE!")
	}
	x, err := evalMust(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	lambda, err := evalMust(ce, ws, parts[1])
	if err != nil {
		return nil, err
	}
	cumV, err := ce.evaluate(ws, parts[2])
	if err != nil {
		return nil, err
	}
	cum, _ := isTruthy(cumV)
	if x < 0 || lambda <= 0 {
		return nil, fmt.Errorf("#NUM!")
	}
	if cum {
		return 1 - math.Exp(-lambda*x), nil
	}
	return lambda * math.Exp(-lambda*x), nil
}

func fnCovariance(ce *CalculationEngine, ws *Worksheet, args string, sample bool) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) != 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	a, err := collectList(ce, ws, parts[0], false)
	if err != nil {
		return nil, err
	}
	b, err := collectList(ce, ws, parts[1], false)
	if err != nil {
		return nil, err
	}
	n := len(a)
	if len(b) < n {
		n = len(b)
	}
	if n == 0 || (sample && n < 2) {
		return nil, fmt.Errorf("#DIV/0!")
	}
	ma, mb := 0.0, 0.0
	for i := 0; i < n; i++ {
		ma += a[i]
		mb += b[i]
	}
	ma /= float64(n)
	mb /= float64(n)
	s := 0.0
	for i := 0; i < n; i++ {
		s += (a[i] - ma) * (b[i] - mb)
	}
	den := float64(n)
	if sample {
		den--
	}
	return s / den, nil
}

func fnCovP(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return fnCovariance(ce, ws, args, false)
}
func fnCovS(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	return fnCovariance(ce, ws, args, true)
}

func fnErfPrecise(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	x, err := evalMust(ce, ws, args)
	if err != nil {
		return nil, err
	}
	return math.Erf(x), nil
}

func fnErfcPrecise(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	x, err := evalMust(ce, ws, args)
	if err != nil {
		return nil, err
	}
	return math.Erfc(x), nil
}

func fnConfidenceNorm(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) != 3 {
		return nil, fmt.Errorf("#VALUE!")
	}
	alpha, err := evalMust(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	sd, err := evalMust(ce, ws, parts[1])
	if err != nil {
		return nil, err
	}
	size, err := evalMust(ce, ws, parts[2])
	if err != nil {
		return nil, err
	}
	if alpha <= 0 || alpha >= 1 || sd <= 0 || size < 1 {
		return nil, fmt.Errorf("#NUM!")
	}
	z := normsInv(1 - alpha/2)
	return z * sd / math.Sqrt(size), nil
}

func normsInv(p float64) float64 {
	// Acklam's rational approximation.
	if p <= 0 || p >= 1 {
		return math.NaN()
	}
	a := []float64{-3.969683028665376e+01, 2.209460984245205e+02, -2.759285104469687e+02, 1.383577518672690e+02, -3.066479806614716e+01, 2.506628277459239e+00}
	b := []float64{-5.447609879822406e+01, 1.615858368580409e+02, -1.556989798598866e+02, 6.680131188771972e+01, -1.328068155288572e+01}
	c := []float64{-7.784894002430293e-03, -3.223964580411365e-01, -2.400758277161838e+00, -2.549732539343734e+00, 4.374664141464968e+00, 2.938163982698783e+00}
	d := []float64{7.784695709041462e-03, 3.224671290700398e-01, 2.445134137142996e+00, 3.754408661907416e+00}
	plow, phigh := 0.02425, 1-0.02425
	var q, r float64
	if p < plow {
		q = math.Sqrt(-2 * math.Log(p))
		return (((((c[0]*q+c[1])*q+c[2])*q+c[3])*q+c[4])*q + c[5]) / ((((d[0]*q+d[1])*q+d[2])*q+d[3])*q + 1)
	}
	if p > phigh {
		q = math.Sqrt(-2 * math.Log(1-p))
		return -(((((c[0]*q+c[1])*q+c[2])*q+c[3])*q+c[4])*q + c[5]) / ((((d[0]*q+d[1])*q+d[2])*q+d[3])*q + 1)
	}
	q = p - 0.5
	r = q * q
	return (((((a[0]*r+a[1])*r+a[2])*r+a[3])*r+a[4])*r + a[5]) * q / (((((b[0]*r+b[1])*r+b[2])*r+b[3])*r+b[4])*r + 1)
}

func fnPercentileExc(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) != 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	nums, err := collectList(ce, ws, parts[0], false)
	if err != nil {
		return nil, err
	}
	k, err := evalMust(ce, ws, parts[1])
	if err != nil {
		return nil, err
	}
	n := len(nums)
	if n == 0 || k <= 0 || k >= 1 {
		return nil, fmt.Errorf("#NUM!")
	}
	for i := 0; i < n; i++ {
		for j := i + 1; j < n; j++ {
			if nums[j] < nums[i] {
				nums[i], nums[j] = nums[j], nums[i]
			}
		}
	}
	pos := k * float64(n+1)
	if pos < 1 || pos > float64(n) {
		return nil, fmt.Errorf("#NUM!")
	}
	lo := int(math.Floor(pos))
	hi := int(math.Ceil(pos))
	if lo < 1 {
		lo = 1
	}
	if hi > n {
		hi = n
	}
	if lo == hi {
		return nums[lo-1], nil
	}
	w := pos - float64(lo)
	return nums[lo-1]*(1-w) + nums[hi-1]*w, nil
}

func fnIsEven(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	n, err := evalMust(ce, ws, args)
	if err != nil {
		return nil, err
	}
	return int(math.Trunc(math.Abs(n)))%2 == 0, nil
}

func fnIsOdd(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	n, err := evalMust(ce, ws, args)
	if err != nil {
		return nil, err
	}
	return int(math.Trunc(math.Abs(n)))%2 == 1, nil
}

func fnISOWeekNum(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	t, err := evalDate(ce, ws, args)
	if err != nil {
		return nil, err
	}
	_, w := t.ISOWeek()
	return float64(w), nil
}

func fnDelta(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 1 || len(parts) > 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	a, err := evalMust(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	b := 0.0
	if len(parts) == 2 {
		b, err = evalMust(ce, ws, parts[1])
		if err != nil {
			return nil, err
		}
	}
	if a == b {
		return 1.0, nil
	}
	return 0.0, nil
}

func fnNegBinom(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 4 {
		return nil, fmt.Errorf("#VALUE!")
	}
	f, err := evalMust(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	s, err := evalMust(ce, ws, parts[1])
	if err != nil {
		return nil, err
	}
	p, err := evalMust(ce, ws, parts[2])
	if err != nil {
		return nil, err
	}
	cumV, err := ce.evaluate(ws, parts[3])
	if err != nil {
		return nil, err
	}
	cum, _ := isTruthy(cumV)
	fi, si := int(f), int(s)
	if fi < 0 || si < 1 || p <= 0 || p > 1 {
		return nil, fmt.Errorf("#NUM!")
	}
	pmf := func(ff int) float64 {
		return math.Exp(logComb(ff+si-1, ff) + float64(si)*math.Log(p) + float64(ff)*math.Log(1-p))
	}
	if !cum {
		return pmf(fi), nil
	}
	sum := 0.0
	for i := 0; i <= fi; i++ {
		sum += pmf(i)
	}
	return sum, nil
}

func fnHypGeom(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 5 {
		return nil, fmt.Errorf("#VALUE!")
	}
	sampleS, err := evalMust(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	numberS, err := evalMust(ce, ws, parts[1])
	if err != nil {
		return nil, err
	}
	popS, err := evalMust(ce, ws, parts[2])
	if err != nil {
		return nil, err
	}
	numberP, err := evalMust(ce, ws, parts[3])
	if err != nil {
		return nil, err
	}
	cumV, err := ce.evaluate(ws, parts[4])
	if err != nil {
		return nil, err
	}
	cum, _ := isTruthy(cumV)
	ks, ns, Ks, Ns := int(sampleS), int(numberS), int(popS), int(numberP)
	pmf := func(k int) float64 {
		if k < 0 || k > ns || k > Ks || ns-k > Ns-Ks {
			return 0
		}
		return math.Exp(logComb(Ks, k) + logComb(Ns-Ks, ns-k) - logComb(Ns, ns))
	}
	if !cum {
		return pmf(ks), nil
	}
	sum := 0.0
	for k := 0; k <= ks; k++ {
		sum += pmf(k)
	}
	return sum, nil
}

func fnGeomean(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	nums, err := collectList(ce, ws, args, false)
	if err != nil || len(nums) == 0 {
		return nil, fmt.Errorf("#NUM!")
	}
	s := 0.0
	for _, n := range nums {
		if n <= 0 {
			return nil, fmt.Errorf("#NUM!")
		}
		s += math.Log(n)
	}
	return math.Exp(s / float64(len(nums))), nil
}

func fnHarmean(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	nums, err := collectList(ce, ws, args, false)
	if err != nil || len(nums) == 0 {
		return nil, fmt.Errorf("#NUM!")
	}
	s := 0.0
	for _, n := range nums {
		if n <= 0 {
			return nil, fmt.Errorf("#NUM!")
		}
		s += 1 / n
	}
	return float64(len(nums)) / s, nil
}

func fnWeekNum(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 1 {
		return nil, fmt.Errorf("#VALUE!")
	}
	t, err := evalDate(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	// Excel WEEKNUM default: week starts Sunday, week 1 contains Jan 1.
	y, m, d := t.Date()
	start := time.Date(y, 1, 1, 0, 0, 0, 0, time.UTC)
	offset := int(start.Weekday())
	return float64((int(time.Date(y, m, d, 0, 0, 0, 0, time.UTC).Sub(start).Hours()/24)+offset)/7 + 1), nil
}
