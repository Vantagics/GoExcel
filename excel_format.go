package gospreadsheet

import (
	"math"
	"strconv"
	"strings"
	"time"
	"unicode"
	"unicode/utf8"
)

// formatByExcelCode formats a value with an Excel number format, including the
// TEXT function and formatted cell reads.
func formatByExcelCode(v interface{}, code string) string {
	code = strings.TrimSpace(code)
	if code == "" || strings.EqualFold(code, "General") {
		if n, ok := asNumber(v, true); ok {
			return formatGeneral(n)
		}
		if t, ok := v.(time.Time); ok {
			return formatGeneral(dateToSerial(t))
		}
		return excelText(v)
	}
	sections := splitFormatSections(code)
	if s, ok := v.(string); ok && excelError(s) == nil {
		if len(sections) >= 4 {
			return renderTextSection(sections[3], s)
		}
		return renderTextSection(firstTextSection(sections), s)
	}
	n, ok := asNumber(v, true)
	if !ok {
		if t, ok := v.(time.Time); ok {
			n = dateToSerial(t)
			ok = true
		}
	}
	if !ok {
		return excelText(v)
	}
	sec, val, classicNeg := chooseSection(n, sections)
	if strings.EqualFold(strings.TrimSpace(stripDecor(sec)), "General") || strings.TrimSpace(stripDecor(sec)) == "" {
		out := formatGeneral(val)
		if classicNeg {
			return out
		}
		if n < 0 && val == n {
			return "-" + formatGeneral(-n)
		}
		return out
	}
	if isDateFormat(sec) || hasElapsed(sec) {
		return formatExcelDate(val, sec)
	}
	if strings.Contains(sec, "@") {
		return renderTextSection(sec, formatGeneral(math.Abs(val)))
	}
	out := formatExcelNumber(math.Abs(val), sec)
	if val < 0 && !classicNeg {
		return "-" + out
	}
	return out
}

func formatGeneral(n float64) string {
	if n == 0 {
		return "0"
	}
	a := math.Abs(n)
	if a >= 1e11 || (a > 0 && a < 1e-9) {
		return excelGeneralSci(n)
	}
	if a >= 1e10 {
		return strconv.FormatFloat(n, 'f', 0, 64)
	}
	return strconv.FormatFloat(n, 'g', 10, 64)
}

func excelGeneralSci(n float64) string {
	sign := ""
	if n < 0 {
		sign = "-"
		n = -n
	}
	exp := int(math.Floor(math.Log10(n)))
	mant := n / math.Pow(10, float64(exp))
	ms := strconv.FormatFloat(mant, 'g', 10, 64)
	return sign + ms + "E" + strconv.Itoa(exp)
}

func firstTextSection(sections []string) string {
	if len(sections) >= 4 {
		return sections[3]
	}
	return "@"
}

func renderTextSection(code, text string) string {
	var b strings.Builder
	inQ := false
	for i := 0; i < len(code); i++ {
		ch := code[i]
		if ch == '"' {
			inQ = !inQ
			continue
		}
		if inQ {
			b.WriteByte(ch)
			continue
		}
		if ch == '\\' && i+1 < len(code) {
			b.WriteByte(code[i+1])
			i++
			continue
		}
		if ch == '_' && i+1 < len(code) {
			b.WriteByte(' ')
			i++
			continue
		}
		if ch == '*' && i+1 < len(code) {
			i++
			continue
		}
		if ch == '@' {
			b.WriteString(text)
			continue
		}
		if ch == '[' {
			j := strings.IndexByte(code[i:], ']')
			if j < 0 {
				break
			}
			b.WriteString(localeSymbol(code[i+1 : i+j]))
			i += j
			continue
		}
		if ch == '0' || ch == '#' || ch == '?' {
			continue
		}
		b.WriteByte(ch)
	}
	if b.Len() == 0 {
		return text
	}
	return b.String()
}

type fmtSection struct {
	raw     string
	op      string
	cond    float64
	hasCond bool
}

func chooseSection(n float64, raw []string) (string, float64, bool) {
	secs := make([]fmtSection, len(raw))
	anyCond := false
	for i, s := range raw {
		secs[i] = parseCond(s)
		if secs[i].hasCond {
			anyCond = true
		}
	}
	if anyCond {
		for _, s := range secs {
			if s.hasCond && condMatch(n, s.op, s.cond) {
				return s.raw, n, false
			}
		}
		for _, s := range secs {
			if !s.hasCond {
				return s.raw, n, false
			}
		}
		return "General", n, false
	}
	if n > 0 || len(secs) == 0 {
		if len(secs) == 0 {
			return "General", n, false
		}
		return secs[0].raw, n, false
	}
	if n < 0 && len(secs) >= 2 && strings.TrimSpace(secs[1].raw) != "" {
		return secs[1].raw, math.Abs(n), true
	}
	if n == 0 && len(secs) >= 3 && strings.TrimSpace(secs[2].raw) != "" {
		return secs[2].raw, 0, true
	}
	if n < 0 {
		code := "General"
		if len(secs) > 0 {
			code = secs[0].raw
		}
		return code, n, false
	}
	return secs[0].raw, n, false
}

func parseCond(s string) fmtSection {
	sec := fmtSection{raw: s}
	t := strings.TrimSpace(s)
	if !strings.HasPrefix(t, "[") {
		return sec
	}
	end := strings.IndexByte(t, ']')
	if end < 0 {
		return sec
	}
	body := t[1:end]
	if len(body) == 0 || body[0] == '$' || body[0] == '-' || isColorName(body) {
		return sec
	}
	op := ""
	rest := body
	for _, cand := range []string{">=", "<=", "<>", ">", "<", "="} {
		if strings.HasPrefix(body, cand) {
			op = cand
			rest = strings.TrimSpace(body[len(cand):])
			break
		}
	}
	if op == "" {
		return sec
	}
	val, err := strconv.ParseFloat(rest, 64)
	if err != nil {
		return sec
	}
	sec.op = op
	sec.cond = val
	sec.hasCond = true
	sec.raw = strings.TrimSpace(t[end+1:])
	return sec
}

func condMatch(n float64, op string, v float64) bool {
	switch op {
	case ">":
		return n > v
	case ">=":
		return n >= v
	case "<":
		return n < v
	case "<=":
		return n <= v
	case "=":
		return n == v
	case "<>":
		return n != v
	default:
		return false
	}
}

func isColorName(s string) bool {
	switch strings.ToLower(s) {
	case "black", "blue", "cyan", "green", "magenta", "red", "white", "yellow":
		return true
	}
	return strings.HasPrefix(strings.ToLower(s), "color")
}

func stripDecor(code string) string {
	var b strings.Builder
	for i := 0; i < len(code); i++ {
		if code[i] == '[' {
			j := strings.IndexByte(code[i:], ']')
			if j < 0 {
				break
			}
			inner := code[i+1 : i+j]
			if isColorName(inner) || strings.HasPrefix(inner, "$-") || strings.HasPrefix(inner, "$") {
				i += j
				continue
			}
		}
		b.WriteByte(code[i])
	}
	return b.String()
}

func localeSymbol(inner string) string {
	inner = strings.TrimSpace(inner)
	if isColorName(inner) {
		return ""
	}
	if strings.HasPrefix(inner, "$") {
		body := inner[1:]
		if i := strings.IndexByte(body, '-'); i >= 0 {
			return body[:i]
		}
		return body
	}
	return ""
}

func hasElapsed(code string) bool {
	low := strings.ToLower(code)
	return strings.Contains(low, "[h") || strings.Contains(low, "[m") || strings.Contains(low, "[s")
}

func isDateFormat(code string) bool {
	inQ := false
	for i := 0; i < len(code); i++ {
		ch := code[i]
		if ch == '"' {
			inQ = !inQ
			continue
		}
		if inQ {
			continue
		}
		if ch == '\\' && i+1 < len(code) {
			i++
			continue
		}
		if ch == '[' {
			j := strings.IndexByte(code[i:], ']')
			if j < 0 {
				break
			}
			inner := strings.ToLower(code[i+1 : i+j])
			if strings.HasPrefix(inner, "h") || strings.HasPrefix(inner, "m") || strings.HasPrefix(inner, "s") {
				return true
			}
			i += j
			continue
		}
		switch ch {
		case 'y', 'Y', 'd', 'D', 'h', 'H', 's', 'S':
			return true
		case 'a', 'A':
			rest := strings.ToLower(code[i:])
			if strings.HasPrefix(rest, "am/pm") || strings.HasPrefix(rest, "a/p") || strings.HasPrefix(rest, "aaa") {
				return true
			}
		}
	}
	return false
}

type tok struct {
	k    byte
	text string
}

func tokenize(code string) []tok {
	var out []tok
	inQ := false
	for i := 0; i < len(code); i++ {
		ch := code[i]
		if ch == '"' {
			inQ = !inQ
			continue
		}
		if inQ {
			out = append(out, tok{'L', string(ch)})
			continue
		}
		switch ch {
		case '\\':
			if i+1 < len(code) {
				out = append(out, tok{'L', string(code[i+1])})
				i++
			}
		case '_':
			if i+1 < len(code) {
				out = append(out, tok{'L', " "})
				i++
			}
		case '*':
			if i+1 < len(code) {
				i++
			}
		case '[':
			j := strings.IndexByte(code[i:], ']')
			if j < 0 {
				return out
			}
			sym := localeSymbol(code[i+1 : i+j])
			if sym != "" {
				out = append(out, tok{'L', sym})
			}
			i += j
		case 'E', 'e':
			sign := ""
			if i+1 < len(code) && (code[i+1] == '+' || code[i+1] == '-') {
				sign = string(code[i+1])
				i++
			}
			out = append(out, tok{'E', string(ch) + sign})
		case '0', '#', '?', '.', ',', '%', '/':
			out = append(out, tok{ch, string(ch)})
		default:
			if ch >= 0x80 {
				r, size := utf8.DecodeRuneInString(code[i:])
				out = append(out, tok{'L', string(r)})
				i += size - 1
				break
			}
			out = append(out, tok{'L', string(ch)})
		}
	}
	return out
}

func formatExcelNumber(n float64, code string) string {
	if math.IsNaN(n) || math.IsInf(n, 0) {
		return formatGeneral(n)
	}
	toks := tokenize(code)
	if slash := fractionSlash(toks); slash >= 0 {
		return formatFraction(n, toks, slash)
	}
	if ei := expIndex(toks); ei >= 0 {
		return formatScientific(n, toks, ei)
	}
	return formatFixed(n, toks)
}

func expIndex(toks []tok) int {
	for i, t := range toks {
		if t.k == 'E' {
			return i
		}
	}
	return -1
}

func fractionSlash(toks []tok) int {
	for i, t := range toks {
		if t.k != '/' {
			continue
		}
		if i == 0 || i+1 >= len(toks) {
			continue
		}
		if isPh(toks[i-1].k) && (isPh(toks[i+1].k) || (toks[i+1].k == 'L' && isPh(nextPh(toks, i+1)))) {
			return i
		}
		if isPh(toks[i-1].k) && toks[i+1].k == 'L' {
			return -1
		}
		if digitToken(toks[i+1]) || isPh(toks[i+1].k) {
			if isPh(toks[i-1].k) {
				return i
			}
		}
	}
	return -1
}

func nextPh(toks []tok, i int) byte {
	for _, t := range toks[i:] {
		if isPh(t.k) {
			return t.k
		}
	}
	return 0
}

func isPh(k byte) bool { return k == '0' || k == '#' || k == '?' }
func digitToken(t tok) bool {
	if t.k != 'L' || t.text == "" {
		return false
	}
	for _, r := range t.text {
		if r < '0' || r > '9' {
			return false
		}
	}
	return true
}

func formatFixed(n float64, toks []tok) string {
	percent := 0
	for _, t := range toks {
		if t.k == '%' {
			percent++
		}
	}
	n *= math.Pow(100, float64(percent))
	intPh, fracPh, grouping, scale, prefix, suffix := splitNum(toks)
	n /= math.Pow(1000, float64(scale))
	fracDigits := len(fracPh)
	pow := math.Pow(10, float64(fracDigits))
	if pow == 0 {
		pow = 1
	}
	scaled := n * pow
	if math.Abs(scaled-math.Round(scaled)) < 1e-8 {
		scaled = math.Round(scaled)
	}
	rounded := math.Round(scaled) / pow
	intPart := math.Floor(rounded + 1e-9)
	frac := rounded - intPart
	if frac < 0 {
		frac = 0
	}
	intStr := strconv.FormatFloat(intPart, 'f', 0, 64)
	minZero := countByte(intPh, '0')
	for len(intStr) < minZero {
		intStr = "0" + intStr
	}
	if intStr == "0" && minZero == 0 && !strings.ContainsRune(string(intPh), '0') && intPart == 0 {
		intStr = ""
	}
	spaces := 0
	if len(intPh) > len(intStr) {
		extra := len(intPh) - len(intStr)
		for i := 0; i < extra && i < len(intPh); i++ {
			if intPh[i] == '?' {
				spaces++
			}
		}
	}
	if grouping && intStr != "" {
		intStr = withCommas(intStr)
	}
	fracStr := ""
	showDot := len(fracPh) > 0 || hasDot(toks)
	if fracDigits > 0 {
		raw := int(math.Round(frac*pow + 1e-9))
		fracStr = strconv.Itoa(raw)
		for len(fracStr) < fracDigits {
			fracStr = "0" + fracStr
		}
		if len(fracStr) > fracDigits {
			fracStr = fracStr[len(fracStr)-fracDigits:]
		}
		end := len(fracStr)
		for end > 0 {
			ph := fracPh[end-1]
			if fracStr[end-1] != '0' {
				break
			}
			if ph == '#' {
				end--
				continue
			}
			if ph == '?' {
				break
			}
			break
		}
		// Turn leftover ? zeros into spaces and drop # zeros already removed.
		var fb strings.Builder
		for i := 0; i < end; i++ {
			if fracStr[i] == '0' && i < len(fracPh) && fracPh[i] == '?' && allZeroAfter(fracStr, i) {
				fb.WriteByte(' ')
			} else {
				fb.WriteByte(fracStr[i])
			}
		}
		fracStr = fb.String()
	}
	var b strings.Builder
	b.WriteString(prefix)
	b.WriteString(strings.Repeat(" ", spaces))
	b.WriteString(intStr)
	if showDot && (fracStr != "" || len(fracPh) > 0) {
		b.WriteByte('.')
		b.WriteString(fracStr)
	}
	b.WriteString(suffix)
	if b.Len() == 0 {
		return "0"
	}
	return b.String()
}

func allZeroAfter(s string, i int) bool {
	for _, c := range s[i:] {
		if c != '0' && c != ' ' {
			return false
		}
	}
	return true
}

func hasDot(toks []tok) bool {
	for _, t := range toks {
		if t.k == '.' {
			return true
		}
	}
	return false
}

func countByte(b []byte, c byte) int {
	n := 0
	for _, x := range b {
		if x == c {
			n++
		}
	}
	return n
}

func splitNum(toks []tok) (intPh, fracPh []byte, grouping bool, scale int, prefix, suffix string) {
	seenDigit := false
	seenDot := false
	var pre, suf strings.Builder
	var pendingComma int
	for _, t := range toks {
		switch t.k {
		case '0', '#', '?':
			if pendingComma > 0 {
				if seenDigit {
					grouping = true
				} else {
					pre.WriteString(strings.Repeat(",", pendingComma))
				}
				pendingComma = 0
			}
			seenDigit = true
			if seenDot {
				fracPh = append(fracPh, t.k)
			} else {
				intPh = append(intPh, t.k)
			}
		case '.':
			if pendingComma > 0 && seenDigit {
				scale += pendingComma
				pendingComma = 0
			}
			seenDot = true
		case ',':
			pendingComma++
		case '%':
			if !seenDigit && !seenDot {
				pre.WriteByte('%')
			} else {
				suf.WriteByte('%')
			}
		default:
			if pendingComma > 0 {
				if seenDigit && !seenDot {
					scale += pendingComma
				} else if !seenDigit {
					pre.WriteString(strings.Repeat(",", pendingComma))
				} else {
					suf.WriteString(strings.Repeat(",", pendingComma))
				}
				pendingComma = 0
			}
			if !seenDigit && !seenDot {
				pre.WriteString(t.text)
			} else {
				suf.WriteString(t.text)
			}
		}
	}
	if pendingComma > 0 && seenDigit {
		scale += pendingComma
	}
	return intPh, fracPh, grouping, scale, pre.String(), suf.String()
}

func withCommas(s string) string {
	neg := strings.HasPrefix(s, "-")
	if neg {
		s = s[1:]
	}
	for i := len(s) - 3; i > 0; i -= 3 {
		s = s[:i] + "," + s[i:]
	}
	if neg {
		return "-" + s
	}
	return s
}

func formatScientific(n float64, toks []tok, ei int) string {
	expToks := toks[ei+1:]
	marker := toks[ei].text
	// Percent and trailing commas do not scale an exponent format; they are literals.
	mant := append([]tok(nil), stripScaleCommas(toks[:ei])...)
	for i := range mant {
		if mant[i].k == '%' {
			mant[i] = tok{'L', "%"}
		}
	}
	intPh := digitPh(mant, false)
	maxInt := len(intPh)
	if maxInt < 1 {
		maxInt = 1
	}
	sign := ""
	if n < 0 {
		sign = "-"
		n = -n
	}
	exp := 0
	if n != 0 {
		ee := int(math.Floor(math.Log10(n) + 1e-12))
		mod := ee % maxInt
		if mod < 0 {
			mod += maxInt
		}
		exp = ee - mod
		n = n / math.Pow(10, float64(exp))
	}
	body := formatFixed(n, mant)
	return sign + body + formatExp(exp, marker, expToks)
}

func stripScaleCommas(toks []tok) []tok {
	end := len(toks)
	for end > 0 && toks[end-1].k == ',' {
		end--
	}
	return toks[:end]
}

func digitPh(toks []tok, frac bool) []byte {
	seenDot := false
	var ph []byte
	for _, t := range toks {
		if t.k == '.' {
			seenDot = true
			continue
		}
		if isPh(t.k) && seenDot == frac {
			ph = append(ph, t.k)
		}
	}
	return ph
}

func formatExp(exp int, marker string, toks []tok) string {
	letter := "E"
	showPlus := false
	if marker != "" {
		letter = marker[:1]
		showPlus = strings.Contains(marker, "+")
	}
	zeros := 0
	places := 0
	for _, t := range toks {
		if t.k == '0' {
			zeros++
			places++
		} else if t.k == '#' || t.k == '?' {
			places++
		}
	}
	if places < 1 {
		places = 1
	}
	abs := exp
	if abs < 0 {
		abs = -abs
	}
	s := strconv.Itoa(abs)
	for len(s) < zeros {
		s = "0" + s
	}
	for len(s) < places && zeros == 0 {
		break
	}
	sig := ""
	if exp < 0 {
		sig = "-"
	} else if showPlus {
		sig = "+"
	}
	// Literals stay in place. The sign sits immediately before the digits.
	var before, after strings.Builder
	seenDigit := false
	for _, t := range toks {
		if isPh(t.k) {
			seenDigit = true
			continue
		}
		if t.k == 'L' || t.k == '%' {
			if seenDigit {
				after.WriteString(t.text)
			} else {
				before.WriteString(t.text)
			}
		}
	}
	return letter + before.String() + sig + s + after.String()
}

func formatFraction(n float64, toks []tok, slash int) string {
	neg := n < 0
	n = math.Abs(n)
	numPh, denPh, denFixed, prefix, sep, suffix, mixed := fractionParts(toks, slash)
	maxDen := 9
	if denFixed > 0 {
		maxDen = denFixed
	} else if len(denPh) > 0 {
		maxDen = 1
		for i := 0; i < len(denPh); i++ {
			maxDen *= 10
		}
		maxDen--
		if maxDen < 1 {
			maxDen = 1
		}
	}
	whole := 0.0
	frac := n
	if mixed {
		whole = math.Floor(n + 1e-12)
		frac = n - whole
	}
	num, den := 0, 1
	if denFixed > 0 {
		den = denFixed
		num = int(math.Round(frac * float64(den)))
	} else {
		num, den = bestFraction(frac, maxDen)
	}
	if den != 0 && num == den {
		whole++
		num = 0
	}
	var b strings.Builder
	if neg {
		b.WriteByte('-')
	}
	b.WriteString(prefix)
	if mixed && (whole != 0 || leadingZero(toks)) {
		b.WriteString(strconv.FormatFloat(whole, 'f', 0, 64))
		if sep == "" {
			sep = " "
		}
		b.WriteString(sep)
	}
	if num == 0 && mixed {
		if whole == 0 {
			b.WriteByte('0')
		}
	} else {
		b.WriteString(padPh(strconv.Itoa(num), numPh))
		b.WriteByte('/')
		b.WriteString(padPh(strconv.Itoa(den), denPh))
	}
	b.WriteString(suffix)
	return b.String()
}

func leadingZero(toks []tok) bool {
	for _, t := range toks {
		if t.k == '0' {
			return true
		}
	}
	return false
}

func padPh(s string, ph []byte) string {
	zeros := countByte(ph, '0')
	for len(s) < zeros {
		s = "0" + s
	}
	return s
}

func fractionParts(toks []tok, slash int) (numPh, denPh []byte, denFixed int, prefix, sep, suffix string, mixed bool) {
	// Numerator is the digit run immediately before the slash.
	i := slash - 1
	for i >= 0 && isPh(toks[i].k) {
		numPh = append([]byte{toks[i].k}, numPh...)
		i--
	}
	orig := i
	walk := i
	for walk >= 0 && !isPh(toks[walk].k) {
		walk--
	}
	if walk >= 0 && isPh(toks[walk].k) {
		mixed = true
		var sp strings.Builder
		for _, t := range toks[walk+1 : orig+1] {
			if t.k == 'L' {
				sp.WriteString(t.text)
			}
		}
		sep = sp.String()
		i = walk
	}
	j := slash + 1
	fixed := ""
	for j < len(toks) && (isPh(toks[j].k) || digitToken(toks[j])) {
		if isPh(toks[j].k) {
			denPh = append(denPh, toks[j].k)
		} else {
			fixed += toks[j].text
		}
		j++
	}
	if fixed != "" && len(denPh) == 0 {
		denFixed, _ = strconv.Atoi(fixed)
	}
	cut := orig + 1
	if mixed {
		cut = i + 1
	}
	var pre strings.Builder
	for _, t := range toks[:max(0, cut)] {
		if t.k == 'L' {
			pre.WriteString(t.text)
		}
	}
	// Drop the separating space from the prefix when it belongs to the mixed number.
	ps := pre.String()
	if mixed {
		ps = strings.TrimRight(ps, " ")
	}
	_ = sep
	var suf strings.Builder
	for _, t := range toks[j:] {
		if t.k == 'L' || t.k == '%' {
			suf.WriteString(t.text)
		}
	}
	return numPh, denPh, denFixed, ps, sep, suf.String(), mixed
}

func max(a, b int) int {
	if a > b {
		return a
	}
	return b
}

func bestFraction(x float64, maxDen int) (num, den int) {
	if x <= 1e-12 || maxDen < 1 {
		return 0, 1
	}
	bestN, bestD := 0, 1
	bestErr := x
	for d := 1; d <= maxDen; d++ {
		n := int(math.Round(x * float64(d)))
		if n < 0 {
			n = 0
		}
		err := math.Abs(x - float64(n)/float64(d))
		if err < bestErr-1e-12 || (math.Abs(err-bestErr) <= 1e-12 && d < bestD) {
			bestErr = err
			bestN, bestD = n, d
		}
	}
	g := gcdInt(bestN, bestD)
	if g > 1 {
		bestN /= g
		bestD /= g
	}
	return bestN, bestD
}

func gcdInt(a, b int) int {
	if a < 0 {
		a = -a
	}
	if b < 0 {
		b = -b
	}
	for b != 0 {
		a, b = b, a%b
	}
	if a == 0 {
		return 1
	}
	return a
}

func formatExcelDate(serial float64, code string) string {
	if hasElapsed(code) {
		return formatElapsed(serial, code)
	}
	serial = roundDateSerial(serial, code)
	t := excelSerialToTime(serial)
	frac := serial - math.Floor(serial)
	if frac < 0 {
		frac += 1
	}
	totalSec := int(math.Round(frac * 86400))
	hour := (totalSec / 3600) % 24
	minute := (totalSec / 60) % 60
	second := totalSec % 60
	ms := int(math.Round((frac*86400 - math.Floor(frac*86400)) * 1000))
	if ms == 1000 {
		ms = 0
	}
	ampm := strings.Contains(strings.ToLower(code), "am/pm") || strings.Contains(strings.ToLower(code), "a/p")
	var b strings.Builder
	inQ := false
	for i := 0; i < len(code); {
		ch := code[i]
		if ch == '"' {
			inQ = !inQ
			i++
			continue
		}
		if inQ {
			b.WriteByte(ch)
			i++
			continue
		}
		if ch == '\\' && i+1 < len(code) {
			b.WriteByte(code[i+1])
			i += 2
			continue
		}
		if ch == '_' && i+1 < len(code) {
			b.WriteByte(' ')
			i += 2
			continue
		}
		if ch == '*' && i+1 < len(code) {
			i += 2
			continue
		}
		if ch == '[' {
			j := strings.IndexByte(code[i:], ']')
			if j < 0 {
				break
			}
			b.WriteString(localeSymbol(code[i+1 : i+j]))
			i += j + 1
			continue
		}
		low := strings.ToLower(code[i:])
		if strings.HasPrefix(low, "am/pm") {
			b.WriteString(ampmWord(code[i:i+5], hour))
			i += 5
			continue
		}
		if strings.HasPrefix(low, "a/p") {
			b.WriteString(ampmWord(code[i:i+3], hour))
			i += 3
			continue
		}
		if strings.HasPrefix(low, "aaaa") {
			b.WriteString(t.Weekday().String())
			i += 4
			continue
		}
		if strings.HasPrefix(low, "aaa") {
			name := t.Weekday().String()
			if len(name) > 3 {
				name = name[:3]
			}
			b.WriteString(name)
			i += 3
			continue
		}
		run := 1
		for i+run < len(code) && code[i+run] == ch {
			run++
		}
		switch ch {
		case 'y', 'Y':
			y := t.Year()
			if run < 3 {
				b.WriteString(pad(y%100, 2))
			} else {
				b.WriteString(pad(y, 4))
			}
		case 'm', 'M':
			if minuteContext(code, i) {
				if run >= 2 {
					b.WriteString(pad(minute, 2))
				} else {
					b.WriteString(strconv.Itoa(minute))
				}
			} else if run >= 5 {
				b.WriteString(t.Month().String()[:1])
			} else if run == 4 {
				b.WriteString(t.Month().String())
			} else if run == 3 {
				b.WriteString(t.Month().String()[:3])
			} else if run == 2 {
				b.WriteString(pad(int(t.Month()), 2))
			} else {
				b.WriteString(strconv.Itoa(int(t.Month())))
			}
		case 'd', 'D':
			if run >= 4 {
				b.WriteString(t.Weekday().String())
			} else if run == 3 {
				b.WriteString(t.Weekday().String()[:3])
			} else if run == 2 {
				b.WriteString(pad(t.Day(), 2))
			} else {
				b.WriteString(strconv.Itoa(t.Day()))
			}
		case 'h', 'H':
			h := hour
			if ampm {
				h = hour % 12
				if h == 0 {
					h = 12
				}
			}
			if run >= 2 {
				b.WriteString(pad(h, 2))
			} else {
				b.WriteString(strconv.Itoa(h))
			}
		case 's', 'S':
			if run >= 2 {
				b.WriteString(pad(second, 2))
			} else {
				b.WriteString(strconv.Itoa(second))
			}
		case '.':
			if i+1 < len(code) && (code[i+1] == '0' || code[i+1] == '#') {
				k := 1
				for i+k < len(code) && (code[i+k] == '0' || code[i+k] == '#') {
					k++
				}
				digits := k - 1
				b.WriteByte('.')
				msStr := pad(ms, 3)
				if digits < len(msStr) {
					b.WriteString(msStr[:digits])
				} else {
					b.WriteString(msStr)
					b.WriteString(strings.Repeat("0", digits-len(msStr)))
				}
				i += k
				continue
			}
			b.WriteByte(ch)
			run = 1
		default:
			if unicode.IsLetter(rune(ch)) && (ch == 'g' || ch == 'G' || ch == 'e' || ch == 'E') {
				// era / weekday tokens we do not translate; skip the run
			} else {
				b.WriteByte(ch)
			}
			if ch != '.' {
				// run already counted
			}
		}
		i += run
	}
	return b.String()
}

func ampmWord(pattern string, hour int) string {
	morning := hour < 12
	low := strings.ToLower(pattern)
	if strings.HasPrefix(low, "am/pm") {
		if morning {
			return "AM"
		}
		return "PM"
	}
	parts := strings.SplitN(pattern, "/", 2)
	if morning {
		return parts[0]
	}
	if len(parts) > 1 {
		return parts[1]
	}
	return pattern
}

func minuteContext(code string, i int) bool {
	end := i
	for end < len(code) && (code[end] == 'm' || code[end] == 'M') {
		end++
	}
	prev := lastDateLetter(code, i-1, -1)
	next := lastDateLetter(code, end, 1)
	return prev == 'h' || prev == 'H' || next == 's' || next == 'S'
}

func lastDateLetter(code string, i, dir int) byte {
	for i >= 0 && i < len(code) {
		ch := code[i]
		if strings.ContainsRune("mMhHsSdDyY", rune(ch)) {
			return ch
		}
		if unicode.IsLetter(rune(ch)) {
			return 0
		}
		i += dir
	}
	return 0
}

func roundDateSerial(serial float64, code string) float64 {
	unit := 0.0
	low := strings.ToLower(code)
	if strings.Contains(low, "s") {
		dec := 0
		if i := strings.Index(low, "s"); i >= 0 {
			j := i
			for j < len(low) && low[j] == 's' {
				j++
			}
			if j < len(low) && low[j] == '.' {
				j++
				for j < len(low) && (low[j] == '0' || low[j] == '#') {
					dec++
					j++
				}
			}
		}
		unit = 1.0 / 86400 / math.Pow(10, float64(dec))
	} else if strings.Contains(low, "h") || minuteInFormat(low) {
		if minuteInFormat(low) && !strings.Contains(low, "s") {
			unit = 1.0 / 1440
		} else if strings.Contains(low, "h") {
			unit = 1.0 / 24
		}
	}
	if unit == 0 {
		return serial
	}
	return math.Round(serial/unit) * unit
}

func minuteInFormat(low string) bool {
	for i := 0; i < len(low); i++ {
		if low[i] == 'm' {
			// crude: m not part of am/pm
			if i > 0 && low[i-1] == 'a' {
				continue
			}
			return true
		}
	}
	return false
}

func formatElapsed(serial float64, code string) string {
	secTotal := serial * 86400
	var b strings.Builder
	inQ := false
	for i := 0; i < len(code); {
		ch := code[i]
		if ch == '"' {
			inQ = !inQ
			i++
			continue
		}
		if inQ {
			b.WriteByte(ch)
			i++
			continue
		}
		if ch == '\\' && i+1 < len(code) {
			b.WriteByte(code[i+1])
			i += 2
			continue
		}
		if ch == '[' {
			j := strings.IndexByte(code[i:], ']')
			if j < 0 {
				break
			}
			token := strings.ToLower(code[i+1 : i+j])
			b.WriteString(elapsedNumber(secTotal, token))
			i += j + 1
			continue
		}
		// remainder minutes / seconds after an elapsed hour or minute
		run := 1
		for i+run < len(code) && code[i+run] == ch {
			run++
		}
		switch ch {
		case 'h', 'H':
			h := int(math.Floor(math.Abs(secTotal) / 3600))
			if run >= 2 {
				b.WriteString(pad(h, run))
			} else {
				b.WriteString(strconv.Itoa(h))
			}
		case 'm', 'M':
			if strings.Contains(strings.ToLower(code), "[h") {
				m := int(math.Floor(math.Abs(secTotal)/60)) % 60
				if run >= 2 {
					b.WriteString(pad(m, 2))
				} else {
					b.WriteString(strconv.Itoa(m))
				}
			} else {
				b.WriteString(strings.Repeat(string(ch), run))
			}
		case 's', 'S':
			s := math.Abs(secTotal)
			if strings.Contains(strings.ToLower(code), "[h") || strings.Contains(strings.ToLower(code), "[m") {
				s = math.Mod(s, 60)
			}
			whole := int(math.Floor(s + 1e-9))
			if run >= 2 {
				b.WriteString(pad(whole, 2))
			} else {
				b.WriteString(strconv.Itoa(whole))
			}
		case '.':
			if i+1 < len(code) && (code[i+1] == '0' || code[i+1] == '#') {
				k := 1
				for i+k < len(code) && (code[i+k] == '0' || code[i+k] == '#') {
					k++
				}
				digits := k - 1
				frac := math.Abs(secTotal) - math.Floor(math.Abs(secTotal))
				if strings.Contains(strings.ToLower(code), "[h") || strings.Contains(strings.ToLower(code), "[m") {
					frac = math.Mod(math.Abs(secTotal), 60)
					frac = frac - math.Floor(frac)
				}
				b.WriteByte('.')
				n := int(math.Round(frac * math.Pow(10, float64(digits))))
				b.WriteString(pad(n, digits))
				i += k
				continue
			}
			b.WriteByte(ch)
			run = 1
		default:
			b.WriteByte(ch)
			run = 1
		}
		i += run
	}
	return b.String()
}

func elapsedNumber(sec float64, token string) string {
	token = strings.TrimSpace(token)
	padWidth := 0
	if len(token) >= 2 {
		padWidth = len(token)
	}
	var n float64
	switch token[:1] {
	case "h":
		n = math.Floor(sec / 3600)
	case "m":
		if strings.Contains(token, "h") {
			n = math.Floor(sec / 60)
		}
		n = math.Floor(sec / 60)
	case "s":
		n = math.Floor(sec)
	default:
		return ""
	}
	if n < 0 {
		n = -n
	}
	s := strconv.FormatFloat(n, 'f', 0, 64)
	if padWidth > 1 {
		for len(s) < padWidth {
			s = "0" + s
		}
	}
	return s
}

func pad(n, w int) string {
	if n < 0 {
		n = -n
	}
	s := strconv.Itoa(n)
	for len(s) < w {
		s = "0" + s
	}
	return s
}

func splitFormatSections(code string) []string {
	var sections []string
	var b strings.Builder
	inQ := false
	for i := 0; i < len(code); i++ {
		ch := code[i]
		if ch == '"' {
			inQ = !inQ
			b.WriteByte(ch)
			continue
		}
		if !inQ && ch == '\\' && i+1 < len(code) {
			b.WriteByte(ch)
			i++
			b.WriteByte(code[i])
			continue
		}
		if !inQ && ch == ';' {
			sections = append(sections, b.String())
			b.Reset()
			continue
		}
		b.WriteByte(ch)
	}
	sections = append(sections, b.String())
	return sections
}
