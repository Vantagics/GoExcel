package gospreadsheet

import (
	"math"
	"strconv"
	"strings"
	"time"
)

// formatByExcelCode formats a value with an Excel TEXT format code.
func formatByExcelCode(v interface{}, code string) string {
	code = strings.TrimSpace(code)
	if code == "" || strings.EqualFold(code, "General") {
		return excelText(v)
	}
	sections := splitFormatSections(code)
	if s, ok := v.(string); ok && !isExcelError(s) {
		if len(sections) >= 4 {
			return applyLiterals(sections[3], s)
		}
		return s
	}
	n, ok := asNumber(v, true)
	if !ok {
		if t, ok := v.(time.Time); ok {
			n = dateToSerial(t)
		} else {
			return excelText(v)
		}
	}
	sec := ""
	if len(sections) > 0 {
		sec = sections[0]
	}
	if n < 0 && len(sections) > 1 && strings.TrimSpace(sections[1]) != "" {
		sec = sections[1]
		n = -n
	} else if n == 0 && len(sections) > 2 && strings.TrimSpace(sections[2]) != "" && !isDateFormat(sections[0]) {
		sec = sections[2]
	}
	if isDateFormat(sec) {
		return formatExcelDate(n, sec)
	}
	if n < 0 && (len(sections) < 2 || strings.TrimSpace(sections[1]) == "") {
		return "-" + formatExcelNumber(-n, sec)
	}
	return formatExcelNumber(n, sec)
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

func isDateFormat(code string) bool {
	inQ := false
	for i := 0; i < len(code); i++ {
		ch := code[i]
		if ch == '"' {
			inQ = !inQ
			continue
		}
		if inQ || ch == '\\' {
			if ch == '\\' && i+1 < len(code) {
				i++
			}
			continue
		}
		if ch == '[' {
			for i < len(code) && code[i] != ']' {
				i++
			}
			continue
		}
		switch ch {
		case 'y', 'Y', 'd', 'D', 'h', 'H', 's', 'S':
			return true
		case 'a', 'A':
			rest := strings.ToUpper(code[i:])
			if strings.HasPrefix(rest, "AM/PM") || strings.HasPrefix(rest, "A/P") {
				return true
			}
		}
	}
	return false
}

func formatExcelNumber(n float64, code string) string {
	percent := 0
	var prefix, suffix strings.Builder
	var pattern strings.Builder
	inQ := false
	sci := false
	for i := 0; i < len(code); i++ {
		ch := code[i]
		if ch == '"' {
			inQ = !inQ
			continue
		}
		if inQ {
			if pattern.Len() == 0 {
				prefix.WriteByte(ch)
			} else {
				suffix.WriteByte(ch)
			}
			continue
		}
		if ch == '\\' && i+1 < len(code) {
			i++
			if pattern.Len() == 0 {
				prefix.WriteByte(code[i])
			} else {
				suffix.WriteByte(code[i])
			}
			continue
		}
		if ch == '[' {
			for i < len(code) && code[i] != ']' {
				i++
			}
			continue
		}
		if ch == '%' {
			percent++
			if pattern.Len() == 0 {
				prefix.WriteByte('%')
			} else {
				suffix.WriteByte('%')
			}
			continue
		}
		if ch == 'E' || ch == 'e' {
			sci = true
			pattern.WriteByte(ch)
			continue
		}
		if strings.ContainsRune("0#?.,", rune(ch)) {
			pattern.WriteByte(ch)
			continue
		}
		if pattern.Len() == 0 {
			prefix.WriteByte(ch)
		} else {
			suffix.WriteByte(ch)
		}
	}
	n *= math.Pow(100, float64(percent))
	body := pattern.String()
	if body == "" {
		return prefix.String() + formatNumber(n) + suffix.String()
	}
	if sci {
		return prefix.String() + formatScientific(n, body) + suffix.String()
	}
	thousands := strings.Contains(body, ",")
	intPat, fracPat := body, ""
	if i := strings.IndexByte(body, '.'); i >= 0 {
		intPat = strings.ReplaceAll(body[:i], ",", "")
		fracPat = body[i+1:]
	} else {
		intPat = strings.ReplaceAll(intPat, ",", "")
	}
	fracDigits := 0
	for _, c := range fracPat {
		if c == '0' || c == '#' || c == '?' {
			fracDigits++
		}
	}
	pow := math.Pow(10, float64(fracDigits))
	rounded := math.Round(n*pow) / pow
	neg := rounded < 0
	if neg {
		rounded = -rounded
	}
	intPart := math.Floor(rounded + 1e-10)
	frac := rounded - intPart
	intStr := strconv.FormatFloat(intPart, 'f', 0, 64)
	if thousands {
		intStr = withCommas(intStr)
	}
	force := strings.ContainsRune(intPat, '0')
	if intPart == 0 && !force && !strings.ContainsRune(intPat, '#') && intPat != "" {
		intStr = ""
	}
	if intPart == 0 && !force && strings.Trim(intPat, "#,?") == "" && !strings.ContainsRune(intPat, '0') {
		if strings.ContainsRune(intPat, '#') {
			intStr = ""
		}
	}
	out := intStr
	if fracDigits > 0 {
		fracStr := strconv.FormatFloat(math.Round(frac*pow), 'f', 0, 64)
		for len(fracStr) < fracDigits {
			fracStr = "0" + fracStr
		}
		if len(fracStr) > fracDigits {
			fracStr = fracStr[:fracDigits]
		}
		// Drop optional trailing zeros when the pattern uses #.
		end := len(fracStr)
		for end > 0 && end <= len(fracPat) && fracPat[end-1] == '#' && fracStr[end-1] == '0' {
			end--
		}
		// Keep zeros required by '0'.
		minKeep := 0
		for i, c := range fracPat {
			if c == '0' && i+1 > minKeep {
				minKeep = i + 1
			}
		}
		if end < minKeep {
			end = minKeep
		}
		if end > 0 {
			out += "." + fracStr[:end]
		}
	}
	if out == "" || out == "." {
		out = "0"
	}
	if neg {
		out = "-" + out
	}
	return prefix.String() + out + suffix.String()
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

func formatScientific(n float64, pattern string) string {
	exp := 0
	sign := ""
	if n < 0 {
		sign = "-"
		n = -n
	}
	if n != 0 {
		exp = int(math.Floor(math.Log10(n)))
		n = n / math.Pow(10, float64(exp))
	}
	mant := pattern
	if i := strings.IndexAny(pattern, "Ee"); i >= 0 {
		mant = pattern[:i]
	}
	return sign + formatExcelNumber(n, mant) + "E" + strconv.Itoa(exp)
}

func formatExcelDate(serial float64, code string) string {
	t := excelSerialToTime(serial)
	if serial == 0 {
		t = time.Date(1900, 1, 0, 0, 0, 0, 0, time.UTC)
	}
	frac := serial - math.Floor(serial)
	if frac < 0 {
		frac = 0
	}
	hour := int(math.Floor(frac*24 + 1e-9))
	minute := int(math.Floor(frac*24*60+1e-9)) % 60
	second := int(math.Floor(frac*24*3600+1e-9)) % 60
	ms := int(math.Round((frac*24*3600 - math.Floor(frac*24*3600)) * 1000))
	ampm := strings.Contains(strings.ToUpper(code), "AM/PM") || strings.Contains(strings.ToUpper(code), "A/P")
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
			token := code[i+1 : i+j]
			b.WriteString(elapsedToken(serial, token))
			i += j + 1
			continue
		}
		upper := strings.ToUpper(code[i:])
		if strings.HasPrefix(upper, "AM/PM") {
			if hour < 12 {
				b.WriteString("AM")
			} else {
				b.WriteString("PM")
			}
			i += 5
			continue
		}
		if strings.HasPrefix(upper, "A/P") {
			if hour < 12 {
				b.WriteString("A")
			} else {
				b.WriteString("P")
			}
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
			if run == 2 {
				b.WriteString(pad(y%100, 2))
			} else {
				b.WriteString(strconv.Itoa(y))
			}
		case 'm', 'M':
			if minuteContext(code, i) {
				if run >= 2 {
					b.WriteString(pad(minute, 2))
				} else {
					b.WriteString(strconv.Itoa(minute))
				}
			} else if run >= 4 {
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
				}
				i += k
				continue
			}
			b.WriteByte(ch)
		default:
			b.WriteByte(ch)
			run = 1
		}
		i += run
	}
	return b.String()
}

func minuteContext(code string, i int) bool {
	prev := lastDateLetter(code, i, -1)
	next := lastDateLetter(code, i+1, 1)
	return prev == 'h' || prev == 'H' || next == 's' || next == 'S'
}

func lastDateLetter(code string, i, dir int) byte {
	for i >= 0 && i < len(code) {
		ch := code[i]
		if ch == 'm' || ch == 'M' || ch == 'h' || ch == 'H' || ch == 's' || ch == 'S' || ch == 'd' || ch == 'D' || ch == 'y' || ch == 'Y' {
			return ch
		}
		if (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z') {
			return 0
		}
		i += dir
	}
	return 0
}

func elapsedToken(serial float64, token string) string {
	sec := serial * 86400
	switch strings.ToLower(strings.TrimSpace(token)) {
	case "h", "hh":
		return strconv.FormatInt(int64(math.Floor(sec/3600)), 10)
	case "m", "mm":
		return strconv.FormatInt(int64(math.Floor(sec/60)), 10)
	case "s", "ss":
		return strconv.FormatInt(int64(math.Floor(sec)), 10)
	default:
		return ""
	}
}

func pad(n, w int) string {
	s := strconv.Itoa(n)
	for len(s) < w {
		s = "0" + s
	}
	return s
}

func applyLiterals(code, text string) string {
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
		if ch == '@' {
			b.WriteString(text)
			continue
		}
		if ch == '[' {
			for i < len(code) && code[i] != ']' {
				i++
			}
			continue
		}
		b.WriteByte(ch)
	}
	if b.Len() == 0 {
		return text
	}
	return b.String()
}
