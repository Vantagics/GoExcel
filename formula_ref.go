package gospreadsheet

import (
	"strings"
	"unicode"
)

// shiftFormula offsets relative A1 references by dCol/dRow.
// Absolute markers ($A, $1) stay put. Text inside double quotes is unchanged.
// A reference that would fall outside the worksheet grid becomes #REF!.
func shiftFormula(formula string, dCol, dRow int) string {
	if formula == "" || (dCol == 0 && dRow == 0) {
		return formula
	}
	var b strings.Builder
	b.Grow(len(formula))
	i := 0
	for i < len(formula) {
		switch formula[i] {
		case '"':
			j := i + 1
			for j < len(formula) {
				if formula[j] == '"' {
					if j+1 < len(formula) && formula[j+1] == '"' {
						j += 2
						continue
					}
					j++
					break
				}
				j++
			}
			b.WriteString(formula[i:j])
			i = j
		case '[':
			j := strings.IndexByte(formula[i:], ']')
			if j < 0 {
				b.WriteByte(formula[i])
				i++
				continue
			}
			b.WriteString(formula[i : i+j+1])
			i += j + 1
		default:
			if ref, n, ok := parseFormulaRef(formula, i); ok {
				b.WriteString(shiftRef(ref, dCol, dRow))
				i += n
				continue
			}
			b.WriteByte(formula[i])
			i++
		}
	}
	return b.String()
}

type formulaRef struct {
	colAbs bool
	col    string
	rowAbs bool
	row    int
}

func parseFormulaRef(s string, i int) (formulaRef, int, bool) {
	if i > 0 {
		prev := rune(s[i-1])
		if unicode.IsLetter(prev) || unicode.IsDigit(prev) || prev == '_' || prev == '.' || prev == '$' || prev == '!' && false {
			// '$' is consumed as part of a ref starting at the dollar.
			if unicode.IsLetter(prev) || unicode.IsDigit(prev) || prev == '_' || prev == '.' {
				return formulaRef{}, 0, false
			}
		}
	}
	start := i
	colAbs := false
	if s[i] == '$' {
		colAbs = true
		i++
		if i >= len(s) {
			return formulaRef{}, 0, false
		}
	}
	if !isASCIILetter(s[i]) {
		return formulaRef{}, 0, false
	}
	colStart := i
	for i < len(s) && isASCIILetter(s[i]) {
		i++
	}
	colLen := i - colStart
	if colLen < 1 || colLen > 3 {
		return formulaRef{}, 0, false
	}
	rowAbs := false
	if i < len(s) && s[i] == '$' {
		rowAbs = true
		i++
	}
	if i >= len(s) || !isASCIIDigit(s[i]) {
		return formulaRef{}, 0, false
	}
	row := 0
	for i < len(s) && isASCIIDigit(s[i]) {
		row = row*10 + int(s[i]-'0')
		if row > 1048576 {
			return formulaRef{}, 0, false
		}
		i++
	}
	if row < 1 {
		return formulaRef{}, 0, false
	}
	if i < len(s) {
		n := s[i]
		if isASCIILetter(n) || isASCIIDigit(n) || n == '_' || n == '.' || n == '(' {
			return formulaRef{}, 0, false
		}
	}
	return formulaRef{
		colAbs: colAbs,
		col:    strings.ToUpper(s[colStart : colStart+colLen]),
		rowAbs: rowAbs,
		row:    row,
	}, i - start, true
}

func shiftRef(ref formulaRef, dCol, dRow int) string {
	col := ref.col
	row := ref.row
	if !ref.colAbs && dCol != 0 {
		idx, err := ColumnNameToIndex(col)
		if err != nil {
			return "#REF!"
		}
		idx += dCol
		if idx < 0 || idx > 16383 {
			return "#REF!"
		}
		name, err := ColumnIndexToName(idx)
		if err != nil {
			return "#REF!"
		}
		col = name
	}
	if !ref.rowAbs && dRow != 0 {
		row += dRow
		if row < 1 || row > 1048576 {
			return "#REF!"
		}
	}
	var b strings.Builder
	if ref.colAbs {
		b.WriteByte('$')
	}
	b.WriteString(col)
	if ref.rowAbs {
		b.WriteByte('$')
	}
	b.WriteString(itoa(row))
	return b.String()
}

func isASCIILetter(b byte) bool { return b >= 'A' && b <= 'Z' || b >= 'a' && b <= 'z' }
func isASCIIDigit(b byte) bool  { return b >= '0' && b <= '9' }

func itoa(n int) string {
	if n == 0 {
		return "0"
	}
	var buf [12]byte
	i := len(buf)
	for n > 0 {
		i--
		buf[i] = byte('0' + n%10)
		n /= 10
	}
	return string(buf[i:])
}
