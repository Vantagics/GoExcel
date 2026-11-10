package gospreadsheet

import (
	"fmt"
	"strings"
)

// ArrayFormula is an Excel array formula. The file stores the expression on
// the anchor (the top-left cell of Ref); every cell in Ref reports that
// same formula, which is what the formula bar shows in Excel.
type ArrayFormula struct {
	Ref     string
	Formula string
}

// ArrayRef returns the range this array formula covers.
// Member cells report the anchor's range even though the package stores the
// formula only on the anchor.
func (c *Cell) ArrayRef() string {
	if c == nil {
		return ""
	}
	if c.arrayAnchor && c.formulaRef != "" {
		return c.formulaRef
	}
	if c.sheet != nil {
		if af, ok := c.sheet.arrayAt(c.row, c.col); ok {
			return af.Ref
		}
	}
	return ""
}

// IsArrayFormula reports whether this cell is an array-formula anchor or a
// cell covered by an array formula's ref.
func (c *Cell) IsArrayFormula() bool {
	if c == nil {
		return false
	}
	if c.arrayAnchor {
		return true
	}
	if c.sheet == nil {
		return false
	}
	_, ok := c.sheet.arrayAt(c.row, c.col)
	return ok
}

// ReportedFormula returns the formula Excel shows for this cell.
// Array-formula members return the anchor expression.
func (c *Cell) ReportedFormula() string {
	if c == nil {
		return ""
	}
	if c.Formula != "" {
		return c.Formula
	}
	if c.sheet == nil {
		return ""
	}
	if af, ok := c.sheet.arrayAt(c.row, c.col); ok {
		return af.Formula
	}
	return ""
}

// SetArrayFormula stores an array formula on the anchor of rangeRef and
// records the full ref. rangeRef is "A1:B2" or a single cell. formula is the
// expression with or without braces. Cells in the range report the formula
// through ReportedFormula and FormulaAt; the package still writes it once,
// on the anchor, with t="array".
func (ws *Worksheet) SetArrayFormula(rangeRef, formula string) error {
	start, end, ref, err := normalizeFormulaRef(rangeRef)
	if err != nil {
		return err
	}
	formula = strings.TrimSpace(formula)
	formula = strings.TrimPrefix(formula, "=")
	cell := ws.GetCell(start.Row-1, start.ColumnIdx)
	inner := formula
	if strings.HasPrefix(inner, "{") && strings.HasSuffix(inner, "}") && len(inner) >= 2 {
		inner = inner[1 : len(inner)-1]
	}
	cell.SetFormulaArray(inner)
	cell.formulaRef = ref
	_ = end
	return nil
}

// ArrayFormulas returns array formulas anchored on this sheet, in cell order.
func (ws *Worksheet) ArrayFormulas() []ArrayFormula {
	var out []ArrayFormula
	for _, c := range ws.AllCells() {
		if c.arrayAnchor && c.Formula != "" && c.formulaRef != "" {
			out = append(out, ArrayFormula{Ref: c.formulaRef, Formula: c.Formula})
		}
	}
	return out
}

// FormulaAt returns the formula Excel would show for ref, and the array ref
// when that formula is an array formula. A cell with no formula returns
// "", "", nil.
func (ws *Worksheet) FormulaAt(ref string) (formula, arrayRef string, err error) {
	cr, err := ParseCellReference(stripSheetBang(ref))
	if err != nil {
		return "", "", err
	}
	row, col := cr.Row-1, cr.ColumnIdx
	if c := ws.GetCellIfExists(row, col); c != nil && c.Formula != "" {
		if c.arrayAnchor {
			return c.Formula, c.formulaRef, nil
		}
		return c.Formula, "", nil
	}
	if af, ok := ws.arrayAt(row, col); ok {
		return af.Formula, af.Ref, nil
	}
	if c := ws.GetCellIfExists(row, col); c != nil {
		return c.Formula, "", nil
	}
	return "", "", nil
}

func (ws *Worksheet) arrayAt(row, col int) (ArrayFormula, bool) {
	for _, c := range ws.cells {
		if !c.arrayAnchor || c.formulaRef == "" || c.Formula == "" {
			continue
		}
		if refContains(c.formulaRef, row, col) {
			return ArrayFormula{Ref: c.formulaRef, Formula: c.Formula}, true
		}
	}
	return ArrayFormula{}, false
}

func normalizeFormulaRef(rangeRef string) (start, end *CellReference, ref string, err error) {
	rangeRef = strings.TrimSpace(stripSheetBang(rangeRef))
	rangeRef = stripDollars(rangeRef)
	if rangeRef == "" {
		return nil, nil, "", fmt.Errorf("invalid range: %s", rangeRef)
	}
	if strings.Contains(rangeRef, ":") {
		start, end, err = ParseRange(rangeRef)
		if err != nil {
			return nil, nil, "", err
		}
		a, _ := CellName(start.Row-1, start.ColumnIdx)
		b, _ := CellName(end.Row-1, end.ColumnIdx)
		return start, end, a + ":" + b, nil
	}
	start, err = ParseCellReference(rangeRef)
	if err != nil {
		return nil, nil, "", err
	}
	name, _ := CellName(start.Row-1, start.ColumnIdx)
	return start, start, name, nil
}

func refContains(ref string, row, col int) bool {
	ref = strings.TrimSpace(ref)
	for _, part := range strings.Fields(ref) {
		part = stripDollars(stripSheetBang(part))
		if part == "" {
			continue
		}
		if strings.Contains(part, ":") {
			a, b, err := ParseRange(part)
			if err != nil {
				continue
			}
			r1, r2 := a.Row-1, b.Row-1
			c1, c2 := a.ColumnIdx, b.ColumnIdx
			if r1 > r2 {
				r1, r2 = r2, r1
			}
			if c1 > c2 {
				c1, c2 = c2, c1
			}
			if row >= r1 && row <= r2 && col >= c1 && col <= c2 {
				return true
			}
			continue
		}
		cr, err := ParseCellReference(part)
		if err != nil {
			continue
		}
		if cr.Row-1 == row && cr.ColumnIdx == col {
			return true
		}
	}
	return false
}

func stripSheetBang(ref string) string {
	ref = strings.TrimSpace(ref)
	if i := strings.LastIndex(ref, "!"); i >= 0 {
		ref = ref[i+1:]
	}
	return strings.Trim(ref, "'")
}

func stripDollars(ref string) string {
	return strings.ReplaceAll(ref, "$", "")
}
