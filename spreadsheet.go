// Package gospreadsheet provides a pure Go library for reading and writing
// spreadsheet files, inspired by PHPOffice/PhpSpreadsheet.
// It supports XLSX and CSV formats with an in-memory spreadsheet model.
package gospreadsheet

import (
	"errors"
	"fmt"
	"time"
)

// Version is the semantic version number of the gospreadsheet library.
const Version = "1.0.0"

// CellType represents the data type of a cell value.
type CellType int

const (
	CellTypeEmpty CellType = iota
	CellTypeString
	CellTypeNumeric
	CellTypeBool
	CellTypeFormula
	CellTypeDate
	CellTypeError
)

// Cell represents a single cell in a worksheet.
type Cell struct {
	Value       interface{}
	Type        CellType
	Formula     string
	Style       *Style
	Hyperlink   *Hyperlink
	Comment     *Comment
	RichText    *RichText
	row         int
	col         int
	formulaRef  string // array-formula ref, when this cell is the anchor
	arrayAnchor bool   // this cell stores the array formula for formulaRef
	sheet       *Worksheet
	cachedType  string // cached formula result type: n, str, b, e
	dt          *dataTable
	cellMeta    string // cm attribute: cell metadata index (dynamic arrays)
	valueMeta   string // vm attribute: value metadata index (rich data)
}

// dataTable is an Excel data-table formula (What-If). It is not a normal
// formula expression; the writer emits it as t="dataTable".
type dataTable struct {
	Ref  string
	R1   string
	R2   string
	TwoD bool
	Row  bool
}

// NewCell creates a new empty cell.
func NewCell(row, col int) *Cell {
	return &Cell{
		Type: CellTypeEmpty,
		row:  row,
		col:  col,
	}
}

// SetValue sets the cell value and auto-detects the type.
func (c *Cell) SetValue(v interface{}) *Cell {
	c.cellMeta = ""
	c.valueMeta = ""
	if v == nil {
		c.Value = nil
		c.Type = CellTypeEmpty
		return c
	}
	switch val := v.(type) {
	case string:
		c.Value = val
		c.Type = CellTypeString
	case int:
		c.Value = float64(val)
		c.Type = CellTypeNumeric
	case int32:
		c.Value = float64(val)
		c.Type = CellTypeNumeric
	case int64:
		c.Value = float64(val)
		c.Type = CellTypeNumeric
	case float32:
		c.Value = float64(val)
		c.Type = CellTypeNumeric
	case float64:
		c.Value = val
		c.Type = CellTypeNumeric
	case bool:
		c.Value = val
		c.Type = CellTypeBool
	case time.Time:
		c.Value = val
		c.Type = CellTypeDate
	default:
		c.Value = v
		c.Type = CellTypeString
	}
	return c
}

// SetFormula sets a formula for the cell.
func (c *Cell) SetFormula(formula string) *Cell {
	c.Formula = formula
	c.Type = CellTypeFormula
	c.dt = nil
	c.arrayAnchor = false
	c.formulaRef = ""
	c.cellMeta = ""
	c.valueMeta = ""
	return c
}

// SetStyle sets the style for the cell.
func (c *Cell) SetStyle(s *Style) *Cell {
	c.Style = s
	return c
}

// GetStringValue returns the cell value as a string.
func (c *Cell) GetStringValue() string {
	if c.Value == nil {
		return ""
	}
	switch v := c.Value.(type) {
	case string:
		return v
	case float64:
		if v == float64(int64(v)) {
			return fmt.Sprintf("%d", int64(v))
		}
		return fmt.Sprintf("%g", v)
	case bool:
		if v {
			return "TRUE"
		}
		return "FALSE"
	case time.Time:
		return v.Format("2006-01-02 15:04:05")
	default:
		return fmt.Sprintf("%v", v)
	}
}

// GetNumericValue returns the cell value as float64.
func (c *Cell) GetNumericValue() (float64, error) {
	if c.Type == CellTypeNumeric {
		if v, ok := c.Value.(float64); ok {
			return v, nil
		}
	}
	return 0, errors.New("cell is not numeric")
}

// GetBoolValue returns the cell value as bool.
func (c *Cell) GetBoolValue() (bool, error) {
	if c.Type == CellTypeBool {
		if v, ok := c.Value.(bool); ok {
			return v, nil
		}
	}
	return false, errors.New("cell is not boolean")
}

// GetDateValue returns the cell value as time.Time.
func (c *Cell) GetDateValue() (time.Time, error) {
	if c.Type == CellTypeDate {
		if v, ok := c.Value.(time.Time); ok {
			return v, nil
		}
	}
	return time.Time{}, errors.New("cell is not a date")
}

// Row returns the cell's row index (0-based).
func (c *Cell) Row() int { return c.row }

// Col returns the cell's column index (0-based).
func (c *Cell) Col() int { return c.col }

// SetHyperlink sets a hyperlink on the cell.
func (c *Cell) SetHyperlink(h *Hyperlink) *Cell {
	c.Hyperlink = h
	return c
}

// SetComment sets a comment on the cell.
func (c *Cell) SetComment(comment *Comment) *Cell {
	c.Comment = comment
	return c
}

// SetRichText sets rich text content on the cell.
func (c *Cell) SetRichText(rt *RichText) *Cell {
	c.RichText = rt
	c.Type = CellTypeString
	if rt != nil {
		c.Value = rt.PlainText()
	}
	return c
}

// Clear resets the cell to an empty state, removing value, formula, style, and annotations.
func (c *Cell) Clear() *Cell {
	c.Value = nil
	c.Type = CellTypeEmpty
	c.Formula = ""
	c.Style = nil
	c.Hyperlink = nil
	c.Comment = nil
	c.RichText = nil
	c.formulaRef = ""
	c.arrayAnchor = false
	c.cachedType = ""
	c.dt = nil
	c.cellMeta = ""
	c.valueMeta = ""
	return c
}

// HasFormula returns true if the cell contains a formula.
func (c *Cell) HasFormula() bool {
	return c.Type == CellTypeFormula && c.Formula != ""
}

// IsEmpty returns true if the cell has no value and no formula.
func (c *Cell) IsEmpty() bool {
	return c.Type == CellTypeEmpty && c.Value == nil && c.Formula == ""
}

// IsNumber returns true if the cell contains a numeric value.
func (c *Cell) IsNumber() bool {
	return c.Type == CellTypeNumeric
}

// IsBool returns true if the cell contains a boolean value.
func (c *Cell) IsBool() bool {
	return c.Type == CellTypeBool
}

// IsDate returns true if the cell contains a date value.
func (c *Cell) IsDate() bool {
	return c.Type == CellTypeDate
}

// IsString returns true if the cell contains a string value.
func (c *Cell) IsString() bool {
	return c.Type == CellTypeString
}

// IsError returns true if the cell contains an error value.
func (c *Cell) IsError() bool {
	return c.Type == CellTypeError
}

// SetDateWithStyle sets a date value and applies the default date number format.
func (c *Cell) SetDateWithStyle(t time.Time) *Cell {
	c.SetValue(t)
	if c.Style == nil {
		c.Style = NewStyle()
	}
	c.Style.NumberFormat = &FormatDate
	return c
}

// SetNumberWithFormat sets a numeric value and applies a number format.
func (c *Cell) SetNumberWithFormat(v float64, nf NumberFormat) *Cell {
	c.SetValue(v)
	if c.Style == nil {
		c.Style = NewStyle()
	}
	c.Style.NumberFormat = &nf
	return c
}

// SetFormulaArray sets an array formula (equivalent to Ctrl+Shift+Enter in Excel).
// The formula is stored on this cell. For a multi-cell array, use
// Worksheet.SetArrayFormula so every cell in the range reports the formula.
func (c *Cell) SetFormulaArray(formula string) *Cell {
	c.SetFormula("{" + formula + "}")
	c.arrayAnchor = true
	if c.formulaRef == "" {
		if name, err := CellName(c.row, c.col); err == nil {
			c.formulaRef = name
		}
	}
	return c
}

// SetInlineString sets the cell value as an inline string.
// This is an alternative to shared strings for cells that have unique text.
func (c *Cell) SetInlineString(s string) *Cell {
	c.Value = s
	c.Type = CellTypeString
	return c
}

// GetFormattedValue returns the cell value formatted according to its number format.
// If no number format is set, it returns the default string representation.
func (c *Cell) GetFormattedValue() string {
	if c.Style != nil && c.Style.NumberFormat != nil {
		code := c.Style.NumberFormat.FormatCode
		switch {
		case c.Type == CellTypeDate:
			if t, ok := c.Value.(time.Time); ok {
				return formatDateByCode(t, code)
			}
		case c.Type == CellTypeNumeric:
			if v, ok := c.Value.(float64); ok {
				return formatNumberByCode(v, code)
			}
		}
	}
	return c.GetStringValue()
}

// formatDateByCode formats a time.Time according to a simple format code.
func formatDateByCode(t time.Time, code string) string {
	switch code {
	case "yyyy-mm-dd":
		return t.Format("2006-01-02")
	case "yyyy-mm-dd hh:mm:ss":
		return t.Format("2006-01-02 15:04:05")
	case "hh:mm:ss":
		return t.Format("15:04:05")
	case "mm/dd/yyyy":
		return t.Format("01/02/2006")
	case "dd/mm/yyyy":
		return t.Format("02/01/2006")
	default:
		return t.Format("2006-01-02 15:04:05")
	}
}

// formatNumberByCode formats a float64 according to a simple format code.
func formatNumberByCode(v float64, code string) string {
	switch code {
	case "General":
		if v == float64(int64(v)) {
			return fmt.Sprintf("%d", int64(v))
		}
		return fmt.Sprintf("%g", v)
	case "0":
		return fmt.Sprintf("%.0f", v)
	case "0.00":
		return fmt.Sprintf("%.2f", v)
	case "0%":
		return fmt.Sprintf("%.0f%%", v*100)
	case "0.00%":
		return fmt.Sprintf("%.2f%%", v*100)
	case "@":
		return fmt.Sprintf("%v", v)
	default:
		return fmt.Sprintf("%g", v)
	}
}
