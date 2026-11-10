package gospreadsheet

import (
	"errors"
	"fmt"
)

// DocumentProperties holds metadata about the spreadsheet.
type DocumentProperties struct {
	Creator        string
	LastModifiedBy string
	Title          string
	Subject        string
	Description    string
	Keywords       string
	Category       string
}

// definedName is a workbook defined name, including sheet-scoped names.
type definedName struct {
	Name         string
	Ref          string
	LocalSheetID int
	HasLocal     bool
	Hidden       bool
}

// Workbook represents a spreadsheet workbook containing one or more worksheets.
type Workbook struct {
	worksheets         []*Worksheet
	activeSheet        int
	Properties         DocumentProperties
	namedRanges        map[string]string // name -> "Sheet1!A1:B2"
	definedNameList    []definedName
	workbookProtection *WorkbookProtection
	date1904           bool
	writePassword      string
	strictOOXML        bool
	loadedStrict       bool
	slicerCaches       []*SlicerCache

	// source is the package this workbook was loaded from, used to preserve
	// parts the model does not edit (theme, drawings, charts, media, pivots).
	source         *ooxmlPackage
	extraRootAttrs string
	preserved      []rawXML
	styleExtras    styleExtras
	workbookRels   []relationship
}

// New creates a new workbook with a single default worksheet.
func New() *Workbook {
	wb := &Workbook{
		worksheets:  make([]*Worksheet, 0),
		namedRanges: make(map[string]string),
	}
	ws := NewWorksheet("Sheet1")
	ws.parent = wb
	ws.index = 0
	wb.worksheets = append(wb.worksheets, ws)
	return wb
}

// NewEmpty creates a new workbook with no worksheets.
func NewEmpty() *Workbook {
	return &Workbook{
		worksheets:  make([]*Worksheet, 0),
		namedRanges: make(map[string]string),
	}
}

// AddSheet adds a new worksheet with the given title.
func (wb *Workbook) AddSheet(title string) (*Worksheet, error) {
	for _, ws := range wb.worksheets {
		if ws.title == title {
			return nil, fmt.Errorf("worksheet with title %q already exists", title)
		}
	}
	ws := NewWorksheet(title)
	ws.parent = wb
	ws.index = len(wb.worksheets)
	wb.worksheets = append(wb.worksheets, ws)
	return ws, nil
}

// GetSheet returns the worksheet at the given index (0-based).
func (wb *Workbook) GetSheet(index int) (*Worksheet, error) {
	if index < 0 || index >= len(wb.worksheets) {
		return nil, fmt.Errorf("sheet index %d out of range (0-%d)", index, len(wb.worksheets)-1)
	}
	return wb.worksheets[index], nil
}

// GetSheetByName returns the worksheet with the given title.
func (wb *Workbook) GetSheetByName(title string) (*Worksheet, error) {
	for _, ws := range wb.worksheets {
		if ws.title == title {
			return ws, nil
		}
	}
	return nil, fmt.Errorf("worksheet %q not found", title)
}

// SheetCount returns the number of worksheets.
func (wb *Workbook) SheetCount() int {
	return len(wb.worksheets)
}

// GetSheetNames returns the titles of all worksheets.
func (wb *Workbook) GetSheetNames() []string {
	names := make([]string, len(wb.worksheets))
	for i, ws := range wb.worksheets {
		names[i] = ws.title
	}
	return names
}

// RemoveSheet removes the worksheet at the given index.
func (wb *Workbook) RemoveSheet(index int) error {
	if index < 0 || index >= len(wb.worksheets) {
		return fmt.Errorf("sheet index %d out of range", index)
	}
	if len(wb.worksheets) <= 1 {
		return errors.New("cannot remove the last worksheet")
	}
	wb.worksheets = append(wb.worksheets[:index], wb.worksheets[index+1:]...)
	// Re-index
	for i, ws := range wb.worksheets {
		ws.index = i
	}
	if wb.activeSheet >= len(wb.worksheets) {
		wb.activeSheet = len(wb.worksheets) - 1
	}
	return nil
}

// GetActiveSheet returns the active worksheet.
func (wb *Workbook) GetActiveSheet() *Worksheet {
	if len(wb.worksheets) == 0 {
		return nil
	}
	return wb.worksheets[wb.activeSheet]
}

// SetActiveSheet sets the active worksheet by index.
func (wb *Workbook) SetActiveSheet(index int) error {
	if index < 0 || index >= len(wb.worksheets) {
		return fmt.Errorf("sheet index %d out of range", index)
	}
	wb.activeSheet = index
	return nil
}

// SetDate1904 selects Excel's 1904 date system (used by workbooks created on
// older Mac Excel). The zero value is the 1900 date system.
func (wb *Workbook) SetDate1904(v bool) { wb.date1904 = v }

// IsDate1904 reports whether the workbook uses the 1904 date system.
func (wb *Workbook) IsDate1904() bool { return wb.date1904 }

// SetPassword encrypts the workbook with ECMA-376 Agile encryption
// (AES-256, SHA-512) when it is saved as xlsx/xlsm/xltx/xltm.
// An empty password writes a normal OOXML package.
func (wb *Workbook) SetPassword(password string) { wb.writePassword = password }

// SetStrictOOXML selects strict ECMA-376 namespaces and conformance="strict"
// when the workbook is saved. The default remains transitional OOXML.
func (wb *Workbook) SetStrictOOXML(strict bool) { wb.strictOOXML = strict }

// StrictOOXML reports whether the next save will use strict OOXML.
func (wb *Workbook) StrictOOXML() bool { return wb.strictOOXML }

// LoadedStrict reports whether the workbook was read from a strict OOXML package.
// Saving still uses transitional OOXML unless SetStrictOOXML(true) is called.
func (wb *Workbook) LoadedStrict() bool { return wb.loadedStrict }

// AddNamedRange adds a named range to the workbook.
func (wb *Workbook) AddNamedRange(name, reference string) {
	wb.namedRanges[name] = reference
	for i := range wb.definedNameList {
		if wb.definedNameList[i].Name == name && !wb.definedNameList[i].HasLocal {
			wb.definedNameList[i].Ref = reference
			return
		}
	}
	wb.definedNameList = append(wb.definedNameList, definedName{Name: name, Ref: reference})
}

// GetNamedRange returns the reference for a named range.
func (wb *Workbook) GetNamedRange(name string) (string, error) {
	ref, ok := wb.namedRanges[name]
	if !ok {
		return "", fmt.Errorf("named range %q not found", name)
	}
	return ref, nil
}

// GetNamedRanges returns all named ranges.
func (wb *Workbook) GetNamedRanges() map[string]string {
	result := make(map[string]string, len(wb.namedRanges))
	for k, v := range wb.namedRanges {
		result[k] = v
	}
	return result
}

// RemoveNamedRange removes a named range by name.
func (wb *Workbook) RemoveNamedRange(name string) error {
	if _, ok := wb.namedRanges[name]; !ok {
		return fmt.Errorf("named range %q not found", name)
	}
	delete(wb.namedRanges, name)
	kept := wb.definedNameList[:0]
	for _, dn := range wb.definedNameList {
		if dn.Name == name && !dn.HasLocal {
			continue
		}
		kept = append(kept, dn)
	}
	wb.definedNameList = kept
	return nil
}

// Sheets returns all worksheets.
func (wb *Workbook) Sheets() []*Worksheet {
	result := make([]*Worksheet, len(wb.worksheets))
	copy(result, wb.worksheets)
	return result
}

// SetWorkbookProtection sets the workbook protection.
func (wb *Workbook) SetWorkbookProtection(wp *WorkbookProtection) {
	wb.workbookProtection = wp
}

// GetWorkbookProtection returns the workbook protection, or nil if not set.
func (wb *Workbook) GetWorkbookProtection() *WorkbookProtection {
	return wb.workbookProtection
}

// ClearProtection removes all workbook protection.
func (wb *Workbook) ClearProtection() {
	wb.workbookProtection = nil
}

// Validate checks the workbook for common issues.
func (wb *Workbook) Validate() error {
	if len(wb.worksheets) == 0 {
		return errors.New("workbook must contain at least one worksheet")
	}
	// Check for duplicate sheet names
	names := make(map[string]bool)
	for _, ws := range wb.worksheets {
		if names[ws.title] {
			return fmt.Errorf("duplicate worksheet name: %q", ws.title)
		}
		names[ws.title] = true
		if err := ws.Validate(); err != nil {
			return fmt.Errorf("worksheet %q: %w", ws.title, err)
		}
	}
	return nil
}

// Close is a no-op for in-memory workbooks but provided for API compatibility.
// It can be used to signal that the workbook is no longer needed.
func (wb *Workbook) Close() error {
	return nil
}

// CopySheet copies a worksheet to a new sheet with the given name.
func (wb *Workbook) CopySheet(srcIndex int, newTitle string) (*Worksheet, error) {
	src, err := wb.GetSheet(srcIndex)
	if err != nil {
		return nil, err
	}
	dst, err := wb.AddSheet(newTitle)
	if err != nil {
		return nil, err
	}
	// Copy cells
	for key, cell := range src.cells {
		newCell := &Cell{
			Value:       cell.Value,
			Type:        cell.Type,
			Formula:     cell.Formula,
			row:         cell.row,
			col:         cell.col,
			formulaRef:  cell.formulaRef,
			arrayAnchor: cell.arrayAnchor,
			sheet:       dst,
		}
		if cell.Style != nil {
			styleCopy := *cell.Style
			newCell.Style = &styleCopy
		}
		if cell.Hyperlink != nil {
			hlCopy := *cell.Hyperlink
			newCell.Hyperlink = &hlCopy
		}
		if cell.Comment != nil {
			cmCopy := *cell.Comment
			newCell.Comment = &cmCopy
		}
		dst.cells[key] = newCell
	}
	// Copy merge cells
	dst.mergeCells = make([]MergeCell, len(src.mergeCells))
	copy(dst.mergeCells, src.mergeCells)
	// Copy dimensions
	for k, v := range src.colWidths {
		dst.colWidths[k] = v
	}
	for k, v := range src.rowHeights {
		dst.rowHeights[k] = v
	}
	for k, v := range src.hiddenRows {
		dst.hiddenRows[k] = v
	}
	for k, v := range src.hiddenCols {
		dst.hiddenCols[k] = v
	}
	// Copy frozen pane
	if src.frozen != nil {
		frozenCopy := *src.frozen
		dst.frozen = &frozenCopy
	}
	dst.tabColor = src.tabColor
	dst.state = src.state
	dst.rightToLeft = src.rightToLeft
	dst.defaultRowHeight = src.defaultRowHeight
	if len(src.colSpans) > 0 {
		dst.colSpans = append([]colSpan(nil), src.colSpans...)
	}
	return dst, nil
}

// MoveSheet moves a worksheet from one index to another.
func (wb *Workbook) MoveSheet(fromIndex, toIndex int) error {
	if fromIndex < 0 || fromIndex >= len(wb.worksheets) {
		return fmt.Errorf("source index %d out of range", fromIndex)
	}
	if toIndex < 0 || toIndex >= len(wb.worksheets) {
		return fmt.Errorf("destination index %d out of range", toIndex)
	}
	if fromIndex == toIndex {
		return nil
	}
	ws := wb.worksheets[fromIndex]
	// Remove from old position
	wb.worksheets = append(wb.worksheets[:fromIndex], wb.worksheets[fromIndex+1:]...)
	// Insert at new position
	rear := make([]*Worksheet, len(wb.worksheets[toIndex:]))
	copy(rear, wb.worksheets[toIndex:])
	wb.worksheets = append(wb.worksheets[:toIndex], ws)
	wb.worksheets = append(wb.worksheets, rear...)
	// Re-index
	for i, w := range wb.worksheets {
		w.index = i
	}
	return nil
}
