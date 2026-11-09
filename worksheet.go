package gospreadsheet

import (
	"errors"
	"fmt"
	"sort"
	"strconv"
)

// MergeCell represents a merged cell range.
type MergeCell struct {
	StartRow int
	StartCol int
	EndRow   int
	EndCol   int
}

// SortOrder represents the sort direction.
type SortOrder int

const (
	SortOrderAscending SortOrder = iota
	SortOrderDescending
)

// SheetView represents view settings for a worksheet.
type SheetView struct {
	ZoomScale         uint // zoom percentage (10-400), default 100
	ShowGridlines     bool
	ShowRowColHeaders bool
	ShowRuler         bool
	TopLeftCell       string // e.g. "A1"
}

// Worksheet represents a single worksheet in a spreadsheet.
type Worksheet struct {
	title              string
	cells              map[string]*Cell // key: "row,col" (0-based)
	mergeCells         []MergeCell
	colWidths          map[int]float64
	rowHeights         map[int]float64
	frozen             *CellReference // frozen pane position
	parent             *Workbook
	index              int
	conditionalFormats []*ConditionalFormatting
	dataValidations    []*DataValidation
	autoFilter         *AutoFilter
	pageSetup          *PageSetup
	sheetProtection    *SheetProtection
	tabColor           string // hex color for sheet tab
	hiddenRows         map[int]bool
	hiddenCols         map[int]bool
	sheetView          *SheetView
	state              string // visible, hidden, veryHidden
	rightToLeft        bool
	defaultRowHeight   float64
	colSpans           []colSpan
	partName           string
	relID              string
	rawPart            []byte
	rawContentType     string
	extraRootAttrs     string
	sheetPrRaw         string
	loadedTabColor     string
	preserved          []rawXML
	sheetRels          []relationship
	autoFilterRaw      string
}

// NewWorksheet creates a new worksheet with the given title.
func NewWorksheet(title string) *Worksheet {
	return &Worksheet{
		title:      title,
		cells:      make(map[string]*Cell),
		colWidths:  make(map[int]float64),
		rowHeights: make(map[int]float64),
		hiddenRows: make(map[int]bool),
		hiddenCols: make(map[int]bool),
	}
}

// Title returns the worksheet title.
func (ws *Worksheet) Title() string {
	return ws.title
}

// SetTitle sets the worksheet title.
func (ws *Worksheet) SetTitle(title string) *Worksheet {
	ws.title = title
	return ws
}

func cellKey(row, col int) string {
	return strconv.Itoa(row) + "," + strconv.Itoa(col)
}

// GetCell returns the cell at the given 0-based row and column.
// If the cell doesn't exist, it creates a new empty cell.
func (ws *Worksheet) GetCell(row, col int) *Cell {
	key := cellKey(row, col)
	if c, ok := ws.cells[key]; ok {
		return c
	}
	c := NewCell(row, col)
	ws.cells[key] = c
	return c
}

// GetCellIfExists returns the cell at the given 0-based row and column,
// or nil if the cell doesn't exist. Does not create new cells.
func (ws *Worksheet) GetCellIfExists(row, col int) *Cell {
	key := cellKey(row, col)
	return ws.cells[key]
}

// GetCellByName returns the cell at the given reference (e.g., "A1").
func (ws *Worksheet) GetCellByName(ref string) (*Cell, error) {
	cr, err := ParseCellReference(ref)
	if err != nil {
		return nil, err
	}
	return ws.GetCell(cr.Row-1, cr.ColumnIdx), nil
}

// SetCellValue sets the value of a cell by reference (e.g., "A1").
func (ws *Worksheet) SetCellValue(ref string, value interface{}) error {
	cell, err := ws.GetCellByName(ref)
	if err != nil {
		return err
	}
	cell.SetValue(value)
	return nil
}

// SetCellFormula sets a formula for a cell by reference.
func (ws *Worksheet) SetCellFormula(ref string, formula string) error {
	cell, err := ws.GetCellByName(ref)
	if err != nil {
		return err
	}
	cell.SetFormula(formula)
	return nil
}

// SetCellStyle sets the style for a cell by reference.
func (ws *Worksheet) SetCellStyle(ref string, style *Style) error {
	cell, err := ws.GetCellByName(ref)
	if err != nil {
		return err
	}
	cell.SetStyle(style)
	return nil
}

// GetCellValue returns the value of a cell by reference.
func (ws *Worksheet) GetCellValue(ref string) (interface{}, error) {
	cell, err := ws.GetCellByName(ref)
	if err != nil {
		return nil, err
	}
	return cell.Value, nil
}

// MergeCells merges a range of cells (e.g., "A1:C3").
func (ws *Worksheet) MergeCells(rangeStr string) error {
	start, end, err := ParseRange(rangeStr)
	if err != nil {
		return err
	}
	ws.mergeCells = append(ws.mergeCells, MergeCell{
		StartRow: start.Row - 1,
		StartCol: start.ColumnIdx,
		EndRow:   end.Row - 1,
		EndCol:   end.ColumnIdx,
	})
	return nil
}

// GetMergeCells returns all merged cell ranges.
func (ws *Worksheet) GetMergeCells() []MergeCell {
	return ws.mergeCells
}

// SetColumnWidth sets the width of a column (0-based index).
func (ws *Worksheet) SetColumnWidth(col int, width float64) *Worksheet {
	ws.colWidths[col] = width
	return ws
}

// GetColumnWidth returns the width of a column.
func (ws *Worksheet) GetColumnWidth(col int) float64 {
	if w, ok := ws.colWidths[col]; ok {
		return w
	}
	for _, sp := range ws.colSpans {
		if col+1 >= sp.Min && col+1 <= sp.Max && sp.Width != 0 {
			return sp.Width
		}
	}
	return 8.43 // default Excel column width
}

// SetRowHeight sets the height of a row (0-based index).
func (ws *Worksheet) SetRowHeight(row int, height float64) *Worksheet {
	ws.rowHeights[row] = height
	return ws
}

// GetRowHeight returns the height of a row.
func (ws *Worksheet) GetRowHeight(row int) float64 {
	if h, ok := ws.rowHeights[row]; ok {
		return h
	}
	return 15.0 // default Excel row height
}

// FreezePane freezes rows and columns at the given cell reference.
func (ws *Worksheet) FreezePane(ref string) error {
	cr, err := ParseCellReference(ref)
	if err != nil {
		return err
	}
	ws.frozen = cr
	return nil
}

// GetFreezePane returns the freeze pane position, or nil if not set.
func (ws *Worksheet) GetFreezePane() *CellReference {
	return ws.frozen
}

// Dimensions returns the used range of the worksheet as (minRow, minCol, maxRow, maxCol), all 0-based.
// Returns an error if the worksheet is empty.
func (ws *Worksheet) Dimensions() (minRow, minCol, maxRow, maxCol int, err error) {
	if len(ws.cells) == 0 {
		return 0, 0, 0, 0, errors.New("worksheet is empty")
	}
	first := true
	for _, c := range ws.cells {
		if c.Type == CellTypeEmpty && c.Value == nil && c.Formula == "" {
			continue
		}
		if first {
			minRow, minCol = c.row, c.col
			maxRow, maxCol = c.row, c.col
			first = false
			continue
		}
		if c.row < minRow {
			minRow = c.row
		}
		if c.col < minCol {
			minCol = c.col
		}
		if c.row > maxRow {
			maxRow = c.row
		}
		if c.col > maxCol {
			maxCol = c.col
		}
	}
	if first {
		return 0, 0, 0, 0, errors.New("worksheet has no data cells")
	}
	return
}

// RowIterator returns rows in order. Each row is a slice of cells (may contain nil for empty cells).
func (ws *Worksheet) RowIterator() ([][](*Cell), error) {
	minRow, minCol, maxRow, maxCol, err := ws.Dimensions()
	if err != nil {
		return nil, err
	}
	rows := make([][]*Cell, maxRow-minRow+1)
	for i := range rows {
		row := make([]*Cell, maxCol-minCol+1)
		for j := range row {
			key := cellKey(minRow+i, minCol+j)
			if c, ok := ws.cells[key]; ok {
				row[j] = c
			}
		}
		rows[i] = row
	}
	return rows, nil
}

// CellCount returns the number of non-empty cells.
func (ws *Worksheet) CellCount() int {
	count := 0
	for _, c := range ws.cells {
		if c.Type != CellTypeEmpty || c.Value != nil || c.Formula != "" {
			count++
		}
	}
	return count
}

// AllCells returns all non-empty cells sorted by row then column.
func (ws *Worksheet) AllCells() []*Cell {
	cells := make([]*Cell, 0)
	for _, c := range ws.cells {
		if c.Type != CellTypeEmpty || c.Value != nil || c.Formula != "" {
			cells = append(cells, c)
		}
	}
	sort.Slice(cells, func(i, j int) bool {
		if cells[i].row != cells[j].row {
			return cells[i].row < cells[j].row
		}
		return cells[i].col < cells[j].col
	})
	return cells
}

// AddConditionalFormatting adds conditional formatting to the worksheet.
func (ws *Worksheet) AddConditionalFormatting(cf *ConditionalFormatting) *Worksheet {
	ws.conditionalFormats = append(ws.conditionalFormats, cf)
	return ws
}

// GetConditionalFormattings returns all conditional formatting rules.
func (ws *Worksheet) GetConditionalFormattings() []*ConditionalFormatting {
	return ws.conditionalFormats
}

// AddDataValidation adds a data validation rule to the worksheet.
func (ws *Worksheet) AddDataValidation(dv *DataValidation) *Worksheet {
	ws.dataValidations = append(ws.dataValidations, dv)
	return ws
}

// GetDataValidations returns all data validation rules.
func (ws *Worksheet) GetDataValidations() []*DataValidation {
	return ws.dataValidations
}

// SetAutoFilter sets the auto filter for the worksheet.
func (ws *Worksheet) SetAutoFilter(af *AutoFilter) *Worksheet {
	ws.autoFilter = af
	ws.autoFilterRaw = ""
	return ws
}

// GetAutoFilter returns the auto filter, or nil if not set.
func (ws *Worksheet) GetAutoFilter() *AutoFilter {
	return ws.autoFilter
}

// SetPageSetup sets the page setup for the worksheet.
func (ws *Worksheet) SetPageSetup(ps *PageSetup) *Worksheet {
	ws.pageSetup = ps
	return ws
}

// GetPageSetup returns the page setup, creating a default one if not set.
func (ws *Worksheet) GetPageSetup() *PageSetup {
	if ws.pageSetup == nil {
		ws.pageSetup = NewPageSetup()
	}
	return ws.pageSetup
}

// SetSheetProtection sets the sheet protection.
func (ws *Worksheet) SetSheetProtection(sp *SheetProtection) *Worksheet {
	ws.sheetProtection = sp
	return ws
}

// GetSheetProtection returns the sheet protection, or nil if not set.
func (ws *Worksheet) GetSheetProtection() *SheetProtection {
	return ws.sheetProtection
}

// SetTabColor sets the sheet tab color (hex string, e.g., "FF0000" or "FFFF0000").
func (ws *Worksheet) SetTabColor(color string) *Worksheet {
	ws.tabColor = color
	return ws
}

// SetSheetState sets the sheet visibility: "visible", "hidden", or "veryHidden".
func (ws *Worksheet) SetSheetState(state string) *Worksheet {
	switch state {
	case "hidden", "veryHidden":
		ws.state = state
	default:
		ws.state = "visible"
	}
	return ws
}

// GetSheetState returns "visible", "hidden", or "veryHidden".
func (ws *Worksheet) GetSheetState() string {
	if ws.state == "" {
		return "visible"
	}
	return ws.state
}

// GetTabColor returns the sheet tab color.
func (ws *Worksheet) GetTabColor() string {
	return ws.tabColor
}

// SetCellHyperlink sets a hyperlink on a cell by reference.
func (ws *Worksheet) SetCellHyperlink(ref string, url string) error {
	cell, err := ws.GetCellByName(ref)
	if err != nil {
		return err
	}
	cell.SetHyperlink(NewHyperlink(url))
	return nil
}

// SetCellComment sets a comment on a cell by reference.
func (ws *Worksheet) SetCellComment(ref string, author, text string) error {
	cell, err := ws.GetCellByName(ref)
	if err != nil {
		return err
	}
	cell.SetComment(NewComment(author, text))
	return nil
}

// InsertRow inserts a new row at the given 0-based index, shifting existing rows down.
func (ws *Worksheet) InsertRow(rowIdx int) {
	newCells := make(map[string]*Cell)
	for key, cell := range ws.cells {
		if cell.row >= rowIdx {
			cell.row++
			newCells[cellKey(cell.row, cell.col)] = cell
		} else {
			newCells[key] = cell
		}
	}
	ws.cells = newCells

	// Shift merge cells
	for i := range ws.mergeCells {
		if ws.mergeCells[i].StartRow >= rowIdx {
			ws.mergeCells[i].StartRow++
		}
		if ws.mergeCells[i].EndRow >= rowIdx {
			ws.mergeCells[i].EndRow++
		}
	}

	// Shift row heights
	newHeights := make(map[int]float64)
	for r, h := range ws.rowHeights {
		if r >= rowIdx {
			newHeights[r+1] = h
		} else {
			newHeights[r] = h
		}
	}
	ws.rowHeights = newHeights
}

// DeleteRow deletes the row at the given 0-based index, shifting rows up.
func (ws *Worksheet) DeleteRow(rowIdx int) {
	newCells := make(map[string]*Cell)
	for _, cell := range ws.cells {
		if cell.row == rowIdx {
			continue // skip deleted row
		}
		if cell.row > rowIdx {
			cell.row--
		}
		newCells[cellKey(cell.row, cell.col)] = cell
	}
	ws.cells = newCells

	// Shift merge cells
	newMerges := make([]MergeCell, 0)
	for _, mc := range ws.mergeCells {
		if mc.StartRow == rowIdx && mc.EndRow == rowIdx {
			continue // remove merge entirely in deleted row
		}
		if mc.StartRow > rowIdx {
			mc.StartRow--
		}
		if mc.EndRow > rowIdx {
			mc.EndRow--
		}
		newMerges = append(newMerges, mc)
	}
	ws.mergeCells = newMerges

	// Shift row heights
	newHeights := make(map[int]float64)
	for r, h := range ws.rowHeights {
		if r == rowIdx {
			continue
		}
		if r > rowIdx {
			newHeights[r-1] = h
		} else {
			newHeights[r] = h
		}
	}
	ws.rowHeights = newHeights
}

// InsertColumn inserts a new column at the given 0-based index, shifting columns right.
func (ws *Worksheet) InsertColumn(colIdx int) {
	newCells := make(map[string]*Cell)
	for key, cell := range ws.cells {
		if cell.col >= colIdx {
			cell.col++
			newCells[cellKey(cell.row, cell.col)] = cell
		} else {
			newCells[key] = cell
		}
	}
	ws.cells = newCells

	// Shift merge cells
	for i := range ws.mergeCells {
		if ws.mergeCells[i].StartCol >= colIdx {
			ws.mergeCells[i].StartCol++
		}
		if ws.mergeCells[i].EndCol >= colIdx {
			ws.mergeCells[i].EndCol++
		}
	}

	// Shift column widths
	newWidths := make(map[int]float64)
	for c, w := range ws.colWidths {
		if c >= colIdx {
			newWidths[c+1] = w
		} else {
			newWidths[c] = w
		}
	}
	ws.colWidths = newWidths
}

// DeleteColumn deletes the column at the given 0-based index, shifting columns left.
func (ws *Worksheet) DeleteColumn(colIdx int) {
	newCells := make(map[string]*Cell)
	for _, cell := range ws.cells {
		if cell.col == colIdx {
			continue
		}
		if cell.col > colIdx {
			cell.col--
		}
		newCells[cellKey(cell.row, cell.col)] = cell
	}
	ws.cells = newCells

	// Shift merge cells
	newMerges := make([]MergeCell, 0)
	for _, mc := range ws.mergeCells {
		if mc.StartCol == colIdx && mc.EndCol == colIdx {
			continue
		}
		if mc.StartCol > colIdx {
			mc.StartCol--
		}
		if mc.EndCol > colIdx {
			mc.EndCol--
		}
		newMerges = append(newMerges, mc)
	}
	ws.mergeCells = newMerges

	// Shift column widths
	newWidths := make(map[int]float64)
	for c, w := range ws.colWidths {
		if c == colIdx {
			continue
		}
		if c > colIdx {
			newWidths[c-1] = w
		} else {
			newWidths[c] = w
		}
	}
	ws.colWidths = newWidths
}

// CopyRow copies a row to a new position.
func (ws *Worksheet) CopyRow(srcRow, dstRow int) {
	_, minCol, _, maxCol, err := ws.Dimensions()
	if err != nil {
		return
	}
	for col := minCol; col <= maxCol; col++ {
		srcKey := cellKey(srcRow, col)
		if srcCell, ok := ws.cells[srcKey]; ok {
			dstCell := ws.GetCell(dstRow, col)
			dstCell.Value = srcCell.Value
			dstCell.Type = srcCell.Type
			dstCell.Formula = srcCell.Formula
			// Deep copy style to avoid shared mutable state
			if srcCell.Style != nil {
				styleCopy := *srcCell.Style
				if srcCell.Style.Font != nil {
					fontCopy := *srcCell.Style.Font
					styleCopy.Font = &fontCopy
				}
				if srcCell.Style.Fill != nil {
					fillCopy := *srcCell.Style.Fill
					styleCopy.Fill = &fillCopy
				}
				if srcCell.Style.Borders != nil {
					bordersCopy := *srcCell.Style.Borders
					styleCopy.Borders = &bordersCopy
				}
				if srcCell.Style.Alignment != nil {
					alignCopy := *srcCell.Style.Alignment
					styleCopy.Alignment = &alignCopy
				}
				if srcCell.Style.NumberFormat != nil {
					nfCopy := *srcCell.Style.NumberFormat
					styleCopy.NumberFormat = &nfCopy
				}
				dstCell.Style = &styleCopy
			}
		}
	}
}

// RemoveMergedCell removes a merged cell region by its range string (e.g. "A1:C3").
func (ws *Worksheet) RemoveMergedCell(rangeStr string) error {
	start, end, err := ParseRange(rangeStr)
	if err != nil {
		return err
	}
	sr, sc := start.Row-1, start.ColumnIdx
	er, ec := end.Row-1, end.ColumnIdx

	newMerges := make([]MergeCell, 0, len(ws.mergeCells))
	found := false
	for _, mc := range ws.mergeCells {
		if mc.StartRow == sr && mc.StartCol == sc && mc.EndRow == er && mc.EndCol == ec {
			found = true
			continue
		}
		newMerges = append(newMerges, mc)
	}
	if !found {
		return fmt.Errorf("merged cell range %q not found", rangeStr)
	}
	ws.mergeCells = newMerges
	return nil
}

// SetRowHidden sets the visibility of a row (0-based index).
func (ws *Worksheet) SetRowHidden(row int, hidden bool) *Worksheet {
	ws.hiddenRows[row] = hidden
	return ws
}

// IsRowHidden returns whether a row is hidden.
func (ws *Worksheet) IsRowHidden(row int) bool {
	return ws.hiddenRows[row]
}

// SetColumnHidden sets the visibility of a column (0-based index).
func (ws *Worksheet) SetColumnHidden(col int, hidden bool) *Worksheet {
	ws.hiddenCols[col] = hidden
	return ws
}

// IsColumnHidden returns whether a column is hidden.
func (ws *Worksheet) IsColumnHidden(col int) bool {
	if ws.hiddenCols[col] {
		return true
	}
	for _, sp := range ws.colSpans {
		if sp.Hidden && col+1 >= sp.Min && col+1 <= sp.Max {
			return true
		}
	}
	return false
}

// SetSheetView sets the sheet view configuration.
func (ws *Worksheet) SetSheetView(sv *SheetView) *Worksheet {
	ws.sheetView = sv
	return ws
}

// GetSheetView returns the sheet view, creating a default one if not set.
func (ws *Worksheet) GetSheetView() *SheetView {
	if ws.sheetView == nil {
		ws.sheetView = &SheetView{
			ZoomScale:         100,
			ShowGridlines:     true,
			ShowRowColHeaders: true,
		}
	}
	return ws.sheetView
}

// Sort sorts the rows in the worksheet by the values in the given column.
// column is a column letter (e.g. "A", "B"), firstRow is 0-based index of the first data row
// (allows skipping header rows), and order specifies ascending or descending.
func (ws *Worksheet) Sort(column string, firstRow int, order SortOrder) error {
	colIdx, err := ColumnNameToIndex(column)
	if err != nil {
		return err
	}

	_, _, maxRow, _, dimErr := ws.Dimensions()
	if dimErr != nil {
		return dimErr
	}

	// Collect rows to sort
	type rowData struct {
		rowIdx int
		cells  map[int]*Cell // col -> cell
	}
	var rows []rowData
	for r := firstRow; r <= maxRow; r++ {
		rd := rowData{rowIdx: r, cells: make(map[int]*Cell)}
		for key, cell := range ws.cells {
			if cell.row == r {
				rd.cells[cell.col] = cell
				_ = key
			}
		}
		if len(rd.cells) > 0 {
			rows = append(rows, rd)
		}
	}

	// Sort rows by the specified column
	sort.SliceStable(rows, func(i, j int) bool {
		ci := rows[i].cells[colIdx]
		cj := rows[j].cells[colIdx]
		vi := sortValue(ci)
		vj := sortValue(cj)
		if order == SortOrderDescending {
			return vi > vj
		}
		return vi < vj
	})

	// Remove old cells for sorted rows
	for r := firstRow; r <= maxRow; r++ {
		for key, cell := range ws.cells {
			if cell.row == r {
				delete(ws.cells, key)
			}
		}
	}

	// Re-insert cells at new row positions
	for newIdx, rd := range rows {
		newRow := firstRow + newIdx
		for col, cell := range rd.cells {
			cell.row = newRow
			ws.cells[cellKey(newRow, col)] = cell
		}
	}

	return nil
}

// sortValue returns a comparable string for sorting.
func sortValue(c *Cell) string {
	if c == nil {
		return ""
	}
	switch c.Type {
	case CellTypeNumeric:
		if v, ok := c.Value.(float64); ok {
			// Pad with zeros for proper string comparison of numbers
			return fmt.Sprintf("%020.6f", v)
		}
	case CellTypeBool:
		if v, ok := c.Value.(bool); ok {
			if v {
				return "1"
			}
			return "0"
		}
	}
	return c.GetStringValue()
}

// SetBorderRange applies a border style to all cells in a range (e.g. "A1:C3").
// This creates a proper box border where edge cells get the appropriate border sides.
func (ws *Worksheet) SetBorderRange(rangeStr string, border Border) error {
	start, end, err := ParseRange(rangeStr)
	if err != nil {
		return err
	}
	sr, sc := start.Row-1, start.ColumnIdx
	er, ec := end.Row-1, end.ColumnIdx

	for r := sr; r <= er; r++ {
		for c := sc; c <= ec; c++ {
			cell := ws.GetCell(r, c)
			if cell.Style == nil {
				cell.Style = NewStyle()
			}
			if cell.Style.Borders == nil {
				cell.Style.Borders = &Borders{}
			}
			b := cell.Style.Borders
			// Top edge
			if r == sr {
				b.Top = border
			}
			// Bottom edge
			if r == er {
				b.Bottom = border
			}
			// Left edge
			if c == sc {
				b.Left = border
			}
			// Right edge
			if c == ec {
				b.Right = border
			}
		}
	}
	return nil
}

// Validate checks the worksheet for common issues and returns an error if invalid.
func (ws *Worksheet) Validate() error {
	if ws.title == "" {
		return fmt.Errorf("worksheet title cannot be empty")
	}
	// Check for overlapping merge cells
	for i, mc1 := range ws.mergeCells {
		for j, mc2 := range ws.mergeCells {
			if i >= j {
				continue
			}
			if mc1.StartRow <= mc2.EndRow && mc1.EndRow >= mc2.StartRow &&
				mc1.StartCol <= mc2.EndCol && mc1.EndCol >= mc2.StartCol {
				return fmt.Errorf("overlapping merged cell regions")
			}
		}
	}
	return nil
}
