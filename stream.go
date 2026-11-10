package gospreadsheet

import (
	"archive/zip"
	"bufio"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"strconv"
	"strings"
	"time"
)

// StreamCell is one cell in a streaming read or write.
// Column is 0-based. On write, Value may be a string, bool, float64, int,
// time.Time, or nil. A Formula without a leading "=" is stored as the formula
// body. Style is optional. Link is an external URL. Formatted is set by
// OpenRows when WithFormattedValues is used.
type StreamCell struct {
	Column    int
	Value     interface{}
	Formula   string
	Type      CellType
	Style     *Style
	Link      string
	Formatted string
}

// StreamRow is one worksheet row. Number is 1-based.
// The slice is overwritten on the next RowIterator.Next call.
type StreamRow struct {
	Number int
	Cells  []StreamCell
}

// RowIterator yields worksheet rows without building the cell map.
// Shared strings are loaded once; each row is discarded after it is returned.
type RowIterator struct {
	file      *os.File
	sheet     io.ReadCloser
	dec       *xml.Decoder
	sst       []string
	row       StreamRow
	err       error
	done      bool
	next      int
	formatted bool
	date1904  bool
	xfs       []parsedXF
}

// WithFormattedValues makes OpenRows apply number and date formats, including
// the 1904 date system. The formatted text is stored on StreamCell.Formatted
// and date cells are returned as time.Time.
func WithFormattedValues() func(*rowConfig) {
	return func(c *rowConfig) { c.formatted = true }
}

type rowConfig struct {
	formatted bool
}

// OpenRows opens a worksheet for streaming. sheet is the sheet name, or empty
// for the first sheet. Close releases the file.
func OpenRows(filename, sheet string, opts ...func(*rowConfig)) (*RowIterator, error) {
	f, err := os.Open(filename)
	if err != nil {
		return nil, fmt.Errorf("opening file: %w", err)
	}
	info, err := f.Stat()
	if err != nil {
		f.Close()
		return nil, err
	}
	zr, err := zip.NewReader(f, info.Size())
	if err != nil {
		f.Close()
		return nil, fmt.Errorf("opening zip: %w", err)
	}
	cfg := rowConfig{}
	for _, opt := range opts {
		if opt != nil {
			opt(&cfg)
		}
	}
	it, err := newRowIterator(zr, sheet, cfg)
	if err != nil {
		f.Close()
		return nil, err
	}
	it.file = f
	return it, nil
}

func newRowIterator(zr *zip.Reader, sheet string, cfg rowConfig) (*RowIterator, error) {
	wb, err := readZipFile(zr, "xl/workbook.xml")
	if err != nil {
		return nil, err
	}
	rels, _ := readZipFile(zr, "xl/_rels/workbook.xml.rels")
	relList, _ := parseRelationships(rels)
	part, err := sheetPart(wb, relList, sheet)
	if err != nil {
		return nil, err
	}
	var sst []string
	if raw, err := readZipFile(zr, "xl/sharedStrings.xml"); err == nil {
		items, err := parseSharedStrings(raw)
		if err != nil {
			return nil, err
		}
		sst = make([]string, len(items))
		for i, item := range items {
			sst[i] = item.text
		}
	}
	zf, ok := zipEntry(zr, part)
	if !ok {
		return nil, fmt.Errorf("sheet part %s is missing", part)
	}
	rc, err := zf.Open()
	if err != nil {
		return nil, err
	}
	it := &RowIterator{sheet: rc, dec: xml.NewDecoder(rc), sst: sst, next: 1, formatted: cfg.formatted}
	if cfg.formatted {
		it.date1904 = workbookDate1904(wb)
		if raw, err := readZipFile(zr, "xl/styles.xml"); err == nil {
			it.xfs = parseStyleSheet(raw, nil).xfs
		}
	}
	return it, nil
}

// Next advances to the next row. False means the sheet is finished or an error
// occurred; check Err.
func (it *RowIterator) Next() bool {
	if it == nil || it.done || it.err != nil {
		return false
	}
	for {
		tok, err := it.dec.Token()
		if err == io.EOF {
			it.done = true
			return false
		}
		if err != nil {
			it.err = err
			return false
		}
		se, ok := tok.(xml.StartElement)
		if !ok || se.Name.Local != "row" {
			continue
		}
		row, err := it.readRow(se)
		if err != nil {
			it.err = err
			return false
		}
		it.row = row
		return true
	}
}

// Row returns the current row. It is valid until the next Next call.
func (it *RowIterator) Row() StreamRow { return it.row }

// Err returns the first iteration error.
func (it *RowIterator) Err() error {
	if it == nil {
		return nil
	}
	return it.err
}

// Close releases the worksheet and the file.
func (it *RowIterator) Close() error {
	if it == nil {
		return nil
	}
	var err error
	if it.sheet != nil {
		err = it.sheet.Close()
		it.sheet = nil
	}
	if it.file != nil {
		if e := it.file.Close(); err == nil {
			err = e
		}
		it.file = nil
	}
	return err
}

func (it *RowIterator) readRow(start xml.StartElement) (StreamRow, error) {
	n := it.next
	if v := attrLocal(start, "r"); v != "" {
		if parsed, err := strconv.Atoi(v); err == nil && parsed > 0 {
			n = parsed
		}
	}
	it.next = n + 1
	row := StreamRow{Number: n}
	for {
		tok, err := it.dec.Token()
		if err != nil {
			return row, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local != "c" {
				if err := skipElement(it.dec); err != nil {
					return row, err
				}
				continue
			}
			cell, err := it.readCell(t)
			if err != nil {
				return row, err
			}
			if cell.Type != CellTypeEmpty || cell.Formula != "" || cell.Value != nil {
				row.Cells = append(row.Cells, cell)
			}
		case xml.EndElement:
			if t.Name.Local == "row" {
				return row, nil
			}
		}
	}
}

func (it *RowIterator) readCell(start xml.StartElement) (StreamCell, error) {
	cell := StreamCell{}
	if ref := attrLocal(start, "r"); ref != "" {
		if cr, err := ParseCellReference(ref); err == nil {
			cell.Column = cr.ColumnIdx
		}
	}
	kind := attrLocal(start, "t")
	styleIdx := -1
	if s := attrLocal(start, "s"); s != "" {
		if n, err := strconv.Atoi(s); err == nil {
			styleIdx = n
		}
	}
	var value, formula string
	for {
		tok, err := it.dec.Token()
		if err != nil {
			return cell, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "f":
				var b strings.Builder
				if err := readCharData(it.dec, &b, "f"); err != nil {
					return cell, err
				}
				formula = strings.TrimSpace(b.String())
			case "v":
				var b strings.Builder
				if err := readCharData(it.dec, &b, "v"); err != nil {
					return cell, err
				}
				value = b.String()
			case "is":
				item, err := parseStringItem(it.dec, nil)
				if err != nil {
					return cell, err
				}
				value = item.text
				kind = "inlineStr"
			default:
				if err := skipElement(it.dec); err != nil {
					return cell, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "c" {
				cell.Formula = formula
				switch kind {
				case "s":
					idx, _ := strconv.Atoi(strings.TrimSpace(value))
					if idx >= 0 && idx < len(it.sst) {
						cell.Value = it.sst[idx]
					}
					cell.Type = CellTypeString
				case "inlineStr", "str":
					cell.Value = unescapeOOXMLText(value)
					cell.Type = CellTypeString
				case "b":
					cell.Value = value == "1" || strings.EqualFold(value, "true")
					cell.Type = CellTypeBool
				case "e":
					cell.Value = value
					cell.Type = CellTypeError
				default:
					if value == "" && formula == "" {
						return cell, nil
					}
					if f, err := strconv.ParseFloat(value, 64); err == nil {
						cell.Value = f
						cell.Type = CellTypeNumeric
					} else if value != "" {
						cell.Value = value
						cell.Type = CellTypeString
					}
				}
				if formula != "" {
					cell.Type = CellTypeFormula
				}
				it.applyFormat(&cell, styleIdx)
				return cell, nil
			}
		}
	}
}

func (it *RowIterator) applyFormat(cell *StreamCell, styleIdx int) {
	if it == nil || !it.formatted || cell == nil {
		return
	}
	code := "General"
	date := false
	if styleIdx >= 0 && styleIdx < len(it.xfs) {
		if it.xfs[styleIdx].formatCode != "" {
			code = it.xfs[styleIdx].formatCode
		}
		date = it.xfs[styleIdx].date
	}
	switch v := cell.Value.(type) {
	case float64:
		if date {
			serial := v
			if it.date1904 {
				serial += 1462
			}
			t := serialToDate(serial)
			cell.Value = t
			cell.Type = CellTypeDate
			cell.Formatted = formatByExcelCode(serial, code)
			return
		}
		cell.Formatted = formatByExcelCode(v, code)
	case time.Time:
		cell.Formatted = formatByExcelCode(v, code)
	case string:
		cell.Formatted = formatByExcelCode(v, code)
	case bool:
		if v {
			cell.Formatted = "TRUE"
		} else {
			cell.Formatted = "FALSE"
		}
	}
}

func workbookDate1904(workbook []byte) bool {
	return strings.Contains(string(workbook), `date1904="1"`) || strings.Contains(string(workbook), `date1904="true"`)
}

func zipEntry(zr *zip.Reader, name string) (*zip.File, bool) {
	for _, f := range zr.File {
		if strings.EqualFold(f.Name, name) {
			return f, true
		}
	}
	return nil, false
}

func sheetPart(workbook []byte, rels []relationship, want string) (string, error) {
	dec := xml.NewDecoder(bytes.NewReader(workbook))
	var name, rid string
	for {
		tok, err := dec.Token()
		if err != nil {
			break
		}
		se, ok := tok.(xml.StartElement)
		if !ok || se.Name.Local != "sheet" {
			continue
		}
		n := attrLocal(se, "name")
		id := attrLocal(se, "id")
		if want == "" || n == want {
			name, rid = n, id
			break
		}
	}
	if rid == "" {
		if want == "" {
			return "", fmt.Errorf("workbook has no sheets")
		}
		return "", fmt.Errorf("sheet %q was not found", want)
	}
	for _, rel := range rels {
		if rel.ID == rid {
			if p := resolvePart("xl/workbook.xml", rel.Target); p != "" {
				return p, nil
			}
		}
	}
	_ = name
	return "", fmt.Errorf("sheet %q has no relationship", want)
}

// StreamWriter writes an xlsx workbook row by row. Memory stays proportional
// to one row, the unique styles, and (when enabled) the unique shared strings.
type StreamWriter struct {
	zw        *zip.Writer
	dst       io.WriteCloser
	sheets    []string
	open      *StreamSheet
	closed    bool
	styles    *styleRegistry
	shared    bool
	sst       map[string]int
	sstList   []string
	formulas  bool
	dateStyle int
}

// UseSharedStrings stores cell text in a shared-string table instead of inline.
// Memory grows with the number of distinct strings.
func (w *StreamWriter) UseSharedStrings() {
	if w != nil {
		w.shared = true
		if w.sst == nil {
			w.sst = map[string]int{}
		}
	}
}

// NewStreamWriter writes an xlsx package to w.
func NewStreamWriter(w io.Writer) *StreamWriter {
	return &StreamWriter{zw: zip.NewWriter(w)}
}

// CreateStream creates filename and returns a writer. Close creates the
// workbook parts and closes the file.
func CreateStream(filename string) (*StreamWriter, error) {
	f, err := os.Create(filename)
	if err != nil {
		return nil, err
	}
	w := NewStreamWriter(f)
	w.dst = f
	return w, nil
}

func (w *StreamWriter) initStyles() {
	if w.styles == nil {
		w.styles = newStyleRegistry(nil)
		w.dateStyle = -1
	}
}

// StreamSheet is one worksheet being written.
type StreamSheet struct {
	parent     *StreamWriter
	bw         *bufio.Writer
	lastRow    int
	closed     bool
	header     bool
	freezeRows int
	freezeCols int
	cols       []streamCol
	merges     []string
	links      []streamLink
}

type streamCol struct {
	min, max int
	width    float64
}

type streamLink struct {
	ref, url, id string
}

// SetColWidth sets the width of one 0-based column. Call it before the first row.
func (s *StreamSheet) SetColWidth(col int, width float64) error {
	return s.SetColWidths(col, col, width)
}

// SetColWidths sets the width of an inclusive 0-based column span.
func (s *StreamSheet) SetColWidths(minCol, maxCol int, width float64) error {
	if s == nil || s.closed {
		return fmt.Errorf("sheet is closed")
	}
	if s.header {
		return fmt.Errorf("column widths must be set before the first row")
	}
	if minCol > maxCol {
		minCol, maxCol = maxCol, minCol
	}
	s.cols = append(s.cols, streamCol{min: minCol, max: maxCol, width: width})
	return nil
}

// MergeCells records a merge such as "A1:B2". It is written when the sheet closes.
func (s *StreamSheet) MergeCells(ref string) error {
	if s == nil || s.closed {
		return fmt.Errorf("sheet is closed")
	}
	ref = strings.TrimSpace(ref)
	if ref == "" {
		return fmt.Errorf("merge reference is required")
	}
	s.merges = append(s.merges, ref)
	return nil
}

// Freeze freezes the top rows and left columns. Call it before the first row.
func (s *StreamSheet) Freeze(rows, cols int) error {
	if s == nil || s.closed {
		return fmt.Errorf("sheet is closed")
	}
	if s.header {
		return fmt.Errorf("freeze panes must be set before the first row")
	}
	if rows < 0 || cols < 0 {
		return fmt.Errorf("freeze counts must be non-negative")
	}
	s.freezeRows, s.freezeCols = rows, cols
	return nil
}

// AddSheet starts a worksheet. The previous sheet is closed first.
// Rows must be written in nondecreasing order.
func (w *StreamWriter) AddSheet(name string) (*StreamSheet, error) {
	if w == nil || w.closed {
		return nil, fmt.Errorf("stream writer is closed")
	}
	name = strings.TrimSpace(name)
	if name == "" {
		return nil, fmt.Errorf("sheet name is required")
	}
	if w.open != nil {
		if err := w.open.Close(); err != nil {
			return nil, err
		}
	}
	part := fmt.Sprintf("xl/worksheets/sheet%d.xml", len(w.sheets)+1)
	entry, err := w.zw.Create(part)
	if err != nil {
		return nil, err
	}
	bw := bufio.NewWriterSize(entry, 64*1024)
	if _, err := bw.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n"); err != nil {
		return nil, err
	}
	if _, err := bw.WriteString(`<worksheet xmlns="` + spreadsheetNS + `" xmlns:r="` + relNS + `">`); err != nil {
		return nil, err
	}
	w.initStyles()
	s := &StreamSheet{parent: w, bw: bw}
	w.open = s
	w.sheets = append(w.sheets, name)
	return s, nil
}

// WriteRow writes a 1-based row. Nil values are skipped.
func (s *StreamSheet) WriteRow(row int, cells []StreamCell) error {
	if s == nil || s.closed {
		return fmt.Errorf("sheet is closed")
	}
	if row < 1 {
		return fmt.Errorf("row %d is out of range", row)
	}
	if s.lastRow != 0 && row < s.lastRow {
		return fmt.Errorf("row %d is before row %d", row, s.lastRow)
	}
	s.lastRow = row
	if err := s.ensureHeader(); err != nil {
		return err
	}
	if _, err := fmt.Fprintf(s.bw, `<row r="%d">`, row); err != nil {
		return err
	}
	for _, c := range cells {
		if c.Value == nil && c.Formula == "" && c.Style == nil && c.Link == "" {
			continue
		}
		ref, err := CellName(row-1, c.Column)
		if err != nil {
			return err
		}
		if c.Link != "" {
			s.links = append(s.links, streamLink{ref: ref, url: c.Link, id: fmt.Sprintf("rId%d", len(s.links)+1)})
		}
		if c.Formula != "" {
			s.parent.formulas = true
		}
		if err := writeStreamCell(s.bw, ref, c, s.parent); err != nil {
			return err
		}
	}
	_, err := s.bw.WriteString(`</row>`)
	return err
}

// Close finishes this worksheet.
func (s *StreamSheet) Close() error {
	if s == nil || s.closed {
		return nil
	}
	s.closed = true
	if s.parent != nil && s.parent.open == s {
		s.parent.open = nil
	}
	if err := s.ensureHeader(); err != nil {
		return err
	}
	if _, err := s.bw.WriteString(`</sheetData>`); err != nil {
		return err
	}
	if len(s.merges) > 0 {
		fmt.Fprintf(s.bw, `<mergeCells count="%d">`, len(s.merges))
		for _, ref := range s.merges {
			fmt.Fprintf(s.bw, `<mergeCell ref="%s"/>`, xmlText(ref))
		}
		s.bw.WriteString(`</mergeCells>`)
	}
	if len(s.links) > 0 {
		fmt.Fprintf(s.bw, `<hyperlinks>`)
		for _, lk := range s.links {
			fmt.Fprintf(s.bw, `<hyperlink ref="%s" r:id="%s"/>`, xmlText(lk.ref), xmlText(lk.id))
		}
		s.bw.WriteString(`</hyperlinks>`)
	}
	if _, err := s.bw.WriteString(`</worksheet>`); err != nil {
		return err
	}
	if err := s.bw.Flush(); err != nil {
		return err
	}
	if len(s.links) == 0 || s.parent == nil {
		return nil
	}
	part := fmt.Sprintf("xl/worksheets/_rels/sheet%d.xml.rels", len(s.parent.sheets))
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	b.WriteString(`<Relationships xmlns="` + pkgRelNS + `">`)
	for _, lk := range s.links {
		fmt.Fprintf(&b, `<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="%s" TargetMode="External"/>`, xmlText(lk.id), xmlText(lk.url))
	}
	b.WriteString(`</Relationships>`)
	return writeZip(s.parent.zw, part, b.String())
}

func (s *StreamSheet) ensureHeader() error {
	if s.header {
		return nil
	}
	s.header = true
	var b strings.Builder
	b.WriteString(`<sheetViews><sheetView workbookViewId="0"`)
	if s.freezeRows == 0 && s.freezeCols == 0 {
		b.WriteString(`/></sheetViews>`)
	} else {
		b.WriteString(`>`)
		pane := "bottomRight"
		switch {
		case s.freezeCols > 0 && s.freezeRows > 0:
			pane = "bottomRight"
		case s.freezeCols > 0:
			pane = "topRight"
		default:
			pane = "bottomLeft"
		}
		b.WriteString(`<pane`)
		if s.freezeCols > 0 {
			fmt.Fprintf(&b, ` xSplit="%d"`, s.freezeCols)
		}
		if s.freezeRows > 0 {
			fmt.Fprintf(&b, ` ySplit="%d"`, s.freezeRows)
		}
		top, _ := CellName(s.freezeRows, s.freezeCols)
		fmt.Fprintf(&b, ` topLeftCell="%s" activePane="%s" state="frozen"/>`, top, pane)
		fmt.Fprintf(&b, `<selection pane="%s" activeCell="%s" sqref="%s"/>`, pane, top, top)
		b.WriteString(`</sheetView></sheetViews>`)
	}
	if len(s.cols) > 0 {
		fmt.Fprintf(&b, `<cols>`)
		for _, c := range s.cols {
			fmt.Fprintf(&b, `<col min="%d" max="%d" width="%s" customWidth="1"/>`, c.min+1, c.max+1, trimFloat(c.width))
		}
		b.WriteString(`</cols>`)
	}
	b.WriteString(`<sheetData>`)
	_, err := s.bw.WriteString(b.String())
	return err
}

// Close writes the workbook, styles, and package relationships.
func (w *StreamWriter) Close() error {
	if w == nil || w.closed {
		return nil
	}
	if w.open != nil {
		if err := w.open.Close(); err != nil {
			return err
		}
	}
	if len(w.sheets) == 0 {
		if _, err := w.AddSheet("Sheet1"); err != nil {
			return err
		}
		if err := w.open.Close(); err != nil {
			return err
		}
	}
	w.closed = true
	if err := writeStreamStatic(w); err != nil {
		return err
	}
	if err := w.zw.Close(); err != nil {
		return err
	}
	if w.dst != nil {
		return w.dst.Close()
	}
	return nil
}

func writeStreamCell(w io.Writer, ref string, c StreamCell, parent *StreamWriter) error {
	if _, err := fmt.Fprintf(w, `<c r="%s"`, ref); err != nil {
		return err
	}
	if c.Style != nil && parent != nil {
		parent.initStyles()
		if id := parent.styles.xfIndex(c.Style); id > 0 {
			fmt.Fprintf(w, ` s="%d"`, id)
		}
	}
	if _, ok := c.Value.(time.Time); ok && (c.Style == nil || c.Style.NumberFormat == nil) && parent != nil {
		parent.initStyles()
		if parent.dateStyle < 0 {
			st := NewStyle()
			st.NumberFormat = &FormatDate
			parent.dateStyle = parent.styles.xfIndex(st)
		}
		if parent.dateStyle > 0 {
			fmt.Fprintf(w, ` s="%d"`, parent.dateStyle)
		}
	}
	formula := strings.TrimSpace(c.Formula)
	if strings.HasPrefix(formula, "=") {
		formula = formula[1:]
	}
	if formula != "" {
		if _, err := fmt.Fprintf(w, `><f>%s</f>`, xmlText(formula)); err != nil {
			return err
		}
		if c.Value != nil {
			if err := writeStreamCached(w, c.Value); err != nil {
				return err
			}
		}
		_, err := io.WriteString(w, `</c>`)
		return err
	}
	switch val := c.Value.(type) {
	case string:
		if parent != nil && parent.shared {
			idx := parent.sharedIndex(val)
			_, err := fmt.Fprintf(w, ` t="s"><v>%d</v></c>`, idx)
			return err
		}
		_, err := fmt.Fprintf(w, ` t="inlineStr"><is><t>%s</t></is></c>`, xmlCellText(val))
		return err
	case time.Time:
		_, err := io.WriteString(w, `><v>`+formatNumber(dateToSerial(val))+`</v></c>`)
		return err
	case bool:
		n := 0
		if val {
			n = 1
		}
		_, err := fmt.Fprintf(w, ` t="b"><v>%d</v></c>`, n)
		return err
	case int:
		_, err := fmt.Fprintf(w, `><v>%d</v></c>`, val)
		return err
	case int64:
		_, err := fmt.Fprintf(w, `><v>%d</v></c>`, val)
		return err
	case float64:
		_, err := io.WriteString(w, `><v>`+formatNumber(val)+`</v></c>`)
		return err
	case float32:
		_, err := io.WriteString(w, `><v>`+formatNumber(float64(val))+`</v></c>`)
		return err
	default:
		if c.Value == nil {
			_, err := io.WriteString(w, `/>`)
			return err
		}
		_, err := fmt.Fprintf(w, `><v>%s</v></c>`, xmlCellText(fmt.Sprint(c.Value)))
		return err
	}
}

func writeStreamCached(w io.Writer, v interface{}) error {
	switch val := v.(type) {
	case string:
		_, err := fmt.Fprintf(w, `<v>%s</v>`, xmlCellText(val))
		return err
	case bool:
		n := 0
		if val {
			n = 1
		}
		_, err := fmt.Fprintf(w, `<v>%d</v>`, n)
		return err
	case int:
		_, err := fmt.Fprintf(w, `<v>%d</v>`, val)
		return err
	case int64:
		_, err := fmt.Fprintf(w, `<v>%d</v>`, val)
		return err
	case float64:
		_, err := io.WriteString(w, `<v>`+formatNumber(val)+`</v>`)
		return err
	default:
		_, err := fmt.Fprintf(w, `<v>%s</v>`, xmlCellText(fmt.Sprint(val)))
		return err
	}
}

func (w *StreamWriter) sharedIndex(s string) int {
	if w.sst == nil {
		w.sst = map[string]int{}
	}
	if idx, ok := w.sst[s]; ok {
		return idx
	}
	idx := len(w.sstList)
	w.sst[s] = idx
	w.sstList = append(w.sstList, s)
	return idx
}

func writeStreamStatic(w *StreamWriter) error {
	zw := w.zw
	sheets := w.sheets
	var ct strings.Builder
	ct.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	ct.WriteString(`<Types xmlns="` + ctNS + `">`)
	ct.WriteString(`<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`)
	ct.WriteString(`<Default Extension="xml" ContentType="application/xml"/>`)
	ct.WriteString(`<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`)
	ct.WriteString(`<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`)
	if len(w.sstList) > 0 {
		ct.WriteString(`<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>`)
	}
	for i := range sheets {
		fmt.Fprintf(&ct, `<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`, i+1)
	}
	ct.WriteString(`</Types>`)
	if err := writeZip(zw, "[Content_Types].xml", ct.String()); err != nil {
		return err
	}
	root := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<Relationships xmlns="` + pkgRelNS + `">` +
		`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>` +
		`</Relationships>`
	if err := writeZip(zw, "_rels/.rels", root); err != nil {
		return err
	}
	var wb strings.Builder
	wb.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	wb.WriteString(`<workbook xmlns="` + spreadsheetNS + `" xmlns:r="` + relNS + `"><sheets>`)
	for i, name := range sheets {
		fmt.Fprintf(&wb, `<sheet name="%s" sheetId="%d" r:id="rId%d"/>`, xmlText(name), i+1, i+1)
	}
	wb.WriteString(`</sheets>`)
	if w.formulas {
		wb.WriteString(`<calcPr calcId="191029" fullCalcOnLoad="1"/>`)
	}
	wb.WriteString(`</workbook>`)
	if err := writeZip(zw, "xl/workbook.xml", wb.String()); err != nil {
		return err
	}
	var rels strings.Builder
	rels.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	rels.WriteString(`<Relationships xmlns="` + pkgRelNS + `">`)
	for i := range sheets {
		fmt.Fprintf(&rels, `<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>`, i+1, i+1)
	}
	next := len(sheets) + 1
	fmt.Fprintf(&rels, `<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`, next)
	next++
	if len(w.sstList) > 0 {
		fmt.Fprintf(&rels, `<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`, next)
	}
	rels.WriteString(`</Relationships>`)
	if err := writeZip(zw, "xl/_rels/workbook.xml.rels", rels.String()); err != nil {
		return err
	}
	if w.styles == nil {
		w.initStyles()
	}
	var styles strings.Builder
	w.styles.writeStyles(&styles)
	if err := writeZip(zw, "xl/styles.xml", styles.String()); err != nil {
		return err
	}
	if len(w.sstList) == 0 {
		return nil
	}
	var sst strings.Builder
	sst.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	fmt.Fprintf(&sst, `<sst xmlns="%s" count="%d" uniqueCount="%d">`, spreadsheetNS, len(w.sstList), len(w.sstList))
	for _, s := range w.sstList {
		if needsPreserveSpace(s) {
			fmt.Fprintf(&sst, `<si><t xml:space="preserve">%s</t></si>`, xmlCellText(s))
		} else {
			fmt.Fprintf(&sst, `<si><t>%s</t></si>`, xmlCellText(s))
		}
	}
	sst.WriteString(`</sst>`)
	return writeZip(zw, "xl/sharedStrings.xml", sst.String())
}
