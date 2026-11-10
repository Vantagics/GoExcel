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
)

// StreamCell is one cell in a streaming read or write.
// Column is 0-based. On write, Value may be a string, bool, float64, int,
// or nil. A Formula without a leading "=" is stored as the formula body.
type StreamCell struct {
	Column  int
	Value   interface{}
	Formula string
	Type    CellType
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
	file  *os.File
	sheet io.ReadCloser
	dec   *xml.Decoder
	sst   []string
	row   StreamRow
	err   error
	done  bool
	next  int
}

// OpenRows opens a worksheet for streaming. sheet is the sheet name, or empty
// for the first sheet. Close releases the file.
func OpenRows(filename, sheet string) (*RowIterator, error) {
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
	it, err := newRowIterator(zr, sheet)
	if err != nil {
		f.Close()
		return nil, err
	}
	it.file = f
	return it, nil
}

func newRowIterator(zr *zip.Reader, sheet string) (*RowIterator, error) {
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
	return &RowIterator{sheet: rc, dec: xml.NewDecoder(rc), sst: sst, next: 1}, nil
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
				return cell, nil
			}
		}
	}
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
// to one row plus the zip compressor, not to the sheet size. Cell text is
// stored inline so a shared-string table is not accumulated.
type StreamWriter struct {
	zw     *zip.Writer
	dst    io.WriteCloser
	sheets []string
	open   *StreamSheet
	closed bool
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

// StreamSheet is one worksheet being written.
type StreamSheet struct {
	parent  *StreamWriter
	bw      *bufio.Writer
	lastRow int
	closed  bool
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
	if _, err := bw.WriteString(`<worksheet xmlns="` + spreadsheetNS + `"><sheetData>`); err != nil {
		return nil, err
	}
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
	if _, err := fmt.Fprintf(s.bw, `<row r="%d">`, row); err != nil {
		return err
	}
	for _, c := range cells {
		if c.Value == nil && c.Formula == "" {
			continue
		}
		ref, err := CellName(row-1, c.Column)
		if err != nil {
			return err
		}
		if err := writeStreamCell(s.bw, ref, c); err != nil {
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
	if _, err := s.bw.WriteString(`</sheetData></worksheet>`); err != nil {
		return err
	}
	return s.bw.Flush()
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
	if err := writeStreamStatic(w.zw, w.sheets); err != nil {
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

func writeStreamCell(w io.Writer, ref string, c StreamCell) error {
	if _, err := fmt.Fprintf(w, `<c r="%s"`, ref); err != nil {
		return err
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
		_, err := fmt.Fprintf(w, ` t="inlineStr"><is><t>%s</t></is></c>`, xmlCellText(val))
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

func writeStreamStatic(zw *zip.Writer, sheets []string) error {
	var ct strings.Builder
	ct.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	ct.WriteString(`<Types xmlns="` + ctNS + `">`)
	ct.WriteString(`<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`)
	ct.WriteString(`<Default Extension="xml" ContentType="application/xml"/>`)
	ct.WriteString(`<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`)
	ct.WriteString(`<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`)
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
	wb.WriteString(`</sheets></workbook>`)
	if err := writeZip(zw, "xl/workbook.xml", wb.String()); err != nil {
		return err
	}
	var rels strings.Builder
	rels.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	rels.WriteString(`<Relationships xmlns="` + pkgRelNS + `">`)
	for i := range sheets {
		fmt.Fprintf(&rels, `<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>`, i+1, i+1)
	}
	fmt.Fprintf(&rels, `<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`, len(sheets)+1)
	rels.WriteString(`</Relationships>`)
	if err := writeZip(zw, "xl/_rels/workbook.xml.rels", rels.String()); err != nil {
		return err
	}
	styles := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<styleSheet xmlns="` + spreadsheetNS + `">` +
		`<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>` +
		`<fills count="1"><fill><patternFill patternType="none"/></fill></fills>` +
		`<borders count="1"><border/></borders>` +
		`<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
		`<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>` +
		`</styleSheet>`
	return writeZip(zw, "xl/styles.xml", styles)
}
