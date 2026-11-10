package gospreadsheet

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"strconv"
	"strings"
)

// XLSXReader reads XLSX files into a Workbook.
type XLSXReader struct{}

// NewXLSXReader creates a new XLSX reader.
func NewXLSXReader() *XLSXReader {
	return &XLSXReader{}
}

// Open reads an XLSX file and returns a Workbook.
func (r *XLSXReader) Open(filename string) (*Workbook, error) {
	f, err := os.Open(filename)
	if err != nil {
		return nil, fmt.Errorf("opening file: %w", err)
	}
	defer f.Close()
	info, err := f.Stat()
	if err != nil {
		return nil, fmt.Errorf("stat file: %w", err)
	}
	return r.Read(f, info.Size())
}

// OpenPassword reads an xlsx file, decrypting it when it is an ECMA-376
// encrypted package. password may be empty for a normal zip package.
func (r *XLSXReader) OpenPassword(filename, password string) (*Workbook, error) {
	b, err := os.ReadFile(filename)
	if err != nil {
		return nil, fmt.Errorf("opening file: %w", err)
	}
	return r.ReadPassword(bytes.NewReader(b), int64(len(b)), password)
}

// ReadPassword reads an xlsx package, decrypting OLE/ECMA-376 encryption when
// the bytes are a compound file. A normal zip package ignores password.
func (r *XLSXReader) ReadPassword(reader io.ReaderAt, size int64, password string) (*Workbook, error) {
	if isOLEReader(reader, size) {
		if password == "" {
			return nil, ErrPasswordRequired
		}
		buf := make([]byte, size)
		n, err := reader.ReadAt(buf, 0)
		if err != nil && err != io.EOF {
			return nil, fmt.Errorf("reading encrypted package: %w", err)
		}
		if int64(n) != size {
			return nil, fmt.Errorf("reading encrypted package: short read %d of %d", n, size)
		}
		plain, err := DecryptPackage(buf, password)
		if err != nil {
			return nil, err
		}
		return r.readPackage(bytes.NewReader(plain), int64(len(plain)))
	}
	return r.readPackage(reader, size)
}

// Read reads an XLSX from an io.ReaderAt with the given size.
// Encrypted packages return ErrPasswordRequired; use ReadPassword.
func (r *XLSXReader) Read(reader io.ReaderAt, size int64) (*Workbook, error) {
	if isOLEReader(reader, size) {
		return nil, ErrPasswordRequired
	}
	return r.readPackage(reader, size)
}

func isOLEReader(reader io.ReaderAt, size int64) bool {
	if reader == nil || size < 8 {
		return false
	}
	magic := make([]byte, 8)
	n, err := reader.ReadAt(magic, 0)
	return err == nil && n == 8 && isOLE(magic)
}

func (r *XLSXReader) readPackage(reader io.ReaderAt, size int64) (*Workbook, error) {
	zr, err := zip.NewReader(reader, size)
	if err != nil {
		return nil, fmt.Errorf("opening zip: %w", err)
	}
	pkg := readPackage(zr)
	wb := NewEmpty()
	wb.source = pkg

	wbData, ok := pkg.file("xl/workbook.xml")
	if !ok {
		return nil, fmt.Errorf("reading workbook: %w", fmt.Errorf("file %q not found in archive", "xl/workbook.xml"))
	}
	relsData, ok := pkg.file("xl/_rels/workbook.xml.rels")
	if !ok {
		return nil, fmt.Errorf("reading workbook rels: %w", fmt.Errorf("file %q not found in archive", "xl/_rels/workbook.xml.rels"))
	}
	rels, err := parseRelationships(relsData)
	if err != nil {
		return nil, fmt.Errorf("reading workbook rels: %w", fmt.Errorf("parsing workbook.xml.rels: %w", err))
	}
	wb.workbookRels = rels
	if data, ok := pkg.file("_rels/.rels"); ok {
		if pkgRels, err := parseRelationships(data); err == nil {
			wb.packageRels = pkgRels
		}
	}
	if err := parseWorkbook(wbData, wb); err != nil {
		return nil, fmt.Errorf("reading workbook: %w", err)
	}

	var theme []string
	if rel, ok := findRel(rels, "theme"); ok {
		if p := resolvePart("xl/workbook.xml", rel.Target); p != "" {
			if data, ok := pkg.file(p); ok {
				theme = loadThemeColors(data)
			}
		}
	}
	styles := styleTable{}
	stylePath := "xl/styles.xml"
	if rel, ok := findRel(rels, "styles"); ok {
		if p := resolvePart("xl/workbook.xml", rel.Target); p != "" {
			stylePath = p
		}
	}
	if data, ok := pkg.file(stylePath); ok {
		styles = parseStyleSheet(data, theme)
		wb.styleExtras = styles.extras
	}

	sst, err := loadSharedStrings(pkg, rels)
	if err != nil {
		return nil, fmt.Errorf("reading shared strings: %w", err)
	}
	readCoreProperties(pkg, wb)

	for _, ws := range wb.worksheets {
		if ws.partName == "" {
			continue
		}
		data, ok := pkg.file(ws.partName)
		if !ok {
			return nil, fmt.Errorf("reading sheet %q: %w", ws.title, fmt.Errorf("file %q not found in archive", ws.partName))
		}
		if root := rootLocalName(data); root != "" && root != "worksheet" {
			ws.kind = root
			ws.rawPart = append([]byte(nil), data...)
			if ctData, ok := pkg.file("[Content_Types].xml"); ok {
				ws.rawContentType = parseContentTypes(ctData).overrideFor(ws.partName)
			}
			if relData, ok := pkg.file(relsPathFor(ws.partName)); ok {
				if sheetRels, err := parseRelationships(relData); err == nil {
					ws.sheetRels = sheetRels
				}
			}
			continue
		}
		if err := readWorksheet(ws, data, pkg, sst, &styles, wb, theme); err != nil {
			return nil, fmt.Errorf("reading sheet %q: %w", ws.title, err)
		}
	}
	indexWorkbookFeatures(wb)
	return wb, nil
}

func readZipFile(zr *zip.Reader, name string) ([]byte, error) {
	for _, f := range zr.File {
		if strings.EqualFold(f.Name, name) {
			rc, err := f.Open()
			if err != nil {
				return nil, err
			}
			defer rc.Close()
			return io.ReadAll(rc)
		}
	}
	return nil, fmt.Errorf("file %q not found in archive", name)
}

func loadSharedStrings(pkg *ooxmlPackage, rels []relationship) ([]sstItem, error) {
	path := "xl/sharedStrings.xml"
	if rel, ok := findRel(rels, "sharedStrings"); ok {
		if p := resolvePart("xl/workbook.xml", rel.Target); p != "" {
			path = p
		}
	}
	data, ok := pkg.file(path)
	if !ok {
		return nil, nil
	}
	items, err := parseSharedStrings(data)
	if err != nil {
		return nil, fmt.Errorf("parsing sharedStrings.xml: %w", err)
	}
	return items, nil
}

type sstItem struct {
	text string
	rich *RichText
}

func parseSharedStrings(data []byte) ([]sstItem, error) {
	dec := xml.NewDecoder(bytes.NewReader(stripBOM(data)))
	var items []sstItem
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			return items, nil
		}
		if err != nil {
			return nil, err
		}
		se, ok := tok.(xml.StartElement)
		if !ok || se.Name.Local != "si" {
			continue
		}
		item, err := parseStringItem(dec, nil)
		if err != nil {
			return nil, err
		}
		items = append(items, item)
	}
}

// parseStringItem reads an si or is element. The start tag is already consumed.
func parseStringItem(dec *xml.Decoder, theme []string) (sstItem, error) {
	var plain strings.Builder
	var runs []RichTextRun
	inRPh := 0
	inR := false
	inT := false
	sawR := false
	var cur strings.Builder
	var font *Font
	for {
		tok, err := dec.Token()
		if err != nil {
			return sstItem{}, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "rPh":
				inRPh++
			case "r":
				if inRPh == 0 {
					inR = true
					sawR = true
					cur.Reset()
					font = nil
				}
			case "rPr":
				if inR && inRPh == 0 {
					f, err := parseRunProps(dec, theme)
					if err != nil {
						return sstItem{}, err
					}
					font = f
				}
			case "t":
				if inRPh == 0 {
					inT = true
				}
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "rPh":
				if inRPh > 0 {
					inRPh--
				}
			case "r":
				if inR && inRPh == 0 {
					runs = append(runs, RichTextRun{Text: cur.String(), Font: font})
					inR = false
				}
			case "t":
				inT = false
			case "si", "is":
				item := sstItem{text: unescapeOOXMLText(plain.String())}
				if sawR {
					for i := range runs {
						runs[i].Text = unescapeOOXMLText(runs[i].Text)
					}
					item.rich = &RichText{Runs: runs}
					item.text = item.rich.PlainText()
				}
				return item, nil
			}
		case xml.CharData:
			if inT && inRPh == 0 {
				s := string(t)
				plain.WriteString(s)
				if inR {
					cur.WriteString(s)
				}
			}
		}
	}
}

func parseRunProps(dec *xml.Decoder, theme []string) (*Font, error) {
	// rPr start already consumed.
	font := &Font{}
	depth := 1
	for depth > 0 {
		tok, err := dec.Token()
		if err != nil {
			return nil, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			switch t.Name.Local {
			case "b":
				if attrLocal(t, "val") != "0" && !strings.EqualFold(attrLocal(t, "val"), "false") {
					font.Bold = true
				}
			case "i":
				if attrLocal(t, "val") != "0" && !strings.EqualFold(attrLocal(t, "val"), "false") {
					font.Italic = true
				}
			case "strike":
				if attrLocal(t, "val") != "0" {
					font.Strikethrough = true
				}
			case "u":
				if !strings.EqualFold(attrLocal(t, "val"), "none") && attrLocal(t, "val") != "0" {
					font.Underline = true
				}
			case "sz":
				font.Size, _ = strconv.ParseFloat(attrLocal(t, "val"), 64)
			case "rFont", "name":
				if v := attrLocal(t, "val"); v != "" {
					font.Name = v
				}
			case "color":
				font.Color = resolveColor(attrLocal(t, "rgb"), attrLocal(t, "theme"), attrLocal(t, "indexed"), parseFloatAttr(attrLocal(t, "tint")), theme)
			}
		case xml.EndElement:
			depth--
		}
	}
	if font.Name == "" && font.Size == 0 && !font.Bold && !font.Italic && !font.Underline && !font.Strikethrough && font.Color == "" {
		return nil, nil
	}
	return font, nil
}

func parseFloatAttr(v string) float64 {
	f, _ := strconv.ParseFloat(v, 64)
	return f
}

func parseWorkbook(data []byte, wb *Workbook) error {
	wb.extraRootAttrs = extraNamespaceDecls(firstStartTag(data, "workbook"))
	if bytes.Contains(data, []byte(`conformance="strict"`)) || bytes.Contains(data, []byte("http://purl.oclc.org/ooxml/spreadsheetml/main")) {
		wb.loadedStrict = true
	}
	data = stripBOM(data)
	dec := xml.NewDecoder(bytes.NewReader(data))
	for {
		tok, err := dec.Token()
		if err != nil {
			return fmt.Errorf("parsing workbook.xml: %w", err)
		}
		if _, ok := tok.(xml.StartElement); ok {
			break
		}
	}
	for {
		start := dec.InputOffset()
		tok, err := dec.Token()
		if err == io.EOF {
			return nil
		}
		if err != nil {
			return fmt.Errorf("parsing workbook.xml: %w", err)
		}
		se, ok := tok.(xml.StartElement)
		if !ok {
			if _, end := tok.(xml.EndElement); end {
				return nil
			}
			continue
		}
		switch se.Name.Local {
		case "workbookPr":
			wb.date1904 = parseBoolAttr(attrLocal(se, "date1904"), false)
			if bytes.Contains(data, []byte(`conformance="strict"`)) || bytes.Contains(data, []byte("http://purl.oclc.org/ooxml/spreadsheetml/main")) {
				wb.loadedStrict = true
			}
			if err := skipElement(dec); err != nil {
				return err
			}
		case "workbookProtection":
			wb.workbookProtection = &WorkbookProtection{
				LockStructure: parseBoolAttr(attrLocal(se, "lockStructure"), false),
				LockWindows:   parseBoolAttr(attrLocal(se, "lockWindows"), false),
				revisions:     parseBoolAttr(attrLocal(se, "lockRevision"), false),
			}
			if p := attrLocal(se, "workbookPassword"); p != "" {
				wb.workbookProtection.Password = p
			}
			if err := skipElement(dec); err != nil {
				return err
			}
		case "bookViews":
			if err := parseBookViews(dec, wb); err != nil {
				return err
			}
		case "sheets":
			if err := parseSheets(dec, wb); err != nil {
				return err
			}
		case "definedNames":
			if err := parseDefinedNames(dec, wb); err != nil {
				return err
			}
		default:
			raw, err := elementOuterXML(dec, data, start)
			if err != nil {
				return err
			}
			if strings.TrimSpace(raw) != "" {
				wb.preserved = append(wb.preserved, rawXML{Name: se.Name.Local, XML: strings.TrimSpace(raw)})
			}
		}
	}
}

func parseBookViews(dec *xml.Decoder, wb *Workbook) error {
	for {
		tok, err := dec.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "workbookView" {
				if v := attrLocal(t, "activeTab"); v != "" {
					if n, err := strconv.Atoi(v); err == nil && n >= 0 {
						wb.activeSheet = n
					}
				}
			}
			if err := skipElement(dec); err != nil {
				return err
			}
		case xml.EndElement:
			if t.Name.Local == "bookViews" {
				return nil
			}
		}
	}
}

func parseSheets(dec *xml.Decoder, wb *Workbook) error {
	for {
		tok, err := dec.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "sheet" {
				name := attrLocal(t, "name")
				rid := attrLocal(t, "id")
				state := attrLocal(t, "state")
				if rel, ok := relByID(wb.workbookRels, rid); ok {
					ws, err := wb.AddSheet(name)
					if err != nil {
						return err
					}
					ws.relID = rid
					ws.relType = rel.Type
					ws.partName = resolvePart("xl/workbook.xml", rel.Target)
					if state == "hidden" || state == "veryHidden" {
						ws.state = state
					}
				}
			}
			if err := skipElement(dec); err != nil {
				return err
			}
		case xml.EndElement:
			if t.Name.Local == "sheets" {
				return nil
			}
		}
	}
}

func relByID(rels []relationship, id string) (relationship, bool) {
	for _, rel := range rels {
		if rel.ID == id {
			return rel, true
		}
	}
	return relationship{}, false
}

func relByIDMust(rels []relationship, id string) relationship {
	rel, _ := relByID(rels, id)
	return rel
}

func parseDefinedNames(dec *xml.Decoder, wb *Workbook) error {
	for {
		tok, err := dec.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local != "definedName" {
				if err := skipElement(dec); err != nil {
					return err
				}
				continue
			}
			var dn struct {
				Name   string `xml:"name,attr"`
				Local  string `xml:"localSheetId,attr"`
				Hidden string `xml:"hidden,attr"`
				Ref    string `xml:",chardata"`
			}
			if err := dec.DecodeElement(&dn, &t); err != nil {
				return err
			}
			item := definedName{
				Name:   dn.Name,
				Ref:    strings.TrimSpace(dn.Ref),
				Hidden: parseBoolAttr(dn.Hidden, false),
			}
			if dn.Local != "" {
				if n, err := strconv.Atoi(dn.Local); err == nil {
					item.HasLocal = true
					item.LocalSheetID = n
				}
			}
			wb.addDefinedName(item)
			if item.HasLocal {
				if ws, err := wb.GetSheet(item.LocalSheetID); err == nil {
					switch item.Name {
					case "_xlnm.Print_Area":
						ws.GetPageSetup().SetPrintArea(item.Ref)
					case "_xlnm.Print_Titles":
						rows, cols := splitPrintTitles(item.Ref)
						if rows != "" {
							ws.GetPageSetup().SetRepeatRows(rows)
						}
						if cols != "" {
							ws.GetPageSetup().SetRepeatColumns(cols)
						}
					}
				}
			}
		case xml.EndElement:
			if t.Name.Local == "definedNames" {
				return nil
			}
		}
	}
}

func splitPrintTitles(ref string) (rows, cols string) {
	for _, part := range strings.Split(ref, ",") {
		part = strings.TrimSpace(part)
		if i := strings.LastIndex(part, "!"); i >= 0 {
			part = part[i+1:]
		}
		part = strings.ReplaceAll(part, "$", "")
		if part == "" {
			continue
		}
		if isRowSpan(part) {
			rows = part
		} else if isColSpan(part) {
			cols = part
		}
	}
	return rows, cols
}

func isRowSpan(s string) bool {
	i := strings.IndexByte(s, ':')
	if i <= 0 {
		return false
	}
	return isDigits(s[:i]) && isDigits(s[i+1:])
}

func isColSpan(s string) bool {
	i := strings.IndexByte(s, ':')
	if i <= 0 {
		return false
	}
	return isColLetters(s[:i]) && isColLetters(s[i+1:])
}

func isDigits(s string) bool {
	if s == "" {
		return false
	}
	for i := 0; i < len(s); i++ {
		if s[i] < '0' || s[i] > '9' {
			return false
		}
	}
	return true
}

func isColLetters(s string) bool {
	if s == "" {
		return false
	}
	for i := 0; i < len(s); i++ {
		if !isASCIILetter(s[i]) {
			return false
		}
	}
	return true
}

// isDateFormatID returns true if the given numFmtId corresponds to a
// built-in Excel date/time format.
// Built-in date format IDs: 14-22, 27-36, 45-47, 50-58.
func isDateFormatID(id int) bool {
	if isTimeOnlyFormatID(id) {
		return false
	}
	return (id >= 14 && id <= 22) ||
		(id >= 27 && id <= 36) ||
		(id >= 45 && id <= 47) ||
		(id >= 50 && id <= 58)
}

// isTimeOnlyFormatID reports built-in formats that are times or elapsed
// durations (h:mm, mm:ss, [h]:mm:ss) rather than calendar dates.
func isTimeOnlyFormatID(id int) bool {
	return (id >= 18 && id <= 21) || (id >= 45 && id <= 47)
}

// isTimeOnlyFormat reports custom formats such as h:mm:ss or [h]:mm:ss.
// Calendar formats that also contain a time (yyyy-mm-dd hh:mm) are dates.
func isTimeOnlyFormat(code string) bool {
	if code == "" {
		return false
	}
	cleaned := stripBracketedExceptElapsed(stripQuotedAndEscaped(strings.ToLower(code)))
	if strings.ContainsAny(cleaned, "yd") || strings.Contains(cleaned, "mmm") {
		return false
	}
	if strings.ContainsAny(cleaned, "hs") {
		return true
	}
	// Elapsed minutes, written [m] or [mm]. A bare "m" is a month.
	lower := strings.ToLower(code)
	return strings.Contains(lower, "[m]") || strings.Contains(lower, "[mm]")
}

// looksLikeDateFormat checks if a custom format code string contains
// date/time tokens like y, m, d, h, s. It strips quoted strings, escaped
// characters, and bracketed color/locale codes before checking, so that
// tokens inside literals (e.g. 0.00"mm") don't cause false positives.
func looksLikeDateFormat(code string) bool {
	if code == "" {
		return false
	}
	lower := strings.ToLower(code)
	if lower == "general" || lower == "@" {
		return false
	}
	cleaned := stripQuotedAndEscaped(lower)
	cleaned = stripBracketedExceptElapsed(cleaned)
	hasDate := strings.ContainsAny(cleaned, "yd")
	hasTime := strings.ContainsAny(cleaned, "hs")
	hasAMPM := strings.Contains(cleaned, "am/pm") || strings.Contains(cleaned, "a/p")
	if hasDate || hasAMPM {
		return true
	}
	if hasTime {
		return true
	}
	if strings.Contains(cleaned, "mmm") {
		return true
	}
	return false
}

func stripQuotedAndEscaped(code string) string {
	var b strings.Builder
	b.Grow(len(code))
	i := 0
	for i < len(code) {
		switch code[i] {
		case '"':
			i++
			for i < len(code) && code[i] != '"' {
				i++
			}
			if i < len(code) {
				i++
			}
		case '\\':
			i += 2
		default:
			b.WriteByte(code[i])
			i++
		}
	}
	return b.String()
}

func stripBracketedExceptElapsed(code string) string {
	var b strings.Builder
	b.Grow(len(code))
	i := 0
	for i < len(code) {
		if code[i] == '[' {
			end := strings.IndexByte(code[i:], ']')
			if end == -1 {
				b.WriteByte(code[i])
				i++
				continue
			}
			inner := code[i+1 : i+end]
			switch strings.ToLower(inner) {
			case "h", "hh", "m", "mm", "s", "ss":
				b.WriteString(inner)
			}
			i += end + 1
		} else {
			b.WriteByte(code[i])
			i++
		}
	}
	return b.String()
}

func (wb *Workbook) addDefinedName(dn definedName) {
	wb.definedNameList = append(wb.definedNameList, dn)
	key := dn.Name
	if dn.HasLocal {
		if _, exists := wb.namedRanges[key]; exists {
			if ws, err := wb.GetSheet(dn.LocalSheetID); err == nil {
				key = ws.title + "!" + dn.Name
			} else {
				key = fmt.Sprintf("#%d!%s", dn.LocalSheetID, dn.Name)
			}
		}
	}
	if _, exists := wb.namedRanges[key]; !exists {
		wb.namedRanges[key] = dn.Ref
	}
}

func readCoreProperties(pkg *ooxmlPackage, wb *Workbook) {
	data, ok := pkg.file("docProps/core.xml")
	if !ok {
		return
	}
	type xmlCore struct {
		Creator        string `xml:"creator"`
		Title          string `xml:"title"`
		Subject        string `xml:"subject"`
		Description    string `xml:"description"`
		Keywords       string `xml:"keywords"`
		Category       string `xml:"category"`
		LastModifiedBy string `xml:"lastModifiedBy"`
	}
	var core xmlCore
	if err := xml.Unmarshal(stripBOM(data), &core); err != nil {
		return
	}
	wb.Properties.Creator = core.Creator
	wb.Properties.Title = core.Title
	wb.Properties.Subject = core.Subject
	wb.Properties.Description = core.Description
	wb.Properties.Keywords = core.Keywords
	wb.Properties.Category = core.Category
	wb.Properties.LastModifiedBy = core.LastModifiedBy
}
