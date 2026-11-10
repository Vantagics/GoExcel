package gospreadsheet

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"strconv"
	"strings"
	"time"
)

type pendingCell struct {
	row, col   int
	t, s       string
	value      string
	formula    string
	fType      string
	fRef       string
	fSi        string
	fR1        string
	fR2        string
	fDT2D      string
	fDTR       string
	inline     string
	inlineRich *RichText
	hasInline  bool
	cm, vm     string
}

type sharedMaster struct {
	formula  string
	row, col int
}

type rowMeta struct {
	row    int
	height float64
	custom bool
	hidden bool
}

func readWorksheet(ws *Worksheet, data []byte, pkg *ooxmlPackage, sst []sstItem, styles *styleTable, wb *Workbook, theme []string) error {
	ws.extraRootAttrs = extraNamespaceDecls(firstStartTag(data, "worksheet"))
	if relData, ok := pkg.file(relsPathFor(ws.partName)); ok {
		if rels, err := parseRelationships(relData); err == nil {
			ws.sheetRels = rels
		}
	}
	data = stripBOM(data)
	dec := xml.NewDecoder(bytes.NewReader(data))
	for {
		tok, err := dec.Token()
		if err != nil {
			return fmt.Errorf("parsing sheet xml: %w", err)
		}
		if _, ok := tok.(xml.StartElement); ok {
			break
		}
	}

	var pending []pendingCell
	masters := map[string]sharedMaster{}
	var rows []rowMeta

	for {
		start := dec.InputOffset()
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return fmt.Errorf("parsing sheet xml: %w", err)
		}
		se, ok := tok.(xml.StartElement)
		if !ok {
			if _, end := tok.(xml.EndElement); end {
				break
			}
			continue
		}
		switch se.Name.Local {
		case "sheetPr":
			raw, err := elementOuterXML(dec, data, start)
			if err != nil {
				return err
			}
			ws.sheetPrRaw = strings.TrimSpace(raw)
			applySheetPr(ws, theme)
		case "dimension":
			if err := skipElement(dec); err != nil {
				return err
			}
		case "sheetViews":
			if err := parseSheetViews(dec, ws); err != nil {
				return err
			}
		case "sheetFormatPr":
			if v := attrLocal(se, "defaultRowHeight"); v != "" {
				ws.defaultRowHeight, _ = strconv.ParseFloat(v, 64)
			}
			if err := skipElement(dec); err != nil {
				return err
			}
		case "cols":
			if err := parseCols(dec, ws, styles); err != nil {
				return err
			}
		case "sheetData":
			cells, meta, err := parseSheetData(dec, theme)
			if err != nil {
				return err
			}
			pending = append(pending, cells...)
			rows = append(rows, meta...)
		case "mergeCells":
			if err := parseMergeCells(dec, ws); err != nil {
				return err
			}
		case "hyperlinks":
			if err := parseHyperlinks(dec, ws); err != nil {
				return err
			}
		case "conditionalFormatting":
			raw, err := elementOuterXML(dec, data, start)
			if err != nil {
				return err
			}
			raw = strings.TrimSpace(raw)
			if raw != "" {
				ws.preserved = append(ws.preserved, rawXML{Name: "conditionalFormatting", XML: raw})
				parseConditionalFormatting(ws, raw)
			}
		case "dataValidations":
			if err := parseDataValidations(dec, ws); err != nil {
				return err
			}
		case "autoFilter":
			raw, err := elementOuterXML(dec, data, start)
			if err != nil {
				return err
			}
			ws.autoFilterRaw = strings.TrimSpace(raw)
			parseAutoFilter(ws, ws.autoFilterRaw)
		case "sheetProtection":
			parseSheetProtection(ws, se)
			if err := skipElement(dec); err != nil {
				return err
			}
		case "printOptions":
			ensurePageSetup(ws)
			ps := ws.pageSetup
			ps.PrintGridlines = parseBoolAttr(attrLocal(se, "gridLines"), false)
			ps.PrintHeadings = parseBoolAttr(attrLocal(se, "headings"), false)
			ps.CenterHorizontally = parseBoolAttr(attrLocal(se, "horizontalCentered"), false)
			ps.CenterVertically = parseBoolAttr(attrLocal(se, "verticalCentered"), false)
			if err := skipElement(dec); err != nil {
				return err
			}
		case "pageMargins":
			ensurePageSetup(ws)
			m := ws.pageSetup.Margins
			if v := attrLocal(se, "left"); v != "" {
				m.Left, _ = strconv.ParseFloat(v, 64)
			}
			if v := attrLocal(se, "right"); v != "" {
				m.Right, _ = strconv.ParseFloat(v, 64)
			}
			if v := attrLocal(se, "top"); v != "" {
				m.Top, _ = strconv.ParseFloat(v, 64)
			}
			if v := attrLocal(se, "bottom"); v != "" {
				m.Bottom, _ = strconv.ParseFloat(v, 64)
			}
			if v := attrLocal(se, "header"); v != "" {
				m.Header, _ = strconv.ParseFloat(v, 64)
			}
			if v := attrLocal(se, "footer"); v != "" {
				m.Footer, _ = strconv.ParseFloat(v, 64)
			}
			ws.pageSetup.Margins = m
			if err := skipElement(dec); err != nil {
				return err
			}
		case "pageSetup":
			ensurePageSetup(ws)
			ps := ws.pageSetup
			if v := attrLocal(se, "paperSize"); v != "" {
				n, _ := strconv.Atoi(v)
				ps.PaperSize = PaperSize(n)
			}
			if v := attrLocal(se, "orientation"); v != "" {
				ps.Orientation = Orientation(v)
			}
			if v := attrLocal(se, "scale"); v != "" {
				ps.Scale, _ = strconv.Atoi(v)
			}
			if v := attrLocal(se, "fitToWidth"); v != "" {
				ps.FitToWidth, _ = strconv.Atoi(v)
			}
			if v := attrLocal(se, "fitToHeight"); v != "" {
				ps.FitToHeight, _ = strconv.Atoi(v)
			}
			ps.rid = attrLocal(se, "id")
			if err := skipElement(dec); err != nil {
				return err
			}
		case "headerFooter":
			if err := parseHeaderFooter(dec, ws); err != nil {
				return err
			}
		default:
			raw, err := elementOuterXML(dec, data, start)
			if err != nil {
				return err
			}
			raw = strings.TrimSpace(raw)
			if raw != "" {
				ws.preserved = append(ws.preserved, rawXML{Name: se.Name.Local, XML: raw})
			}
		}
	}

	for i := range pending {
		c := &pending[i]
		if c.fType == "shared" {
			if c.formula != "" {
				masters[c.fSi] = sharedMaster{formula: c.formula, row: c.row, col: c.col}
			}
		}
	}
	for i := range pending {
		c := &pending[i]
		if c.fType == "shared" && c.formula == "" {
			if m, ok := masters[c.fSi]; ok {
				c.formula = shiftFormula(m.formula, c.col-m.col, c.row-m.row)
			}
		}
		if c.fType == "array" && c.formula != "" && !strings.HasPrefix(c.formula, "{") {
			c.formula = "{" + c.formula + "}"
		}
	}
	for _, meta := range rows {
		if meta.custom && meta.height > 0 {
			ws.SetRowHeight(meta.row, meta.height)
		}
		if meta.hidden {
			ws.SetRowHidden(meta.row, true)
		}
	}
	for i := range pending {
		applyPendingCell(ws, &pending[i], sst, styles, wb)
	}
	readComments(ws, pkg)
	return nil
}

func ensurePageSetup(ws *Worksheet) {
	if ws.pageSetup == nil {
		ws.pageSetup = NewPageSetup()
		ws.pageSetup.Scale = 0
		ws.pageSetup.PaperSize = 0
		ws.pageSetup.Orientation = ""
	}
}

func applySheetPr(ws *Worksheet, theme []string) {
	if ws.sheetPrRaw == "" {
		return
	}
	var pr struct {
		TabColor *struct {
			RGB     string `xml:"rgb,attr"`
			Theme   string `xml:"theme,attr"`
			Indexed string `xml:"indexed,attr"`
			Tint    string `xml:"tint,attr"`
		} `xml:"tabColor"`
	}
	if err := xml.Unmarshal([]byte(ws.sheetPrRaw), &pr); err != nil || pr.TabColor == nil {
		return
	}
	tint, _ := strconv.ParseFloat(pr.TabColor.Tint, 64)
	color := resolveColor(pr.TabColor.RGB, pr.TabColor.Theme, pr.TabColor.Indexed, tint, theme)
	if color != "" {
		if len(color) == 6 {
			color = "FF" + color
		}
		ws.tabColor = color
		ws.loadedTabColor = color
	}
}

func parseSheetViews(dec *xml.Decoder, ws *Worksheet) error {
	for {
		tok, err := dec.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "sheetView" {
				zoom := 100
				if v := attrLocal(t, "zoomScale"); v != "" {
					if n, err := strconv.Atoi(v); err == nil {
						zoom = n
					}
				}
				grid := parseBoolAttr(attrLocal(t, "showGridLines"), true)
				top := attrLocal(t, "topLeftCell")
				ws.rightToLeft = parseBoolAttr(attrLocal(t, "rightToLeft"), false)
				if zoom != 100 || !grid || top != "" {
					ws.sheetView = &SheetView{
						ZoomScale:         uint(zoom),
						ShowGridlines:     grid,
						ShowRowColHeaders: parseBoolAttr(attrLocal(t, "showRowColHeaders"), true),
						TopLeftCell:       top,
					}
				}
				if err := parseSheetViewChildren(dec, ws); err != nil {
					return err
				}
				continue
			}
			if err := skipElement(dec); err != nil {
				return err
			}
		case xml.EndElement:
			if t.Name.Local == "sheetViews" {
				return nil
			}
		}
	}
}

func parseSheetViewChildren(dec *xml.Decoder, ws *Worksheet) error {
	for {
		tok, err := dec.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "pane" {
				xSplit, _ := strconv.ParseFloat(attrLocal(t, "xSplit"), 64)
				ySplit, _ := strconv.ParseFloat(attrLocal(t, "ySplit"), 64)
				state := attrLocal(t, "state")
				if state == "frozen" || state == "frozenSplit" || xSplit > 0 || ySplit > 0 {
					col := int(xSplit)
					row := int(ySplit) + 1
					if row < 1 {
						row = 1
					}
					name, err := ColumnIndexToName(col)
					if err == nil {
						ws.frozen = &CellReference{Column: name, ColumnIdx: col, Row: row}
					}
				}
			}
			if err := skipElement(dec); err != nil {
				return err
			}
		case xml.EndElement:
			if t.Name.Local == "sheetView" {
				return nil
			}
		}
	}
}

func parseCols(dec *xml.Decoder, ws *Worksheet, styles *styleTable) error {
	for {
		tok, err := dec.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "col" {
				minV, _ := strconv.Atoi(attrLocal(t, "min"))
				maxV, _ := strconv.Atoi(attrLocal(t, "max"))
				if minV < 1 {
					minV = 1
				}
				if maxV < minV {
					maxV = minV
				}
				sp := colSpan{Min: minV, Max: maxV}
				if w := attrLocal(t, "width"); w != "" {
					sp.Width, _ = strconv.ParseFloat(w, 64)
					sp.CustomWidth = true
				}
				if attrLocal(t, "customWidth") == "0" || strings.EqualFold(attrLocal(t, "customWidth"), "false") {
					// Width was present but not marked custom. Keep it on the span
					// so it can be written back, but don't override the default API
					// width for enormous trailing spans below.
					if attrLocal(t, "width") == "" {
						sp.CustomWidth = false
					}
				}
				sp.Hidden = parseBoolAttr(attrLocal(t, "hidden"), false)
				sp.BestFit = parseBoolAttr(attrLocal(t, "bestFit"), false)
				if s := attrLocal(t, "style"); s != "" && styles != nil {
					if idx, err := strconv.Atoi(s); err == nil && idx >= 0 && idx < len(styles.xfs) {
						sp.Style = cloneStyle(styles.xfs[idx].style)
					}
				}
				ws.colSpans = append(ws.colSpans, sp)
				if maxV-minV <= 256 {
					for c := minV; c <= maxV; c++ {
						if sp.Width != 0 || sp.Hidden {
							if attrLocal(t, "width") != "" {
								ws.colWidths[c-1] = sp.Width
							}
						}
						if sp.Hidden {
							ws.hiddenCols[c-1] = true
						}
					}
				}
			}
			if err := skipElement(dec); err != nil {
				return err
			}
		case xml.EndElement:
			if t.Name.Local == "cols" {
				return nil
			}
		}
	}
}

func parseSheetData(dec *xml.Decoder, theme []string) ([]pendingCell, []rowMeta, error) {
	var cells []pendingCell
	var rows []rowMeta
	nextRow := 1
	for {
		tok, err := dec.Token()
		if err != nil {
			return nil, nil, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local != "row" {
				if err := skipElement(dec); err != nil {
					return nil, nil, err
				}
				continue
			}
			rowNum := nextRow
			if v := attrLocal(t, "r"); v != "" {
				if n, err := strconv.Atoi(v); err == nil && n > 0 {
					rowNum = n
				}
			}
			nextRow = rowNum + 1
			meta := rowMeta{row: rowNum - 1}
			if v := attrLocal(t, "ht"); v != "" {
				meta.height, _ = strconv.ParseFloat(v, 64)
				meta.custom = parseBoolAttr(attrLocal(t, "customHeight"), true)
			}
			meta.hidden = parseBoolAttr(attrLocal(t, "hidden"), false)
			rows = append(rows, meta)
			rowCells, err := parseRow(dec, rowNum-1, theme)
			if err != nil {
				return nil, nil, err
			}
			cells = append(cells, rowCells...)
		case xml.EndElement:
			if t.Name.Local == "sheetData" {
				return cells, rows, nil
			}
		}
	}
}

func parseRow(dec *xml.Decoder, row int, theme []string) ([]pendingCell, error) {
	var cells []pendingCell
	nextCol := 0
	for {
		tok, err := dec.Token()
		if err != nil {
			return nil, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local != "c" {
				if err := skipElement(dec); err != nil {
					return nil, err
				}
				continue
			}
			col := nextCol
			if ref := attrLocal(t, "r"); ref != "" {
				cr, err := ParseCellReference(ref)
				if err != nil {
					if err := skipElement(dec); err != nil {
						return nil, err
					}
					continue
				}
				row = cr.Row - 1
				col = cr.ColumnIdx
			}
			nextCol = col + 1
			cell, err := parseCell(dec, t, row, col, theme)
			if err != nil {
				return nil, err
			}
			cells = append(cells, cell)
		case xml.EndElement:
			if t.Name.Local == "row" {
				return cells, nil
			}
		}
	}
}

func parseCell(dec *xml.Decoder, start xml.StartElement, row, col int, theme []string) (pendingCell, error) {
	pc := pendingCell{row: row, col: col, t: attrLocal(start, "t"), s: attrLocal(start, "s"), cm: attrLocal(start, "cm"), vm: attrLocal(start, "vm")}
	for {
		tok, err := dec.Token()
		if err != nil {
			return pc, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "f":
				pc.fType = attrLocal(t, "t")
				pc.fRef = attrLocal(t, "ref")
				pc.fSi = attrLocal(t, "si")
				pc.fR1 = attrLocal(t, "r1")
				pc.fR2 = attrLocal(t, "r2")
				pc.fDT2D = attrLocal(t, "dt2D")
				pc.fDTR = attrLocal(t, "dtr")
				var text strings.Builder
				if err := readCharData(dec, &text, "f"); err != nil {
					return pc, err
				}
				pc.formula = strings.TrimSpace(text.String())
			case "v":
				var text strings.Builder
				if err := readCharData(dec, &text, "v"); err != nil {
					return pc, err
				}
				pc.value = text.String()
			case "is":
				item, err := parseStringItem(dec, theme)
				if err != nil {
					return pc, err
				}
				pc.hasInline = true
				pc.inline = item.text
				pc.inlineRich = item.rich
			default:
				if err := skipElement(dec); err != nil {
					return pc, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "c" {
				return pc, nil
			}
		}
	}
}

func readCharData(dec *xml.Decoder, b *strings.Builder, end string) error {
	depth := 1
	for depth > 0 {
		tok, err := dec.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
		case xml.EndElement:
			depth--
		case xml.CharData:
			if depth == 1 {
				b.Write(t)
			}
		}
	}
	return nil
}

func applyPendingCell(ws *Worksheet, pc *pendingCell, sst []sstItem, styles *styleTable, wb *Workbook) {
	var st *Style
	isDate := false
	if pc.s != "" && styles != nil {
		if idx, err := strconv.Atoi(pc.s); err == nil && idx >= 0 && idx < len(styles.xfs) {
			st = cloneStyle(styles.xfs[idx].style)
			isDate = styles.xfs[idx].date
		}
	}
	hasFormula := pc.formula != "" || pc.fType == "array" || pc.fType == "shared" || pc.fType == "dataTable"
	hasValue := pc.value != "" || pc.hasInline || pc.t == "b"
	if !hasFormula && !hasValue && st == nil && pc.cm == "" && pc.vm == "" {
		return
	}
	cell := ws.GetCell(pc.row, pc.col)
	defer func() {
		cell.cellMeta = pc.cm
		cell.valueMeta = pc.vm
	}()
	if st != nil {
		cell.Style = st
	}
	if hasFormula {
		if pc.fType == "dataTable" {
			cell.SetFormula(dataTableFormula(pc.fR1, pc.fR2))
			cell.dt = &dataTable{
				Ref:  pc.fRef,
				R1:   pc.fR1,
				R2:   pc.fR2,
				TwoD: parseBoolAttr(pc.fDT2D, false),
				Row:  parseBoolAttr(pc.fDTR, false),
			}
			applyCached(cell, pc, false, wb)
			return
		}
		formula := pc.formula
		if strings.HasPrefix(formula, "=") {
			formula = strings.TrimPrefix(formula, "=")
		}
		cell.SetFormula(formula)
		cell.formulaRef = pc.fRef
		cell.arrayAnchor = pc.fType == "array"
		applyCached(cell, pc, isDate, wb)
		return
	}
	switch pc.t {
	case "s":
		idx, err := strconv.Atoi(strings.TrimSpace(pc.value))
		if err == nil && idx >= 0 && idx < len(sst) {
			cell.SetValue(sst[idx].text)
			if sst[idx].rich != nil {
				rt := *sst[idx].rich
				runs := append([]RichTextRun(nil), sst[idx].rich.Runs...)
				rt.Runs = runs
				cell.RichText = &rt
			}
		}
	case "inlineStr":
		cell.SetValue(pc.inline)
		if pc.inlineRich != nil {
			cell.RichText = pc.inlineRich
		}
	case "b":
		cell.SetValue(pc.value == "1" || strings.EqualFold(pc.value, "true"))
	case "e":
		cell.Value = pc.value
		cell.Type = CellTypeError
	case "str":
		cell.SetValue(unescapeOOXMLText(pc.value))
	case "d":
		if t, ok := parseExcelDate(pc.value); ok {
			cell.SetValue(t)
		} else {
			cell.SetValue(pc.value)
		}
	default:
		if pc.value == "" {
			return
		}
		if f, err := strconv.ParseFloat(pc.value, 64); err == nil {
			if isDate {
				cell.SetValue(wb.timeFromSerial(f))
			} else {
				cell.SetValue(f)
			}
		} else {
			cell.SetValue(pc.value)
		}
	}
}

func applyCached(cell *Cell, pc *pendingCell, isDate bool, wb *Workbook) {
	if pc.t == "str" || pc.t == "e" {
		cell.cachedType = pc.t
		cell.Value = pc.value
		if pc.t == "str" {
			cell.Value = unescapeOOXMLText(pc.value)
		}
		return
	}
	if pc.value == "" {
		return
	}
	switch pc.t {
	case "b":
		cell.Value = pc.value == "1" || strings.EqualFold(pc.value, "true")
		cell.cachedType = "b"
	case "e":
		cell.Value = pc.value
		cell.cachedType = "e"
	case "str":
		cell.Value = unescapeOOXMLText(pc.value)
		cell.cachedType = "str"
	case "s":
		cell.cachedType = "str"
		cell.Value = pc.value
	default:
		if f, err := strconv.ParseFloat(pc.value, 64); err == nil {
			if isDate {
				cell.Value = wb.timeFromSerial(f)
				cell.cachedType = "d"
			} else {
				cell.Value = f
				cell.cachedType = "n"
			}
		} else if pc.value != "" {
			cell.Value = unescapeOOXMLText(pc.value)
			cell.cachedType = "str"
		}
	}
}

func parseExcelDate(v string) (time.Time, bool) {
	layouts := []string{
		"2006-01-02",
		"2006-01-02T15:04:05",
		"2006-01-02T15:04:05Z",
		"2006-01-02T15:04:05.000",
		time.RFC3339,
	}
	for _, layout := range layouts {
		if t, err := time.Parse(layout, v); err == nil {
			return t, true
		}
	}
	return time.Time{}, false
}

func parseMergeCells(dec *xml.Decoder, ws *Worksheet) error {
	for {
		tok, err := dec.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "mergeCell" {
				if ref := attrLocal(t, "ref"); ref != "" {
					if !strings.Contains(ref, ":") {
						ref = ref + ":" + ref
					}
					_ = ws.MergeCells(ref)
				}
			}
			if err := skipElement(dec); err != nil {
				return err
			}
		case xml.EndElement:
			if t.Name.Local == "mergeCells" {
				return nil
			}
		}
	}
}

func parseHyperlinks(dec *xml.Decoder, ws *Worksheet) error {
	for {
		tok, err := dec.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "hyperlink" {
				ref := attrLocal(t, "ref")
				target := ""
				if id := attrLocal(t, "id"); id != "" {
					if rel, ok := relByID(ws.sheetRels, id); ok {
						target = rel.Target
					}
				}
				if loc := attrLocal(t, "location"); loc != "" && target == "" {
					target = loc
				}
				if ref != "" && target != "" {
					link := NewHyperlink(target)
					if tip := attrLocal(t, "tooltip"); tip != "" {
						link.SetTooltip(tip)
					}
					for _, addr := range expandRef(ref) {
						if cell, err := ws.GetCellByName(addr); err == nil {
							// Each cell needs its own hyperlink value; the
							// relationship id is filled in per cell at write time.
							cp := *link
							cell.SetHyperlink(&cp)
						}
					}
				}
			}
			if err := skipElement(dec); err != nil {
				return err
			}
		case xml.EndElement:
			if t.Name.Local == "hyperlinks" {
				return nil
			}
		}
	}
}

func parseConditionalFormatting(ws *Worksheet, raw string) {
	var block struct {
		SQRef string `xml:"sqref,attr"`
		Rules []struct {
			Type       string   `xml:"type,attr"`
			Operator   string   `xml:"operator,attr"`
			Priority   int      `xml:"priority,attr"`
			StopIfTrue string   `xml:"stopIfTrue,attr"`
			Formula    []string `xml:"formula"`
		} `xml:"cfRule"`
	}
	if err := xml.Unmarshal([]byte(raw), &block); err != nil {
		return
	}
	cf := NewConditionalFormatting(block.SQRef)
	cf.fromFile = true
	for _, rule := range block.Rules {
		cf.Rules = append(cf.Rules, ConditionalRule{
			Type:       ConditionalType(rule.Type),
			Operator:   ConditionalOperator(rule.Operator),
			Formula:    append([]string(nil), rule.Formula...),
			Priority:   rule.Priority,
			StopIfTrue: parseBoolAttr(rule.StopIfTrue, false),
		})
	}
	ws.conditionalFormats = append(ws.conditionalFormats, cf)
}

func parseDataValidations(dec *xml.Decoder, ws *Worksheet) error {
	for {
		tok, err := dec.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local != "dataValidation" {
				if err := skipElement(dec); err != nil {
					return err
				}
				continue
			}
			dv := NewDataValidation(attrLocal(t, "sqref"))
			if typ := attrLocal(t, "type"); typ != "" {
				dv.Type = ValidationType(typ)
			}
			if op := attrLocal(t, "operator"); op != "" {
				dv.Operator = ValidationOperator(op)
			}
			dv.AllowBlank = parseBoolAttr(attrLocal(t, "allowBlank"), true)
			dv.ShowInputMsg = parseBoolAttr(attrLocal(t, "showInputMessage"), true)
			dv.ShowErrorMsg = parseBoolAttr(attrLocal(t, "showErrorMessage"), true)
			if es := attrLocal(t, "errorStyle"); es != "" {
				dv.ErrorStyle = ValidationErrorStyle(es)
			}
			dv.ErrorTitle = attrLocal(t, "errorTitle")
			dv.ErrorMessage = attrLocal(t, "error")
			dv.PromptTitle = attrLocal(t, "promptTitle")
			dv.PromptMessage = attrLocal(t, "prompt")
			f1, f2, err := readValidationFormulas(dec)
			if err != nil {
				return err
			}
			dv.Formula1 = f1
			dv.Formula2 = f2
			ws.dataValidations = append(ws.dataValidations, dv)
		case xml.EndElement:
			if t.Name.Local == "dataValidations" {
				return nil
			}
		}
	}
}

func readValidationFormulas(dec *xml.Decoder) (string, string, error) {
	var f1, f2 string
	for {
		tok, err := dec.Token()
		if err != nil {
			return "", "", err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			var b strings.Builder
			if err := readCharData(dec, &b, t.Name.Local); err != nil {
				return "", "", err
			}
			switch t.Name.Local {
			case "formula1":
				f1 = b.String()
			case "formula2":
				f2 = b.String()
			}
		case xml.EndElement:
			if t.Name.Local == "dataValidation" {
				return f1, f2, nil
			}
		}
	}
}

func parseAutoFilter(ws *Worksheet, raw string) {
	if raw == "" {
		return
	}
	var block struct {
		Ref string `xml:"ref,attr"`
		Col []struct {
			ColID   int `xml:"colId,attr"`
			Filters *struct {
				Filter []struct {
					Val string `xml:"val,attr"`
				} `xml:"filter"`
			} `xml:"filters"`
			Custom *struct {
				Filter []struct {
					Op  string `xml:"operator,attr"`
					Val string `xml:"val,attr"`
				} `xml:"customFilter"`
			} `xml:"customFilters"`
		} `xml:"filterColumn"`
	}
	if err := xml.Unmarshal([]byte(raw), &block); err != nil {
		return
	}
	af := NewAutoFilter(block.Ref)
	for _, col := range block.Col {
		fc := AutoFilterColumn{ColumnIndex: col.ColID, ShowButton: true}
		if col.Filters != nil {
			fc.FilterType = FilterValues
			for _, f := range col.Filters.Filter {
				fc.Values = append(fc.Values, f.Val)
			}
		}
		if col.Custom != nil {
			fc.FilterType = FilterCustom
			for _, f := range col.Custom.Filter {
				op := f.Op
				if op == "" {
					op = "equal"
				}
				fc.Conditions = append(fc.Conditions, FilterCondition{Operator: FilterOperator(op), Value: f.Val})
			}
		}
		af.Columns = append(af.Columns, fc)
	}
	ws.autoFilter = af
}

func parseSheetProtection(ws *Worksheet, se xml.StartElement) {
	sp := NewSheetProtection()
	sp.Sheet = parseBoolAttr(attrLocal(se, "sheet"), false)
	sp.Objects = parseBoolAttr(attrLocal(se, "objects"), false)
	sp.Scenarios = parseBoolAttr(attrLocal(se, "scenarios"), false)
	// OOXML allow-flags default to true (permitted). The library stores true
	// when the action is denied.
	sp.FormatCells = !parseBoolAttr(attrLocal(se, "formatCells"), true)
	sp.FormatColumns = !parseBoolAttr(attrLocal(se, "formatColumns"), true)
	sp.FormatRows = !parseBoolAttr(attrLocal(se, "formatRows"), true)
	sp.InsertColumns = !parseBoolAttr(attrLocal(se, "insertColumns"), true)
	sp.InsertRows = !parseBoolAttr(attrLocal(se, "insertRows"), true)
	sp.InsertHyperlinks = !parseBoolAttr(attrLocal(se, "insertHyperlinks"), true)
	sp.DeleteColumns = !parseBoolAttr(attrLocal(se, "deleteColumns"), true)
	sp.DeleteRows = !parseBoolAttr(attrLocal(se, "deleteRows"), true)
	sp.SelectLockedCells = !parseBoolAttr(attrLocal(se, "selectLockedCells"), true)
	sp.Sort = !parseBoolAttr(attrLocal(se, "sort"), true)
	sp.AutoFilter = !parseBoolAttr(attrLocal(se, "autoFilter"), true)
	sp.PivotTables = !parseBoolAttr(attrLocal(se, "pivotTables"), true)
	sp.SelectUnlockedCells = !parseBoolAttr(attrLocal(se, "selectUnlockedCells"), true)
	if p := attrLocal(se, "password"); p != "" {
		sp.Password = p
	}
	sp.algoName = attrLocal(se, "algorithmName")
	sp.hashValue = attrLocal(se, "hashValue")
	sp.saltValue = attrLocal(se, "saltValue")
	sp.spinCount = attrLocal(se, "spinCount")
	ws.sheetProtection = sp
}

func parseHeaderFooter(dec *xml.Decoder, ws *Worksheet) error {
	ensurePageSetup(ws)
	hf := &HeaderFooter{}
	// The start tag was consumed. Read children.
	var current string
	var buf strings.Builder
	for {
		tok, err := dec.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			current = t.Name.Local
			buf.Reset()
		case xml.CharData:
			if current != "" {
				buf.Write(t)
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "oddHeader":
				hf.OddHeader = buf.String()
			case "oddFooter":
				hf.OddFooter = buf.String()
			case "evenHeader":
				hf.EvenHeader = buf.String()
			case "evenFooter":
				hf.EvenFooter = buf.String()
			case "headerFooter":
				if hf.OddHeader != "" || hf.OddFooter != "" || hf.EvenHeader != "" || hf.EvenFooter != "" {
					ws.pageSetup.HeaderFooter = hf
				}
				return nil
			}
			current = ""
		}
	}
}

func readComments(ws *Worksheet, pkg *ooxmlPackage) {
	rel, ok := findRel(ws.sheetRels, "comments")
	if !ok {
		return
	}
	part := resolvePart(ws.partName, rel.Target)
	data, ok := pkg.file(part)
	if !ok {
		return
	}
	dec := xml.NewDecoder(bytes.NewReader(stripBOM(data)))
	var authors []string
	inAuthors := false
	for {
		tok, err := dec.Token()
		if err != nil {
			return
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "authors":
				inAuthors = true
			case "author":
				if inAuthors {
					var b strings.Builder
					if err := readCharData(dec, &b, "author"); err != nil {
						return
					}
					authors = append(authors, unescapeOOXMLText(b.String()))
				}
			case "comment":
				ref := attrLocal(t, "ref")
				authorID, _ := strconv.Atoi(attrLocal(t, "authorId"))
				text := readCommentText(dec)
				author := ""
				if authorID >= 0 && authorID < len(authors) {
					author = authors[authorID]
				}
				if strings.HasPrefix(text, author+":") {
					text = strings.TrimPrefix(text, author+":")
					text = strings.TrimPrefix(text, "\n")
				}
				if cell, err := ws.GetCellByName(ref); err == nil {
					cell.SetComment(NewComment(author, text))
				}
			}
		case xml.EndElement:
			if t.Name.Local == "authors" {
				inAuthors = false
			}
		}
	}
}

func readCommentText(dec *xml.Decoder) string {
	var b strings.Builder
	depth := 1
	inT := false
	for depth > 0 {
		tok, err := dec.Token()
		if err != nil {
			return b.String()
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			if t.Name.Local == "t" {
				inT = true
			}
		case xml.EndElement:
			if t.Name.Local == "t" {
				inT = false
			}
			depth--
		case xml.CharData:
			if inT {
				b.Write(t)
			}
		}
	}
	return unescapeOOXMLText(b.String())
}

func dataTableFormula(r1, r2 string) string {
	if r2 != "" {
		return "TABLE(" + r1 + "," + r2 + ")"
	}
	if r1 == "" {
		return "TABLE()"
	}
	return "TABLE(" + r1 + ")"
}

// expandRef lists every cell in a single reference or a rectangular range.
// Very large ranges keep only the top-left cell so a whole-column hyperlink
// cannot materialize a million cells.
func expandRef(ref string) []string {
	ref = strings.TrimSpace(ref)
	if ref == "" {
		return nil
	}
	if !strings.Contains(ref, ":") {
		if _, err := ParseCellReference(ref); err != nil {
			return nil
		}
		return []string{ref}
	}
	start, end, err := ParseRange(ref)
	if err != nil {
		return nil
	}
	rows := end.Row - start.Row + 1
	cols := end.ColumnIdx - start.ColumnIdx + 1
	if rows < 1 || cols < 1 || rows*cols > 4096 {
		return []string{start.Column + strconv.Itoa(start.Row)}
	}
	out := make([]string, 0, rows*cols)
	for r := start.Row; r <= end.Row; r++ {
		for c := start.ColumnIdx; c <= end.ColumnIdx; c++ {
			name, err := CellName(r-1, c)
			if err == nil {
				out = append(out, name)
			}
		}
	}
	return out
}

func (wb *Workbook) timeFromSerial(serial float64) time.Time {
	if wb != nil && wb.date1904 {
		serial += 1462
	}
	return serialToDate(serial)
}

func (wb *Workbook) serialFromTime(t time.Time) float64 {
	s := dateToSerial(t)
	if wb != nil && wb.date1904 {
		s -= 1462
	}
	return s
}
