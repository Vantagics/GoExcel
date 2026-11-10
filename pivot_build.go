package gospreadsheet

import (
	"archive/zip"
	"fmt"
	"strconv"
	"strings"
)

// pivotAuthor is the snapshot used to write a pivot cache that was created
// in memory rather than loaded from a package.
type pivotAuthor struct {
	sourceSheet string
	sourceRef   string
	rowName     string
	dataName    string
	headers     []string
	textCols    []bool
	// records are data rows, one value per header. Numeric columns use num.
	texts  [][]string
	nums   [][]float64
	groups []pivotGroup

	cacheID    int
	cachePart  string
	recPart    string
	tablePart  string
	cacheRelID string
	sheetRelID string
}

type pivotGroup struct {
	key string
	sum float64
}

type slicerAuthor struct {
	pivotName  string
	field      string
	items      []string
	cachePart  string
	slicerPart string
	cacheRelID string
	sheetRelID string
}

type contentOverride struct {
	name string
	ct   string
}

// AddPivotTable creates a pivot cache from source (a range such as "A1:C10"
// or "Sheet1!A1:C10") and a pivot table whose top-left cell is dest.
// rowField and dataField are header names from the first row of source.
// dataField is summed for each distinct rowField value.
func (ws *Worksheet) AddPivotTable(name, source, dest, rowField, dataField string) (*PivotTable, error) {
	if ws == nil {
		return nil, fmt.Errorf("no worksheet")
	}
	name = strings.TrimSpace(name)
	if name == "" {
		return nil, fmt.Errorf("pivot table name is required")
	}
	for _, p := range ws.pivots {
		if p.Name == name {
			return nil, fmt.Errorf("pivot table %q already exists", name)
		}
	}
	srcSheet, ref, err := ws.resolveSource(source)
	if err != nil {
		return nil, err
	}
	start, end, err := ParseRange(ref)
	if err != nil {
		return nil, err
	}
	if end.Row <= start.Row {
		return nil, fmt.Errorf("pivot source %s needs a header row and at least one data row", ref)
	}
	headers := make([]string, 0, end.ColumnIdx-start.ColumnIdx+1)
	for col := start.ColumnIdx; col <= end.ColumnIdx; col++ {
		cell := srcSheet.GetCellIfExists(start.Row-1, col)
		h := ""
		if cell != nil {
			h = strings.TrimSpace(cell.GetStringValue())
		}
		if h == "" {
			h = fmt.Sprintf("Column%d", col+1)
		}
		headers = append(headers, h)
	}
	rowIdx, dataIdx := -1, -1
	for i, h := range headers {
		if h == rowField {
			rowIdx = i
		}
		if h == dataField {
			dataIdx = i
		}
	}
	if rowIdx < 0 {
		return nil, fmt.Errorf("row field %q is not a header in %s", rowField, ref)
	}
	if dataIdx < 0 {
		return nil, fmt.Errorf("data field %q is not a header in %s", dataField, ref)
	}
	if rowIdx == dataIdx {
		return nil, fmt.Errorf("row field and data field must be different columns")
	}
	author := &pivotAuthor{
		sourceSheet: srcSheet.Title(),
		sourceRef:   ref,
		rowName:     rowField,
		dataName:    dataField,
		headers:     headers,
		textCols:    make([]bool, len(headers)),
	}
	for i := range author.textCols {
		author.textCols[i] = i != dataIdx
	}
	order := []string{}
	seen := map[string]int{}
	for r := start.Row + 1; r <= end.Row; r++ {
		texts := make([]string, len(headers))
		nums := make([]float64, len(headers))
		for i := range headers {
			cell := srcSheet.GetCellIfExists(r-1, start.ColumnIdx+i)
			if cell == nil || cell.Value == nil {
				continue
			}
			if i == dataIdx {
				switch v := cell.Value.(type) {
				case float64:
					nums[i] = v
				case int:
					nums[i] = float64(v)
				default:
					return nil, fmt.Errorf("data field %q row %d is not numeric (%T)", dataField, r, cell.Value)
				}
				continue
			}
			texts[i] = cell.GetStringValue()
		}
		key := texts[rowIdx]
		if key == "" {
			continue
		}
		author.texts = append(author.texts, texts)
		author.nums = append(author.nums, nums)
		if _, ok := seen[key]; !ok {
			seen[key] = len(order)
			order = append(order, key)
			author.groups = append(author.groups, pivotGroup{key: key})
		}
		author.groups[seen[key]].sum += nums[dataIdx]
	}
	if len(author.groups) == 0 {
		return nil, fmt.Errorf("pivot source %s has no data rows", ref)
	}
	destRef, err := ParseCellReference(dest)
	if err != nil {
		return nil, err
	}
	// Header + one row per group + grand total, two columns.
	last, _ := CellName(destRef.Row-1+len(author.groups)+1, destRef.ColumnIdx+1)
	first, _ := CellName(destRef.Row-1, destRef.ColumnIdx)
	loc := first + ":" + last
	ws.SetCellValue(first, rowField)
	sumHeader, _ := CellName(destRef.Row-1, destRef.ColumnIdx+1)
	ws.SetCellValue(sumHeader, "Sum of "+dataField)
	var grand float64
	for i, g := range author.groups {
		label, _ := CellName(destRef.Row+i, destRef.ColumnIdx)
		value, _ := CellName(destRef.Row+i, destRef.ColumnIdx+1)
		ws.SetCellValue(label, g.key)
		ws.SetCellValue(value, g.sum)
		grand += g.sum
	}
	totalLabel, _ := CellName(destRef.Row+len(author.groups), destRef.ColumnIdx)
	totalValue, _ := CellName(destRef.Row+len(author.groups), destRef.ColumnIdx+1)
	ws.SetCellValue(totalLabel, "Grand Total")
	ws.SetCellValue(totalValue, grand)

	pt := &PivotTable{
		Name:        name,
		DataCaption: "Values",
		Location:    loc,
		Style:       "PivotStyleLight16",
		RowFields:   []string{rowField},
		CacheFields: append([]string(nil), headers...),
		DataFields: []PivotDataField{{
			Name:       "Sum of " + dataField,
			FieldIndex: dataIdx,
			FieldName:  dataField,
		}},
		author: author,
	}
	pt.origName = pt.Name
	pt.origCaption = pt.DataCaption
	pt.origLoc = pt.Location
	pt.origStyle = pt.Style
	ws.pivots = append(ws.pivots, pt)
	return pt, nil
}

// AddSlicer creates a slicer and slicer cache for a field of a pivot table
// on this sheet. The cache lists the distinct values from the pivot's source.
func (ws *Worksheet) AddSlicer(pivotName, field, caption string) (*Slicer, error) {
	var pt *PivotTable
	for _, p := range ws.pivots {
		if p.Name == pivotName {
			pt = p
			break
		}
	}
	if pt == nil || pt.author == nil {
		return nil, fmt.Errorf("pivot table %q was not created on this sheet", pivotName)
	}
	col := -1
	for i, h := range pt.author.headers {
		if h == field {
			col = i
			break
		}
	}
	if col < 0 {
		return nil, fmt.Errorf("field %q is not in pivot %q", field, pivotName)
	}
	var items []string
	seen := map[string]bool{}
	if col == indexOf(pt.author.headers, pt.author.rowName) {
		for _, g := range pt.author.groups {
			items = append(items, g.key)
		}
	} else {
		for _, row := range pt.author.texts {
			v := row[col]
			if v == "" || seen[v] {
				continue
			}
			seen[v] = true
			items = append(items, v)
		}
	}
	if caption == "" {
		caption = field
	}
	cacheName := "Slicer_" + strings.ReplaceAll(field, " ", "")
	name := cacheName
	s := &Slicer{
		Name:        name,
		Cache:       cacheName,
		Caption:     caption,
		Style:       "SlicerStyleLight1",
		ColumnCount: 1,
		RowHeight:   24130,
		author: &slicerAuthor{
			pivotName: pivotName,
			field:     field,
			items:     items,
		},
	}
	s.origName = s.Name
	s.origCaption = s.Caption
	s.origStyle = s.Style
	s.origColumnCount = s.ColumnCount
	s.origRowHeight = s.RowHeight
	ws.slicers = append(ws.slicers, s)
	if ws.parent != nil {
		ws.parent.slicerCaches = append(ws.parent.slicerCaches, &SlicerCache{
			Name:       cacheName,
			SourceName: field,
			author:     s.author,
		})
	}
	return s, nil
}

func indexOf(ss []string, want string) int {
	for i, s := range ss {
		if s == want {
			return i
		}
	}
	return -1
}

func (ws *Worksheet) resolveSource(source string) (*Worksheet, string, error) {
	source = strings.TrimSpace(source)
	sheet := ws
	ref := source
	if i := strings.LastIndex(source, "!"); i >= 0 {
		title := strings.Trim(source[:i], "'")
		ref = source[i+1:]
		if ws.parent == nil {
			return nil, "", fmt.Errorf("worksheet %q is not in a workbook", title)
		}
		found, err := ws.parent.GetSheetByName(title)
		if err != nil {
			return nil, "", err
		}
		sheet = found
	}
	ref = stripDollars(ref)
	if !strings.Contains(ref, ":") {
		return nil, "", fmt.Errorf("pivot source must be a range, got %q", source)
	}
	return sheet, ref, nil
}

func (wb *Workbook) layoutAuthored() {
	if wb == nil {
		return
	}
	n := 1
	for _, ws := range wb.worksheets {
		for _, p := range ws.pivots {
			if p.author == nil {
				continue
			}
			p.author.tablePart = fmt.Sprintf("xl/pivotTables/pivotTable%d.xml", n)
			p.author.cachePart = fmt.Sprintf("xl/pivotCache/pivotCacheDefinition%d.xml", n)
			p.author.recPart = fmt.Sprintf("xl/pivotCache/pivotCacheRecords%d.xml", n)
			p.author.cacheID = n - 1
			p.CacheID = p.author.cacheID
			n++
		}
	}
	s := 1
	for _, ws := range wb.worksheets {
		for _, sl := range ws.slicers {
			if sl.author == nil {
				continue
			}
			sl.author.slicerPart = fmt.Sprintf("xl/slicers/slicer%d.xml", s)
			sl.author.cachePart = fmt.Sprintf("xl/slicerCaches/slicerCache%d.xml", s)
			s++
		}
	}
}

func (wb *Workbook) authoredContentTypes() []contentOverride {
	var out []contentOverride
	if wb == nil {
		return nil
	}
	for _, ws := range wb.worksheets {
		for _, p := range ws.pivots {
			if p.author == nil {
				continue
			}
			out = append(out,
				contentOverride{"/" + p.author.cachePart, "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"},
				contentOverride{"/" + p.author.recPart, "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"},
				contentOverride{"/" + p.author.tablePart, "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"},
			)
		}
		for _, sl := range ws.slicers {
			if sl.author == nil {
				continue
			}
			out = append(out,
				contentOverride{"/" + sl.author.slicerPart, "application/vnd.ms-excel.slicer+xml"},
				contentOverride{"/" + sl.author.cachePart, "application/vnd.ms-excel.slicerCache+xml"},
			)
		}
	}
	return out
}

func writePivotCacheList(b *strings.Builder, wb *Workbook) {
	var extra strings.Builder
	if wb != nil {
		for _, ws := range wb.worksheets {
			for _, p := range ws.pivots {
				if p.author == nil || p.author.cacheRelID == "" {
					continue
				}
				fmt.Fprintf(&extra, `<pivotCache cacheId="%d" r:id="%s"/>`, p.author.cacheID, p.author.cacheRelID)
			}
		}
	}
	preserved := preservedNamed(wb.preserved, "pivotCaches")
	switch {
	case preserved == "" && extra.Len() == 0:
		return
	case preserved == "":
		b.WriteString(`<pivotCaches>`)
		b.WriteString(extra.String())
		b.WriteString(`</pivotCaches>`)
	case extra.Len() == 0:
		b.WriteString(preserved)
	default:
		b.WriteString(strings.Replace(preserved, "</pivotCaches>", extra.String()+"</pivotCaches>", 1))
	}
}

func (w *XLSXWriter) writeAuthoredParts(zw *zip.Writer, wb *Workbook) error {
	if wb == nil {
		return nil
	}
	for si, ws := range wb.worksheets {
		tab := si + 1
		for _, p := range ws.pivots {
			if p.author == nil {
				continue
			}
			def, rec, table := p.authoredXML(tab)
			if err := writeZip(zw, p.author.cachePart, def); err != nil {
				return err
			}
			if err := writeZip(zw, p.author.recPart, rec); err != nil {
				return err
			}
			if err := writeZip(zw, p.author.tablePart, table); err != nil {
				return err
			}
			relsName := relsPathFor(p.author.cachePart)
			if err := writeZip(zw, relsName, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+"\n"+
				`<Relationships xmlns="`+pkgRelNS+`">`+
				`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="`+xmlText(pathBase(p.author.recPart))+`"/>`+
				`</Relationships>`); err != nil {
				return err
			}
			ptRels := relsPathFor(p.author.tablePart)
			if err := writeZip(zw, ptRels, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+"\n"+
				`<Relationships xmlns="`+pkgRelNS+`">`+
				`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../`+xmlText(strings.TrimPrefix(p.author.cachePart, "xl/"))+`"/>`+
				`</Relationships>`); err != nil {
				return err
			}
		}
		for _, s := range ws.slicers {
			if s.author == nil {
				continue
			}
			cacheID := 0
			for _, p := range ws.pivots {
				if p.Name == s.author.pivotName && p.author != nil {
					cacheID = p.author.cacheID
				}
			}
			if err := writeZip(zw, s.author.cachePart, slicerCacheXML(s, tab, cacheID)); err != nil {
				return err
			}
			if err := writeZip(zw, s.author.slicerPart, slicerXML(s)); err != nil {
				return err
			}
		}
	}
	return nil
}

func pathBase(p string) string {
	if i := strings.LastIndex(p, "/"); i >= 0 {
		return p[i+1:]
	}
	return p
}

func (p *PivotTable) authoredXML(tab int) (def, rec, table string) {
	a := p.author
	_ = tab
	var recB strings.Builder
	recB.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	fmt.Fprintf(&recB, `<pivotCacheRecords xmlns="%s" xmlns:r="%s" count="%d">`, spreadsheetNS, relNS, len(a.texts))
	rowIdx := indexOf(a.headers, a.rowName)
	dataIdx := indexOf(a.headers, a.dataName)
	shared := map[int]map[string]int{}
	for i, textCol := range a.textCols {
		if !textCol {
			continue
		}
		shared[i] = map[string]int{}
		for _, row := range a.texts {
			v := row[i]
			if _, ok := shared[i][v]; !ok {
				shared[i][v] = len(shared[i])
			}
		}
	}
	for r := range a.texts {
		recB.WriteString(`<r>`)
		for c := range a.headers {
			if a.textCols[c] {
				fmt.Fprintf(&recB, `<x v="%d"/>`, shared[c][a.texts[r][c]])
			} else {
				fmt.Fprintf(&recB, `<n v="%s"/>`, formatNumber(a.nums[r][c]))
			}
		}
		recB.WriteString(`</r>`)
	}
	recB.WriteString(`</pivotCacheRecords>`)

	var defB strings.Builder
	defB.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	fmt.Fprintf(&defB, `<pivotCacheDefinition xmlns="%s" xmlns:r="%s" r:id="rId1" refreshedBy="GoExcel" createdVersion="8" refreshedVersion="8" minRefreshableVersion="3" recordCount="%d">`, spreadsheetNS, relNS, len(a.texts))
	fmt.Fprintf(&defB, `<cacheSource type="worksheet"><worksheetSource ref="%s" sheet="%s"/></cacheSource>`, xmlText(a.sourceRef), xmlText(a.sourceSheet))
	fmt.Fprintf(&defB, `<cacheFields count="%d">`, len(a.headers))
	for c, name := range a.headers {
		if a.textCols[c] {
			fmt.Fprintf(&defB, `<cacheField name="%s" numFmtId="0"><sharedItems count="%d">`, xmlText(name), len(shared[c]))
			// Emit shared strings in index order.
			ordered := make([]string, len(shared[c]))
			for v, i := range shared[c] {
				ordered[i] = v
			}
			for _, v := range ordered {
				fmt.Fprintf(&defB, `<s v="%s"/>`, xmlText(v))
			}
			defB.WriteString(`</sharedItems></cacheField>`)
			continue
		}
		min, max := a.nums[0][c], a.nums[0][c]
		allInt := true
		for _, row := range a.nums {
			if row[c] < min {
				min = row[c]
			}
			if row[c] > max {
				max = row[c]
			}
			if row[c] != float64(int64(row[c])) {
				allInt = false
			}
		}
		intAttr := "0"
		if allInt {
			intAttr = "1"
		}
		fmt.Fprintf(&defB, `<cacheField name="%s" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="%s" minValue="%s" maxValue="%s"/></cacheField>`, xmlText(name), intAttr, formatNumber(min), formatNumber(max))
	}
	defB.WriteString(`</cacheFields>`)
	defB.WriteString(`<extLst><ext uri="{725AE2AE-9491-48be-B2B4-4EB974FC3084}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:pivotCacheDefinition/></ext></extLst>`)
	defB.WriteString(`</pivotCacheDefinition>`)

	var tb strings.Builder
	tb.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	fmt.Fprintf(&tb, `<pivotTableDefinition xmlns="%s" name="%s" cacheId="%d" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="%s" updatedVersion="8" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="8" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">`,
		spreadsheetNS, xmlText(p.Name), a.cacheID, xmlText(p.DataCaption))
	fmt.Fprintf(&tb, `<location ref="%s" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>`, xmlText(p.Location))
	fmt.Fprintf(&tb, `<pivotFields count="%d">`, len(a.headers))
	for c := range a.headers {
		switch {
		case c == rowIdx:
			fmt.Fprintf(&tb, `<pivotField axis="axisRow" showAll="0"><items count="%d">`, len(a.groups)+1)
			for i := range a.groups {
				if i == 0 {
					tb.WriteString(`<item x="0"/>`)
				} else {
					fmt.Fprintf(&tb, `<item x="%d"/>`, i)
				}
			}
			tb.WriteString(`<item t="default"/></items></pivotField>`)
		case c == dataIdx:
			tb.WriteString(`<pivotField dataField="1" showAll="0"/>`)
		default:
			tb.WriteString(`<pivotField showAll="0"/>`)
		}
	}
	tb.WriteString(`</pivotFields>`)
	tb.WriteString(`<rowFields count="1"><field x="` + strconv.Itoa(rowIdx) + `"/></rowFields>`)
	fmt.Fprintf(&tb, `<rowItems count="%d">`, len(a.groups)+1)
	for i := range a.groups {
		if i == 0 {
			tb.WriteString(`<i><x/></i>`)
		} else {
			fmt.Fprintf(&tb, `<i><x v="%d"/></i>`, i)
		}
	}
	tb.WriteString(`<i t="grand"><x/></i></rowItems>`)
	tb.WriteString(`<colItems count="1"><i/></colItems>`)
	fmt.Fprintf(&tb, `<dataFields count="1"><dataField name="%s" fld="%d" baseField="0" baseItem="0"/></dataFields>`, xmlText("Sum of "+a.dataName), dataIdx)
	style := p.Style
	if style == "" {
		style = "PivotStyleLight16"
	}
	fmt.Fprintf(&tb, `<pivotTableStyleInfo name="%s" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>`, xmlText(style))
	tb.WriteString(`<extLst><ext uri="{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:pivotTableDefinition hideValuesRow="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"/></ext><ext uri="{747A6164-185A-40DC-8AA5-F01512510D54}" xmlns:xpdl="http://schemas.microsoft.com/office/spreadsheetml/2016/pivotdefaultlayout"><xpdl:pivotTableDefinition16/></ext></extLst>`)
	tb.WriteString(`</pivotTableDefinition>`)
	return defB.String(), recB.String(), tb.String()
}

func slicerCacheXML(s *Slicer, tab, cacheID int) string {
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	fmt.Fprintf(&b, `<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="%s" sourceName="%s">`, xmlText(s.Cache), xmlText(s.author.field))
	fmt.Fprintf(&b, `<pivotTables><pivotTable tabId="%d" name="%s"/></pivotTables>`, tab, xmlText(s.author.pivotName))
	fmt.Fprintf(&b, `<data><tabular pivotCacheId="%d"><items count="%d">`, cacheID, len(s.author.items))
	for _, item := range s.author.items {
		fmt.Fprintf(&b, `<i v="%s"/>`, xmlText(item))
	}
	b.WriteString(`</items></tabular></data></slicerCacheDefinition>`)
	return b.String()
}

func slicerXML(s *Slicer) string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n" +
		fmt.Sprintf(`<slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><slicer name="%s" cache="%s" caption="%s" startItem="0" columnCount="%d" rowHeight="%d" style="%s"/></slicers>`,
			xmlText(s.Name), xmlText(s.Cache), xmlText(s.Caption), s.ColumnCount, s.RowHeight, xmlText(s.Style))
}

func (ws *Worksheet) ensureAuthoredSlicerExt() {
	var inner strings.Builder
	anySlicer := false
	for _, s := range ws.slicers {
		if s.author == nil || s.author.sheetRelID == "" {
			continue
		}
		anySlicer = true
		fmt.Fprintf(&inner, `<x14:slicer r:id="%s"/>`, s.author.sheetRelID)
	}
	if !anySlicer {
		return
	}
	const marker = `{A8765BA9-456A-4dab-B4F3-ACF838C121DE}`
	list := `<ext uri="` + marker + `" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerList>` + inner.String() + `</x14:slicerList></ext>`
	for i := range ws.preserved {
		if ws.preserved[i].Name != "extLst" {
			continue
		}
		xml := ws.preserved[i].XML
		if idx := strings.Index(xml, marker); idx >= 0 {
			start := strings.LastIndex(xml[:idx], "<ext ")
			endRel := strings.Index(xml[idx:], "</ext>")
			if start >= 0 && endRel >= 0 {
				xml = xml[:start] + xml[idx+endRel+len("</ext>"):]
			}
		}
		ws.preserved[i].XML = strings.Replace(xml, "</extLst>", list+"</extLst>", 1)
		return
	}
	ws.preserved = append(ws.preserved, rawXML{Name: "extLst", XML: `<extLst>` + list + `</extLst>`})
}
