package gospreadsheet

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"strconv"
	"strings"
)

// Sparkline is one sparkline: the data formula and the cell that displays it.
type Sparkline struct {
	Formula string
	Ref     string
}

// SparklineGroup is an x14:sparklineGroup. Zero-value booleans are the Excel
// defaults (off). Type is "line", "column", or "stacked"; empty means line.
type SparklineGroup struct {
	Type                string
	DisplayEmptyCellsAs string
	Markers             bool
	High                bool
	Low                 bool
	First               bool
	Last                bool
	Negative            bool
	DisplayHidden       bool
	DisplayXAxis        bool
	RightToLeft         bool
	DateAxis            bool
	LineWeight          float64
	ManualMax           *float64
	ManualMin           *float64
	ColorSeries         string
	ColorNegative       string
	ColorAxis           string
	ColorMarkers        string
	ColorFirst          string
	ColorLast           string
	ColorHigh           string
	ColorLow            string
	Sparklines          []Sparkline
}

// Slicer is one slicer on a worksheet. Name, Caption, Style, RowHeight, and
// ColumnCount can be edited; the rest of the slicer part is preserved.
type Slicer struct {
	Name        string
	Cache       string
	Caption     string
	Style       string
	RowHeight   int
	ColumnCount int

	part            string
	origCaption     string
	origStyle       string
	origRowHeight   int
	origColumnCount int
	origName        string
}

// SlicerCache is a slicer cache. SourceName is the pivot field or table column.
type SlicerCache struct {
	Name       string
	SourceName string

	part       string
	origSource string
}

// PivotDataField is a data field on a pivot table.
type PivotDataField struct {
	Name       string
	FieldIndex int
	FieldName  string
}

// PivotTable is the layout of one pivot table. Name, DataCaption, Location,
// and Style can be edited; the pivot cache and the rest of the definition
// are written back unchanged.
type PivotTable struct {
	Name        string
	CacheID     int
	DataCaption string
	Location    string
	Style       string
	RowFields   []string
	ColFields   []string
	PageFields  []string
	DataFields  []PivotDataField
	CacheFields []string

	part        string
	origName    string
	origCaption string
	origLoc     string
	origStyle   string
}

// SparklineGroups returns a copy of the sheet's sparkline groups.
func (ws *Worksheet) SparklineGroups() []SparklineGroup {
	out := make([]SparklineGroup, len(ws.sparklineGroups))
	for i, g := range ws.sparklineGroups {
		out[i] = g
		out[i].Sparklines = append([]Sparkline(nil), g.Sparklines...)
	}
	return out
}

// SetSparklineGroups replaces the sheet's sparkline groups. The next save
// rewrites the sparkline extension from this model.
func (ws *Worksheet) SetSparklineGroups(groups []SparklineGroup) {
	ws.sparklineGroups = nil
	for _, g := range groups {
		g.Sparklines = append([]Sparkline(nil), g.Sparklines...)
		ws.sparklineGroups = append(ws.sparklineGroups, g)
	}
	ws.sparklinesEdited = true
}

// Slicers returns the slicers attached to this sheet.
func (ws *Worksheet) Slicers() []*Slicer { return ws.slicers }

// PivotTables returns the pivot tables whose parts are linked from this sheet.
func (ws *Worksheet) PivotTables() []*PivotTable { return ws.pivots }

// SlicerCaches returns workbook slicer caches.
func (wb *Workbook) SlicerCaches() []*SlicerCache { return wb.slicerCaches }

// PivotTables returns every pivot table in the workbook.
func (wb *Workbook) PivotTables() []*PivotTable {
	var out []*PivotTable
	for _, ws := range wb.worksheets {
		out = append(out, ws.pivots...)
	}
	return out
}

// SetCaption updates the slicer caption written back to the slicer part.
func (s *Slicer) SetCaption(v string) { s.Caption = v }

// SetStyle updates the slicer style name.
func (s *Slicer) SetStyle(v string) { s.Style = v }

// SetColumnCount updates how many columns of buttons the slicer shows.
func (s *Slicer) SetColumnCount(n int) { s.ColumnCount = n }

// SetRowHeight updates the slicer row height in points.
func (s *Slicer) SetRowHeight(n int) { s.RowHeight = n }

// SetSourceName updates the field or column the slicer cache reads.
func (c *SlicerCache) SetSourceName(v string) { c.SourceName = v }

// SetName updates the pivot table name.
func (p *PivotTable) SetName(v string) { p.Name = v }

// SetDataCaption updates the pivot table data caption.
func (p *PivotTable) SetDataCaption(v string) { p.DataCaption = v }

// SetLocation updates the pivot table's destination ref (for example "A3:D20").
func (p *PivotTable) SetLocation(ref string) { p.Location = ref }

// SetStyle updates the pivot table style name.
func (p *PivotTable) SetStyle(v string) { p.Style = v }

func indexWorkbookFeatures(wb *Workbook) {
	if wb == nil || wb.source == nil {
		return
	}
	caches := indexPivotCaches(wb)
	for _, ws := range wb.worksheets {
		ws.sparklineGroups = parseSparklineGroups(preservedNamed(ws.preserved, "extLst"))
		ws.slicers = findSlicers(wb, ws)
		ws.pivots = findPivots(wb, ws, caches)
	}
	wb.slicerCaches = findSlicerCaches(wb)
}

func preservedNamed(nodes []rawXML, name string) string {
	var b strings.Builder
	for _, n := range nodes {
		if n.Name == name {
			b.WriteString(n.XML)
		}
	}
	return b.String()
}

func (wb *Workbook) applyFeatureEdits() {
	if wb == nil {
		return
	}
	for _, ws := range wb.worksheets {
		ws.applySparklineEdits()
		for _, p := range ws.pivots {
			p.applyEdits(wb.source)
		}
		for _, s := range ws.slicers {
			s.applyEdits(wb.source)
		}
	}
	for _, c := range wb.slicerCaches {
		c.applyEdits(wb.source)
	}
}

func parseSparklineGroups(extXML string) []SparklineGroup {
	if !strings.Contains(extXML, "sparklineGroup") {
		return nil
	}
	dec := xml.NewDecoder(strings.NewReader(extXML))
	var groups []SparklineGroup
	for {
		tok, err := dec.Token()
		if err != nil {
			break
		}
		se, ok := tok.(xml.StartElement)
		if !ok || se.Name.Local != "sparklineGroup" {
			continue
		}
		groups = append(groups, parseOneSparklineGroup(dec, se))
	}
	return groups
}

func parseOneSparklineGroup(dec *xml.Decoder, se xml.StartElement) SparklineGroup {
	g := SparklineGroup{
		Type:                attrLocal(se, "type"),
		DisplayEmptyCellsAs: attrLocal(se, "displayEmptyCellsAs"),
		Markers:             parseBoolAttr(attrLocal(se, "markers"), false),
		High:                parseBoolAttr(attrLocal(se, "high"), false),
		Low:                 parseBoolAttr(attrLocal(se, "low"), false),
		First:               parseBoolAttr(attrLocal(se, "first"), false),
		Last:                parseBoolAttr(attrLocal(se, "last"), false),
		Negative:            parseBoolAttr(attrLocal(se, "negative"), false),
		DisplayHidden:       parseBoolAttr(attrLocal(se, "displayHidden"), false),
		DisplayXAxis:        parseBoolAttr(attrLocal(se, "displayXAxis"), false),
		RightToLeft:         parseBoolAttr(attrLocal(se, "rightToLeft"), false),
		DateAxis:            parseBoolAttr(attrLocal(se, "dateAxis"), false),
	}
	if v := attrLocal(se, "lineWeight"); v != "" {
		g.LineWeight, _ = strconv.ParseFloat(v, 64)
	}
	if v := attrLocal(se, "manualMax"); v != "" {
		f, _ := strconv.ParseFloat(v, 64)
		g.ManualMax = &f
	}
	if v := attrLocal(se, "manualMin"); v != "" {
		f, _ := strconv.ParseFloat(v, 64)
		g.ManualMin = &f
	}
	depth := 1
	for depth > 0 {
		tok, err := dec.Token()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			switch t.Name.Local {
			case "colorSeries":
				g.ColorSeries = attrLocal(t, "rgb")
			case "colorNegative":
				g.ColorNegative = attrLocal(t, "rgb")
			case "colorAxis":
				g.ColorAxis = attrLocal(t, "rgb")
			case "colorMarkers":
				g.ColorMarkers = attrLocal(t, "rgb")
			case "colorFirst":
				g.ColorFirst = attrLocal(t, "rgb")
			case "colorLast":
				g.ColorLast = attrLocal(t, "rgb")
			case "colorHigh":
				g.ColorHigh = attrLocal(t, "rgb")
			case "colorLow":
				g.ColorLow = attrLocal(t, "rgb")
			case "sparkline":
				g.Sparklines = append(g.Sparklines, parseSparkline(dec, t))
				depth--
			}
		case xml.EndElement:
			depth--
		}
	}
	if g.Type == "" {
		g.Type = "line"
	}
	return g
}

func parseSparkline(dec *xml.Decoder, start xml.StartElement) Sparkline {
	var s Sparkline
	depth := 1
	var text strings.Builder
	capture := ""
	for depth > 0 {
		tok, err := dec.Token()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			if t.Name.Local == "f" || t.Name.Local == "sqref" {
				capture = t.Name.Local
				text.Reset()
			} else {
				capture = ""
			}
		case xml.EndElement:
			if capture != "" && depth == 2 {
				if capture == "f" {
					s.Formula = strings.TrimSpace(text.String())
				} else {
					s.Ref = strings.TrimSpace(text.String())
				}
				capture = ""
			}
			depth--
		case xml.CharData:
			if capture != "" {
				text.Write(t)
			}
		}
	}
	_ = start
	return s
}

func (ws *Worksheet) applySparklineEdits() {
	if !ws.sparklinesEdited {
		return
	}
	xml := sparklineGroupsXML(ws.sparklineGroups)
	for i := range ws.preserved {
		if ws.preserved[i].Name != "extLst" {
			continue
		}
		if updated, ok := replaceLocalElement(ws.preserved[i].XML, "sparklineGroups", xml); ok {
			ws.preserved[i].XML = updated
			return
		}
	}
	ext := `<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">` + xml + `</ext></extLst>`
	ws.preserved = append(ws.preserved, rawXML{Name: "extLst", XML: ext})
}

func sparklineGroupsXML(groups []SparklineGroup) string {
	var b strings.Builder
	b.WriteString(`<x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">`)
	for _, g := range groups {
		b.WriteString(`<x14:sparklineGroup`)
		if g.Type != "" && g.Type != "line" {
			fmt.Fprintf(&b, ` type="%s"`, xmlText(g.Type))
		}
		if g.DisplayEmptyCellsAs != "" {
			fmt.Fprintf(&b, ` displayEmptyCellsAs="%s"`, xmlText(g.DisplayEmptyCellsAs))
		}
		writeBoolAttr(&b, "markers", g.Markers)
		writeBoolAttr(&b, "high", g.High)
		writeBoolAttr(&b, "low", g.Low)
		writeBoolAttr(&b, "first", g.First)
		writeBoolAttr(&b, "last", g.Last)
		writeBoolAttr(&b, "negative", g.Negative)
		writeBoolAttr(&b, "displayHidden", g.DisplayHidden)
		writeBoolAttr(&b, "displayXAxis", g.DisplayXAxis)
		writeBoolAttr(&b, "rightToLeft", g.RightToLeft)
		writeBoolAttr(&b, "dateAxis", g.DateAxis)
		if g.LineWeight != 0 {
			fmt.Fprintf(&b, ` lineWeight="%s"`, trimFloat(g.LineWeight))
		}
		if g.ManualMax != nil {
			fmt.Fprintf(&b, ` manualMax="%s"`, trimFloat(*g.ManualMax))
		}
		if g.ManualMin != nil {
			fmt.Fprintf(&b, ` manualMin="%s"`, trimFloat(*g.ManualMin))
		}
		b.WriteByte('>')
		writeSparkColor(&b, "colorSeries", g.ColorSeries)
		writeSparkColor(&b, "colorNegative", g.ColorNegative)
		writeSparkColor(&b, "colorAxis", g.ColorAxis)
		writeSparkColor(&b, "colorMarkers", g.ColorMarkers)
		writeSparkColor(&b, "colorFirst", g.ColorFirst)
		writeSparkColor(&b, "colorLast", g.ColorLast)
		writeSparkColor(&b, "colorHigh", g.ColorHigh)
		writeSparkColor(&b, "colorLow", g.ColorLow)
		b.WriteString(`<x14:sparklines>`)
		for _, s := range g.Sparklines {
			fmt.Fprintf(&b, `<x14:sparkline><xm:f>%s</xm:f><xm:sqref>%s</xm:sqref></x14:sparkline>`, xmlText(s.Formula), xmlText(s.Ref))
		}
		b.WriteString(`</x14:sparklines></x14:sparklineGroup>`)
	}
	b.WriteString(`</x14:sparklineGroups>`)
	return b.String()
}

func writeBoolAttr(b *strings.Builder, name string, v bool) {
	if v {
		fmt.Fprintf(b, ` %s="1"`, name)
	}
}

func writeSparkColor(b *strings.Builder, name, rgb string) {
	if rgb == "" {
		return
	}
	fmt.Fprintf(b, `<x14:%s rgb="%s"/>`, name, xmlText(rgb))
}

func findSlicers(wb *Workbook, ws *Worksheet) []*Slicer {
	var out []*Slicer
	for _, rel := range ws.sheetRels {
		if !strings.Contains(strings.ToLower(rel.Type), "slicer") || strings.Contains(strings.ToLower(rel.Type), "cache") {
			continue
		}
		part := resolvePart(ws.partName, rel.Target)
		data, ok := wb.source.file(part)
		if !ok {
			continue
		}
		dec := xml.NewDecoder(bytes.NewReader(data))
		for {
			tok, err := dec.Token()
			if err != nil {
				break
			}
			se, ok := tok.(xml.StartElement)
			if !ok || se.Name.Local != "slicer" {
				continue
			}
			s := &Slicer{
				Name:        attrLocal(se, "name"),
				Cache:       attrLocal(se, "cache"),
				Caption:     attrLocal(se, "caption"),
				Style:       attrLocal(se, "style"),
				RowHeight:   atoiString(attrLocal(se, "rowHeight")),
				ColumnCount: atoiString(attrLocal(se, "columnCount")),
				part:        part,
			}
			s.origName = s.Name
			s.origCaption = s.Caption
			s.origStyle = s.Style
			s.origRowHeight = s.RowHeight
			s.origColumnCount = s.ColumnCount
			out = append(out, s)
		}
	}
	return out
}

func findSlicerCaches(wb *Workbook) []*SlicerCache {
	var out []*SlicerCache
	for _, name := range wb.source.order {
		l := strings.ToLower(name)
		if !strings.Contains(l, "slicercache") || !strings.HasSuffix(l, ".xml") || strings.Contains(l, "/_rels/") {
			continue
		}
		data, ok := wb.source.file(name)
		if !ok {
			continue
		}
		dec := xml.NewDecoder(bytes.NewReader(data))
		for {
			tok, err := dec.Token()
			if err != nil {
				break
			}
			se, ok := tok.(xml.StartElement)
			if !ok || se.Name.Local != "slicerCacheDefinition" {
				continue
			}
			c := &SlicerCache{
				Name:       attrLocal(se, "name"),
				SourceName: attrLocal(se, "sourceName"),
				part:       name,
			}
			c.origSource = c.SourceName
			out = append(out, c)
			break
		}
	}
	return out
}

func (s *Slicer) applyEdits(pkg *ooxmlPackage) {
	if s == nil || pkg == nil || s.part == "" {
		return
	}
	if s.Caption == s.origCaption && s.Style == s.origStyle && s.Name == s.origName && s.RowHeight == s.origRowHeight && s.ColumnCount == s.origColumnCount {
		return
	}
	data, ok := pkg.file(s.part)
	if !ok {
		return
	}
	doc := string(data)
	doc = setMatchingAttr(doc, "slicer", "name", s.origName, "caption", s.Caption)
	doc = setMatchingAttr(doc, "slicer", "name", s.origName, "style", s.Style)
	doc = setMatchingAttr(doc, "slicer", "name", s.origName, "name", s.Name)
	if s.RowHeight != 0 {
		doc = setMatchingAttr(doc, "slicer", "name", s.origName, "rowHeight", strconv.Itoa(s.RowHeight))
	}
	if s.ColumnCount != 0 {
		doc = setMatchingAttr(doc, "slicer", "name", s.origName, "columnCount", strconv.Itoa(s.ColumnCount))
	}
	pkg.setFile(s.part, []byte(doc))
	s.origName = s.Name
	s.origCaption = s.Caption
	s.origStyle = s.Style
	s.origRowHeight = s.RowHeight
	s.origColumnCount = s.ColumnCount
}

func (c *SlicerCache) applyEdits(pkg *ooxmlPackage) {
	if c == nil || pkg == nil || c.part == "" || c.SourceName == c.origSource {
		return
	}
	data, ok := pkg.file(c.part)
	if !ok {
		return
	}
	doc := setElementAttr(string(data), "slicerCacheDefinition", "sourceName", c.SourceName)
	pkg.setFile(c.part, []byte(doc))
	c.origSource = c.SourceName
}

type pivotCacheInfo struct {
	id     int
	part   string
	fields []string
}

func indexPivotCaches(wb *Workbook) map[int]pivotCacheInfo {
	out := map[int]pivotCacheInfo{}
	data, ok := wb.source.file("xl/workbook.xml")
	if !ok {
		return out
	}
	dec := xml.NewDecoder(bytes.NewReader(data))
	for {
		tok, err := dec.Token()
		if err != nil {
			break
		}
		se, ok := tok.(xml.StartElement)
		if !ok || se.Name.Local != "pivotCache" {
			continue
		}
		id := atoiString(attrLocal(se, "cacheId"))
		rid := attrLocal(se, "id")
		info := pivotCacheInfo{id: id}
		for _, rel := range wb.workbookRels {
			if rel.ID == rid {
				info.part = resolvePart("xl/workbook.xml", rel.Target)
				break
			}
		}
		if info.part != "" {
			if raw, ok := wb.source.file(info.part); ok {
				info.fields = cacheFieldNames(raw)
			}
		}
		out[id] = info
	}
	return out
}

func cacheFieldNames(data []byte) []string {
	dec := xml.NewDecoder(bytes.NewReader(data))
	var names []string
	for {
		tok, err := dec.Token()
		if err != nil {
			break
		}
		se, ok := tok.(xml.StartElement)
		if !ok || se.Name.Local != "cacheField" {
			continue
		}
		names = append(names, attrLocal(se, "name"))
	}
	return names
}

func findPivots(wb *Workbook, ws *Worksheet, caches map[int]pivotCacheInfo) []*PivotTable {
	var out []*PivotTable
	for _, rel := range ws.sheetRels {
		t := strings.ToLower(rel.Type)
		if !strings.Contains(t, "pivottable") {
			continue
		}
		part := resolvePart(ws.partName, rel.Target)
		data, ok := wb.source.file(part)
		if !ok {
			continue
		}
		p := parsePivotTable(data)
		p.part = part
		if c, ok := caches[p.CacheID]; ok {
			p.CacheFields = append([]string(nil), c.fields...)
			p.RowFields = nameFields(p.RowFields, c.fields)
			p.ColFields = nameFields(p.ColFields, c.fields)
			p.PageFields = nameFields(p.PageFields, c.fields)
			for i := range p.DataFields {
				idx := p.DataFields[i].FieldIndex
				if idx >= 0 && idx < len(c.fields) {
					p.DataFields[i].FieldName = c.fields[idx]
				}
			}
		}
		p.origName = p.Name
		p.origCaption = p.DataCaption
		p.origLoc = p.Location
		p.origStyle = p.Style
		out = append(out, p)
	}
	return out
}

func nameFields(indexes []string, fields []string) []string {
	out := make([]string, len(indexes))
	for i, s := range indexes {
		n, err := strconv.Atoi(s)
		if err == nil && n >= 0 && n < len(fields) && fields[n] != "" {
			out[i] = fields[n]
			continue
		}
		out[i] = s
	}
	return out
}

func parsePivotTable(data []byte) *PivotTable {
	p := &PivotTable{}
	dec := xml.NewDecoder(bytes.NewReader(data))
	var section string
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "pivotTableDefinition":
				// x14:pivotTableDefinition inside extLst repeats the local
				// name and must not clear attributes taken from the root.
				if v := attrLocal(t, "name"); v != "" {
					p.Name = v
				}
				if v := attrLocal(t, "cacheId"); v != "" {
					p.CacheID = atoiString(v)
				}
				if v := attrLocal(t, "dataCaption"); v != "" {
					p.DataCaption = v
				}
			case "location":
				p.Location = attrLocal(t, "ref")
			case "pivotTableStyleInfo":
				p.Style = attrLocal(t, "name")
			case "rowFields", "colFields", "pageFields", "dataFields":
				section = t.Name.Local
			case "field":
				idx := attrLocal(t, "x")
				switch section {
				case "rowFields":
					p.RowFields = append(p.RowFields, idx)
				case "colFields":
					p.ColFields = append(p.ColFields, idx)
				case "pageFields":
					p.PageFields = append(p.PageFields, idx)
				}
			case "dataField":
				p.DataFields = append(p.DataFields, PivotDataField{
					Name:       attrLocal(t, "name"),
					FieldIndex: atoiString(attrLocal(t, "fld")),
				})
			}
		case xml.EndElement:
			if t.Name.Local == section {
				section = ""
			}
		}
	}
	return p
}

func (p *PivotTable) applyEdits(pkg *ooxmlPackage) {
	if p == nil || pkg == nil || p.part == "" {
		return
	}
	if p.Name == p.origName && p.DataCaption == p.origCaption && p.Location == p.origLoc && p.Style == p.origStyle {
		return
	}
	data, ok := pkg.file(p.part)
	if !ok {
		return
	}
	doc := string(data)
	if p.Name != p.origName {
		doc = setElementAttr(doc, "pivotTableDefinition", "name", p.Name)
	}
	if p.DataCaption != p.origCaption {
		doc = setElementAttr(doc, "pivotTableDefinition", "dataCaption", p.DataCaption)
	}
	if p.Location != p.origLoc {
		doc = setElementAttr(doc, "location", "ref", p.Location)
	}
	if p.Style != p.origStyle {
		doc = setElementAttr(doc, "pivotTableStyleInfo", "name", p.Style)
	}
	pkg.setFile(p.part, []byte(doc))
	p.origName = p.Name
	p.origCaption = p.DataCaption
	p.origLoc = p.Location
	p.origStyle = p.Style
}

func atoiString(s string) int {
	n, _ := strconv.Atoi(strings.TrimSpace(s))
	return n
}

func (p *ooxmlPackage) setFile(name string, data []byte) {
	if p == nil {
		return
	}
	for k := range p.files {
		if strings.EqualFold(k, name) {
			p.files[k] = data
			return
		}
	}
	p.files[name] = data
	p.order = append(p.order, name)
}

// replaceLocalElement replaces the first element with the given local name.
func replaceLocalElement(doc, local, replacement string) (string, bool) {
	start, end, ok := elementSpan(doc, local)
	if !ok {
		return doc, false
	}
	return doc[:start] + replacement + doc[end:], true
}

func elementSpan(doc, local string) (start, end int, ok bool) {
	for i := 0; i < len(doc); i++ {
		if doc[i] != '<' || (i+1 < len(doc) && (doc[i+1] == '/' || doc[i+1] == '!' || doc[i+1] == '?')) {
			continue
		}
		j := i + 1
		for j < len(doc) && doc[j] != ' ' && doc[j] != '>' && doc[j] != '/' && doc[j] != '\n' && doc[j] != '\t' {
			j++
		}
		tag := doc[i+1 : j]
		name := tag
		if k := strings.IndexByte(tag, ':'); k >= 0 {
			name = tag[k+1:]
		}
		if name != local {
			continue
		}
		// Find the end of the start tag.
		k := j
		inQuote := byte(0)
		selfClose := false
		for k < len(doc) {
			c := doc[k]
			if inQuote != 0 {
				if c == inQuote {
					inQuote = 0
				}
				k++
				continue
			}
			if c == '"' || c == '\'' {
				inQuote = c
				k++
				continue
			}
			if c == '>' {
				if k > 0 && doc[k-1] == '/' {
					selfClose = true
				}
				k++
				break
			}
			k++
		}
		if selfClose {
			return i, k, true
		}
		closer := "</" + tag + ">"
		depth := 1
		pos := k
		for pos < len(doc) && depth > 0 {
			next := strings.Index(doc[pos:], "<")
			if next < 0 {
				return 0, 0, false
			}
			pos += next
			if strings.HasPrefix(doc[pos:], closer) {
				depth--
				if depth == 0 {
					return i, pos + len(closer), true
				}
				pos += len(closer)
				continue
			}
			if strings.HasPrefix(doc[pos:], "<"+tag) {
				// another start tag with the same qualified name
				rest := pos + 1 + len(tag)
				if rest < len(doc) && (doc[rest] == ' ' || doc[rest] == '>' || doc[rest] == '/') {
					depth++
				}
			}
			pos++
		}
	}
	return 0, 0, false
}

func setElementAttr(doc, local, attr, value string) string {
	return setMatchingAttr(doc, local, "", "", attr, value)
}

// setMatchingAttr sets attr on the first element whose local name matches and,
// when matchAttr is set, whose matchAttr equals matchValue.
func setMatchingAttr(doc, local, matchAttr, matchValue, attr, value string) string {
	for i := 0; i < len(doc); i++ {
		if doc[i] != '<' || (i+1 < len(doc) && (doc[i+1] == '/' || doc[i+1] == '!' || doc[i+1] == '?')) {
			continue
		}
		j := i + 1
		for j < len(doc) && doc[j] != ' ' && doc[j] != '>' && doc[j] != '/' && doc[j] != '\n' && doc[j] != '\t' {
			j++
		}
		tag := doc[i+1 : j]
		name := tag
		if k := strings.IndexByte(tag, ':'); k >= 0 {
			name = tag[k+1:]
		}
		if name != local {
			continue
		}
		k := j
		inQuote := byte(0)
		end := -1
		for k < len(doc) {
			c := doc[k]
			if inQuote != 0 {
				if c == inQuote {
					inQuote = 0
				}
				k++
				continue
			}
			if c == '"' || c == '\'' {
				inQuote = c
				k++
				continue
			}
			if c == '>' {
				end = k
				break
			}
			k++
		}
		if end < 0 {
			return doc
		}
		startTag := doc[i : end+1]
		if matchAttr != "" && attrValue(startTag, matchAttr) != matchValue {
			i = end
			continue
		}
		updated := putAttr(startTag, attr, value)
		return doc[:i] + updated + doc[end+1:]
	}
	return doc
}

func attrValue(startTag, attr string) string {
	key := attr + `="`
	p := strings.Index(startTag, key)
	if p < 0 {
		return ""
	}
	rest := startTag[p+len(key):]
	q := strings.IndexByte(rest, '"')
	if q < 0 {
		return ""
	}
	return rest[:q]
}

func putAttr(startTag, attr, value string) string {
	key := attr + `="`
	escaped := xmlText(value)
	if p := strings.Index(startTag, key); p >= 0 {
		rest := startTag[p+len(key):]
		q := strings.IndexByte(rest, '"')
		if q >= 0 {
			return startTag[:p+len(key)] + escaped + rest[q:]
		}
	}
	// Insert before the closing > or />.
	insert := fmt.Sprintf(` %s="%s"`, attr, escaped)
	if strings.HasSuffix(startTag, "/>") {
		return startTag[:len(startTag)-2] + insert + "/>"
	}
	return startTag[:len(startTag)-1] + insert + ">"
}
