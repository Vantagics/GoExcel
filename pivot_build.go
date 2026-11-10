package gospreadsheet

import (
	"archive/zip"
	"fmt"
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

	// gridReady selects the full pivot renderer.
	gridReady    bool
	cells        [][]pvCell
	cache        []pvCacheField
	rowAxis      []int
	colAxis      []int
	pageAxis     []int
	values       []pvValue
	showRowGrand bool
	showColGrand bool
	showSub      bool
	grid         pivotGrid
	selections   map[int]pivotSelection

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
// For several fields, filters, aggregations, or grouping, use BuildPivotTable.
func (ws *Worksheet) AddPivotTable(name, source, dest, rowField, dataField string) (*PivotTable, error) {
	on := true
	return ws.BuildPivotTable(PivotBuild{
		Name: name, Source: source, Dest: dest,
		RowFields:         []string{rowField},
		Values:            []PivotValueField{{Field: dataField, Function: "sum"}},
		RowGrandTotals:    &on,
		ColumnGrandTotals: &on,
		Subtotals:         &on,
	})
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
	if wb.hasCellImages() {
		out = append(out,
			contentOverride{"/xl/metadata.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"},
			contentOverride{"/xl/richData/rdrichvalue.xml", "application/vnd.ms-excel.rdrichvalue+xml"},
			contentOverride{"/xl/richData/rdrichvaluestructure.xml", "application/vnd.ms-excel.rdrichvaluestructure+xml"},
			contentOverride{"/xl/richData/rdRichValueTypes.xml", "application/vnd.ms-excel.rdrichvaluetypes+xml"},
			contentOverride{"/xl/richData/richValueRel.xml", "application/vnd.ms-excel.richvaluerel+xml"},
		)
	}
	if wb.hasThreads() {
		out = append(out, contentOverride{"/xl/persons/person.xml", "application/vnd.ms-excel.person+xml"})
		for i, ws := range wb.worksheets {
			if len(ws.threads) == 0 {
				continue
			}
			out = append(out, contentOverride{"/" + threadPartName(i), "application/vnd.ms-excel.threadedcomments+xml"})
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
	_ = tab
	if p.author == nil {
		return "", "", ""
	}
	return p.author.renderPivot(p)
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
