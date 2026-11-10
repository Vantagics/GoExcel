package gospreadsheet

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"math"
	"os"
	"sort"
	"strconv"
	"strings"
	"sync"
	"time"
)

// writeStrict is set for the duration of XLSXWriter.Write. Write holds writeMu.
var (
	writeMu     sync.Mutex
	writeStrict bool
)

// XLSXWriter writes a Workbook to XLSX format.
type XLSXWriter struct{}

// NewXLSXWriter creates a new XLSX writer.
func NewXLSXWriter() *XLSXWriter {
	return &XLSXWriter{}
}

// Save writes the workbook to a file.
func (w *XLSXWriter) Save(wb *Workbook, filename string) error {
	f, err := os.Create(filename)
	if err != nil {
		return fmt.Errorf("creating file: %w", err)
	}
	writeErr := w.Write(wb, f)
	closeErr := f.Close()
	if writeErr != nil {
		return writeErr
	}
	if closeErr != nil {
		return fmt.Errorf("closing file: %w", closeErr)
	}
	return nil
}

// Write writes the workbook to an io.Writer.
// When the workbook has a password, the bytes are an ECMA-376 Agile encrypted
// OLE package rather than a zip. When strict OOXML is requested, generated
// and preserved XML parts use strict namespaces.
func (w *XLSXWriter) Write(wb *Workbook, writer io.Writer) error {
	writeMu.Lock()
	defer writeMu.Unlock()
	writeStrict = wb != nil && wb.strictOOXML
	defer func() { writeStrict = false }()
	if wb != nil && wb.writePassword != "" {
		var buf bytes.Buffer
		if err := w.writeUnencrypted(wb, &buf); err != nil {
			return err
		}
		enc, err := EncryptPackage(buf.Bytes(), wb.writePassword)
		if err != nil {
			return err
		}
		if _, err := writer.Write(enc); err != nil {
			return fmt.Errorf("writing encrypted package: %w", err)
		}
		return nil
	}
	return w.writeUnencrypted(wb, writer)
}

func (w *XLSXWriter) writeUnencrypted(wb *Workbook, writer io.Writer) error {
	zw := zip.NewWriter(writer)
	if err := w.writePackage(zw, wb); err != nil {
		_ = zw.Close()
		return err
	}
	if err := zw.Close(); err != nil {
		return fmt.Errorf("closing zip: %w", err)
	}
	return nil
}

func (w *XLSXWriter) writePackage(zw *zip.Writer, wb *Workbook) error {
	if wb != nil {
		wb.applyFeatureEdits()
	}
	reg := newStyleRegistry(wb)
	sst := collectSharedStrings(wb)
	plans := make([]sheetPlan, len(wb.worksheets))
	for i, ws := range wb.worksheets {
		plans[i] = planSheet(wb, ws, i)
	}
	assignWorkbookRIDs(wb, plans)
	written := map[string]bool{}
	var originalCT contentTypesDoc
	if wb.source != nil {
		if data, ok := wb.source.file("[Content_Types].xml"); ok {
			originalCT = parseContentTypes(data)
		}
	}

	if err := w.writeContentTypes(zw, wb, plans, originalCT); err != nil {
		return err
	}
	written["[Content_Types].xml"] = true
	if err := w.writeRootRels(zw); err != nil {
		return err
	}
	written["_rels/.rels"] = true
	if err := w.writeWorkbookRels(zw, wb, plans); err != nil {
		return err
	}
	written["xl/_rels/workbook.xml.rels"] = true
	if err := w.writeWorkbookXML(zw, wb, plans); err != nil {
		return err
	}
	written["xl/workbook.xml"] = true
	if err := w.writeStylesXML(zw, reg); err != nil {
		return err
	}
	written["xl/styles.xml"] = true
	if err := w.writeSharedStringsXML(zw, sst); err != nil {
		return err
	}
	written["xl/sharedStrings.xml"] = true

	for i, ws := range wb.worksheets {
		plan := plans[i]
		if err := w.writeOneSheet(zw, wb, ws, plan, sst, reg); err != nil {
			return err
		}
		written[plan.partName] = true
		if plan.relsXML != "" {
			if err := writeZip(zw, plan.relsName, plan.relsXML); err != nil {
				return err
			}
			written[plan.relsName] = true
		}
		if plan.commentsXML != "" {
			if err := writeZip(zw, plan.commentsName, plan.commentsXML); err != nil {
				return err
			}
			written[plan.commentsName] = true
		}
		if plan.vmlXML != "" {
			if err := writeZip(zw, plan.vmlName, plan.vmlXML); err != nil {
				return err
			}
			written[plan.vmlName] = true
		}
	}
	if err := w.writeCoreProperties(zw, wb); err != nil {
		return err
	}
	written["docProps/core.xml"] = true
	if err := w.writeAppProperties(zw, wb); err != nil {
		return err
	}
	written["docProps/app.xml"] = true

	if wb.source != nil {
		for _, name := range wb.source.order {
			if written[name] || !keepPackagePart(name) {
				continue
			}
			// Sheet-level comment/vml parts are rewritten when this package
			// emits its own. Skip originals that a plan replaced.
			if replacedOriginal(plans, name) {
				continue
			}
			data, ok := wb.source.file(name)
			if !ok {
				continue
			}
			if err := writeZipBytes(zw, name, data); err != nil {
				return err
			}
			written[name] = true
		}
	}
	return nil
}

func replacedOriginal(plans []sheetPlan, name string) bool {
	l := strings.ToLower(name)
	for _, p := range plans {
		if p.replacedComments != "" && strings.EqualFold(p.replacedComments, name) {
			return true
		}
		if p.replacedVML != "" && strings.EqualFold(p.replacedVML, name) {
			return true
		}
		if strings.HasPrefix(l, "xl/worksheets/") && strings.Contains(l, "/sheet") {
			// Rewritten worksheet parts and their rels are emitted explicitly.
			return false
		}
	}
	return false
}

func writeZip(zw *zip.Writer, name, content string) error {
	return writeZipBytes(zw, name, []byte(content))
}

func writeZipBytes(zw *zip.Writer, name string, content []byte) error {
	if writeStrict && strictXMLPart(name) {
		content = []byte(toStrictOOXML(string(content)))
	}
	f, err := zw.Create(name)
	if err != nil {
		return err
	}
	_, err = f.Write(content)
	return err
}

func strictXMLPart(name string) bool {
	l := strings.ToLower(name)
	return strings.HasSuffix(l, ".xml") || strings.HasSuffix(l, ".rels")
}

// toStrictOOXML rewrites transitional document namespaces to the strict
// ECMA-376 URIs Excel writes. Package relationship and content-type
// namespaces stay on the transitional OPC URIs, matching real strict files.
func toStrictOOXML(s string) string {
	replacer := strings.NewReplacer(
		"http://schemas.openxmlformats.org/spreadsheetml/2006/main", "http://purl.oclc.org/ooxml/spreadsheetml/main",
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships", "http://purl.oclc.org/ooxml/officeDocument/relationships",
		"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties", "http://purl.oclc.org/ooxml/officeDocument/extendedProperties",
		"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes", "http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes",
		"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing", "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing",
		"http://schemas.openxmlformats.org/drawingml/2006/chart", "http://purl.oclc.org/ooxml/drawingml/chart",
		"http://schemas.openxmlformats.org/drawingml/2006/main", "http://purl.oclc.org/ooxml/drawingml/main",
	)
	return replacer.Replace(s)
}

type sheetPlan struct {
	partName         string
	relTarget        string
	relsName         string
	relsXML          string
	contentType      string
	commentsName     string
	commentsXML      string
	vmlName          string
	vmlXML           string
	legacyRID        string
	replacedComments string
	replacedVML      string
	passthrough      bool
	workbookRID      string
	relType          string
}

func planSheet(wb *Workbook, ws *Worksheet, index int) sheetPlan {
	plan := sheetPlan{}
	if ws.rawPart != nil && ws.CellCount() == 0 && len(ws.mergeCells) == 0 && len(cellsToWrite(ws)) == 0 {
		plan.passthrough = true
		plan.partName = ws.partName
		if plan.partName == "" {
			plan.partName = fmt.Sprintf("xl/worksheets/sheet%d.xml", index+1)
		}
		plan.relTarget = strings.TrimPrefix(plan.partName, "xl/")
		plan.contentType = ws.rawContentType
		if plan.contentType == "" {
			plan.contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"
		}
		if ws.relType != "" {
			plan.relType = ws.relType
		} else if strings.Contains(strings.ToLower(plan.contentType), "chartsheet") {
			plan.relType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet"
		} else if strings.Contains(strings.ToLower(plan.contentType), "dialogsheet") {
			plan.relType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet"
		} else {
			plan.relType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
		}
		plan.relsName = relsPathFor(plan.partName)
		return plan
	}
	plan.partName = fmt.Sprintf("xl/worksheets/sheet%d.xml", index+1)
	plan.relTarget = fmt.Sprintf("worksheets/sheet%d.xml", index+1)
	plan.contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
	plan.relType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
	plan.relsName = relsPathFor(plan.partName)

	var rels []relationship
	used := map[string]bool{}
	comments := sheetComments(ws)
	rewriteComments := len(comments) > 0
	for _, rel := range ws.sheetRels {
		suffix := relTypeSuffix(rel.Type)
		if suffix == "hyperlink" {
			continue
		}
		if rewriteComments && (suffix == "comments" || suffix == "vmlDrawing") {
			if suffix == "comments" {
				plan.replacedComments = resolvePart(ws.partName, rel.Target)
			}
			if suffix == "vmlDrawing" {
				plan.replacedVML = resolvePart(ws.partName, rel.Target)
			}
			continue
		}
		if strings.EqualFold(rel.TargetMode, "External") || rel.Target == "" {
			rels = append(rels, rel)
			used[rel.ID] = true
			continue
		}
		// Keep relationship ids stable so preserved drawing/table markup still
		// points at them.
		rels = append(rels, rel)
		used[rel.ID] = true
	}
	for _, c := range cellsToWrite(ws) {
		if c.Hyperlink == nil || c.Hyperlink.URL == "" || !isExternalURL(c.Hyperlink.URL) {
			continue
		}
		id := nextRelID(used)
		rels = append(rels, relationship{
			ID:         id,
			Type:       "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
			Target:     c.Hyperlink.URL,
			TargetMode: "External",
		})
		c.Hyperlink.rid = id
	}
	if rewriteComments {
		plan.commentsName = fmt.Sprintf("xl/comments%d.xml", index+1)
		plan.vmlName = fmt.Sprintf("xl/drawings/vmlDrawing%d.vml", index+1)
		cid := nextRelID(used)
		vid := nextRelID(used)
		plan.legacyRID = vid
		rels = append(rels, relationship{
			ID:     cid,
			Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
			Target: fmt.Sprintf("../comments%d.xml", index+1),
		})
		rels = append(rels, relationship{
			ID:     vid,
			Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
			Target: fmt.Sprintf("../drawings/vmlDrawing%d.vml", index+1),
		})
		plan.commentsXML = commentsXML(comments)
		plan.vmlXML = vmlDrawingXML(comments)
	}
	if len(rels) > 0 {
		plan.relsXML = relationshipsXML(rels)
	}
	if ps := ws.pageSetup; ps != nil && ps.rid != "" {
		// printerSettings relationship is already in the preserved list.
	}
	return plan
}

func (w *XLSXWriter) writeContentTypes(zw *zip.Writer, wb *Workbook, plans []sheetPlan, original contentTypesDoc) error {
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	b.WriteString(`<Types xmlns="` + ctNS + `">`)
	defaults := map[string]string{
		"rels": "application/vnd.openxmlformats-package.relationships+xml",
		"xml":  "application/xml",
	}
	for ext, ct := range original.defaults {
		if _, ok := defaults[ext]; !ok {
			defaults[ext] = ct
		}
	}
	// Stable order: rels, xml, then the rest.
	b.WriteString(`<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`)
	b.WriteString(`<Default Extension="xml" ContentType="application/xml"/>`)
	var exts []string
	for ext := range defaults {
		if ext == "rels" || ext == "xml" {
			continue
		}
		exts = append(exts, ext)
	}
	sort.Strings(exts)
	needsVML := false
	for _, p := range plans {
		if p.vmlXML != "" {
			needsVML = true
		}
	}
	if needsVML {
		if _, ok := defaults["vml"]; !ok {
			exts = append(exts, "vml")
			defaults["vml"] = "application/vnd.openxmlformats-officedocument.vmlDrawing"
		}
	}
	seenExt := map[string]bool{"rels": true, "xml": true}
	for _, ext := range exts {
		if seenExt[ext] {
			continue
		}
		seenExt[ext] = true
		fmt.Fprintf(&b, `<Default Extension="%s" ContentType="%s"/>`, xmlText(ext), xmlText(defaults[ext]))
	}

	wbCT := "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
	if wb.source != nil {
		if _, ok := wb.source.file("xl/vbaProject.bin"); ok {
			wbCT = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
		}
	}
	fmt.Fprintf(&b, `<Override PartName="/xl/workbook.xml" ContentType="%s"/>`, wbCT)
	b.WriteString(`<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`)
	b.WriteString(`<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>`)
	b.WriteString(`<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>`)
	b.WriteString(`<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>`)
	seen := map[string]bool{
		"/xl/workbook.xml": true, "/xl/styles.xml": true, "/xl/sharedstrings.xml": true,
		"/docprops/core.xml": true, "/docprops/app.xml": true,
	}
	for _, p := range plans {
		part := "/" + p.partName
		fmt.Fprintf(&b, `<Override PartName="%s" ContentType="%s"/>`, part, p.contentType)
		seen[strings.ToLower(part)] = true
		if p.commentsXML != "" {
			cp := "/" + p.commentsName
			fmt.Fprintf(&b, `<Override PartName="%s" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>`, cp)
			seen[strings.ToLower(cp)] = true
		}
	}
	if wb.source != nil {
		var extras []string
		for _, name := range wb.source.order {
			if !keepPackagePart(name) {
				continue
			}
			part := "/" + name
			if seen[strings.ToLower(part)] {
				continue
			}
			ct := original.overrideFor(name)
			if ct == "" || ct == "application/xml" {
				continue
			}
			// Defaults already cover the extension.
			ext := strings.TrimPrefix(strings.ToLower(pathExt(name)), ".")
			if def, ok := defaults[ext]; ok && def == ct {
				continue
			}
			extras = append(extras, fmt.Sprintf(`<Override PartName="%s" ContentType="%s"/>`, part, xmlText(ct)))
			seen[strings.ToLower(part)] = true
		}
		for _, e := range extras {
			b.WriteString(e)
		}
	}
	b.WriteString(`</Types>`)
	return writeZip(zw, "[Content_Types].xml", b.String())
}

func pathExt(name string) string {
	if i := strings.LastIndex(name, "."); i >= 0 {
		return name[i:]
	}
	return ""
}

func (w *XLSXWriter) writeRootRels(zw *zip.Writer) error {
	rels := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n"
	rels += `<Relationships xmlns="` + pkgRelNS + `">`
	rels += `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>`
	rels += `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>`
	rels += `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>`
	rels += `</Relationships>`
	return writeZip(zw, "_rels/.rels", rels)
}

func (w *XLSXWriter) writeWorkbookRels(zw *zip.Writer, wb *Workbook, plans []sheetPlan) error {
	used := map[string]bool{}
	var kept []relationship
	for _, rel := range wb.workbookRels {
		suffix := relTypeSuffix(rel.Type)
		switch suffix {
		case "worksheet", "chartsheet", "styles", "sharedStrings", "calcChain":
			continue
		}
		if rel.ID == "" {
			continue
		}
		kept = append(kept, rel)
		used[rel.ID] = true
	}
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	b.WriteString(`<Relationships xmlns="` + pkgRelNS + `">`)
	for i := range wb.worksheets {
		id := plans[i].workbookRID
		if id == "" {
			id = nextRelID(used)
			plans[i].workbookRID = id
		}
		used[id] = true
		relType := plans[i].relType
		if relType == "" {
			relType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
		}
		fmt.Fprintf(&b, `<Relationship Id="%s" Type="%s" Target="%s"/>`, id, relType, xmlText(plans[i].relTarget))
	}
	styleID := nextRelID(used)
	sstID := nextRelID(used)
	fmt.Fprintf(&b, `<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`, styleID)
	fmt.Fprintf(&b, `<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`, sstID)
	for _, rel := range kept {
		mode := ""
		if strings.EqualFold(rel.TargetMode, "External") {
			mode = ` TargetMode="External"`
		}
		fmt.Fprintf(&b, `<Relationship Id="%s" Type="%s" Target="%s"%s/>`, xmlText(rel.ID), xmlText(rel.Type), xmlText(rel.Target), mode)
	}
	b.WriteString(`</Relationships>`)
	return writeZip(zw, "xl/_rels/workbook.xml.rels", b.String())
}

func (w *XLSXWriter) writeWorkbookXML(zw *zip.Writer, wb *Workbook, plans []sheetPlan) error {
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	b.WriteString(`<workbook xmlns="` + spreadsheetNS + `" xmlns:r="` + relNS + `"`)
	if wb.strictOOXML {
		b.WriteString(` conformance="strict"`)
	}
	b.WriteString(wb.extraRootAttrs)
	b.WriteString(`>`)
	writePreservedNamed(&b, wb.preserved, "fileVersion")
	writePreservedNamed(&b, wb.preserved, "fileSharing")
	b.WriteString(`<workbookPr`)
	if wb.date1904 {
		b.WriteString(` date1904="1"`)
	}
	b.WriteString(`/>`)
	if wp := wb.workbookProtection; wp != nil {
		b.WriteString(`<workbookProtection`)
		if wp.LockStructure {
			b.WriteString(` lockStructure="1"`)
		}
		if wp.LockWindows {
			b.WriteString(` lockWindows="1"`)
		}
		if wp.revisions {
			b.WriteString(` lockRevision="1"`)
		}
		if isExcelPassword(wp.Password) {
			fmt.Fprintf(&b, ` workbookPassword="%s"`, xmlText(wp.Password))
		}
		b.WriteString(`/>`)
	}
	active := wb.activeSheet
	if active < 0 || active >= len(wb.worksheets) {
		active = 0
	}
	fmt.Fprintf(&b, `<bookViews><workbookView activeTab="%d"/></bookViews>`, active)
	b.WriteString(`<sheets>`)
	for i, ws := range wb.worksheets {
		state := ""
		if ws.state == "hidden" || ws.state == "veryHidden" {
			state = fmt.Sprintf(` state="%s"`, ws.state)
		}
		fmt.Fprintf(&b, `<sheet name="%s" sheetId="%d"%s r:id="%s"/>`, xmlText(ws.title), i+1, state, plans[i].workbookRID)
	}
	b.WriteString(`</sheets>`)
	writePreservedNamed(&b, wb.preserved, "functionGroups")
	writePreservedNamed(&b, wb.preserved, "externalReferences")
	if names := definedNamesForWrite(wb); len(names) > 0 {
		b.WriteString(`<definedNames>`)
		for _, dn := range names {
			b.WriteString(`<definedName name="`)
			b.WriteString(xmlText(dn.Name))
			b.WriteByte('"')
			if dn.HasLocal {
				fmt.Fprintf(&b, ` localSheetId="%d"`, dn.LocalSheetID)
			}
			if dn.Hidden {
				b.WriteString(` hidden="1"`)
			}
			b.WriteByte('>')
			b.WriteString(xmlText(dn.Ref))
			b.WriteString(`</definedName>`)
		}
		b.WriteString(`</definedNames>`)
	}
	b.WriteString(`<calcPr calcId="191029" fullCalcOnLoad="1"/>`)
	for _, name := range []string{"oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"} {
		writePreservedNamed(&b, wb.preserved, name)
	}
	// Anything else we preserved and did not emit above.
	emitted := map[string]bool{
		"fileVersion": true, "fileSharing": true, "workbookPr": true, "workbookProtection": true,
		"bookViews": true, "sheets": true, "functionGroups": true, "externalReferences": true,
		"definedNames": true, "calcPr": true, "oleSize": true, "customWorkbookViews": true,
		"pivotCaches": true, "smartTagPr": true, "smartTagTypes": true, "webPublishing": true,
		"fileRecoveryPr": true, "webPublishObjects": true, "extLst": true,
	}
	for _, raw := range wb.preserved {
		if emitted[raw.Name] {
			continue
		}
		b.WriteString(raw.XML)
	}
	b.WriteString(`</workbook>`)
	return writeZip(zw, "xl/workbook.xml", b.String())
}

func assignWorkbookRIDs(wb *Workbook, plans []sheetPlan) {
	used := map[string]bool{}
	for _, rel := range wb.workbookRels {
		suffix := relTypeSuffix(rel.Type)
		switch suffix {
		case "worksheet", "chartsheet", "styles", "sharedStrings", "calcChain":
			continue
		}
		if rel.ID != "" {
			used[rel.ID] = true
		}
	}
	for i, ws := range wb.worksheets {
		id := ""
		if ws.relID != "" && !used[ws.relID] {
			id = ws.relID
		}
		if id == "" {
			id = nextRelID(used)
		} else {
			used[id] = true
		}
		plans[i].workbookRID = id
	}
}

func writePreservedNamed(b *strings.Builder, nodes []rawXML, name string) {
	for _, n := range nodes {
		if n.Name == name {
			b.WriteString(n.XML)
		}
	}
}

func definedNamesForWrite(wb *Workbook) []definedName {
	list := append([]definedName(nil), wb.definedNameList...)
	have := map[string]bool{}
	for _, dn := range list {
		if !dn.HasLocal {
			have[dn.Name] = true
		}
	}
	var extras []string
	for name := range wb.namedRanges {
		if !have[name] && !strings.Contains(name, "!") && !strings.HasPrefix(name, "#") {
			extras = append(extras, name)
		}
	}
	sort.Strings(extras)
	for _, name := range extras {
		list = append(list, definedName{Name: name, Ref: wb.namedRanges[name]})
	}
	for i, ws := range wb.worksheets {
		if ws.pageSetup == nil {
			continue
		}
		if ws.pageSetup.PrintArea != nil && ws.pageSetup.PrintArea.Range != "" && !hasLocalName(list, "_xlnm.Print_Area", i) {
			ref := ws.pageSetup.PrintArea.Range
			if !strings.Contains(ref, "!") {
				ref = quoteSheetName(ws.title) + "!" + ref
			}
			list = append(list, definedName{Name: "_xlnm.Print_Area", Ref: ref, HasLocal: true, LocalSheetID: i})
		}
		if (ws.pageSetup.RepeatRows != "" || ws.pageSetup.RepeatColumns != "") && !hasLocalName(list, "_xlnm.Print_Titles", i) {
			var parts []string
			if ws.pageSetup.RepeatColumns != "" {
				parts = append(parts, quoteSheetName(ws.title)+"!"+absoluteSpan(ws.pageSetup.RepeatColumns))
			}
			if ws.pageSetup.RepeatRows != "" {
				parts = append(parts, quoteSheetName(ws.title)+"!"+absoluteSpan(ws.pageSetup.RepeatRows))
			}
			list = append(list, definedName{Name: "_xlnm.Print_Titles", Ref: strings.Join(parts, ","), HasLocal: true, LocalSheetID: i})
		}
	}
	return list
}

func hasLocalName(list []definedName, name string, sheet int) bool {
	for _, dn := range list {
		if dn.Name == name && dn.HasLocal && dn.LocalSheetID == sheet {
			return true
		}
	}
	return false
}

func absoluteSpan(s string) string {
	if strings.Contains(s, "$") {
		return s
	}
	parts := strings.Split(s, ":")
	for i, p := range parts {
		parts[i] = "$" + p
	}
	return strings.Join(parts, ":")
}

func (w *XLSXWriter) writeStylesXML(zw *zip.Writer, reg *styleRegistry) error {
	var b strings.Builder
	reg.writeStyles(&b)
	return writeZip(zw, "xl/styles.xml", b.String())
}

type sstTable struct {
	strings []string
	index   map[string]int
	count   int
}

func collectSharedStrings(wb *Workbook) *sstTable {
	ss := &sstTable{index: map[string]int{}}
	for _, ws := range wb.worksheets {
		for _, c := range cellsToWrite(ws) {
			if c.RichText != nil && len(c.RichText.Runs) > 0 {
				continue
			}
			if c.Type == CellTypeString {
				if s, ok := c.Value.(string); ok {
					ss.add(s)
				}
			}
		}
	}
	return ss
}

func (ss *sstTable) add(s string) int {
	ss.count++
	if idx, ok := ss.index[s]; ok {
		return idx
	}
	idx := len(ss.strings)
	ss.strings = append(ss.strings, s)
	ss.index[s] = idx
	return idx
}

func (ss *sstTable) get(s string) int {
	if idx, ok := ss.index[s]; ok {
		return idx
	}
	return -1
}

func (w *XLSXWriter) writeSharedStringsXML(zw *zip.Writer, ss *sstTable) error {
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	count := ss.count
	if count < len(ss.strings) {
		count = len(ss.strings)
	}
	fmt.Fprintf(&b, `<sst xmlns="%s" count="%d" uniqueCount="%d">`, spreadsheetNS, count, len(ss.strings))
	for _, s := range ss.strings {
		if needsPreserveSpace(s) {
			fmt.Fprintf(&b, `<si><t xml:space="preserve">%s</t></si>`, xmlText(s))
		} else {
			fmt.Fprintf(&b, `<si><t>%s</t></si>`, xmlText(s))
		}
	}
	b.WriteString(`</sst>`)
	return writeZip(zw, "xl/sharedStrings.xml", b.String())
}

func (w *XLSXWriter) writeOneSheet(zw *zip.Writer, wb *Workbook, ws *Worksheet, plan sheetPlan, sst *sstTable, reg *styleRegistry) error {
	if plan.passthrough {
		return writeZipBytes(zw, plan.partName, ws.rawPart)
	}
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	b.WriteString(`<worksheet xmlns="` + spreadsheetNS + `" xmlns:r="` + relNS + `"`)
	b.WriteString(ws.extraRootAttrs)
	b.WriteString(`>`)

	if ws.sheetPrRaw != "" && ws.tabColor == ws.loadedTabColor {
		b.WriteString(ws.sheetPrRaw)
	} else if ws.tabColor != "" || (ws.pageSetup != nil && (ws.pageSetup.FitToWidth > 0 || ws.pageSetup.FitToHeight > 0)) {
		b.WriteString(`<sheetPr>`)
		if ws.pageSetup != nil && (ws.pageSetup.FitToWidth > 0 || ws.pageSetup.FitToHeight > 0) {
			b.WriteString(`<pageSetUpPr fitToPage="1"/>`)
		}
		if c := normalizeARGB(ws.tabColor); c != "" {
			fmt.Fprintf(&b, `<tabColor rgb="%s"/>`, c)
		}
		b.WriteString(`</sheetPr>`)
	}

	b.WriteString(`<dimension ref="`)
	b.WriteString(sheetDimension(ws))
	b.WriteString(`"/>`)
	writeSheetView(&b, wb, ws)
	height := ws.defaultRowHeight
	if height == 0 {
		height = 15
	}
	fmt.Fprintf(&b, `<sheetFormatPr defaultRowHeight="%s"/>`, trimFloat(height))
	writeCols(&b, ws, reg)
	writeSheetData(&b, wb, ws, sst, reg)
	writePreservedNamed(&b, ws.preserved, "sheetCalcPr")
	writeSheetProtection(&b, ws)
	writePreservedNamed(&b, ws.preserved, "protectedRanges")
	writePreservedNamed(&b, ws.preserved, "scenarios")
	if ws.autoFilterRaw != "" && ws.autoFilter != nil {
		b.WriteString(ws.autoFilterRaw)
	} else if ws.autoFilter != nil {
		writeAutoFilter(&b, ws.autoFilter)
	}
	writePreservedNamed(&b, ws.preserved, "sortState")
	writePreservedNamed(&b, ws.preserved, "dataConsolidate")
	writePreservedNamed(&b, ws.preserved, "customSheetViews")
	writeMerges(&b, ws)
	writePreservedNamed(&b, ws.preserved, "phoneticPr")
	writePreservedNamed(&b, ws.preserved, "conditionalFormatting")
	writeNewConditionalFormatting(&b, ws, reg)
	if len(ws.dataValidations) > 0 {
		writeDataValidations(&b, ws)
	}
	writeHyperlinks(&b, ws)
	writePageSetup(&b, ws)
	for _, name := range []string{"rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts"} {
		writePreservedNamed(&b, ws.preserved, name)
	}
	if plan.legacyRID != "" {
		fmt.Fprintf(&b, `<legacyDrawing r:id="%s"/>`, plan.legacyRID)
	} else {
		writePreservedNamed(&b, ws.preserved, "legacyDrawing")
	}
	writePreservedNamed(&b, ws.preserved, "extLst")
	b.WriteString(`</worksheet>`)
	return writeZip(zw, plan.partName, b.String())
}

func writeSheetView(b *strings.Builder, wb *Workbook, ws *Worksheet) {
	b.WriteString(`<sheetViews><sheetView`)
	if wb != nil && wb.activeSheet >= 0 && wb.activeSheet < len(wb.worksheets) && wb.worksheets[wb.activeSheet] == ws {
		b.WriteString(` tabSelected="1"`)
	}
	b.WriteString(` workbookViewId="0"`)
	if ws.sheetView != nil {
		if ws.sheetView.ZoomScale != 0 && ws.sheetView.ZoomScale != 100 {
			fmt.Fprintf(b, ` zoomScale="%d"`, ws.sheetView.ZoomScale)
		}
		if !ws.sheetView.ShowGridlines {
			b.WriteString(` showGridLines="0"`)
		}
		if ws.sheetView.TopLeftCell != "" {
			fmt.Fprintf(b, ` topLeftCell="%s"`, xmlText(ws.sheetView.TopLeftCell))
		}
	}
	if ws.rightToLeft {
		b.WriteString(` rightToLeft="1"`)
	}
	if ws.frozen == nil {
		b.WriteString(`/>`)
	} else {
		b.WriteString(`>`)
		xSplit := ws.frozen.ColumnIdx
		ySplit := ws.frozen.Row - 1
		if ySplit < 0 {
			ySplit = 0
		}
		pane := "bottomRight"
		switch {
		case xSplit > 0 && ySplit > 0:
			pane = "bottomRight"
		case xSplit > 0:
			pane = "topRight"
		default:
			pane = "bottomLeft"
		}
		b.WriteString(`<pane`)
		if xSplit > 0 {
			fmt.Fprintf(b, ` xSplit="%d"`, xSplit)
		}
		if ySplit > 0 {
			fmt.Fprintf(b, ` ySplit="%d"`, ySplit)
		}
		top, _ := CellName(ws.frozen.Row-1, ws.frozen.ColumnIdx)
		fmt.Fprintf(b, ` topLeftCell="%s" activePane="%s" state="frozen"/>`, top, pane)
		fmt.Fprintf(b, `<selection pane="%s" activeCell="%s" sqref="%s"/>`, pane, top, top)
		b.WriteString(`</sheetView>`)
	}
	if ws.frozen == nil {
		// sheetView was self-closed; close sheetViews only.
		b.WriteString(`</sheetViews>`)
		return
	}
	b.WriteString(`</sheetViews>`)
}

func writeCols(b *strings.Builder, ws *Worksheet, reg *styleRegistry) {
	spans := normalizeCols(ws)
	if len(spans) == 0 {
		return
	}
	b.WriteString(`<cols>`)
	for _, sp := range spans {
		fmt.Fprintf(b, `<col min="%d" max="%d"`, sp.Min, sp.Max)
		if sp.Width != 0 || sp.CustomWidth {
			fmt.Fprintf(b, ` width="%s" customWidth="1"`, trimFloat(sp.Width))
		}
		if sp.Hidden {
			b.WriteString(` hidden="1"`)
		}
		if sp.BestFit {
			b.WriteString(` bestFit="1"`)
		}
		if sp.Style != nil {
			if id := reg.xfIndex(sp.Style); id > 0 {
				fmt.Fprintf(b, ` style="%d"`, id)
			}
		}
		b.WriteString(`/>`)
	}
	b.WriteString(`</cols>`)
}

func normalizeCols(ws *Worksheet) []colSpan {
	var spans []colSpan
	spans = append(spans, ws.colSpans...)
	overlay := map[int]colSpan{}
	for col, width := range ws.colWidths {
		sp := overlay[col+1]
		sp.Min, sp.Max = col+1, col+1
		sp.Width = width
		sp.CustomWidth = true
		overlay[col+1] = sp
	}
	for col, hidden := range ws.hiddenCols {
		if !hidden {
			continue
		}
		sp := overlay[col+1]
		sp.Min, sp.Max = col+1, col+1
		sp.Hidden = true
		overlay[col+1] = sp
	}
	if len(overlay) == 0 && len(spans) == 0 {
		return nil
	}
	// Apply single-column overrides onto existing spans or append them.
	for col, sp := range overlay {
		replaced := false
		for i := range spans {
			if spans[i].Min == col && spans[i].Max == col {
				if sp.CustomWidth {
					spans[i].Width = sp.Width
					spans[i].CustomWidth = true
				}
				if sp.Hidden {
					spans[i].Hidden = true
				}
				replaced = true
				break
			}
		}
		if !replaced {
			spans = append(spans, sp)
		}
	}
	sort.Slice(spans, func(i, j int) bool {
		if spans[i].Min == spans[j].Min {
			return spans[i].Max < spans[j].Max
		}
		return spans[i].Min < spans[j].Min
	})
	return spans
}

func cellsToWrite(ws *Worksheet) []*Cell {
	cells := make([]*Cell, 0)
	for _, c := range ws.cells {
		if c.Type != CellTypeEmpty || c.Value != nil || c.Formula != "" || c.Style != nil || c.Hyperlink != nil || c.Comment != nil || (c.RichText != nil && len(c.RichText.Runs) > 0) {
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

func writeSheetData(b *strings.Builder, wb *Workbook, ws *Worksheet, sst *sstTable, reg *styleRegistry) {
	b.WriteString(`<sheetData>`)
	cells := cellsToWrite(ws)
	rowSet := map[int]bool{}
	for _, c := range cells {
		rowSet[c.row] = true
	}
	for row, h := range ws.rowHeights {
		if h != 0 {
			rowSet[row] = true
		}
	}
	for row, hidden := range ws.hiddenRows {
		if hidden {
			rowSet[row] = true
		}
	}
	rows := make([]int, 0, len(rowSet))
	for row := range rowSet {
		rows = append(rows, row)
	}
	sort.Ints(rows)
	idx := 0
	for _, row := range rows {
		b.WriteString(`<row`)
		fmt.Fprintf(b, ` r="%d"`, row+1)
		if h, ok := ws.rowHeights[row]; ok && h != 0 {
			fmt.Fprintf(b, ` ht="%s" customHeight="1"`, trimFloat(h))
		}
		if ws.hiddenRows[row] {
			b.WriteString(` hidden="1"`)
		}
		b.WriteByte('>')
		for idx < len(cells) && cells[idx].row == row {
			writeCell(b, wb, cells[idx], sst, reg)
			idx++
		}
		b.WriteString(`</row>`)
	}
	b.WriteString(`</sheetData>`)
}

func writeCell(b *strings.Builder, wb *Workbook, c *Cell, sst *sstTable, reg *styleRegistry) {
	ref, err := CellName(c.row, c.col)
	if err != nil {
		return
	}
	styleID := 0
	if st := effectiveStyle(c); st != nil {
		styleID = reg.xfIndex(st)
	}
	b.WriteString(`<c r="`)
	b.WriteString(ref)
	b.WriteByte('"')
	if styleID > 0 {
		fmt.Fprintf(b, ` s="%d"`, styleID)
	}

	if c.RichText != nil && len(c.RichText.Runs) > 0 {
		b.WriteString(` t="inlineStr"><is>`)
		for _, run := range c.RichText.Runs {
			b.WriteString(`<r>`)
			if run.Font != nil {
				b.WriteString(richRunProps(run.Font))
			}
			if needsPreserveSpace(run.Text) {
				fmt.Fprintf(b, `<t xml:space="preserve">%s</t>`, xmlText(run.Text))
			} else {
				fmt.Fprintf(b, `<t>%s</t>`, xmlText(run.Text))
			}
			b.WriteString(`</r>`)
		}
		b.WriteString(`</is></c>`)
		return
	}

	if c.Type == CellTypeFormula || c.Formula != "" {
		writeFormulaCell(b, wb, c)
		return
	}
	switch c.Type {
	case CellTypeString:
		if s, ok := c.Value.(string); ok {
			fmt.Fprintf(b, ` t="s"><v>%d</v></c>`, sst.get(s))
			return
		}
	case CellTypeNumeric:
		if v, ok := c.Value.(float64); ok {
			fmt.Fprintf(b, `><v>%s</v></c>`, formatNumber(v))
			return
		}
	case CellTypeBool:
		if v, ok := c.Value.(bool); ok {
			n := 0
			if v {
				n = 1
			}
			fmt.Fprintf(b, ` t="b"><v>%d</v></c>`, n)
			return
		}
	case CellTypeDate:
		if t, ok := c.Value.(time.Time); ok {
			fmt.Fprintf(b, `><v>%s</v></c>`, formatNumber(wb.serialFromTime(t)))
			return
		}
	case CellTypeError:
		fmt.Fprintf(b, ` t="e"><v>%s</v></c>`, xmlText(fmt.Sprint(c.Value)))
		return
	}
	if styleID > 0 {
		b.WriteString(`/>`)
		return
	}
	b.WriteString(`/>`)
}

func writeFormulaCell(b *strings.Builder, wb *Workbook, c *Cell) {
	if c.dt != nil {
		if c.cachedType == "str" || c.cachedType == "e" || c.cachedType == "b" {
			fmt.Fprintf(b, ` t="%s"`, c.cachedType)
		}
		b.WriteString(`><f t="dataTable"`)
		if c.dt.Ref != "" {
			fmt.Fprintf(b, ` ref="%s"`, xmlText(c.dt.Ref))
		}
		if c.dt.TwoD {
			b.WriteString(` dt2D="1"`)
		} else {
			b.WriteString(` dt2D="0"`)
		}
		if c.dt.Row {
			b.WriteString(` dtr="1"`)
		} else {
			b.WriteString(` dtr="0"`)
		}
		if c.dt.R1 != "" {
			fmt.Fprintf(b, ` r1="%s"`, xmlText(c.dt.R1))
		}
		if c.dt.R2 != "" {
			fmt.Fprintf(b, ` r2="%s"`, xmlText(c.dt.R2))
		}
		b.WriteString(`/>`)
		writeCachedValue(b, wb, c)
		b.WriteString(`</c>`)
		return
	}
	formula := strings.TrimSpace(c.Formula)
	if strings.HasPrefix(formula, "=") {
		formula = strings.TrimPrefix(formula, "=")
	}
	kind := ""
	ref := c.formulaRef
	if c.arrayAnchor {
		kind = "array"
	}
	if strings.HasPrefix(formula, "{") && strings.HasSuffix(formula, "}") && len(formula) >= 2 {
		formula = formula[1 : len(formula)-1]
		kind = "array"
		if ref == "" {
			ref, _ = CellName(c.row, c.col)
		}
	}
	cachedType := c.cachedType
	switch c.Value.(type) {
	case bool:
		cachedType = "b"
	case time.Time:
		cachedType = "d"
	}
	if cachedType == "str" || cachedType == "e" || cachedType == "b" {
		fmt.Fprintf(b, ` t="%s"`, cachedType)
	}
	b.WriteString(`><f`)
	if kind == "array" {
		fmt.Fprintf(b, ` t="array" ref="%s"`, xmlText(ref))
	}
	b.WriteByte('>')
	b.WriteString(xmlText(formula))
	b.WriteString(`</f>`)
	writeCachedValue(b, wb, c)
	b.WriteString(`</c>`)
}

func writeCachedValue(b *strings.Builder, wb *Workbook, c *Cell) {
	switch v := c.Value.(type) {
	case float64:
		fmt.Fprintf(b, `<v>%s</v>`, formatNumber(v))
	case bool:
		n := 0
		if v {
			n = 1
		}
		fmt.Fprintf(b, `<v>%d</v>`, n)
	case time.Time:
		fmt.Fprintf(b, `<v>%s</v>`, formatNumber(wb.serialFromTime(v)))
	case string:
		if v != "" {
			fmt.Fprintf(b, `<v>%s</v>`, xmlText(v))
		}
	}
}

func richRunProps(f *Font) string {
	var b strings.Builder
	b.WriteString("<rPr>")
	if f.Bold {
		b.WriteString("<b/>")
	}
	if f.Italic {
		b.WriteString("<i/>")
	}
	if f.Underline {
		b.WriteString(`<u/>`)
	}
	if f.Strikethrough {
		b.WriteString("<strike/>")
	}
	size := f.Size
	if size == 0 {
		size = 11
	}
	fmt.Fprintf(&b, `<sz val="%s"/>`, trimFloat(size))
	if c := normalizeARGB(f.Color); c != "" {
		fmt.Fprintf(&b, `<color rgb="%s"/>`, c)
	}
	name := f.Name
	if name == "" {
		name = "Calibri"
	}
	fmt.Fprintf(&b, `<rFont val="%s"/>`, xmlText(name))
	b.WriteString("</rPr>")
	return b.String()
}

func formatNumber(v float64) string {
	if math.IsNaN(v) || math.IsInf(v, 0) {
		return "0"
	}
	return strconv.FormatFloat(v, 'f', -1, 64)
}

func sheetDimension(ws *Worksheet) string {
	cells := cellsToWrite(ws)
	if len(cells) == 0 && len(ws.mergeCells) == 0 {
		return "A1"
	}
	minR, minC := 1<<30, 1<<30
	maxR, maxC := 0, 0
	for _, c := range cells {
		if c.row < minR {
			minR = c.row
		}
		if c.col < minC {
			minC = c.col
		}
		if c.row > maxR {
			maxR = c.row
		}
		if c.col > maxC {
			maxC = c.col
		}
	}
	for _, mc := range ws.mergeCells {
		if mc.StartRow < minR {
			minR = mc.StartRow
		}
		if mc.StartCol < minC {
			minC = mc.StartCol
		}
		if mc.EndRow > maxR {
			maxR = mc.EndRow
		}
		if mc.EndCol > maxC {
			maxC = mc.EndCol
		}
	}
	if minR == 1<<30 {
		return "A1"
	}
	a, _ := CellName(minR, minC)
	b, _ := CellName(maxR, maxC)
	if a == b {
		return a
	}
	return a + ":" + b
}

func writeMerges(b *strings.Builder, ws *Worksheet) {
	if len(ws.mergeCells) == 0 {
		return
	}
	fmt.Fprintf(b, `<mergeCells count="%d">`, len(ws.mergeCells))
	for _, mc := range ws.mergeCells {
		start, _ := CellName(mc.StartRow, mc.StartCol)
		end, _ := CellName(mc.EndRow, mc.EndCol)
		fmt.Fprintf(b, `<mergeCell ref="%s:%s"/>`, start, end)
	}
	b.WriteString(`</mergeCells>`)
}

func writeSheetProtection(b *strings.Builder, ws *Worksheet) {
	sp := ws.sheetProtection
	if sp == nil {
		return
	}
	b.WriteString(`<sheetProtection`)
	if sp.Sheet {
		b.WriteString(` sheet="1"`)
	}
	if sp.Objects {
		b.WriteString(` objects="1"`)
	}
	if sp.Scenarios {
		b.WriteString(` scenarios="1"`)
	}
	writeAllow := func(name string, denied bool) {
		if denied {
			fmt.Fprintf(b, ` %s="0"`, name)
		}
	}
	writeAllow("formatCells", sp.FormatCells)
	writeAllow("formatColumns", sp.FormatColumns)
	writeAllow("formatRows", sp.FormatRows)
	writeAllow("insertColumns", sp.InsertColumns)
	writeAllow("insertRows", sp.InsertRows)
	writeAllow("insertHyperlinks", sp.InsertHyperlinks)
	writeAllow("deleteColumns", sp.DeleteColumns)
	writeAllow("deleteRows", sp.DeleteRows)
	writeAllow("selectLockedCells", sp.SelectLockedCells)
	writeAllow("sort", sp.Sort)
	writeAllow("autoFilter", sp.AutoFilter)
	writeAllow("pivotTables", sp.PivotTables)
	writeAllow("selectUnlockedCells", sp.SelectUnlockedCells)
	if isExcelPassword(sp.Password) {
		fmt.Fprintf(b, ` password="%s"`, xmlText(sp.Password))
	}
	if sp.algoName != "" {
		fmt.Fprintf(b, ` algorithmName="%s" hashValue="%s" saltValue="%s" spinCount="%s"`, xmlText(sp.algoName), xmlText(sp.hashValue), xmlText(sp.saltValue), xmlText(sp.spinCount))
	}
	b.WriteString(`/>`)
}

func isExcelPassword(p string) bool {
	if p == "" || strings.Contains(p, ":") {
		return false
	}
	for _, r := range p {
		if (r < '0' || r > '9') && (r < 'A' || r > 'F') && (r < 'a' || r > 'f') {
			return false
		}
	}
	return len(p)%2 == 0
}

func writeAutoFilter(b *strings.Builder, af *AutoFilter) {
	if af == nil || af.Range == "" {
		return
	}
	fmt.Fprintf(b, `<autoFilter ref="%s">`, xmlText(af.Range))
	for _, col := range af.Columns {
		fmt.Fprintf(b, `<filterColumn colId="%d">`, col.ColumnIndex)
		switch col.FilterType {
		case FilterCustom:
			b.WriteString(`<customFilters>`)
			for _, cond := range col.Conditions {
				op := cond.Operator
				if op == "" || op == FilterOpEqual {
					fmt.Fprintf(b, `<customFilter val="%s"/>`, xmlText(cond.Value))
				} else {
					fmt.Fprintf(b, `<customFilter operator="%s" val="%s"/>`, xmlText(string(op)), xmlText(cond.Value))
				}
			}
			b.WriteString(`</customFilters>`)
		default:
			b.WriteString(`<filters>`)
			for _, v := range col.Values {
				fmt.Fprintf(b, `<filter val="%s"/>`, xmlText(v))
			}
			b.WriteString(`</filters>`)
		}
		b.WriteString(`</filterColumn>`)
	}
	b.WriteString(`</autoFilter>`)
}

func writeNewConditionalFormatting(b *strings.Builder, ws *Worksheet, reg *styleRegistry) {
	maxPriority := 0
	for _, raw := range ws.preserved {
		if raw.Name != "conditionalFormatting" {
			continue
		}
		// Priorities already live in the preserved markup.
		_ = raw
	}
	for _, cf := range ws.conditionalFormats {
		if cf == nil || cf.fromFile {
			continue
		}
		for _, rule := range cf.Rules {
			if rule.Priority > maxPriority {
				maxPriority = rule.Priority
			}
		}
	}
	base := 1000
	if maxPriority > base {
		base = maxPriority
	}
	n := 0
	for _, cf := range ws.conditionalFormats {
		if cf == nil || cf.fromFile || cf.Range == "" || len(cf.Rules) == 0 {
			continue
		}
		fmt.Fprintf(b, `<conditionalFormatting sqref="%s">`, xmlText(cf.Range))
		for _, rule := range cf.Rules {
			n++
			priority := rule.Priority
			if priority == 0 {
				priority = base + n
			}
			typ := rule.Type
			if typ == "" {
				typ = ConditionalCellIs
			}
			fmt.Fprintf(b, `<cfRule type="%s" priority="%d"`, xmlText(string(typ)), priority)
			if rule.StopIfTrue {
				b.WriteString(` stopIfTrue="1"`)
			}
			if rule.Operator != "" && typ == ConditionalCellIs {
				fmt.Fprintf(b, ` operator="%s"`, xmlText(string(rule.Operator)))
			}
			if rule.Style != nil {
				fmt.Fprintf(b, ` dxfId="%d"`, reg.dxfIndex(rule.Style))
			}
			b.WriteByte('>')
			for _, f := range rule.Formula {
				fmt.Fprintf(b, `<formula>%s</formula>`, xmlText(f))
			}
			b.WriteString(`</cfRule>`)
		}
		b.WriteString(`</conditionalFormatting>`)
	}
}

func writeDataValidations(b *strings.Builder, ws *Worksheet) {
	fmt.Fprintf(b, `<dataValidations count="%d">`, len(ws.dataValidations))
	for _, dv := range ws.dataValidations {
		if dv == nil {
			continue
		}
		b.WriteString(`<dataValidation`)
		if dv.Type != "" && dv.Type != ValidationNone {
			fmt.Fprintf(b, ` type="%s"`, xmlText(string(dv.Type)))
		}
		if dv.Operator != "" {
			fmt.Fprintf(b, ` operator="%s"`, xmlText(string(dv.Operator)))
		}
		if dv.AllowBlank {
			b.WriteString(` allowBlank="1"`)
		}
		if dv.ShowInputMsg {
			b.WriteString(` showInputMessage="1"`)
		}
		if dv.ShowErrorMsg {
			b.WriteString(` showErrorMessage="1"`)
		}
		if dv.ErrorStyle != "" {
			fmt.Fprintf(b, ` errorStyle="%s"`, xmlText(string(dv.ErrorStyle)))
		}
		if dv.ErrorTitle != "" {
			fmt.Fprintf(b, ` errorTitle="%s"`, xmlText(dv.ErrorTitle))
		}
		if dv.ErrorMessage != "" {
			fmt.Fprintf(b, ` error="%s"`, xmlText(dv.ErrorMessage))
		}
		if dv.PromptTitle != "" {
			fmt.Fprintf(b, ` promptTitle="%s"`, xmlText(dv.PromptTitle))
		}
		if dv.PromptMessage != "" {
			fmt.Fprintf(b, ` prompt="%s"`, xmlText(dv.PromptMessage))
		}
		fmt.Fprintf(b, ` sqref="%s"`, xmlText(dv.Range))
		if dv.Formula1 == "" && dv.Formula2 == "" {
			b.WriteString(`/>`)
			continue
		}
		b.WriteByte('>')
		if dv.Formula1 != "" {
			fmt.Fprintf(b, `<formula1>%s</formula1>`, xmlText(dv.Formula1))
		}
		if dv.Formula2 != "" {
			fmt.Fprintf(b, `<formula2>%s</formula2>`, xmlText(dv.Formula2))
		}
		b.WriteString(`</dataValidation>`)
	}
	b.WriteString(`</dataValidations>`)
}

func writeHyperlinks(b *strings.Builder, ws *Worksheet) {
	type link struct {
		ref, rid, loc, tip string
	}
	var links []link
	for _, c := range cellsToWrite(ws) {
		if c.Hyperlink == nil || c.Hyperlink.URL == "" {
			continue
		}
		ref, err := CellName(c.row, c.col)
		if err != nil {
			continue
		}
		lk := link{ref: ref, tip: c.Hyperlink.Tooltip}
		url := c.Hyperlink.URL
		if isExternalURL(url) {
			lk.rid = c.Hyperlink.rid
		} else {
			lk.loc = strings.TrimPrefix(url, "#")
		}
		links = append(links, lk)
	}
	if len(links) == 0 {
		return
	}
	b.WriteString(`<hyperlinks>`)
	for _, lk := range links {
		fmt.Fprintf(b, `<hyperlink ref="%s"`, lk.ref)
		if lk.rid != "" {
			fmt.Fprintf(b, ` r:id="%s"`, lk.rid)
		}
		if lk.loc != "" {
			fmt.Fprintf(b, ` location="%s"`, xmlText(lk.loc))
		}
		if lk.tip != "" {
			fmt.Fprintf(b, ` tooltip="%s"`, xmlText(lk.tip))
		}
		b.WriteString(`/>`)
	}
	b.WriteString(`</hyperlinks>`)
}

func writePageSetup(b *strings.Builder, ws *Worksheet) {
	ps := ws.pageSetup
	if ps == nil {
		return
	}
	if ps.PrintGridlines || ps.PrintHeadings || ps.CenterHorizontally || ps.CenterVertically {
		b.WriteString(`<printOptions`)
		if ps.PrintHeadings {
			b.WriteString(` headings="1"`)
		}
		if ps.PrintGridlines {
			b.WriteString(` gridLines="1"`)
		}
		if ps.CenterHorizontally {
			b.WriteString(` horizontalCentered="1"`)
		}
		if ps.CenterVertically {
			b.WriteString(` verticalCentered="1"`)
		}
		b.WriteString(`/>`)
	}
	m := ps.Margins
	fmt.Fprintf(b, `<pageMargins left="%s" right="%s" top="%s" bottom="%s" header="%s" footer="%s"/>`,
		trimFloat(m.Left), trimFloat(m.Right), trimFloat(m.Top), trimFloat(m.Bottom), trimFloat(m.Header), trimFloat(m.Footer))
	b.WriteString(`<pageSetup`)
	if ps.PaperSize > 0 {
		fmt.Fprintf(b, ` paperSize="%d"`, ps.PaperSize)
	}
	if ps.Orientation != "" {
		fmt.Fprintf(b, ` orientation="%s"`, xmlText(string(ps.Orientation)))
	}
	if ps.Scale > 0 {
		fmt.Fprintf(b, ` scale="%d"`, ps.Scale)
	}
	if ps.FitToWidth > 0 {
		fmt.Fprintf(b, ` fitToWidth="%d"`, ps.FitToWidth)
	}
	if ps.FitToHeight > 0 {
		fmt.Fprintf(b, ` fitToHeight="%d"`, ps.FitToHeight)
	}
	if ps.rid != "" {
		fmt.Fprintf(b, ` r:id="%s"`, xmlText(ps.rid))
	}
	b.WriteString(`/>`)
	if hf := ps.HeaderFooter; hf != nil && (hf.OddHeader != "" || hf.OddFooter != "" || hf.EvenHeader != "" || hf.EvenFooter != "") {
		b.WriteString(`<headerFooter`)
		if hf.DifferentOddEven {
			b.WriteString(` differentOddEven="1"`)
		}
		b.WriteByte('>')
		if hf.OddHeader != "" {
			fmt.Fprintf(b, `<oddHeader>%s</oddHeader>`, xmlText(hf.OddHeader))
		}
		if hf.OddFooter != "" {
			fmt.Fprintf(b, `<oddFooter>%s</oddFooter>`, xmlText(hf.OddFooter))
		}
		if hf.EvenHeader != "" {
			fmt.Fprintf(b, `<evenHeader>%s</evenHeader>`, xmlText(hf.EvenHeader))
		}
		if hf.EvenFooter != "" {
			fmt.Fprintf(b, `<evenFooter>%s</evenFooter>`, xmlText(hf.EvenFooter))
		}
		b.WriteString(`</headerFooter>`)
	}
}

type commentRef struct {
	ref, author, text string
	row, col          int
	shapeID           int
}

func sheetComments(ws *Worksheet) []commentRef {
	var out []commentRef
	for _, c := range cellsToWrite(ws) {
		if c.Comment == nil {
			continue
		}
		ref, err := CellName(c.row, c.col)
		if err != nil {
			continue
		}
		out = append(out, commentRef{
			ref: ref, author: c.Comment.Author, text: c.Comment.Text,
			row: c.row, col: c.col,
		})
	}
	for i := range out {
		out[i].shapeID = 1024 + i
	}
	return out
}

func commentsXML(comments []commentRef) string {
	authors := []string{}
	index := map[string]int{}
	for _, c := range comments {
		if _, ok := index[c.author]; !ok {
			index[c.author] = len(authors)
			authors = append(authors, c.author)
		}
	}
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	b.WriteString(`<comments xmlns="` + spreadsheetNS + `"><authors>`)
	for _, a := range authors {
		fmt.Fprintf(&b, `<author>%s</author>`, xmlText(a))
	}
	b.WriteString(`</authors><commentList>`)
	for _, c := range comments {
		fmt.Fprintf(&b, `<comment ref="%s" authorId="%d" shapeId="%d"><text>`, c.ref, index[c.author], c.shapeID)
		if c.author != "" {
			fmt.Fprintf(&b, `<r><rPr><b/><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/></rPr><t>%s:</t></r>`, xmlText(c.author))
		}
		fmt.Fprintf(&b, `<r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/></rPr><t xml:space="preserve">%s</t></r>`, xmlText("\n"+c.text))
		b.WriteString(`</text></comment>`)
	}
	b.WriteString(`</commentList></comments>`)
	return b.String()
}

func vmlDrawingXML(comments []commentRef) string {
	var b strings.Builder
	b.WriteString(`<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">`)
	b.WriteString(`<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>`)
	b.WriteString(`<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>`)
	for _, c := range comments {
		fmt.Fprintf(&b, `<v:shape id="_x0000_s%d" type="#_x0000_t202" style="position:absolute;margin-left:80pt;margin-top:20pt;width:104pt;height:78pt;z-index:1;visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">`, c.shapeID)
		b.WriteString(`<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>`)
		b.WriteString(`<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>`)
		b.WriteString(`<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>`)
		fmt.Fprintf(&b, `<x:Anchor>%d, 15, %d, 2, %d, 31, %d, 16</x:Anchor>`, c.col+1, c.row, c.col+3, c.row+4)
		b.WriteString(`<x:AutoFill>False</x:AutoFill>`)
		fmt.Fprintf(&b, `<x:Row>%d</x:Row><x:Column>%d</x:Column>`, c.row, c.col)
		b.WriteString(`</x:ClientData></v:shape>`)
	}
	b.WriteString(`</xml>`)
	return b.String()
}

func relationshipsXML(rels []relationship) string {
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	b.WriteString(`<Relationships xmlns="` + pkgRelNS + `">`)
	for _, rel := range rels {
		mode := ""
		if strings.EqualFold(rel.TargetMode, "External") {
			mode = ` TargetMode="External"`
		}
		fmt.Fprintf(&b, `<Relationship Id="%s" Type="%s" Target="%s"%s/>`, xmlText(rel.ID), xmlText(rel.Type), xmlText(rel.Target), mode)
	}
	b.WriteString(`</Relationships>`)
	return b.String()
}

func (w *XLSXWriter) writeCoreProperties(zw *zip.Writer, wb *Workbook) error {
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	b.WriteString(`<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">`)
	if wb.Properties.Creator != "" {
		fmt.Fprintf(&b, `<dc:creator>%s</dc:creator>`, xmlText(wb.Properties.Creator))
	}
	if wb.Properties.Title != "" {
		fmt.Fprintf(&b, `<dc:title>%s</dc:title>`, xmlText(wb.Properties.Title))
	}
	if wb.Properties.Subject != "" {
		fmt.Fprintf(&b, `<dc:subject>%s</dc:subject>`, xmlText(wb.Properties.Subject))
	}
	if wb.Properties.Description != "" {
		fmt.Fprintf(&b, `<dc:description>%s</dc:description>`, xmlText(wb.Properties.Description))
	}
	if wb.Properties.Keywords != "" {
		fmt.Fprintf(&b, `<cp:keywords>%s</cp:keywords>`, xmlText(wb.Properties.Keywords))
	}
	if wb.Properties.Category != "" {
		fmt.Fprintf(&b, `<cp:category>%s</cp:category>`, xmlText(wb.Properties.Category))
	}
	if wb.Properties.LastModifiedBy != "" {
		fmt.Fprintf(&b, `<cp:lastModifiedBy>%s</cp:lastModifiedBy>`, xmlText(wb.Properties.LastModifiedBy))
	}
	b.WriteString(`</cp:coreProperties>`)
	return writeZip(zw, "docProps/core.xml", b.String())
}

func (w *XLSXWriter) writeAppProperties(zw *zip.Writer, wb *Workbook) error {
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	b.WriteString(`<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">`)
	b.WriteString(`<Application>GoExcel</Application>`)
	fmt.Fprintf(&b, `<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>%d</vt:i4></vt:variant></vt:vector></HeadingPairs>`, len(wb.worksheets))
	fmt.Fprintf(&b, `<TitlesOfParts><vt:vector size="%d" baseType="lpstr">`, len(wb.worksheets))
	for _, ws := range wb.worksheets {
		fmt.Fprintf(&b, `<vt:lpstr>%s</vt:lpstr>`, xmlText(ws.title))
	}
	b.WriteString(`</vt:vector></TitlesOfParts></Properties>`)
	return writeZip(zw, "docProps/app.xml", b.String())
}

// dateToSerial converts a time.Time to an Excel serial date number.
// Excel's epoch is January 1, 1900 (serial number 1).
// Note: Excel incorrectly treats 1900 as a leap year (Lotus 1-2-3 bug).
func dateToSerial(t time.Time) float64 {
	epoch := time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
	return t.Sub(epoch).Hours() / 24.0
}

// serialToDate converts an Excel serial date number to time.Time.
func serialToDate(serial float64) time.Time {
	epoch := time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
	days := int(math.Floor(serial + 1e-10))
	frac := serial - float64(days)
	if frac < 0 {
		frac = 0
	}
	ms := math.Round(frac * 86400 * 1000)
	return epoch.AddDate(0, 0, days).Add(time.Duration(ms) * time.Millisecond)
}
