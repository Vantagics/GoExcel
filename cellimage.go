package gospreadsheet

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"path"
	"strconv"
	"strings"
)

// CellImage is a picture stored in a cell (a rich-value local image).
// Excel records the cell as an error #VALUE! plus a value-metadata index.
type CellImage struct {
	Data      []byte
	Extension string
}

// SetImage stores a picture in the cell. ext is a file extension such as "png".
// The cell value becomes the #VALUE! sentinel Excel writes for a local image.
func (c *Cell) SetImage(data []byte, ext string) *Cell {
	if c == nil {
		return nil
	}
	ext = strings.TrimPrefix(strings.ToLower(strings.TrimSpace(ext)), ".")
	if ext == "" {
		ext = "png"
	}
	c.image = &CellImage{Data: append([]byte(nil), data...), Extension: ext}
	c.Type = CellTypeError
	c.Value = "#VALUE!"
	c.Formula = ""
	c.valueMeta = ""
	return c
}

// Image returns the cell's picture, when it has one.
func (c *Cell) Image() (CellImage, bool) {
	if c == nil || c.image == nil || len(c.image.Data) == 0 {
		return CellImage{}, false
	}
	return CellImage{Data: append([]byte(nil), c.image.Data...), Extension: c.image.Extension}, true
}

func (wb *Workbook) hasCellImages() bool {
	if wb == nil {
		return false
	}
	for _, ws := range wb.worksheets {
		for _, c := range ws.cells {
			if c.image != nil && len(c.image.Data) > 0 {
				return true
			}
		}
	}
	return false
}

func assignCellImages(wb *Workbook) {
	if wb == nil {
		return
	}
	n := 1
	for _, ws := range wb.worksheets {
		for _, c := range cellsToWrite(ws) {
			if c.image == nil || len(c.image.Data) == 0 {
				continue
			}
			c.valueMeta = strconv.Itoa(n)
			c.Type = CellTypeError
			c.Value = "#VALUE!"
			n++
		}
	}
}

func indexRichImages(wb *Workbook) {
	if wb == nil || wb.source == nil {
		return
	}
	meta, ok := wb.source.file("xl/metadata.xml")
	if !ok || !bytes.Contains(meta, []byte("XLRICHVALUE")) {
		return
	}
	vmToRV := valueMetadataToRichValue(meta)
	relIndex := richValueRelIndexes(wb.source.fileBytes("xl/richData/rdrichvalue.xml"))
	rels := richValueRelationships(wb)
	for _, ws := range wb.worksheets {
		for _, c := range ws.cells {
			vm, err := strconv.Atoi(c.valueMeta)
			if err != nil || vm < 1 || vm > len(vmToRV) {
				continue
			}
			rv := vmToRV[vm-1]
			if rv < 0 || rv >= len(relIndex) {
				continue
			}
			rid := ""
			idx := relIndex[rv]
			if idx >= 0 && idx < len(rels.order) {
				rid = rels.order[idx]
			}
			target := rels.byID[rid]
			if target == "" {
				continue
			}
			part := resolvePart("xl/richData/richValueRel.xml", target)
			data, ok := wb.source.file(part)
			if !ok {
				continue
			}
			ext := strings.TrimPrefix(strings.ToLower(path.Ext(part)), ".")
			if ext == "" {
				ext = "png"
			}
			c.image = &CellImage{Data: append([]byte(nil), data...), Extension: ext}
		}
	}
}

type richRelList struct {
	order []string
	byID  map[string]string
}

func richValueRelationships(wb *Workbook) richRelList {
	out := richRelList{byID: map[string]string{}}
	raw, ok := wb.source.file("xl/richData/_rels/richValueRel.xml.rels")
	if !ok {
		return out
	}
	rels, err := parseRelationships(raw)
	if err != nil {
		return out
	}
	for _, rel := range rels {
		out.order = append(out.order, rel.ID)
		out.byID[rel.ID] = rel.Target
	}
	return out
}

func (pkg *ooxmlPackage) fileBytes(name string) []byte {
	if pkg == nil {
		return nil
	}
	b, _ := pkg.file(name)
	return b
}

// valueMetadataToRichValue maps each 1-based vm index to an rdrichvalue index.
func valueMetadataToRichValue(meta []byte) []int {
	dec := xml.NewDecoder(bytes.NewReader(meta))
	var rvb []int
	var vmRV []int
	inFuture := false
	inValue := false
	for {
		tok, err := dec.Token()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "futureMetadata":
				inFuture = attrLocal(t, "name") == "XLRICHVALUE"
			case "rvb":
				if inFuture {
					n, _ := strconv.Atoi(attrLocal(t, "i"))
					rvb = append(rvb, n)
				}
			case "valueMetadata":
				inValue = true
			case "rc":
				if inValue {
					n, _ := strconv.Atoi(attrLocal(t, "v"))
					if n >= 0 && n < len(rvb) {
						vmRV = append(vmRV, rvb[n])
					} else {
						vmRV = append(vmRV, n)
					}
				}
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "futureMetadata":
				inFuture = false
			case "valueMetadata":
				inValue = false
			}
		}
	}
	return vmRV
}

// richValueRelIndexes returns the richValueRel index stored in each rv.
func richValueRelIndexes(data []byte) []int {
	if len(data) == 0 {
		return nil
	}
	dec := xml.NewDecoder(bytes.NewReader(data))
	var out []int
	inRV := false
	first := true
	for {
		tok, err := dec.Token()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "rv" {
				inRV = true
				first = true
				out = append(out, 0)
			}
			if t.Name.Local == "v" && inRV && first {
				var b strings.Builder
				_ = readCharData(dec, &b, "v")
				n, _ := strconv.Atoi(strings.TrimSpace(b.String()))
				out[len(out)-1] = n
				first = false
			}
		case xml.EndElement:
			if t.Name.Local == "rv" {
				inRV = false
			}
		}
	}
	return out
}

func (w *XLSXWriter) writeCellImages(zw *zip.Writer, wb *Workbook) error {
	images := collectCellImages(wb)
	if len(images) == 0 {
		return nil
	}
	if err := writeZip(zw, "xl/metadata.xml", imageMetadataXML(len(images))); err != nil {
		return err
	}
	if err := writeZip(zw, "xl/richData/rdrichvaluestructure.xml", imageStructureXML()); err != nil {
		return err
	}
	if err := writeZip(zw, "xl/richData/rdrichvalue.xml", imageValuesXML(len(images))); err != nil {
		return err
	}
	if err := writeZip(zw, "xl/richData/rdRichValueTypes.xml", imageTypesXML()); err != nil {
		return err
	}
	var rels strings.Builder
	rels.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	rels.WriteString(`<richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel" xmlns:r="` + relNS + `">`)
	var parts strings.Builder
	parts.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	parts.WriteString(`<Relationships xmlns="` + pkgRelNS + `">`)
	for i, img := range images {
		ext := img.Extension
		if ext == "" {
			ext = "png"
		}
		name := fmt.Sprintf("xl/media/rich%d.%s", i+1, ext)
		if err := writeZipBytes(zw, name, img.Data); err != nil {
			return err
		}
		fmt.Fprintf(&rels, `<rel r:id="rId%d"/>`, i+1)
		fmt.Fprintf(&parts, `<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/rich%d.%s"/>`, i+1, i+1, ext)
	}
	rels.WriteString(`</richValueRels>`)
	parts.WriteString(`</Relationships>`)
	if err := writeZip(zw, "xl/richData/richValueRel.xml", rels.String()); err != nil {
		return err
	}
	return writeZip(zw, "xl/richData/_rels/richValueRel.xml.rels", parts.String())
}

func collectCellImages(wb *Workbook) []CellImage {
	if wb == nil {
		return nil
	}
	var out []CellImage
	for _, ws := range wb.worksheets {
		for _, c := range cellsToWrite(ws) {
			if c.image == nil || len(c.image.Data) == 0 {
				continue
			}
			out = append(out, *c.image)
		}
	}
	return out
}

func imageMetadataXML(n int) string {
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	b.WriteString(`<metadata xmlns="` + spreadsheetNS + `" xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">`)
	b.WriteString(`<metadataTypes count="1"><metadataType name="XLRICHVALUE" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1"/></metadataTypes>`)
	fmt.Fprintf(&b, `<futureMetadata name="XLRICHVALUE" count="%d">`, n)
	for i := 0; i < n; i++ {
		fmt.Fprintf(&b, `<bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="%d"/></ext></extLst></bk>`, i)
	}
	b.WriteString(`</futureMetadata>`)
	fmt.Fprintf(&b, `<valueMetadata count="%d">`, n)
	for i := 0; i < n; i++ {
		fmt.Fprintf(&b, `<bk><rc t="1" v="%d"/></bk>`, i)
	}
	b.WriteString(`</valueMetadata></metadata>`)
	return b.String()
}

func imageStructureXML() string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1">` +
		`<s t="_localImage"><k n="_rvRel:LocalImageIdentifier" t="i"/><k n="CalcOrigin" t="i"/></s>` +
		`</rvStructures>`
}

func imageValuesXML(n int) string {
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	fmt.Fprintf(&b, `<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="%d">`, n)
	for i := 0; i < n; i++ {
		fmt.Fprintf(&b, `<rv s="0"><v>%d</v><v>5</v></rv>`, i)
	}
	b.WriteString(`</rvData>`)
	return b.String()
}

func imageTypesXML() string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<rvTypesInfo xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x" xmlns:x="` + spreadsheetNS + `">` +
		`<global><keyFlags>` +
		`<key name="_Self"><flag name="ExcludeFromFile" value="1"/><flag name="ExcludeFromCalcComparison" value="1"/></key>` +
		`<key name="_DisplayString"><flag name="ExcludeFromCalcComparison" value="1"/></key>` +
		`</keyFlags></global></rvTypesInfo>`
}

func skipRichImagePart(name string) bool {
	l := strings.ToLower(strings.ReplaceAll(name, "\\", "/"))
	return l == "xl/metadata.xml" || strings.HasPrefix(l, "xl/richdata/")
}
