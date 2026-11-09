package gospreadsheet

import (
	"encoding/xml"
	"fmt"
	"strconv"
	"strings"
	"time"
)

type parsedXF struct {
	style      *Style
	formatCode string
	date       bool
}

type styleTable struct {
	xfs    []parsedXF
	extras styleExtras
}

type styleExtras struct {
	dxfsInner      string
	dxfCount       int
	tableStylesXML string
	colorsXML      string
	extLstXML      string
}

func parseStyleSheet(data []byte, theme []string) styleTable {
	var table styleTable
	if len(data) == 0 {
		return table
	}
	rawKids, err := captureRootChildren(data, map[string]bool{
		"dxfs":        true,
		"tableStyles": true,
		"colors":      true,
		"extLst":      true,
	})
	if err == nil {
		for _, kid := range rawKids {
			switch kid.Name {
			case "dxfs":
				table.extras.dxfsInner, table.extras.dxfCount = innerXML(kid.XML)
			case "tableStyles":
				table.extras.tableStylesXML = kid.XML
			case "colors":
				table.extras.colorsXML = kid.XML
			case "extLst":
				table.extras.extLstXML = kid.XML
			}
		}
	}

	var ss xmlStyleSheet
	if err := xml.Unmarshal(stripBOM(data), &ss); err != nil {
		return table
	}
	codes := map[int]string{}
	for id, code := range builtinNumFmts {
		codes[id] = code
	}
	for _, nf := range ss.NumFmts.NumFmt {
		codes[nf.ID] = nf.Code
	}
	table.xfs = make([]parsedXF, len(ss.CellXfs.Xf))
	for i, xf := range ss.CellXfs.Xf {
		table.xfs[i] = buildParsedXF(xf, ss, codes, theme)
	}
	return table
}

func innerXML(outer string) (string, int) {
	i := strings.IndexByte(outer, '>')
	if i < 0 {
		return "", 0
	}
	if i > 0 && outer[i-1] == '/' {
		return "", 0
	}
	j := strings.LastIndex(outer, "</")
	if j < 0 || j <= i {
		return "", 0
	}
	inner := outer[i+1 : j]
	return inner, strings.Count(inner, "<dxf")
}

type xmlStyleSheet struct {
	NumFmts struct {
		NumFmt []struct {
			ID   int    `xml:"numFmtId,attr"`
			Code string `xml:"formatCode,attr"`
		} `xml:"numFmt"`
	} `xml:"numFmts"`
	Fonts struct {
		Font []xmlFont `xml:"font"`
	} `xml:"fonts"`
	Fills struct {
		Fill []xmlFill `xml:"fill"`
	} `xml:"fills"`
	Borders struct {
		Border []xmlBorder `xml:"border"`
	} `xml:"borders"`
	CellXfs struct {
		Xf []xmlXF `xml:"xf"`
	} `xml:"cellXfs"`
}

type xmlOnOff struct {
	Val *string `xml:"val,attr"`
}

func (o *xmlOnOff) on() bool {
	if o == nil {
		return false
	}
	if o.Val == nil {
		return true
	}
	return *o.Val != "0" && !strings.EqualFold(*o.Val, "false")
}

type xmlUnderline struct {
	Val string `xml:"val,attr"`
}

func (u *xmlUnderline) on() bool {
	if u == nil {
		return false
	}
	return !strings.EqualFold(u.Val, "none") && u.Val != "0"
}

type xmlColor struct {
	RGB     string  `xml:"rgb,attr"`
	Theme   string  `xml:"theme,attr"`
	Indexed string  `xml:"indexed,attr"`
	Tint    float64 `xml:"tint,attr"`
}

func (c *xmlColor) resolve(theme []string) string {
	if c == nil {
		return ""
	}
	return resolveColor(c.RGB, c.Theme, c.Indexed, c.Tint, theme)
}

type xmlFont struct {
	B      *xmlOnOff     `xml:"b"`
	I      *xmlOnOff     `xml:"i"`
	U      *xmlUnderline `xml:"u"`
	Strike *xmlOnOff     `xml:"strike"`
	Sz     *struct {
		Val float64 `xml:"val,attr"`
	} `xml:"sz"`
	Color *xmlColor `xml:"color"`
	Name  *struct {
		Val string `xml:"val,attr"`
	} `xml:"name"`
	RFont *struct {
		Val string `xml:"val,attr"`
	} `xml:"rFont"`
}

type xmlFill struct {
	Pattern *struct {
		PatternType string    `xml:"patternType,attr"`
		FgColor     *xmlColor `xml:"fgColor"`
		BgColor     *xmlColor `xml:"bgColor"`
	} `xml:"patternFill"`
}

type xmlBorderEdge struct {
	Style string    `xml:"style,attr"`
	Color *xmlColor `xml:"color"`
}

type xmlBorder struct {
	DiagonalUp   string        `xml:"diagonalUp,attr"`
	DiagonalDown string        `xml:"diagonalDown,attr"`
	Left         xmlBorderEdge `xml:"left"`
	Right        xmlBorderEdge `xml:"right"`
	Top          xmlBorderEdge `xml:"top"`
	Bottom       xmlBorderEdge `xml:"bottom"`
	Diagonal     xmlBorderEdge `xml:"diagonal"`
}

type xmlXF struct {
	NumFmtID          int    `xml:"numFmtId,attr"`
	FontID            int    `xml:"fontId,attr"`
	FillID            int    `xml:"fillId,attr"`
	BorderID          int    `xml:"borderId,attr"`
	ApplyNumberFormat string `xml:"applyNumberFormat,attr"`
	ApplyFont         string `xml:"applyFont,attr"`
	ApplyFill         string `xml:"applyFill,attr"`
	ApplyBorder       string `xml:"applyBorder,attr"`
	ApplyAlignment    string `xml:"applyAlignment,attr"`
	Alignment         *struct {
		Horizontal   string `xml:"horizontal,attr"`
		Vertical     string `xml:"vertical,attr"`
		WrapText     string `xml:"wrapText,attr"`
		ShrinkToFit  string `xml:"shrinkToFit,attr"`
		TextRotation string `xml:"textRotation,attr"`
		Indent       string `xml:"indent,attr"`
	} `xml:"alignment"`
}

func buildParsedXF(xf xmlXF, ss xmlStyleSheet, codes map[int]string, theme []string) parsedXF {
	st := &Style{}
	code := codes[xf.NumFmtID]
	useNum := xf.NumFmtID != 0 && (xf.ApplyNumberFormat == "" || xf.ApplyNumberFormat == "1" || xf.ApplyNumberFormat == "true")
	if useNum && code != "" && !strings.EqualFold(code, "General") {
		c := code
		st.NumberFormat = &NumberFormat{FormatCode: c}
	}
	useFont := xf.ApplyFont == "1" || xf.ApplyFont == "true" || (xf.ApplyFont == "" && xf.FontID > 0)
	if useFont && xf.FontID >= 0 && xf.FontID < len(ss.Fonts.Font) {
		if f := fontFromXML(ss.Fonts.Font[xf.FontID], theme); f != nil {
			st.Font = f
		}
	}
	useFill := xf.ApplyFill == "1" || xf.ApplyFill == "true" || xf.FillID >= 2
	if useFill && xf.FillID >= 0 && xf.FillID < len(ss.Fills.Fill) {
		if fill := fillFromXML(ss.Fills.Fill[xf.FillID], theme); fill != nil {
			st.Fill = fill
		}
	}
	useBorder := xf.ApplyBorder == "1" || xf.ApplyBorder == "true" || xf.BorderID > 0
	if useBorder && xf.BorderID >= 0 && xf.BorderID < len(ss.Borders.Border) {
		if b := borderFromXML(ss.Borders.Border[xf.BorderID], theme); b != nil {
			st.Borders = b
		}
	}
	if xf.Alignment != nil && (xf.ApplyAlignment == "" || xf.ApplyAlignment == "1" || xf.ApplyAlignment == "true") {
		if al := alignmentFromXML(xf.Alignment); al != nil {
			st.Alignment = al
		}
	}
	if st.Font == nil && st.Fill == nil && st.Borders == nil && st.Alignment == nil && st.NumberFormat == nil {
		st = nil
	}
	date := false
	if code != "" {
		date = looksLikeDateFormat(code)
	} else {
		date = isDateFormatID(xf.NumFmtID)
	}
	// Text and general formats are not dates even if a builtin id was remapped.
	if strings.EqualFold(code, "@") || strings.EqualFold(code, "General") {
		date = false
	}
	return parsedXF{style: st, formatCode: code, date: date}
}

func fontFromXML(f xmlFont, theme []string) *Font {
	font := &Font{}
	if f.Name != nil {
		font.Name = f.Name.Val
	} else if f.RFont != nil {
		font.Name = f.RFont.Val
	}
	if f.Sz != nil {
		font.Size = f.Sz.Val
	}
	font.Bold = f.B.on()
	font.Italic = f.I.on()
	font.Underline = f.U.on()
	font.Strikethrough = f.Strike.on()
	font.Color = f.Color.resolve(theme)
	if font.Name == "" && font.Size == 0 && !font.Bold && !font.Italic && !font.Underline && !font.Strikethrough && font.Color == "" {
		return nil
	}
	// Calibri 11 with no other props is the default font.
	if strings.EqualFold(font.Name, "Calibri") && font.Size == 11 && !font.Bold && !font.Italic && !font.Underline && !font.Strikethrough && font.Color == "" {
		return nil
	}
	return font
}

func fillFromXML(f xmlFill, theme []string) *Fill {
	if f.Pattern == nil {
		return nil
	}
	p := strings.ToLower(f.Pattern.PatternType)
	if p == "" || p == "none" || p == "gray125" {
		return nil
	}
	color := f.Pattern.FgColor.resolve(theme)
	if color == "" {
		color = f.Pattern.BgColor.resolve(theme)
	}
	fillType := "pattern"
	if p == "solid" {
		fillType = "solid"
	}
	return &Fill{Type: fillType, Color: color, Pattern: f.Pattern.PatternType}
}

func borderFromXML(b xmlBorder, theme []string) *Borders {
	borders := &Borders{
		Left:         edgeFromXML(b.Left, theme),
		Right:        edgeFromXML(b.Right, theme),
		Top:          edgeFromXML(b.Top, theme),
		Bottom:       edgeFromXML(b.Bottom, theme),
		Diagonal:     edgeFromXML(b.Diagonal, theme),
		DiagonalUp:   parseBoolAttr(b.DiagonalUp, false),
		DiagonalDown: parseBoolAttr(b.DiagonalDown, false),
	}
	if borders.Left.Style == "" && borders.Right.Style == "" && borders.Top.Style == "" && borders.Bottom.Style == "" && borders.Diagonal.Style == "" {
		return nil
	}
	return borders
}

func edgeFromXML(e xmlBorderEdge, theme []string) Border {
	if e.Style == "" || e.Style == "none" {
		return Border{}
	}
	return Border{Style: BorderStyle(e.Style), Color: e.Color.resolve(theme)}
}

func alignmentFromXML(a *struct {
	Horizontal   string `xml:"horizontal,attr"`
	Vertical     string `xml:"vertical,attr"`
	WrapText     string `xml:"wrapText,attr"`
	ShrinkToFit  string `xml:"shrinkToFit,attr"`
	TextRotation string `xml:"textRotation,attr"`
	Indent       string `xml:"indent,attr"`
}) *Alignment {
	if a == nil {
		return nil
	}
	al := &Alignment{}
	if a.Horizontal != "" && a.Horizontal != "general" {
		al.Horizontal = HorizontalAlignment(a.Horizontal)
	}
	if a.Vertical != "" && a.Vertical != "bottom" {
		al.Vertical = VerticalAlignment(a.Vertical)
	}
	al.WrapText = parseBoolAttr(a.WrapText, false)
	al.ShrinkToFit = parseBoolAttr(a.ShrinkToFit, false)
	if a.Indent != "" {
		al.Indent, _ = strconv.Atoi(a.Indent)
	}
	if a.TextRotation != "" {
		rot, _ := strconv.Atoi(a.TextRotation)
		if rot > 90 {
			al.TextRotation = 90 - rot
		} else {
			al.TextRotation = rot
		}
	}
	if al.Horizontal == "" && al.Vertical == "" && !al.WrapText && !al.ShrinkToFit && al.Indent == 0 && al.TextRotation == 0 {
		return nil
	}
	return al
}

func cloneStyle(s *Style) *Style {
	if s == nil {
		return nil
	}
	cp := *s
	if s.Font != nil {
		f := *s.Font
		cp.Font = &f
	}
	if s.Fill != nil {
		f := *s.Fill
		cp.Fill = &f
	}
	if s.Borders != nil {
		b := *s.Borders
		cp.Borders = &b
	}
	if s.Alignment != nil {
		a := *s.Alignment
		cp.Alignment = &a
	}
	if s.NumberFormat != nil {
		n := *s.NumberFormat
		cp.NumberFormat = &n
	}
	return &cp
}

// styleRegistry assigns cellXfs indexes for the writer.
type styleRegistry struct {
	fonts         []string
	fills         []string
	borders       []string
	numFmts       []numFmtEntry
	xfs           []string
	index         map[string]int
	extras        styleExtras
	extraDxfs     string
	extraDxfCount int
	dxfBodies     []string
}

type numFmtEntry struct {
	id   int
	code string
}

func newStyleRegistry(wb *Workbook) *styleRegistry {
	reg := &styleRegistry{
		index: map[string]int{},
		fonts: []string{`<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>`},
		fills: []string{
			`<fill><patternFill patternType="none"/></fill>`,
			`<fill><patternFill patternType="gray125"/></fill>`,
		},
		borders: []string{`<border><left/><right/><top/><bottom/><diagonal/></border>`},
		xfs:     []string{`<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`},
	}
	if wb == nil {
		return reg
	}
	reg.extras = wb.styleExtras
	for _, ws := range wb.worksheets {
		for _, sp := range ws.colSpans {
			if sp.Style != nil {
				reg.xfIndex(sp.Style)
			}
		}
		for _, c := range ws.cells {
			if st := effectiveStyle(c); st != nil {
				reg.xfIndex(st)
			}
		}
		for _, cf := range ws.conditionalFormats {
			if cf.fromFile {
				continue
			}
			for i := range cf.Rules {
				if cf.Rules[i].Style != nil {
					reg.dxfFor(cf.Rules[i].Style)
				}
			}
		}
	}
	return reg
}

func effectiveStyle(c *Cell) *Style {
	if c == nil {
		return nil
	}
	if c.Type != CellTypeDate {
		return c.Style
	}
	if c.Style != nil && c.Style.NumberFormat != nil && c.Style.NumberFormat.FormatCode != "" {
		return c.Style
	}
	base := Style{}
	if c.Style != nil {
		base = *c.Style
	}
	nf := FormatDate
	if t, ok := c.Value.(time.Time); ok && (t.Hour() != 0 || t.Minute() != 0 || t.Second() != 0 || t.Nanosecond() != 0) {
		nf = FormatDateTime
	}
	base.NumberFormat = &nf
	return &base
}

func (reg *styleRegistry) xfIndex(st *Style) int {
	if st == nil {
		return 0
	}
	key := styleKey(st)
	if id, ok := reg.index[key]; ok {
		return id
	}
	numFmtID := 0
	applyNum := false
	if st.NumberFormat != nil && st.NumberFormat.FormatCode != "" && !strings.EqualFold(st.NumberFormat.FormatCode, "General") {
		numFmtID = reg.numFmtID(st.NumberFormat.FormatCode)
		applyNum = true
	}
	fontID := 0
	applyFont := false
	if st.Font != nil && !fontIsDefault(st.Font) {
		fontID = reg.add("font", reg.fonts, fontXML(st.Font), &reg.fonts)
		applyFont = true
	}
	fillID := 0
	applyFill := false
	if st.Fill != nil && st.Fill.Type != "" && st.Fill.Type != "none" {
		fillID = reg.add("fill", reg.fills, fillXML(st.Fill), &reg.fills)
		applyFill = true
	}
	borderID := 0
	applyBorder := false
	if st.Borders != nil && !bordersEmpty(st.Borders) {
		borderID = reg.add("border", reg.borders, borderXML(st.Borders), &reg.borders)
		applyBorder = true
	}
	var b strings.Builder
	fmt.Fprintf(&b, `<xf numFmtId="%d" fontId="%d" fillId="%d" borderId="%d" xfId="0"`, numFmtID, fontID, fillID, borderID)
	if applyNum {
		b.WriteString(` applyNumberFormat="1"`)
	}
	if applyFont {
		b.WriteString(` applyFont="1"`)
	}
	if applyFill {
		b.WriteString(` applyFill="1"`)
	}
	if applyBorder {
		b.WriteString(` applyBorder="1"`)
	}
	align := ""
	if st.Alignment != nil {
		align = alignmentXML(st.Alignment)
	}
	if align != "" {
		b.WriteString(` applyAlignment="1">`)
		b.WriteString(align)
		b.WriteString(`</xf>`)
	} else {
		b.WriteString(`/>`)
	}
	id := len(reg.xfs)
	reg.xfs = append(reg.xfs, b.String())
	reg.index[key] = id
	return id
}

// dxfIndex maps a style to a dxf index, counting original file dxfs first.
func (reg *styleRegistry) dxfFor(st *Style) int {
	// Stored on the registry via a side slice. Implemented in dxfIndex.
	return reg.dxfIndex(st)
}

func (reg *styleRegistry) add(kind string, existing []string, xml string, dest *[]string) int {
	for i, e := range existing {
		if e == xml {
			return i
		}
	}
	*dest = append(*dest, xml)
	return len(*dest) - 1
}

func (reg *styleRegistry) numFmtID(code string) int {
	if id, ok := builtinNumFmtID(code); ok {
		return id
	}
	for _, nf := range reg.numFmts {
		if nf.code == code {
			return nf.id
		}
	}
	id := 164
	if len(reg.numFmts) > 0 {
		id = reg.numFmts[len(reg.numFmts)-1].id + 1
	}
	reg.numFmts = append(reg.numFmts, numFmtEntry{id: id, code: code})
	return id
}

func fontIsDefault(f *Font) bool {
	if f == nil {
		return true
	}
	return f.Name == "" && f.Size == 0 && !f.Bold && !f.Italic && !f.Underline && !f.Strikethrough && f.Color == ""
}

func bordersEmpty(b *Borders) bool {
	return b.Left.Style == "" && b.Right.Style == "" && b.Top.Style == "" && b.Bottom.Style == "" && b.Diagonal.Style == ""
}

func fontXML(f *Font) string {
	var b strings.Builder
	b.WriteString("<font>")
	if f.Bold {
		b.WriteString("<b/>")
	}
	if f.Italic {
		b.WriteString("<i/>")
	}
	if f.Underline {
		b.WriteString(`<u val="single"/>`)
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
	fmt.Fprintf(&b, `<name val="%s"/>`, xmlText(name))
	b.WriteString("</font>")
	return b.String()
}

func fillXML(f *Fill) string {
	p := f.Pattern
	if f.Type == "solid" || p == "" || p == "solid" {
		p = "solid"
	}
	c := normalizeARGB(f.Color)
	if c == "" {
		c = "FF000000"
	}
	if p == "solid" {
		return fmt.Sprintf(`<fill><patternFill patternType="solid"><fgColor rgb="%s"/><bgColor indexed="64"/></patternFill></fill>`, c)
	}
	return fmt.Sprintf(`<fill><patternFill patternType="%s"><fgColor rgb="%s"/></patternFill></fill>`, xmlText(p), c)
}

func borderXML(b *Borders) string {
	var buf strings.Builder
	buf.WriteString("<border")
	if b.DiagonalUp {
		buf.WriteString(` diagonalUp="1"`)
	}
	if b.DiagonalDown {
		buf.WriteString(` diagonalDown="1"`)
	}
	buf.WriteByte('>')
	writeEdge(&buf, "left", b.Left)
	writeEdge(&buf, "right", b.Right)
	writeEdge(&buf, "top", b.Top)
	writeEdge(&buf, "bottom", b.Bottom)
	writeEdge(&buf, "diagonal", b.Diagonal)
	buf.WriteString("</border>")
	return buf.String()
}

func writeEdge(b *strings.Builder, name string, edge Border) {
	if edge.Style == "" || edge.Style == BorderNone {
		b.WriteByte('<')
		b.WriteString(name)
		b.WriteString("/>")
		return
	}
	fmt.Fprintf(b, `<%s style="%s"`, name, xmlText(string(edge.Style)))
	if c := normalizeARGB(edge.Color); c != "" {
		fmt.Fprintf(b, `><color rgb="%s"/></%s>`, c, name)
		return
	}
	b.WriteString("/>")
}

func alignmentXML(a *Alignment) string {
	if a == nil {
		return ""
	}
	var b strings.Builder
	b.WriteString("<alignment")
	h := string(a.Horizontal)
	if h == "middle" {
		h = "center"
	}
	if h != "" && h != "general" {
		fmt.Fprintf(&b, ` horizontal="%s"`, xmlText(h))
	}
	v := string(a.Vertical)
	if v == "middle" {
		v = "center"
	}
	if v != "" {
		fmt.Fprintf(&b, ` vertical="%s"`, xmlText(v))
	}
	if a.WrapText {
		b.WriteString(` wrapText="1"`)
	}
	if a.ShrinkToFit {
		b.WriteString(` shrinkToFit="1"`)
	}
	if a.TextRotation != 0 {
		rot := a.TextRotation
		if rot < 0 {
			rot = 90 - rot
		}
		if rot < 0 {
			rot = 0
		}
		if rot > 180 {
			rot = 180
		}
		fmt.Fprintf(&b, ` textRotation="%d"`, rot)
	}
	if a.Indent > 0 {
		fmt.Fprintf(&b, ` indent="%d"`, a.Indent)
	}
	if b.Len() == len("<alignment") {
		return ""
	}
	b.WriteString("/>")
	return b.String()
}

func styleKey(st *Style) string {
	var b strings.Builder
	if st.Font != nil {
		fmt.Fprintf(&b, "f:%s|%g|%t|%t|%t|%t|%s;", st.Font.Name, st.Font.Size, st.Font.Bold, st.Font.Italic, st.Font.Underline, st.Font.Strikethrough, st.Font.Color)
	}
	if st.Fill != nil {
		fmt.Fprintf(&b, "l:%s|%s|%s;", st.Fill.Type, st.Fill.Color, st.Fill.Pattern)
	}
	if st.Borders != nil {
		br := st.Borders
		fmt.Fprintf(&b, "b:%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%t,%t;",
			br.Left.Style, br.Left.Color, br.Right.Style, br.Right.Color,
			br.Top.Style, br.Top.Color, br.Bottom.Style, br.Bottom.Color,
			br.Diagonal.Style, br.Diagonal.Color, br.DiagonalUp, br.DiagonalDown)
	}
	if st.Alignment != nil {
		a := st.Alignment
		fmt.Fprintf(&b, "a:%s|%s|%t|%t|%d|%d;", a.Horizontal, a.Vertical, a.WrapText, a.ShrinkToFit, a.TextRotation, a.Indent)
	}
	if st.NumberFormat != nil {
		b.WriteString("n:" + st.NumberFormat.FormatCode)
	}
	return b.String()
}

func trimFloat(v float64) string {
	return strconv.FormatFloat(v, 'f', -1, 64)
}

func (reg *styleRegistry) writeStyles(buf *strings.Builder) {
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	buf.WriteString(`<styleSheet xmlns="` + spreadsheetNS + `">`)
	if len(reg.numFmts) > 0 {
		fmt.Fprintf(buf, `<numFmts count="%d">`, len(reg.numFmts))
		for _, nf := range reg.numFmts {
			fmt.Fprintf(buf, `<numFmt numFmtId="%d" formatCode="%s"/>`, nf.id, xmlText(nf.code))
		}
		buf.WriteString(`</numFmts>`)
	}
	fmt.Fprintf(buf, `<fonts count="%d">`, len(reg.fonts))
	for _, f := range reg.fonts {
		buf.WriteString(f)
	}
	buf.WriteString(`</fonts>`)
	fmt.Fprintf(buf, `<fills count="%d">`, len(reg.fills))
	for _, f := range reg.fills {
		buf.WriteString(f)
	}
	buf.WriteString(`</fills>`)
	fmt.Fprintf(buf, `<borders count="%d">`, len(reg.borders))
	for _, f := range reg.borders {
		buf.WriteString(f)
	}
	buf.WriteString(`</borders>`)
	buf.WriteString(`<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>`)
	fmt.Fprintf(buf, `<cellXfs count="%d">`, len(reg.xfs))
	for _, xf := range reg.xfs {
		buf.WriteString(xf)
	}
	buf.WriteString(`</cellXfs>`)
	buf.WriteString(`<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>`)
	dxfs := reg.extras.dxfsInner + reg.extraDxfs
	dxfCount := reg.extras.dxfCount + reg.extraDxfCount
	if dxfCount > 0 || dxfs != "" {
		fmt.Fprintf(buf, `<dxfs count="%d">%s</dxfs>`, dxfCount, dxfs)
	}
	if reg.extras.tableStylesXML != "" {
		buf.WriteString(reg.extras.tableStylesXML)
	}
	if reg.extras.colorsXML != "" {
		buf.WriteString(reg.extras.colorsXML)
	}
	if reg.extras.extLstXML != "" {
		buf.WriteString(reg.extras.extLstXML)
	}
	buf.WriteString(`</styleSheet>`)
}

// dxf bookkeeping. extraDxfs is appended after original dxfs so existing
// dxfId values in preserved conditional formatting stay valid.
func (reg *styleRegistry) dxfIndex(st *Style) int {
	body := dxfXML(st)
	for i, existing := range reg.dxfBodies {
		if existing == body {
			return reg.extras.dxfCount + i
		}
	}
	reg.dxfBodies = append(reg.dxfBodies, body)
	reg.extraDxfs += body
	reg.extraDxfCount++
	return reg.extras.dxfCount + reg.extraDxfCount - 1
}

func dxfXML(st *Style) string {
	if st == nil {
		return "<dxf/>"
	}
	var b strings.Builder
	b.WriteString("<dxf>")
	if st.Font != nil && !fontIsDefault(st.Font) {
		b.WriteString(fontXML(st.Font))
	}
	if st.NumberFormat != nil && st.NumberFormat.FormatCode != "" {
		id, ok := builtinNumFmtID(st.NumberFormat.FormatCode)
		if !ok {
			id = 164
		}
		fmt.Fprintf(&b, `<numFmt numFmtId="%d" formatCode="%s"/>`, id, xmlText(st.NumberFormat.FormatCode))
	}
	if st.Fill != nil && st.Fill.Type != "" && st.Fill.Type != "none" {
		b.WriteString(fillXML(st.Fill))
	}
	if st.Alignment != nil {
		if al := alignmentXML(st.Alignment); al != "" {
			b.WriteString(al)
		}
	}
	if st.Borders != nil && !bordersEmpty(st.Borders) {
		b.WriteString(borderXML(st.Borders))
	}
	b.WriteString("</dxf>")
	return b.String()
}
