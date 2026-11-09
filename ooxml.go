package gospreadsheet

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"io"
	"path"
	"strconv"
	"strings"
	"unicode/utf8"
)

// rawXML is an original element kept so round-trips do not drop features the
// in-memory model does not edit (drawings, tables, pivot caches, extLst, …).
type rawXML struct {
	Name string
	XML  string
}

type relationship struct {
	ID         string
	Type       string
	Target     string
	TargetMode string
}

// colSpan is a <col> record. Min and Max are 1-based inclusive.
type colSpan struct {
	Min         int
	Max         int
	Width       float64
	CustomWidth bool
	Hidden      bool
	BestFit     bool
	Style       *Style
}

type ooxmlPackage struct {
	files map[string][]byte
	order []string
}

func readPackage(zr *zip.Reader) *ooxmlPackage {
	p := &ooxmlPackage{files: make(map[string][]byte, len(zr.File))}
	for _, f := range zr.File {
		rc, err := f.Open()
		if err != nil {
			continue
		}
		data, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			continue
		}
		name := strings.ReplaceAll(f.Name, "\\", "/")
		if _, ok := p.files[name]; ok {
			continue
		}
		p.files[name] = data
		p.order = append(p.order, name)
	}
	return p
}

func (p *ooxmlPackage) file(name string) ([]byte, bool) {
	if p == nil {
		return nil, false
	}
	if data, ok := p.files[name]; ok {
		return data, true
	}
	for n, data := range p.files {
		if strings.EqualFold(n, name) {
			return data, true
		}
	}
	return nil, false
}

func stripBOM(b []byte) []byte {
	return bytes.TrimPrefix(b, []byte{0xEF, 0xBB, 0xBF})
}

func attrLocal(se xml.StartElement, local string) string {
	for _, a := range se.Attr {
		if a.Name.Local == local {
			return a.Value
		}
	}
	return ""
}

// resolvePart resolves a relationship target against the source part
// (for example xl/workbook.xml), returning a package-relative path.
// External targets return "".
func resolvePart(sourcePart, target string) string {
	target = strings.TrimSpace(strings.ReplaceAll(target, "\\", "/"))
	if target == "" || strings.Contains(target, "://") {
		return ""
	}
	if strings.HasPrefix(target, "/") {
		return strings.TrimPrefix(path.Clean(target), "/")
	}
	// Some producers put a package path in a relative Target.
	if strings.HasPrefix(target, "xl/") || strings.HasPrefix(target, "docProps/") || strings.HasPrefix(target, "customXml/") {
		return path.Clean(target)
	}
	dir := path.Dir(sourcePart)
	if dir == "." {
		dir = ""
	}
	return path.Clean(path.Join(dir, target))
}

func relsPathFor(part string) string {
	dir := path.Dir(part)
	base := path.Base(part)
	return path.Join(dir, "_rels", base+".rels")
}

func parseRelationships(data []byte) ([]relationship, error) {
	dec := xml.NewDecoder(bytes.NewReader(stripBOM(data)))
	var rels []relationship
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			return rels, nil
		}
		if err != nil {
			return nil, err
		}
		se, ok := tok.(xml.StartElement)
		if !ok || se.Name.Local != "Relationship" {
			continue
		}
		rels = append(rels, relationship{
			ID:         attrLocal(se, "Id"),
			Type:       attrLocal(se, "Type"),
			Target:     attrLocal(se, "Target"),
			TargetMode: attrLocal(se, "TargetMode"),
		})
	}
}

func relTypeSuffix(relType string) string {
	if i := strings.LastIndex(relType, "/"); i >= 0 {
		return relType[i+1:]
	}
	return relType
}

func findRel(rels []relationship, suffix string) (relationship, bool) {
	for _, rel := range rels {
		if relTypeSuffix(rel.Type) == suffix && !strings.EqualFold(rel.TargetMode, "External") {
			return rel, true
		}
	}
	return relationship{}, false
}

// xmlText strips characters XML 1.0 forbids and escapes the rest.
func xmlText(s string) string {
	s = sanitizeXML(s)
	var buf strings.Builder
	_ = xml.EscapeText(&buf, []byte(s))
	return buf.String()
}

func sanitizeXML(s string) string {
	if s == "" || utf8.ValidString(s) && !hasIllegalXML(s) {
		if utf8.ValidString(s) && !hasIllegalXML(s) {
			return s
		}
	}
	var b strings.Builder
	b.Grow(len(s))
	for _, r := range s {
		if r == utf8.RuneError && !utf8.ValidString(string(r)) {
			continue
		}
		if isLegalXMLRune(r) {
			b.WriteRune(r)
		}
	}
	return b.String()
}

func hasIllegalXML(s string) bool {
	for _, r := range s {
		if !isLegalXMLRune(r) {
			return true
		}
	}
	return false
}

func isLegalXMLRune(r rune) bool {
	return r == 0x9 || r == 0xA || r == 0xD || (r >= 0x20 && r <= 0xD7FF) || (r >= 0xE000 && r <= 0xFFFD) || (r >= 0x10000 && r <= 0x10FFFF)
}

func needsPreserveSpace(s string) bool {
	if s == "" {
		return false
	}
	if strings.TrimSpace(s) != s {
		return true
	}
	return strings.ContainsAny(s, "\n\r\t")
}

func firstStartTag(data []byte, local string) string {
	s := string(stripBOM(data))
	for i := 0; i < len(s); i++ {
		if s[i] != '<' {
			continue
		}
		if i+1 < len(s) && (s[i+1] == '?' || s[i+1] == '!' || s[i+1] == '/') {
			continue
		}
		j := i + 1
		for j < len(s) && s[j] != ' ' && s[j] != '>' && s[j] != '/' && s[j] != '\n' && s[j] != '\r' && s[j] != '\t' {
			j++
		}
		name := s[i+1 : j]
		if name == local || strings.HasSuffix(name, ":"+local) {
			k := strings.IndexByte(s[i:], '>')
			if k < 0 {
				return ""
			}
			return s[i : i+k+1]
		}
		k := strings.IndexByte(s[i:], '>')
		if k < 0 {
			return ""
		}
		i += k
	}
	return ""
}

// extraNamespaceDecls returns xmlns declarations and mc:Ignorable from an
// element's start tag, skipping the declarations this library always writes.
func extraNamespaceDecls(startTag string) string {
	if startTag == "" {
		return ""
	}
	var b strings.Builder
	for i := 0; i < len(startTag); i++ {
		if startTag[i] != ' ' && startTag[i] != '\n' && startTag[i] != '\t' && startTag[i] != '\r' {
			continue
		}
		j := i
		for j < len(startTag) && (startTag[j] == ' ' || startTag[j] == '\n' || startTag[j] == '\t' || startTag[j] == '\r') {
			j++
		}
		k := j
		for k < len(startTag) && startTag[k] != '=' && startTag[k] != '>' && startTag[k] != ' ' {
			k++
		}
		if k >= len(startTag) || startTag[k] != '=' || k+1 >= len(startTag) || startTag[k+1] != '"' {
			i = j
			continue
		}
		name := startTag[j:k]
		v := strings.IndexByte(startTag[k+2:], '"')
		if v < 0 {
			break
		}
		value := startTag[k+2 : k+2+v]
		end := k + 2 + v
		keep := false
		switch {
		case name == "xmlns":
			keep = value != spreadsheetNS
		case strings.HasPrefix(name, "xmlns:"):
			keep = !(name == "xmlns:r" && value == relNS)
		case strings.HasSuffix(name, ":Ignorable"):
			keep = true
		}
		if keep {
			b.WriteByte(' ')
			b.WriteString(name)
			b.WriteString(`="`)
			b.WriteString(xmlText(value))
			b.WriteByte('"')
		}
		i = end
	}
	return b.String()
}

const (
	spreadsheetNS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
	relNS         = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	pkgRelNS      = "http://schemas.openxmlformats.org/package/2006/relationships"
	ctNS          = "http://schemas.openxmlformats.org/package/2006/content-types"
)

// elementOuterXML returns the raw XML of the element that starts at the
// decoder's current position. startPos is the byte offset of the start tag.
func elementOuterXML(dec *xml.Decoder, data []byte, startPos int64) (string, error) {
	depth := 1
	for depth > 0 {
		tok, err := dec.Token()
		if err != nil {
			return "", err
		}
		switch tok.(type) {
		case xml.StartElement:
			depth++
		case xml.EndElement:
			depth--
		}
	}
	end := dec.InputOffset()
	if startPos < 0 || int(end) > len(data) || int(startPos) > len(data) {
		return "", nil
	}
	return string(data[startPos:end]), nil
}

func skipElement(dec *xml.Decoder) error {
	depth := 1
	for depth > 0 {
		tok, err := dec.Token()
		if err != nil {
			return err
		}
		switch tok.(type) {
		case xml.StartElement:
			depth++
		case xml.EndElement:
			depth--
		}
	}
	return nil
}

func quoteSheetName(name string) string {
	escaped := strings.ReplaceAll(name, "'", "''")
	return "'" + escaped + "'"
}

func contentTypeByExt(name string) string {
	ext := strings.ToLower(path.Ext(name))
	switch ext {
	case ".rels":
		return "application/vnd.openxmlformats-package.relationships+xml"
	case ".xml":
		return "application/xml"
	case ".jpeg", ".jpg":
		return "image/jpeg"
	case ".png":
		return "image/png"
	case ".gif":
		return "image/gif"
	case ".emf":
		return "image/x-emf"
	case ".wmf":
		return "image/x-wmf"
	case ".bin":
		return "application/vnd.ms-office.vbaProject"
	case ".vml":
		return "application/vnd.openxmlformats-officedocument.vmlDrawing"
	default:
		return ""
	}
}

type contentTypesDoc struct {
	defaults  map[string]string
	overrides map[string]string
}

func parseContentTypes(data []byte) contentTypesDoc {
	ct := contentTypesDoc{defaults: map[string]string{}, overrides: map[string]string{}}
	dec := xml.NewDecoder(bytes.NewReader(stripBOM(data)))
	for {
		tok, err := dec.Token()
		if err != nil {
			return ct
		}
		se, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		switch se.Name.Local {
		case "Default":
			ext := attrLocal(se, "Extension")
			c := attrLocal(se, "ContentType")
			if ext != "" && c != "" {
				ct.defaults[strings.ToLower(ext)] = c
			}
		case "Override":
			part := attrLocal(se, "PartName")
			c := attrLocal(se, "ContentType")
			if part != "" && c != "" {
				ct.overrides[part] = c
			}
		}
	}
}

func (ct contentTypesDoc) overrideFor(part string) string {
	if !strings.HasPrefix(part, "/") {
		part = "/" + part
	}
	if c, ok := ct.overrides[part]; ok {
		return c
	}
	for p, c := range ct.overrides {
		if strings.EqualFold(p, part) {
			return c
		}
	}
	ext := strings.TrimPrefix(strings.ToLower(path.Ext(part)), ".")
	if c, ok := ct.defaults[ext]; ok {
		return c
	}
	return contentTypeByExt(part)
}

func keepPackagePart(name string) bool {
	l := strings.ToLower(strings.ReplaceAll(name, "\\", "/"))
	switch {
	case strings.HasPrefix(l, "xl/theme/"),
		strings.HasPrefix(l, "xl/drawings/"),
		strings.HasPrefix(l, "xl/charts/"),
		strings.HasPrefix(l, "xl/chartsheets/"),
		strings.HasPrefix(l, "xl/media/"),
		strings.HasPrefix(l, "xl/embeddings/"),
		strings.HasPrefix(l, "xl/pivottables/"),
		strings.HasPrefix(l, "xl/pivotcache/"),
		strings.HasPrefix(l, "xl/tables/"),
		strings.HasPrefix(l, "xl/printersettings/"),
		strings.HasPrefix(l, "xl/threadedcomments/"),
		strings.HasPrefix(l, "xl/persons/"),
		strings.HasPrefix(l, "xl/ctrlprops/"),
		strings.HasPrefix(l, "xl/activex/"),
		strings.HasPrefix(l, "xl/externallinks/"),
		strings.HasPrefix(l, "xl/querytables/"),
		strings.HasPrefix(l, "xl/customxml/"),
		strings.HasPrefix(l, "customxml/"),
		l == "xl/vbaproject.bin",
		l == "xl/volatiledependencies.xml",
		l == "xl/connections.xml",
		strings.HasPrefix(l, "xl/comments"),
		(strings.HasPrefix(l, "docprops/") && l != "docprops/core.xml" && l != "docprops/app.xml"):
		return true
	default:
		return false
	}
}

func nextRelID(used map[string]bool) string {
	for n := 1; ; n++ {
		id := "rId" + strconv.Itoa(n)
		if !used[id] {
			used[id] = true
			return id
		}
	}
}

func markRelIDs(rels []relationship, used map[string]bool) {
	for _, rel := range rels {
		if rel.ID != "" {
			used[rel.ID] = true
		}
	}
}

var indexedColors = []string{
	"000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
	"000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
	"800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
	"9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
	"000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
	"00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
	"3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
	"003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
}

func indexedColor(i int) (string, bool) {
	if i >= 0 && i < len(indexedColors) {
		return indexedColors[i], true
	}
	switch i {
	case 64:
		return "000000", true
	case 65:
		return "FFFFFF", true
	default:
		return "", false
	}
}

func loadThemeColors(data []byte) []string {
	dec := xml.NewDecoder(bytes.NewReader(stripBOM(data)))
	var colors []string
	inScheme := false
	depth := 0
	schemeDepth := 0
	for {
		tok, err := dec.Token()
		if err != nil {
			return colors
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			if t.Name.Local == "clrScheme" {
				inScheme = true
				schemeDepth = depth
				continue
			}
			if !inScheme || depth != schemeDepth+1 {
				continue
			}
			colors = append(colors, readSchemeColor(dec, t))
		case xml.EndElement:
			if inScheme && t.Name.Local == "clrScheme" {
				return colors
			}
			depth--
		}
	}
}

func readSchemeColor(dec *xml.Decoder, start xml.StartElement) string {
	color := "000000"
	depth := 1
	for depth > 0 {
		tok, err := dec.Token()
		if err != nil {
			return color
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			switch t.Name.Local {
			case "srgbClr":
				if v := attrLocal(t, "val"); v != "" {
					color = strings.ToUpper(v)
				}
			case "sysClr":
				if v := attrLocal(t, "lastClr"); v != "" {
					color = strings.ToUpper(v)
				}
			}
		case xml.EndElement:
			depth--
		}
	}
	return color
}

func resolveColor(rgb, themeIdx, indexed string, tint float64, theme []string) string {
	if rgb != "" {
		return strings.ToUpper(strings.TrimPrefix(rgb, "#"))
	}
	if themeIdx != "" {
		i, err := strconv.Atoi(themeIdx)
		if err == nil && i >= 0 && i < len(theme) {
			c := theme[i]
			if tint != 0 {
				c = applyTint(c, tint)
			}
			return strings.ToUpper(c)
		}
	}
	if indexed != "" {
		i, err := strconv.Atoi(indexed)
		if err == nil {
			if c, ok := indexedColor(i); ok {
				return c
			}
		}
	}
	return ""
}

func applyTint(hex string, tint float64) string {
	hex = strings.TrimPrefix(strings.ToUpper(hex), "#")
	if len(hex) == 8 {
		hex = hex[2:]
	}
	if len(hex) != 6 || tint == 0 {
		return hex
	}
	rv, err1 := strconv.ParseUint(hex[0:2], 16, 8)
	gv, err2 := strconv.ParseUint(hex[2:4], 16, 8)
	bv, err3 := strconv.ParseUint(hex[4:6], 16, 8)
	if err1 != nil || err2 != nil || err3 != nil {
		return hex
	}
	h, s, l := rgbToHSL(uint8(rv), uint8(gv), uint8(bv))
	if tint < 0 {
		l *= 1 + tint
	} else {
		l = l*(1-tint) + tint
	}
	r, g, b := hslToRGB(h, s, l)
	return strings.ToUpper(hexRGB(r, g, b))
}

func rgbToHSL(r, g, b uint8) (h, s, l float64) {
	rf, gf, bf := float64(r)/255, float64(g)/255, float64(b)/255
	maxv := rf
	if gf > maxv {
		maxv = gf
	}
	if bf > maxv {
		maxv = bf
	}
	minv := rf
	if gf < minv {
		minv = gf
	}
	if bf < minv {
		minv = bf
	}
	l = (maxv + minv) / 2
	if maxv == minv {
		return 0, 0, l
	}
	d := maxv - minv
	if l > 0.5 {
		s = d / (2 - maxv - minv)
	} else {
		s = d / (maxv + minv)
	}
	switch maxv {
	case rf:
		h = (gf - bf) / d
		if gf < bf {
			h += 6
		}
	case gf:
		h = (bf-rf)/d + 2
	default:
		h = (rf-gf)/d + 4
	}
	h /= 6
	return h, s, l
}

func hslToRGB(h, s, l float64) (r, g, b uint8) {
	if s == 0 {
		v := uint8(l*255 + 0.5)
		return v, v, v
	}
	var q float64
	if l < 0.5 {
		q = l * (1 + s)
	} else {
		q = l + s - l*s
	}
	p := 2*l - q
	return hueToRGB(p, q, h+1.0/3.0), hueToRGB(p, q, h), hueToRGB(p, q, h-1.0/3.0)
}

func hueToRGB(p, q, t float64) uint8 {
	if t < 0 {
		t++
	}
	if t > 1 {
		t--
	}
	var v float64
	switch {
	case t < 1.0/6.0:
		v = p + (q-p)*6*t
	case t < 0.5:
		v = q
	case t < 2.0/3.0:
		v = p + (q-p)*(2.0/3.0-t)*6
	default:
		v = p
	}
	if v < 0 {
		v = 0
	}
	if v > 1 {
		v = 1
	}
	return uint8(v*255 + 0.5)
}

func hexRGB(r, g, b uint8) string {
	const digits = "0123456789ABCDEF"
	return string([]byte{digits[r>>4], digits[r&0xF], digits[g>>4], digits[g&0xF], digits[b>>4], digits[b&0xF]})
}

func normalizeARGB(c string) string {
	c = strings.TrimPrefix(strings.ToUpper(strings.TrimSpace(c)), "#")
	switch len(c) {
	case 6:
		return "FF" + c
	case 8:
		return c
	default:
		return ""
	}
}

func builtinNumFmtCode(id int) (string, bool) {
	code, ok := builtinNumFmts[id]
	return code, ok
}

func builtinNumFmtID(code string) (int, bool) {
	id, ok := builtinNumFmtIDs[code]
	return id, ok
}

var builtinNumFmts = map[int]string{
	0:  "General",
	1:  "0",
	2:  "0.00",
	3:  "#,##0",
	4:  "#,##0.00",
	9:  "0%",
	10: "0.00%",
	11: "0.00E+00",
	12: "# ?/?",
	13: "# ??/??",
	14: "m/d/yyyy",
	15: "d-mmm-yy",
	16: "d-mmm",
	17: "mmm-yy",
	18: "h:mm AM/PM",
	19: "h:mm:ss AM/PM",
	20: "h:mm",
	21: "h:mm:ss",
	22: "m/d/yyyy h:mm",
	37: "#,##0 ;(#,##0)",
	38: "#,##0 ;[Red](#,##0)",
	39: "#,##0.00;(#,##0.00)",
	40: "#,##0.00;[Red](#,##0.00)",
	45: "mm:ss",
	46: "[h]:mm:ss",
	47: "mmss.0",
	48: "##0.0E+0",
	49: "@",
}

var builtinNumFmtIDs = func() map[string]int {
	m := make(map[string]int, len(builtinNumFmts))
	for id, code := range builtinNumFmts {
		m[code] = id
	}
	return m
}()

func isExternalURL(u string) bool {
	l := strings.ToLower(strings.TrimSpace(u))
	return strings.Contains(l, "://") || strings.HasPrefix(l, "mailto:")
}

func parseBoolAttr(v string, def bool) bool {
	switch strings.ToLower(strings.TrimSpace(v)) {
	case "1", "true":
		return true
	case "0", "false":
		return false
	default:
		return def
	}
}

// captureRootChildren returns the raw XML of selected direct children of the
// document element, in document order. Names not listed are skipped.
func captureRootChildren(data []byte, names map[string]bool) ([]rawXML, error) {
	data = stripBOM(data)
	dec := xml.NewDecoder(bytes.NewReader(data))
	// Consume the document element start tag.
	for {
		tok, err := dec.Token()
		if err != nil {
			return nil, err
		}
		if _, ok := tok.(xml.StartElement); ok {
			break
		}
	}
	var out []rawXML
	for {
		startPos := dec.InputOffset()
		tok, err := dec.Token()
		if err == io.EOF {
			return out, nil
		}
		if err != nil {
			return nil, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if names[t.Name.Local] {
				raw, err := elementOuterXML(dec, data, startPos)
				if err != nil {
					return nil, err
				}
				out = append(out, rawXML{Name: t.Name.Local, XML: strings.TrimSpace(raw)})
			} else if err := skipElement(dec); err != nil {
				return nil, err
			}
		case xml.EndElement:
			return out, nil
		}
	}
}

func rootLocalName(data []byte) string {
	dec := xml.NewDecoder(bytes.NewReader(stripBOM(data)))
	for {
		tok, err := dec.Token()
		if err != nil {
			return ""
		}
		if se, ok := tok.(xml.StartElement); ok {
			return se.Name.Local
		}
	}
}
