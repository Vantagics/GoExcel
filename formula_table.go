package gospreadsheet

import (
	"encoding/xml"
	"fmt"
	"strings"
)

type excelTable struct {
	name    string
	sheet   *Worksheet
	r1, c1  int
	r2, c2  int
	headers []string
	header  bool
	totals  bool
}

func (wb *Workbook) excelTables() []excelTable {
	if wb == nil {
		return nil
	}
	var out []excelTable
	for _, ws := range wb.worksheets {
		for _, rel := range ws.sheetRels {
			if !strings.Contains(strings.ToLower(rel.Type), "/table") {
				continue
			}
			part := resolvePart(ws.partName, rel.Target)
			data, ok := wb.source.file(part)
			if !ok {
				continue
			}
			if t, ok := parseExcelTable(data, ws); ok {
				out = append(out, t)
			}
		}
	}
	return out
}

func parseExcelTable(data []byte, ws *Worksheet) (excelTable, bool) {
	dec := xml.NewDecoder(strings.NewReader(string(data)))
	var t excelTable
	t.sheet = ws
	t.header = true
	var colNames []string
	for {
		tok, err := dec.Token()
		if err != nil {
			break
		}
		se, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		switch se.Name.Local {
		case "table":
			for _, a := range se.Attr {
				switch a.Name.Local {
				case "displayName":
					if t.name == "" {
						t.name = a.Value
					}
				case "name":
					t.name = a.Value
				case "ref":
					r1, c1, r2, c2, ok := parseAreaRef(a.Value)
					if ok {
						t.r1, t.c1, t.r2, t.c2 = r1, c1, r2, c2
					}
				case "headerRowCount":
					t.header = a.Value != "0"
				case "totalsRowCount":
					t.totals = a.Value != "" && a.Value != "0"
				}
			}
		case "tableColumn":
			for _, a := range se.Attr {
				if a.Name.Local == "name" {
					colNames = append(colNames, a.Value)
				}
			}
		}
	}
	t.headers = colNames
	if t.name == "" || t.r2 < t.r1 {
		return t, false
	}
	return t, true
}

func parseAreaRef(ref string) (r1, c1, r2, c2 int, ok bool) {
	ref = stripAbs(strings.TrimSpace(ref))
	parts := strings.SplitN(ref, ":", 2)
	a, err := ParseCellReference(parts[0])
	if err != nil {
		return 0, 0, 0, 0, false
	}
	b := a
	if len(parts) == 2 {
		b, err = ParseCellReference(parts[1])
		if err != nil {
			return 0, 0, 0, 0, false
		}
	}
	r1, r2 = a.Row, b.Row
	if r1 > r2 {
		r1, r2 = r2, r1
	}
	c1, c2 = a.ColumnIdx, b.ColumnIdx
	if c1 > c2 {
		c1, c2 = c2, c1
	}
	return r1, c1, r2, c2, true
}

func (p *formulaParser) parseStructured(name string) (interface{}, error) {
	// name is already consumed; p.i points at '['.
	if p.i >= len(p.s) || p.s[p.i] != '[' {
		return nil, fmt.Errorf("#NAME?")
	}
	body, end, err := readBracket(p.s, p.i)
	if err != nil {
		return nil, err
	}
	p.i = end
	table, ok := p.lookupTable(name)
	if !ok {
		return nil, fmt.Errorf("#NAME?")
	}
	spec, cols, err := parseTableSpec(body)
	if err != nil {
		return nil, err
	}
	return table.columnRange(p.ce, spec, cols)
}

func readBracket(s string, i int) (string, int, error) {
	if i >= len(s) || s[i] != '[' {
		return "", i, fmt.Errorf("#NAME?")
	}
	depth := 0
	start := i
	inStr := false
	for i < len(s) {
		ch := s[i]
		if ch == '"' {
			inStr = !inStr
		} else if !inStr {
			if ch == '[' {
				depth++
			} else if ch == ']' {
				depth--
				if depth == 0 {
					return s[start : i+1], i + 1, nil
				}
			}
		}
		i++
	}
	return "", i, fmt.Errorf("#NAME?")
}

func (p *formulaParser) lookupTable(name string) (excelTable, bool) {
	if p.ce.workbook == nil {
		return excelTable{}, false
	}
	for _, t := range p.ce.workbook.excelTables() {
		if strings.EqualFold(t.name, name) {
			return t, true
		}
	}
	return excelTable{}, false
}

func parseTableSpec(body string) (string, []string, error) {
	body = strings.TrimSpace(body)
	if len(body) < 2 || body[0] != '[' || body[len(body)-1] != ']' {
		return "", nil, fmt.Errorf("#NAME?")
	}
	inner := body[1 : len(body)-1]
	if !strings.Contains(inner, "[") {
		return "#Data", []string{cleanCol(inner)}, nil
	}
	var items []string
	for len(inner) > 0 {
		inner = strings.TrimSpace(inner)
		if inner == "" {
			break
		}
		if inner[0] == ',' {
			inner = strings.TrimSpace(inner[1:])
			continue
		}
		if inner[0] == '[' {
			item, next, err := readBracket(inner, 0)
			if err != nil {
				return "", nil, err
			}
			items = append(items, item)
			inner = inner[next:]
			continue
		}
		return "", nil, fmt.Errorf("#NAME?")
	}
	spec := "#Data"
	var cols []string
	for _, item := range items {
		raw := strings.TrimSpace(item[1 : len(item)-1])
		if strings.HasPrefix(raw, "#") || strings.HasPrefix(raw, "@") {
			if strings.HasPrefix(raw, "@") && len(raw) > 1 {
				spec = "#This Row"
				cols = append(cols, cleanCol(raw[1:]))
				continue
			}
			spec = raw
			continue
		}
		cols = append(cols, cleanCol(raw))
	}
	if spec == "@" {
		spec = "#This Row"
	}
	return spec, cols, nil
}

func cleanCol(s string) string {
	s = strings.TrimSpace(s)
	s = strings.Trim(s, "'")
	return strings.ReplaceAll(s, "''", "'")
}

func (t excelTable) columnRange(ce *CalculationEngine, spec string, cols []string) (interface{}, error) {
	r1, r2 := t.r1, t.r2
	if t.header {
		switch strings.ToLower(spec) {
		case "#headers":
			r1, r2 = t.r1, t.r1
		case "#data", "#this row":
			r1 = t.r1 + 1
			if t.totals && r2 >= r1 {
				r2--
			}
		case "#totals":
			if !t.totals {
				return nil, fmt.Errorf("#REF!")
			}
			r1, r2 = t.r2, t.r2
		case "#all":
		default:
			r1 = t.r1 + 1
			if t.totals && r2 >= r1 {
				r2--
			}
		}
	}
	if strings.EqualFold(spec, "#This Row") {
		row := ce.row + 1
		if row < r1 || row > r2 {
			return nil, fmt.Errorf("#VALUE!")
		}
		r1, r2 = row, row
	}
	c1, c2 := t.c1, t.c2
	if len(cols) > 0 {
		i1, ok := t.colIndex(cols[0])
		if !ok {
			return nil, fmt.Errorf("#REF!")
		}
		i2 := i1
		if len(cols) > 1 {
			i2, ok = t.colIndex(cols[len(cols)-1])
			if !ok {
				return nil, fmt.Errorf("#REF!")
			}
		}
		if i1 > i2 {
			i1, i2 = i2, i1
		}
		c1, c2 = t.c1+i1, t.c1+i2
	}
	if r2 < r1 || c2 < c1 {
		return nil, fmt.Errorf("#REF!")
	}
	ref := cellRef(r1, c1) + ":" + cellRef(r2, c2)
	return ce.readArea(t.sheet, ref)
}

func (t excelTable) colIndex(name string) (int, bool) {
	for i, h := range t.headers {
		if strings.EqualFold(strings.TrimSpace(h), strings.TrimSpace(name)) {
			return i, true
		}
	}
	return 0, false
}

func cellRef(row, col int) string {
	name, _ := CellName(row-1, col)
	return name
}
