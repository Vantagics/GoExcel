package gospreadsheet

import (
	"fmt"
	"math"
	"sort"
	"strconv"
	"strings"
	"time"
)

// PivotValueField is one aggregated data column.
// Function is sum, count, average, max, min, product, countNums, stdDev, stdDevp, var, or varp.
// NumberFormat is an Excel built-in format code (for example "0.00" or "#,##0.00"); empty leaves the field general.
type PivotValueField struct {
	Field        string
	Function     string
	Name         string
	NumberFormat string
}

// PivotGrouping groups a date or numeric source field.
// Dates is coarse-to-fine: years, quarters, months, days.
// A positive Interval builds numeric bins of that width over [Start, End).
type PivotGrouping struct {
	Field    string
	Dates    []string
	Start    float64
	End      float64
	Interval float64
}

// PivotFilter chooses which items of a field stay in the summary.
// Include lists the visible item labels. On a page field, one included item
// is a single selection (pageField item="n"). Multiple is true, or more than
// one included item, sets multipleItemSelectionAllowed and hides the rest.
// On a row or column field, items not listed are written with h="1" and left
// out of the grid.
type PivotFilter struct {
	Field    string
	Include  []string
	Multiple bool
}

// PivotBuild describes a pivot table. Nil grand-total and subtotal pointers
// mean Excel's defaults (all shown).
type PivotBuild struct {
	Name              string
	Source            string
	Dest              string
	RowFields         []string
	ColumnFields      []string
	PageFields        []string
	Values            []PivotValueField
	Groupings         []PivotGrouping
	Filters           []PivotFilter
	RowGrandTotals    *bool
	ColumnGrandTotals *bool
	Subtotals         *bool
}

type pivotSelection struct {
	include map[string]bool
	single  bool
	multi   bool
	onAxis  bool
}

type pvCell struct {
	blank  bool
	text   string
	num    float64
	isNum  bool
	isDate bool
	date   time.Time
}

type pvCacheField struct {
	name      string
	srcCol    int
	synthetic bool
	textLike  bool
	items     []string
	numMin    float64
	numMax    float64
	allInt    bool
	dateMin   time.Time
	dateMax   time.Time
	hasDate   bool
	groupBy   string
	par       int
	base      int
	startNum  float64
	endNum    float64
	interval  float64
}

type pvValue struct {
	srcCol   int
	fn       string
	name     string
	numFmtID int
}

type pivotLine struct {
	labels []string
	idxs   []int
	level  int    // fields represented; grand uses 0
	kind   string // data, subtotal, grand
	recs   []int
}

type pivotCol struct {
	labels []string
	idxs   []int
	data   int
	grand  bool
	recs   []int
}

type pivotGrid struct {
	rows []pivotLine
	cols []pivotCol
	// values[row][col]
	vals           [][]float64
	headerRows     int
	rowFieldLabels []string
	valueNames     []string
}

// BuildPivotTable creates a pivot cache and pivot table from b.
// Row, column, and page fields are header names. Values lists the aggregations.
// Date and numeric groupings rewrite the grouped field into Excel group items.
func (ws *Worksheet) BuildPivotTable(b PivotBuild) (*PivotTable, error) {
	if ws == nil {
		return nil, fmt.Errorf("no worksheet")
	}
	b.Name = strings.TrimSpace(b.Name)
	if b.Name == "" {
		return nil, fmt.Errorf("pivot table name is required")
	}
	for _, p := range ws.pivots {
		if p.Name == b.Name {
			return nil, fmt.Errorf("pivot table %q already exists", b.Name)
		}
	}
	if len(b.RowFields) == 0 {
		return nil, fmt.Errorf("pivot table needs at least one row field")
	}
	if len(b.Values) == 0 {
		return nil, fmt.Errorf("pivot table needs at least one data field")
	}
	src, ref, err := ws.resolveSource(b.Source)
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
		h := ""
		if cell := src.GetCellIfExists(start.Row-1, col); cell != nil {
			h = strings.TrimSpace(cell.GetStringValue())
		}
		if h == "" {
			h = fmt.Sprintf("Column%d", col+1)
		}
		headers = append(headers, h)
	}
	colOf := map[string]int{}
	for i, h := range headers {
		colOf[h] = i
	}
	need := append(append(append([]string{}, b.RowFields...), b.ColumnFields...), b.PageFields...)
	for _, v := range b.Values {
		need = append(need, v.Field)
	}
	for _, g := range b.Groupings {
		need = append(need, g.Field)
	}
	for _, name := range need {
		if _, ok := colOf[name]; !ok {
			return nil, fmt.Errorf("field %q is not a header in %s", name, ref)
		}
	}
	var records [][]pvCell
	for r := start.Row + 1; r <= end.Row; r++ {
		row := make([]pvCell, len(headers))
		for c := range headers {
			row[c] = readPivotCell(src.GetCellIfExists(r-1, start.ColumnIdx+c))
		}
		records = append(records, row)
	}
	if len(records) == 0 {
		return nil, fmt.Errorf("pivot source %s has no data rows", ref)
	}
	cache, err := buildCacheFields(headers, records, b)
	if err != nil {
		return nil, err
	}
	rowAxis, err := resolveAxis(cache, expandFields(b.RowFields, b.Groupings))
	if err != nil {
		return nil, err
	}
	colAxis, err := resolveAxis(cache, expandFields(b.ColumnFields, b.Groupings))
	if err != nil {
		return nil, err
	}
	pageAxis, err := resolveAxis(cache, expandFields(b.PageFields, b.Groupings))
	if err != nil {
		return nil, err
	}
	values := make([]pvValue, len(b.Values))
	for i, v := range b.Values {
		fn := strings.ToLower(strings.TrimSpace(v.Function))
		if fn == "" {
			fn = "sum"
		}
		if !validAgg(fn) {
			return nil, fmt.Errorf("unknown aggregation %q", v.Function)
		}
		srcCol := colOf[v.Field]
		if fn != "count" && fn != "countnums" && !columnNumeric(records, srcCol) {
			return nil, fmt.Errorf("aggregation %s requires a numeric field, %q is not numeric", fn, v.Field)
		}
		name := v.Name
		if name == "" {
			name = aggTitle(fn) + " of " + v.Field
		}
		fmtID, _ := builtinNumFmt(v.NumberFormat)
		values[i] = pvValue{srcCol: srcCol, fn: fn, name: name, numFmtID: fmtID}
	}
	showRow := flagOr(b.RowGrandTotals, true)
	showCol := flagOr(b.ColumnGrandTotals, true)
	showSub := flagOr(b.Subtotals, true)
	selections, err := resolveSelections(cache, rowAxis, colAxis, pageAxis, b.Filters)
	if err != nil {
		return nil, err
	}
	view := records
	if len(selections) > 0 {
		view = filterPivotRecords(cache, records, selections)
	}
	grid := layoutPivot(cache, view, rowAxis, colAxis, values, showRow, showCol, showSub)
	names := make([]string, len(values))
	for i, v := range values {
		names[i] = v.name
	}
	grid.valueNames = names
	author := &pivotAuthor{
		sourceSheet:  src.Title(),
		sourceRef:    ref,
		rowName:      b.RowFields[0],
		dataName:     b.Values[0].Field,
		headers:      headers,
		gridReady:    true,
		cells:        records,
		cache:        cache,
		rowAxis:      rowAxis,
		colAxis:      colAxis,
		pageAxis:     pageAxis,
		values:       values,
		showRowGrand: showRow,
		showColGrand: showCol,
		showSub:      showSub,
		grid:         grid,
		selections:   selections,
	}
	author.textCols = make([]bool, len(headers))
	for i := range author.textCols {
		author.textCols[i] = !columnNumeric(records, i)
	}
	author.texts = make([][]string, len(records))
	author.nums = make([][]float64, len(records))
	for r, row := range records {
		author.texts[r] = make([]string, len(headers))
		author.nums[r] = make([]float64, len(headers))
		for c, cell := range row {
			author.texts[r][c] = cell.text
			if cell.isNum {
				author.nums[r][c] = cell.num
			}
		}
	}
	if len(rowAxis) > 0 {
		for _, item := range cache[rowAxis[0]].items {
			author.groups = append(author.groups, pivotGroup{key: item})
		}
	}
	destRef, err := ParseCellReference(b.Dest)
	if err != nil {
		return nil, err
	}
	if len(pageAxis) > 0 {
		// Excel shows each report filter on its own row, then a blank row,
		// above the pivot location.
		need := len(pageAxis) + 1
		if destRef.Row-1 < need {
			destRef.Row += need - (destRef.Row - 1)
		}
		top := destRef.Row - 1 - need
		for i, idx := range pageAxis {
			label, _ := CellName(top+i, destRef.ColumnIdx)
			all, _ := CellName(top+i, destRef.ColumnIdx+1)
			ws.SetCellValue(label, cache[idx].name)
			ws.SetCellValue(all, pageCaption(cache[idx], selections[idx]))
		}
	}
	writePivotCells(ws, destRef, grid)
	lastRow := destRef.Row - 1 + grid.headerRows + len(grid.rows) - 1
	lastCol := destRef.ColumnIdx + len(rowAxis) + len(grid.cols) - 1
	if lastRow < destRef.Row-1 {
		lastRow = destRef.Row - 1
	}
	if lastCol < destRef.ColumnIdx {
		lastCol = destRef.ColumnIdx
	}
	a1, _ := CellName(destRef.Row-1, destRef.ColumnIdx)
	a2, _ := CellName(lastRow, lastCol)
	loc := a1 + ":" + a2
	rowNames := make([]string, len(rowAxis))
	for i, idx := range rowAxis {
		rowNames[i] = cache[idx].name
	}
	colNames := make([]string, len(colAxis))
	for i, idx := range colAxis {
		colNames[i] = cache[idx].name
	}
	pageNames := make([]string, len(pageAxis))
	for i, idx := range pageAxis {
		pageNames[i] = cache[idx].name
	}
	cacheNames := make([]string, len(cache))
	for i, f := range cache {
		cacheNames[i] = f.name
	}
	dataFields := make([]PivotDataField, len(values))
	for i, v := range values {
		dataFields[i] = PivotDataField{
			Name: v.name, FieldIndex: v.srcCol, FieldName: b.Values[i].Field,
			Subtotal: v.fn, NumFmt: b.Values[i].NumberFormat,
		}
	}
	pt := &PivotTable{
		Name: b.Name, DataCaption: "Values", Location: loc, Style: "PivotStyleLight16",
		RowFields: rowNames, ColFields: colNames, PageFields: pageNames,
		DataFields: dataFields, CacheFields: cacheNames, author: author,
	}
	pt.origName, pt.origCaption, pt.origLoc, pt.origStyle = pt.Name, pt.DataCaption, pt.Location, pt.Style
	ws.pivots = append(ws.pivots, pt)
	return pt, nil
}

func flagOr(p *bool, def bool) bool {
	if p == nil {
		return def
	}
	return *p
}

func readPivotCell(cell *Cell) pvCell {
	if cell == nil || cell.Value == nil {
		return pvCell{blank: true}
	}
	switch v := cell.Value.(type) {
	case float64:
		return pvCell{text: formatNumber(v), num: v, isNum: true}
	case int:
		return pvCell{text: strconv.Itoa(v), num: float64(v), isNum: true}
	case bool:
		if v {
			return pvCell{text: "TRUE"}
		}
		return pvCell{text: "FALSE"}
	case time.Time:
		return pvCell{text: v.Format("2006-01-02"), isDate: true, date: v, isNum: true, num: dateToSerial(v)}
	default:
		s := strings.TrimSpace(cell.GetStringValue())
		if s == "" {
			return pvCell{blank: true}
		}
		return pvCell{text: s}
	}
}

func columnNumeric(rows [][]pvCell, col int) bool {
	n := 0
	for _, row := range rows {
		if row[col].blank {
			continue
		}
		n++
		if !row[col].isNum {
			return false
		}
	}
	return n > 0
}

func validAgg(fn string) bool {
	switch fn {
	case "sum", "count", "average", "max", "min", "product", "countnums", "stddev", "stddevp", "var", "varp":
		return true
	default:
		return false
	}
}

func aggTitle(fn string) string {
	switch fn {
	case "sum":
		return "Sum"
	case "count":
		return "Count"
	case "average":
		return "Average"
	case "max":
		return "Max"
	case "min":
		return "Min"
	case "product":
		return "Product"
	case "countnums":
		return "Count Numbers"
	case "stddev":
		return "StdDev"
	case "stddevp":
		return "StdDevp"
	case "var":
		return "Var"
	case "varp":
		return "Varp"
	default:
		return fn
	}
}

func canonicalAgg(fn string) string {
	switch fn {
	case "countnums":
		return "countNums"
	case "stddev":
		return "stdDev"
	case "stddevp":
		return "stdDevp"
	default:
		return fn
	}
}

func builtinNumFmt(code string) (int, bool) {
	switch strings.TrimSpace(code) {
	case "":
		return 0, true
	case "0":
		return 1, true
	case "0.00":
		return 2, true
	case "#,##0":
		return 3, true
	case "#,##0.00":
		return 4, true
	case "0%":
		return 9, true
	case "0.00%":
		return 10, true
	case "yyyy-mm-dd", "yyyy-mm-dd hh:mm:ss":
		return 14, true
	default:
		return 0, false
	}
}

func expandFields(fields []string, groups []PivotGrouping) []string {
	var out []string
	for _, f := range fields {
		g := groupingFor(groups, f)
		if g == nil || len(g.Dates) == 0 {
			out = append(out, f)
			continue
		}
		for _, level := range g.Dates {
			if strings.EqualFold(level, finestDate(g.Dates)) {
				out = append(out, f)
				continue
			}
			out = append(out, groupLevelName(f, level))
		}
	}
	return out
}

func finestDate(levels []string) string {
	order := map[string]int{"years": 1, "quarters": 2, "months": 3, "days": 4}
	best, name := 0, ""
	for _, l := range levels {
		if order[strings.ToLower(l)] >= best {
			best = order[strings.ToLower(l)]
			name = strings.ToLower(l)
		}
	}
	return name
}

func groupLevelName(field, level string) string {
	switch strings.ToLower(level) {
	case "years":
		return "Years"
	case "quarters":
		return "Quarters"
	case "months":
		return "Months"
	case "days":
		return field
	default:
		return field + " " + level
	}
}

func groupingFor(gs []PivotGrouping, field string) *PivotGrouping {
	for i := range gs {
		if gs[i].Field == field {
			return &gs[i]
		}
	}
	return nil
}

func buildCacheFields(headers []string, rows [][]pvCell, b PivotBuild) ([]pvCacheField, error) {
	fields := make([]pvCacheField, len(headers))
	onAxis := map[string]bool{}
	for _, name := range append(append(append([]string{}, b.RowFields...), b.ColumnFields...), b.PageFields...) {
		onAxis[name] = true
	}
	for i, h := range headers {
		f := pvCacheField{name: h, srcCol: i, base: -1, par: -1}
		// Axis fields keep one item per distinct value so filters and page
		// selections can address them. A numeric data field stays continuous.
		if columnNumeric(rows, i) && !columnDates(rows, i) && !onAxis[h] {
			f.textLike = false
			f.numMin, f.numMax = math.Inf(1), math.Inf(-1)
			f.allInt = true
			for _, row := range rows {
				if row[i].blank || !row[i].isNum {
					continue
				}
				if row[i].num < f.numMin {
					f.numMin = row[i].num
				}
				if row[i].num > f.numMax {
					f.numMax = row[i].num
				}
				if row[i].num != math.Trunc(row[i].num) {
					f.allInt = false
				}
			}
		} else {
			f.textLike = true
			seen := map[string]bool{}
			for _, row := range rows {
				label := row[i].text
				if row[i].blank {
					label = "(blank)"
				}
				if !seen[label] {
					seen[label] = true
					f.items = append(f.items, label)
				}
			}
			sort.Strings(f.items)
		}
		fields[i] = f
	}
	// Apply groupings from the end so insertions stay stable relative to later columns.
	for i := len(headers) - 1; i >= 0; i-- {
		g := groupingFor(b.Groupings, headers[i])
		if g == nil {
			continue
		}
		var err error
		fields, err = applyGrouping(fields, headers[i], rows, *g)
		if err != nil {
			return nil, err
		}
	}
	return fields, nil
}

func columnDates(rows [][]pvCell, col int) bool {
	n := 0
	for _, row := range rows {
		if row[col].blank {
			continue
		}
		n++
		if !row[col].isDate {
			return false
		}
	}
	return n > 0
}

func applyGrouping(fields []pvCacheField, name string, rows [][]pvCell, g PivotGrouping) ([]pvCacheField, error) {
	idx := -1
	for i := range fields {
		if fields[i].name == name && !fields[i].synthetic {
			idx = i
			break
		}
	}
	if idx < 0 {
		return nil, fmt.Errorf("cannot group unknown field %q", name)
	}
	src := fields[idx].srcCol
	if len(g.Dates) > 0 {
		if !columnDates(rows, src) {
			return nil, fmt.Errorf("date grouping requires date values in %q", name)
		}
		fine := finestDate(g.Dates)
		fields[idx].textLike = true
		fields[idx].groupBy = fine
		fields[idx].base = idx
		fields[idx].items = distinctDates(rows, src, fine)
		fields[idx].hasDate = true
		var dmin, dmax time.Time
		first := true
		for _, row := range rows {
			if row[src].blank || !row[src].isDate {
				continue
			}
			if first || row[src].date.Before(dmin) {
				dmin = row[src].date
			}
			if first || row[src].date.After(dmax) {
				dmax = row[src].date
			}
			first = false
		}
		fields[idx].dateMin, fields[idx].dateMax = dmin, dmax
		var extra []pvCacheField
		for _, level := range g.Dates {
			if strings.EqualFold(level, fine) {
				continue
			}
			items := distinctDates(rows, src, strings.ToLower(level))
			extra = append(extra, pvCacheField{
				name: groupLevelName(name, level), srcCol: src, synthetic: true, textLike: true,
				items: items, groupBy: strings.ToLower(level), par: idx + len(extra), base: -1,
				hasDate: true, dateMin: dmin, dateMax: dmax,
			})
		}
		// par must point at the base field's index after insertion.
		baseAfter := idx + len(extra)
		for i := range extra {
			extra[i].par = baseAfter
		}
		fields[idx].base = baseAfter
		fields = append(fields[:idx], append(extra, fields[idx:]...)...)
		return fields, nil
	}
	if g.Interval > 0 {
		items := distinctBuckets(rows, src, g.Start, g.End, g.Interval)
		fields[idx].textLike = true
		fields[idx].groupBy = "range"
		fields[idx].items = items
		fields[idx].base = idx
		fields[idx].startNum, fields[idx].endNum, fields[idx].interval = g.Start, g.End, g.Interval
		return fields, nil
	}
	return nil, fmt.Errorf("grouping for %q needs date levels or a positive interval", name)
}

func distinctDates(rows [][]pvCell, col int, level string) []string {
	seen := map[string]bool{}
	var out []string
	for _, row := range rows {
		if row[col].blank || !row[col].isDate {
			continue
		}
		label := dateLabel(row[col].date, level)
		if !seen[label] {
			seen[label] = true
			out = append(out, label)
		}
	}
	sort.Slice(out, func(i, j int) bool { return dateLabelLess(out[i], out[j], level) })
	return out
}

func dateLabel(t time.Time, level string) string {
	switch strings.ToLower(level) {
	case "years":
		return strconv.Itoa(t.Year())
	case "quarters":
		return fmt.Sprintf("Qtr%d", (int(t.Month())-1)/3+1)
	case "months":
		return t.Month().String()[:3]
	default:
		return t.Format("2006-01-02")
	}
}

func dateLabelLess(a, b, level string) bool {
	months := map[string]int{"Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6, "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12}
	switch strings.ToLower(level) {
	case "months":
		return months[a] < months[b]
	case "quarters":
		return a < b
	default:
		return a < b
	}
}

func distinctBuckets(rows [][]pvCell, col int, start, end, interval float64) []string {
	seen := map[string]bool{}
	var out []string
	for _, row := range rows {
		if row[col].blank || !row[col].isNum {
			continue
		}
		label := bucketLabel(row[col].num, start, end, interval)
		if !seen[label] {
			seen[label] = true
			out = append(out, label)
		}
	}
	sort.Slice(out, func(i, j int) bool { return bucketOrder(out[i]) < bucketOrder(out[j]) })
	return out
}

func bucketLabel(v, start, end, interval float64) string {
	if v < start {
		return "<" + trimNum(start)
	}
	if end > start && v >= end {
		return ">" + trimNum(end)
	}
	i := math.Floor((v - start) / interval)
	lo := start + i*interval
	hi := lo + interval
	if interval == math.Trunc(interval) && lo == math.Trunc(lo) {
		return fmt.Sprintf("%d-%d", int(lo), int(hi)-1)
	}
	return trimNum(lo) + "-" + trimNum(hi)
}

func bucketOrder(label string) float64 {
	if strings.HasPrefix(label, "<") {
		return math.Inf(-1)
	}
	if strings.HasPrefix(label, ">") {
		return math.Inf(1)
	}
	n, _ := strconv.ParseFloat(strings.Split(label, "-")[0], 64)
	return n
}

func trimNum(v float64) string { return formatNumber(v) }

func resolveAxis(fields []pvCacheField, names []string) ([]int, error) {
	var out []int
	for _, name := range names {
		found := -1
		for i, f := range fields {
			if f.name == name {
				found = i
				break
			}
		}
		if found < 0 {
			return nil, fmt.Errorf("pivot axis field %q is not in the cache", name)
		}
		out = append(out, found)
	}
	return out, nil
}

func cellLabel(fields []pvCacheField, field int, row []pvCell) string {
	f := fields[field]
	src := f.srcCol
	c := row[src]
	if f.groupBy == "range" {
		if c.blank || !c.isNum {
			return "(blank)"
		}
		return bucketLabel(c.num, f.startNum, f.endNum, f.interval)
	}
	if f.groupBy != "" && c.isDate {
		return dateLabel(c.date, f.groupBy)
	}
	if c.blank {
		return "(blank)"
	}
	return c.text
}

func itemIndex(f pvCacheField, label string) int {
	for i, it := range f.items {
		if it == label {
			return i
		}
	}
	return 0
}

type treeNode struct {
	label string
	index int
	recs  []int
	kids  []*treeNode
}

func buildTree(fields []pvCacheField, axis []int, rows [][]pvCell) *treeNode {
	root := &treeNode{}
	for i := range rows {
		node := root
		node.recs = append(node.recs, i)
		for _, fi := range axis {
			label := cellLabel(fields, fi, rows[i])
			idx := itemIndex(fields[fi], label)
			var child *treeNode
			for _, k := range node.kids {
				if k.label == label {
					child = k
					break
				}
			}
			if child == nil {
				child = &treeNode{label: label, index: idx}
				node.kids = append(node.kids, child)
			}
			child.recs = append(child.recs, i)
			node = child
		}
	}
	sortTree(root)
	return root
}

func sortTree(n *treeNode) {
	sort.SliceStable(n.kids, func(i, j int) bool {
		return bucketOrder(n.kids[i].label) < bucketOrder(n.kids[j].label) || (bucketOrder(n.kids[i].label) == bucketOrder(n.kids[j].label) && n.kids[i].label < n.kids[j].label)
	})
	for _, k := range n.kids {
		sortTree(k)
	}
}

func resolveSelections(cache []pvCacheField, rowAxis, colAxis, pageAxis []int, filters []PivotFilter) (map[int]pivotSelection, error) {
	if len(filters) == 0 {
		return nil, nil
	}
	byName := map[string]int{}
	for i, f := range cache {
		byName[f.name] = i
	}
	out := map[int]pivotSelection{}
	for _, f := range filters {
		idx, ok := byName[f.Field]
		if !ok {
			return nil, fmt.Errorf("pivot filter field %q is not in the cache", f.Field)
		}
		if len(f.Include) == 0 {
			return nil, fmt.Errorf("pivot filter %q needs at least one item", f.Field)
		}
		inc := map[string]bool{}
		for _, item := range f.Include {
			found := false
			for _, have := range cache[idx].items {
				if have == item {
					found = true
					break
				}
			}
			if !found {
				return nil, fmt.Errorf("pivot filter %q has no item %q", f.Field, item)
			}
			inc[item] = true
		}
		page := containsInt(pageAxis, idx)
		onAxis := page || containsInt(rowAxis, idx) || containsInt(colAxis, idx)
		single := page && len(inc) == 1 && !f.Multiple
		multi := page && (f.Multiple || len(inc) > 1)
		out[idx] = pivotSelection{include: inc, single: single, multi: multi, onAxis: onAxis}
	}
	return out, nil
}

func filterPivotRecords(fields []pvCacheField, rows [][]pvCell, sel map[int]pivotSelection) [][]pvCell {
	out := make([][]pvCell, 0, len(rows))
	for _, row := range rows {
		keep := true
		for idx, s := range sel {
			if !s.include[cellLabel(fields, idx, row)] {
				keep = false
				break
			}
		}
		if keep {
			out = append(out, row)
		}
	}
	return out
}

func pageCaption(f pvCacheField, sel pivotSelection) string {
	if sel.single {
		for _, item := range f.items {
			if sel.include[item] {
				return item
			}
		}
	}
	if sel.multi {
		return "(Multiple Items)"
	}
	return "(All)"
}

func layoutPivot(fields []pvCacheField, rows [][]pvCell, rowAxis, colAxis []int, values []pvValue, showRow, showCol, showSub bool) pivotGrid {
	rowTree := buildTree(fields, rowAxis, rows)
	var lines []pivotLine
	var walk func(n *treeNode, depth int, labels []string, idxs []int)
	walk = func(n *treeNode, depth int, labels []string, idxs []int) {
		if depth == len(rowAxis) {
			lines = append(lines, pivotLine{labels: append([]string{}, labels...), idxs: append([]int{}, idxs...), level: depth, kind: "data", recs: append([]int{}, n.recs...)})
			return
		}
		for _, k := range n.kids {
			walk(k, depth+1, append(labels, k.label), append(idxs, k.index))
			if showSub && depth+1 < len(rowAxis) {
				lines = append(lines, pivotLine{labels: append([]string{}, append(labels, k.label)...), idxs: append([]int{}, append(idxs, k.index)...), level: depth + 1, kind: "subtotal", recs: append([]int{}, k.recs...)})
			}
		}
	}
	walk(rowTree, 0, nil, nil)
	if showRow {
		all := make([]int, len(rows))
		for i := range rows {
			all[i] = i
		}
		lines = append(lines, pivotLine{labels: []string{"Grand Total"}, kind: "grand", recs: all})
	}
	colTree := buildTree(fields, colAxis, rows)
	var cols []pivotCol
	var walkC func(n *treeNode, depth int, labels []string, idxs []int)
	walkC = func(n *treeNode, depth int, labels []string, idxs []int) {
		if len(colAxis) == 0 || depth == len(colAxis) {
			recs := n.recs
			if len(colAxis) == 0 {
				recs = make([]int, len(rows))
				for i := range rows {
					recs[i] = i
				}
			}
			for di := range values {
				cols = append(cols, pivotCol{labels: append([]string{}, labels...), idxs: append([]int{}, idxs...), data: di, recs: recs})
			}
			return
		}
		for _, k := range n.kids {
			walkC(k, depth+1, append(labels, k.label), append(idxs, k.index))
		}
	}
	walkC(colTree, 0, nil, nil)
	if showCol && len(colAxis) > 0 {
		all := make([]int, len(rows))
		for i := range rows {
			all[i] = i
		}
		for di := range values {
			cols = append(cols, pivotCol{labels: []string{"Grand Total"}, data: di, grand: true, recs: all})
		}
	}
	if len(cols) == 0 {
		all := make([]int, len(rows))
		for i := range rows {
			all[i] = i
		}
		for di := range values {
			cols = append(cols, pivotCol{data: di, recs: all})
		}
	}
	header := 1
	if len(colAxis) > 0 {
		header = len(colAxis)
		if len(values) > 1 {
			header++
		}
	}
	vals := make([][]float64, len(lines))
	for r, line := range lines {
		vals[r] = make([]float64, len(cols))
		for c, col := range cols {
			ids := intersect(line.recs, col.recs)
			if len(ids) == 0 {
				vals[r][c] = math.NaN()
				continue
			}
			vals[r][c] = aggregate(rows, ids, values[col.data])
		}
	}
	names := make([]string, len(rowAxis))
	for i, idx := range rowAxis {
		names[i] = fields[idx].name
	}
	return pivotGrid{rows: lines, cols: cols, vals: vals, headerRows: header, rowFieldLabels: names}
}

func intersect(a, b []int) []int {
	set := map[int]bool{}
	for _, v := range b {
		set[v] = true
	}
	var out []int
	for _, v := range a {
		if set[v] {
			out = append(out, v)
		}
	}
	return out
}

func aggregate(rows [][]pvCell, ids []int, v pvValue) float64 {
	var nums []float64
	count := 0
	for _, id := range ids {
		cell := rows[id][v.srcCol]
		if cell.blank {
			continue
		}
		count++
		if cell.isNum {
			nums = append(nums, cell.num)
		}
	}
	switch v.fn {
	case "count":
		return float64(count)
	case "countnums":
		return float64(len(nums))
	case "sum":
		s := 0.0
		for _, n := range nums {
			s += n
		}
		return s
	case "average":
		if len(nums) == 0 {
			return 0
		}
		s := 0.0
		for _, n := range nums {
			s += n
		}
		return s / float64(len(nums))
	case "max":
		if len(nums) == 0 {
			return 0
		}
		m := nums[0]
		for _, n := range nums[1:] {
			if n > m {
				m = n
			}
		}
		return m
	case "min":
		if len(nums) == 0 {
			return 0
		}
		m := nums[0]
		for _, n := range nums[1:] {
			if n < m {
				m = n
			}
		}
		return m
	case "product":
		if len(nums) == 0 {
			return 0
		}
		p := 1.0
		for _, n := range nums {
			p *= n
		}
		return p
	case "stddev", "stddevp", "var", "varp":
		if len(nums) == 0 {
			return 0
		}
		mean := 0.0
		for _, n := range nums {
			mean += n
		}
		mean /= float64(len(nums))
		ss := 0.0
		for _, n := range nums {
			d := n - mean
			ss += d * d
		}
		den := float64(len(nums))
		if v.fn == "stddev" || v.fn == "var" {
			if len(nums) < 2 {
				return 0
			}
			den = float64(len(nums) - 1)
		}
		variance := ss / den
		if v.fn == "var" || v.fn == "varp" {
			return variance
		}
		return math.Sqrt(variance)
	default:
		return 0
	}
}

func writePivotCells(ws *Worksheet, dest *CellReference, g pivotGrid) {
	baseR := dest.Row - 1
	baseC := dest.ColumnIdx
	nRow := len(g.rowFieldLabels)
	if nRow == 0 {
		nRow = 1
	}
	// Column header rows sit above the body. The last of them also names the row fields.
	for h := 0; h < g.headerRows; h++ {
		if h == g.headerRows-1 {
			for i, name := range g.rowFieldLabels {
				cell, _ := CellName(baseR+h, baseC+i)
				ws.SetCellValue(cell, name)
			}
		}
		for c, col := range g.cols {
			label := ""
			if h < len(col.labels) {
				label = col.labels[h]
			} else if col.data >= 0 && col.data < len(g.valueNames) {
				label = g.valueNames[col.data]
			}
			if label == "" {
				continue
			}
			cell, _ := CellName(baseR+h, baseC+nRow+c)
			ws.SetCellValue(cell, label)
		}
	}
	for r, line := range g.rows {
		for i := 0; i < len(line.labels) && i < nRow; i++ {
			cell, _ := CellName(baseR+g.headerRows+r, baseC+i)
			ws.SetCellValue(cell, line.labels[i])
		}
		if line.kind == "grand" {
			cell, _ := CellName(baseR+g.headerRows+r, baseC)
			ws.SetCellValue(cell, "Grand Total")
		}
		for c := range g.cols {
			if math.IsNaN(g.vals[r][c]) {
				continue
			}
			cell, _ := CellName(baseR+g.headerRows+r, baseC+nRow+c)
			ws.SetCellValue(cell, g.vals[r][c])
		}
	}
}

func (a *pivotAuthor) renderPivot(p *PivotTable) (def, rec, table string) {
	var recB strings.Builder
	recB.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	fmt.Fprintf(&recB, `<pivotCacheRecords xmlns="%s" xmlns:r="%s" count="%d">`, spreadsheetNS, relNS, len(a.cells))
	for _, row := range a.cells {
		recB.WriteString(`<r>`)
		for _, f := range a.cache {
			if !f.textLike && f.groupBy == "" {
				cell := row[f.srcCol]
				if cell.blank || !cell.isNum {
					recB.WriteString(`<m/>`)
					continue
				}
				fmt.Fprintf(&recB, `<n v="%s"/>`, formatNumber(cell.num))
				continue
			}
			label := "(blank)"
			if f.srcCol >= 0 && f.srcCol < len(row) {
				label = cellLabel(a.cache, indexOfField(a.cache, f), row)
			}
			fmt.Fprintf(&recB, `<x v="%d"/>`, itemIndex(f, label))
		}
		recB.WriteString(`</r>`)
	}
	recB.WriteString(`</pivotCacheRecords>`)

	var defB strings.Builder
	defB.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	fmt.Fprintf(&defB, `<pivotCacheDefinition xmlns="%s" xmlns:r="%s" r:id="rId1" refreshedBy="GoExcel" createdVersion="8" refreshedVersion="8" minRefreshableVersion="3" recordCount="%d">`, spreadsheetNS, relNS, len(a.cells))
	fmt.Fprintf(&defB, `<cacheSource type="worksheet"><worksheetSource ref="%s" sheet="%s"/></cacheSource>`, xmlText(a.sourceRef), xmlText(a.sourceSheet))
	fmt.Fprintf(&defB, `<cacheFields count="%d">`, len(a.cache))
	for _, f := range a.cache {
		db := ""
		if f.synthetic {
			db = ` databaseField="0"`
		}
		fmt.Fprintf(&defB, `<cacheField name="%s" numFmtId="0"%s>`, xmlText(f.name), db)
		if f.textLike || f.groupBy != "" {
			fmt.Fprintf(&defB, `<sharedItems count="%d">`, len(f.items))
			for _, it := range f.items {
				fmt.Fprintf(&defB, `<s v="%s"/>`, xmlText(it))
			}
			defB.WriteString(`</sharedItems>`)
		} else {
			intAttr := "0"
			if f.allInt {
				intAttr = "1"
			}
			fmt.Fprintf(&defB, `<sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="%s" minValue="%s" maxValue="%s"/>`, intAttr, formatNumber(f.numMin), formatNumber(f.numMax))
		}
		if f.groupBy != "" {
			if f.par >= 0 {
				fmt.Fprintf(&defB, `<fieldGroup par="%d">`, f.par)
			} else if f.base >= 0 {
				fmt.Fprintf(&defB, `<fieldGroup base="%d">`, f.base)
			} else {
				defB.WriteString(`<fieldGroup>`)
			}
			if f.groupBy == "range" {
				fmt.Fprintf(&defB, `<rangePr groupBy="range" startNum="%s" endNum="%s" groupInterval="%s"/>`, formatNumber(f.startNum), formatNumber(f.endNum), formatNumber(f.interval))
			} else if f.hasDate {
				fmt.Fprintf(&defB, `<rangePr groupBy="%s" startDate="%s" endDate="%s"/>`, f.groupBy, f.dateMin.Format("2006-01-02T15:04:05"), f.dateMax.Format("2006-01-02T15:04:05"))
			}
			fmt.Fprintf(&defB, `<groupItems count="%d">`, len(f.items))
			for _, it := range f.items {
				fmt.Fprintf(&defB, `<s v="%s"/>`, xmlText(it))
			}
			defB.WriteString(`</groupItems></fieldGroup>`)
		}
		defB.WriteString(`</cacheField>`)
	}
	defB.WriteString(`</cacheFields><extLst><ext uri="{725AE2AE-9491-48be-B2B4-4EB974FC3084}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:pivotCacheDefinition/></ext></extLst></pivotCacheDefinition>`)

	g := a.grid
	var tb strings.Builder
	tb.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n")
	fmt.Fprintf(&tb, `<pivotTableDefinition xmlns="%s" name="%s" cacheId="%d" dataCaption="%s" updatedVersion="8" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="8" indent="0" outline="1" outlineData="1" compact="0" compactData="0" multipleFieldFilters="0"`,
		spreadsheetNS, xmlText(p.Name), a.cacheID, xmlText(p.DataCaption))
	if !a.showRowGrand {
		tb.WriteString(` rowGrandTotals="0"`)
	}
	if !a.showColGrand {
		tb.WriteString(` colGrandTotals="0"`)
	}
	tb.WriteString(`>`)
	pageAttr := ""
	if len(a.pageAxis) > 0 {
		pageAttr = fmt.Sprintf(` rowPageCount="%d" colPageCount="1"`, len(a.pageAxis))
	}
	fmt.Fprintf(&tb, `<location ref="%s" firstHeaderRow="%d" firstDataRow="%d" firstDataCol="%d"%s/>`, xmlText(p.Location), g.headerRows, g.headerRows, len(a.rowAxis), pageAttr)
	fmt.Fprintf(&tb, `<pivotFields count="%d">`, len(a.cache))
	for i, f := range a.cache {
		axis := ""
		if containsInt(a.rowAxis, i) {
			axis = "axisRow"
		} else if containsInt(a.colAxis, i) {
			axis = "axisCol"
		} else if containsInt(a.pageAxis, i) {
			axis = "axisPage"
		}
		data := false
		for _, v := range a.values {
			if v.srcCol >= 0 && a.cacheFieldForSource(v.srcCol) == i {
				data = true
			}
		}
		if axis == "" && !data {
			tb.WriteString(`<pivotField showAll="0"/>`)
			continue
		}
		tb.WriteString(`<pivotField`)
		if axis != "" {
			fmt.Fprintf(&tb, ` axis="%s"`, axis)
		}
		if data {
			tb.WriteString(` dataField="1"`)
		}
		tb.WriteString(` showAll="0"`)
		sel := a.selections[i]
		if sel.multi {
			tb.WriteString(` multipleItemSelectionAllowed="1"`)
		}
		if axis != "" && !a.showSub {
			tb.WriteString(` defaultSubtotal="0"`)
		}
		if axis == "" {
			tb.WriteString(`/>`)
			continue
		}
		hideExcluded := sel.include != nil && !sel.single
		fmt.Fprintf(&tb, `><items count="%d">`, len(f.items)+1)
		anyHidden := false
		for n, label := range f.items {
			h := ""
			if hideExcluded && !sel.include[label] {
				h = ` h="1"`
				anyHidden = true
			}
			if n == 0 {
				fmt.Fprintf(&tb, `<item x="0"%s/>`, h)
			} else {
				fmt.Fprintf(&tb, `<item x="%d"%s/>`, n, h)
			}
		}
		defH := ""
		if anyHidden {
			defH = ` h="1"`
		}
		fmt.Fprintf(&tb, `<item t="default"%s/></items></pivotField>`, defH)
	}
	tb.WriteString(`</pivotFields>`)
	if len(a.rowAxis) > 0 {
		fmt.Fprintf(&tb, `<rowFields count="%d">`, len(a.rowAxis))
		for _, idx := range a.rowAxis {
			fmt.Fprintf(&tb, `<field x="%d"/>`, idx)
		}
		tb.WriteString(`</rowFields>`)
	}
	fmt.Fprintf(&tb, `<rowItems count="%d">`, len(g.rows))
	for _, line := range g.rows {
		switch line.kind {
		case "grand":
			tb.WriteString(`<i t="grand"><x/></i>`)
		case "subtotal":
			tb.WriteString(`<i t="default">`)
			for _, idx := range line.idxs {
				writeX(&tb, idx)
			}
			tb.WriteString(`</i>`)
		default:
			tb.WriteString(`<i>`)
			for _, idx := range line.idxs {
				writeX(&tb, idx)
			}
			tb.WriteString(`</i>`)
		}
	}
	tb.WriteString(`</rowItems>`)
	if len(a.colAxis) > 0 || len(a.values) > 1 {
		n := len(a.colAxis)
		if len(a.values) > 1 {
			n++
		}
		fmt.Fprintf(&tb, `<colFields count="%d">`, n)
		for _, idx := range a.colAxis {
			fmt.Fprintf(&tb, `<field x="%d"/>`, idx)
		}
		if len(a.values) > 1 {
			tb.WriteString(`<field x="-2"/>`)
		}
		tb.WriteString(`</colFields>`)
	}
	fmt.Fprintf(&tb, `<colItems count="%d">`, len(g.cols))
	for _, col := range g.cols {
		if col.grand {
			tb.WriteString(`<i t="grand">`)
			if len(a.values) > 1 {
				writeX(&tb, col.data)
			} else {
				tb.WriteString(`<x/>`)
			}
			tb.WriteString(`</i>`)
			continue
		}
		tb.WriteString(`<i>`)
		for _, idx := range col.idxs {
			writeX(&tb, idx)
		}
		if len(a.values) > 1 {
			writeX(&tb, col.data)
		}
		if len(col.idxs) == 0 && len(a.values) <= 1 {
			tb.WriteString(`<x/>`)
		}
		tb.WriteString(`</i>`)
	}
	tb.WriteString(`</colItems>`)
	if len(a.pageAxis) > 0 {
		fmt.Fprintf(&tb, `<pageFields count="%d">`, len(a.pageAxis))
		for _, idx := range a.pageAxis {
			itemAttr := ""
			if sel := a.selections[idx]; sel.single {
				for n, label := range a.cache[idx].items {
					if sel.include[label] {
						itemAttr = fmt.Sprintf(` item="%d"`, n)
						break
					}
				}
			}
			fmt.Fprintf(&tb, `<pageField fld="%d"%s hier="-1"/>`, idx, itemAttr)
		}
		tb.WriteString(`</pageFields>`)
	}
	fmt.Fprintf(&tb, `<dataFields count="%d">`, len(a.values))
	for _, v := range a.values {
		fmt.Fprintf(&tb, `<dataField name="%s" fld="%d" subtotal="%s" baseField="0" baseItem="0"`, xmlText(v.name), a.cacheFieldForSource(v.srcCol), canonicalAgg(v.fn))
		if v.numFmtID > 0 {
			fmt.Fprintf(&tb, ` numFmtId="%d"`, v.numFmtID)
		}
		tb.WriteString(`/>`)
	}
	tb.WriteString(`</dataFields>`)
	style := p.Style
	if style == "" {
		style = "PivotStyleLight16"
	}
	fmt.Fprintf(&tb, `<pivotTableStyleInfo name="%s" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>`, xmlText(style))
	hide := "1"
	if len(a.values) > 1 {
		hide = "0"
	}
	fmt.Fprintf(&tb, `<extLst><ext uri="{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:pivotTableDefinition hideValuesRow="%s" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"/></ext><ext uri="{747A6164-185A-40DC-8AA5-F01512510D54}" xmlns:xpdl="http://schemas.microsoft.com/office/spreadsheetml/2016/pivotdefaultlayout"><xpdl:pivotTableDefinition16/></ext></extLst>`, hide)
	tb.WriteString(`</pivotTableDefinition>`)
	return defB.String(), recB.String(), tb.String()
}

func (a *pivotAuthor) cacheFieldForSource(src int) int {
	for i, f := range a.cache {
		if !f.synthetic && f.srcCol == src && f.groupBy == "" {
			return i
		}
	}
	for i, f := range a.cache {
		if !f.synthetic && f.srcCol == src {
			return i
		}
	}
	return src
}

func indexOfField(fields []pvCacheField, f pvCacheField) int {
	for i := range fields {
		if fields[i].name == f.name && fields[i].srcCol == f.srcCol && fields[i].groupBy == f.groupBy {
			return i
		}
	}
	return 0
}

func containsInt(xs []int, v int) bool {
	for _, x := range xs {
		if x == v {
			return true
		}
	}
	return false
}

func writeX(b *strings.Builder, idx int) {
	if idx == 0 {
		b.WriteString(`<x/>`)
		return
	}
	fmt.Fprintf(b, `<x v="%d"/>`, idx)
}
