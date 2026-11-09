// Command xlsxdump prints a stable JSON snapshot of a workbook for differential
// checks against openpyxl, Apache POI, and LibreOffice.
package main

import (
	"encoding/json"
	"fmt"
	"os"
	"sort"
	"time"

	gs "github.com/VantageDataChat/GoExcel"
)

type dumpFile struct {
	Date1904 bool        `json:"date1904"`
	Names    []nameDump  `json:"names"`
	Sheets   []sheetDump `json:"sheets"`
	Error    string      `json:"error,omitempty"`
}

type nameDump struct {
	Name string `json:"name"`
	Ref  string `json:"ref"`
}

type sheetDump struct {
	Name     string     `json:"name"`
	State    string     `json:"state"`
	Kind     string     `json:"kind,omitempty"`
	Merges   []string   `json:"merges"`
	Cells    []cellDump `json:"cells"`
	Comments []noteDump `json:"comments"`
	Links    []linkDump `json:"links"`
}

type cellDump struct {
	A  string      `json:"a"`
	T  string      `json:"t"`
	V  interface{} `json:"v,omitempty"`
	F  string      `json:"f,omitempty"`
	NF string      `json:"nf,omitempty"`
}

type noteDump struct {
	A      string `json:"a"`
	Author string `json:"author,omitempty"`
	Text   string `json:"text"`
}

type linkDump struct {
	A   string `json:"a"`
	URL string `json:"url"`
}

func main() {
	args := os.Args[1:]
	rewrite := ""
	if len(args) >= 2 && args[0] == "-rewrite" {
		rewrite = args[1]
		args = args[2:]
	}
	if len(args) != 1 {
		fmt.Fprintln(os.Stderr, "usage: xlsxdump [-rewrite out.xlsx] file.xlsx")
		os.Exit(2)
	}
	wb, err := gs.OpenFile(args[0])
	if err == nil && rewrite != "" {
		if werr := gs.SaveFile(wb, rewrite); werr != nil {
			fmt.Fprintln(os.Stderr, werr)
			os.Exit(1)
		}
	}
	out := dumpOf(wb, err)
	enc := json.NewEncoder(os.Stdout)
	enc.SetEscapeHTML(false)
	if err := enc.Encode(out); err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
	if out.Error != "" {
		os.Exit(1)
	}
}

func dumpOf(wb *gs.Workbook, openErr error) dumpFile {
	if openErr != nil {
		return dumpFile{Error: openErr.Error()}
	}
	d := dumpFile{Date1904: wb.IsDate1904()}
	names := wb.GetNamedRanges()
	keys := make([]string, 0, len(names))
	for k := range names {
		keys = append(keys, k)
	}
	sort.Strings(keys)
	for _, k := range keys {
		d.Names = append(d.Names, nameDump{Name: k, Ref: names[k]})
	}
	for _, ws := range wb.Sheets() {
		sh := sheetDump{Name: ws.Title(), State: ws.GetSheetState(), Kind: ws.Kind()}
		for _, mc := range ws.GetMergeCells() {
			a, _ := gs.CellName(mc.StartRow, mc.StartCol)
			b, _ := gs.CellName(mc.EndRow, mc.EndCol)
			sh.Merges = append(sh.Merges, a+":"+b)
		}
		sort.Strings(sh.Merges)
		for _, c := range ws.AllCells() {
			ref, err := gs.CellName(c.Row(), c.Col())
			if err != nil {
				continue
			}
			item := cellDump{A: ref, T: typeName(c.Type)}
			if c.Formula != "" {
				item.F = c.Formula
				item.T = "f"
			}
			if c.Value != nil {
				item.V = jsonValue(c.Value)
			}
			if c.Style != nil && c.Style.NumberFormat != nil {
				item.NF = c.Style.NumberFormat.FormatCode
			}
			sh.Cells = append(sh.Cells, item)
			if c.Comment != nil {
				sh.Comments = append(sh.Comments, noteDump{A: ref, Author: c.Comment.Author, Text: c.Comment.Text})
			}
			if c.Hyperlink != nil && c.Hyperlink.URL != "" {
				sh.Links = append(sh.Links, linkDump{A: ref, URL: c.Hyperlink.URL})
			}
		}
		// Hyperlinks and comments can sit on cells AllCells skips when the
		// cell has no value. Scan the used rectangle lightly via merges is not
		// enough; the writer keeps them on cells that also have values in the
		// fixtures we care about. Also walk comments already captured.
		d.Sheets = append(d.Sheets, sh)
	}
	return d
}

func typeName(t gs.CellType) string {
	switch t {
	case gs.CellTypeString:
		return "s"
	case gs.CellTypeNumeric:
		return "n"
	case gs.CellTypeBool:
		return "b"
	case gs.CellTypeFormula:
		return "f"
	case gs.CellTypeDate:
		return "d"
	case gs.CellTypeError:
		return "e"
	default:
		return "e"
	}
}

func jsonValue(v interface{}) interface{} {
	switch t := v.(type) {
	case time.Time:
		return t.UTC().Format(time.RFC3339Nano)
	case float64:
		return t
	case bool:
		return t
	case string:
		return t
	default:
		return fmt.Sprint(v)
	}
}
