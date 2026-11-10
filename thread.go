package gospreadsheet

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"strconv"
	"strings"
	"time"
)

// Person is an author in a threaded-comment person list.
type Person struct {
	ID          string
	DisplayName string
	UserID      string
	Provider    string
}

// ThreadedComment is one comment in a cell thread. A reply has ParentID set
// to the root comment's ID. Resolved marks the thread done.
type ThreadedComment struct {
	ID       string
	ParentID string
	Cell     string
	PersonID string
	Text     string
	Created  time.Time
	Resolved bool
}

// Persons returns the workbook person list.
func (wb *Workbook) Persons() []Person {
	if wb == nil {
		return nil
	}
	return append([]Person(nil), wb.persons...)
}

// AddPerson appends a person and assigns an id when ID is empty.
func (wb *Workbook) AddPerson(displayName, userID string) Person {
	if wb == nil {
		return Person{}
	}
	p := Person{DisplayName: displayName, UserID: userID, Provider: "None"}
	p.ID = wb.newGUID()
	wb.persons = append(wb.persons, p)
	return p
}

// ThreadedComments returns the sheet's threaded comments, roots first.
func (ws *Worksheet) ThreadedComments() []ThreadedComment {
	if ws == nil {
		return nil
	}
	return append([]ThreadedComment(nil), ws.threads...)
}

// AddThreadedComment starts a thread on cell (for example "A1").
// A legacy note is also stored so older Excel and LibreOffice still show the text.
func (ws *Worksheet) AddThreadedComment(cell, personID, text string) (ThreadedComment, error) {
	if ws == nil {
		return ThreadedComment{}, fmt.Errorf("no worksheet")
	}
	cell = strings.ToUpper(strings.TrimSpace(cell))
	if _, err := ParseCellReference(cell); err != nil {
		return ThreadedComment{}, err
	}
	if strings.TrimSpace(personID) == "" {
		return ThreadedComment{}, fmt.Errorf("person id is required")
	}
	c := ThreadedComment{
		ID: wbGUID(ws.parent), Cell: cell, PersonID: personID,
		Text: text, Created: time.Now().UTC().Truncate(time.Millisecond),
	}
	ws.threads = append(ws.threads, c)
	ws.markThreadNote(cell)
	return c, nil
}

// ReplyToThreadedComment adds a reply under an existing comment.
func (ws *Worksheet) ReplyToThreadedComment(parentID, personID, text string) (ThreadedComment, error) {
	if ws == nil {
		return ThreadedComment{}, fmt.Errorf("no worksheet")
	}
	var parent *ThreadedComment
	for i := range ws.threads {
		if ws.threads[i].ID == parentID {
			parent = &ws.threads[i]
			break
		}
	}
	if parent == nil {
		return ThreadedComment{}, fmt.Errorf("threaded comment %q was not found", parentID)
	}
	root := parent.ID
	cell := parent.Cell
	if parent.ParentID != "" {
		root = parent.ParentID
	}
	c := ThreadedComment{
		ID: wbGUID(ws.parent), ParentID: root, Cell: cell, PersonID: personID,
		Text: text, Created: time.Now().UTC().Truncate(time.Millisecond),
	}
	ws.threads = append(ws.threads, c)
	ws.markThreadNote(cell)
	return c, nil
}

// SetThreadedCommentText replaces the text of a comment or reply.
func (ws *Worksheet) SetThreadedCommentText(id, text string) error {
	for i := range ws.threads {
		if ws.threads[i].ID == id {
			ws.threads[i].Text = text
			ws.markThreadNote(ws.threads[i].Cell)
			return nil
		}
	}
	return fmt.Errorf("threaded comment %q was not found", id)
}

// SetThreadedCommentResolved marks a thread resolved or open.
// id may be the root comment or any reply in the thread.
func (ws *Worksheet) SetThreadedCommentResolved(id string, resolved bool) error {
	root := id
	for _, c := range ws.threads {
		if c.ID == id && c.ParentID != "" {
			root = c.ParentID
			break
		}
	}
	found := false
	for i := range ws.threads {
		if ws.threads[i].ID == root && ws.threads[i].ParentID == "" {
			ws.threads[i].Resolved = resolved
			found = true
		}
	}
	if !found {
		return fmt.Errorf("threaded comment %q was not found", id)
	}
	return nil
}

func (wb *Workbook) newGUID() string {
	if wb == nil {
		return "{00000001-0000-4000-8000-000000000001}"
	}
	wb.nextID++
	return fmt.Sprintf("{%08X-0000-4000-8000-%012X}", wb.nextID, wb.nextID)
}

func wbGUID(wb *Workbook) string {
	if wb == nil {
		return "{00000001-0000-4000-8000-000000000001}"
	}
	return wb.newGUID()
}

func (wb *Workbook) hasThreads() bool {
	if wb == nil {
		return false
	}
	for _, ws := range wb.worksheets {
		if len(ws.threads) > 0 {
			return true
		}
	}
	return false
}

func (ws *Worksheet) markThreadNote(cell string) {
	if c, err := ws.GetCellByName(cell); err == nil {
		c.threadNote = true
	}
	ws.syncThreadNote(cell)
}

func (ws *Worksheet) syncThreadNote(cell string) {
	c, err := ws.GetCellByName(cell)
	if err != nil {
		return
	}
	// A note loaded from the file is left alone. A thread created or edited
	// through the API (threadNote) keeps a legacy note in sync.
	if !c.threadNote {
		return
	}
	c.SetComment(NewComment(ws.threadAuthor(cell), ws.threadPlain(cell)))
	c.threadNote = true
}

func (ws *Worksheet) threadAuthor(cell string) string {
	for _, t := range ws.threads {
		if strings.EqualFold(t.Cell, cell) && t.ParentID == "" {
			return ws.personName(t.PersonID)
		}
	}
	return ""
}

func (ws *Worksheet) personName(id string) string {
	if ws.parent == nil {
		return id
	}
	for _, p := range ws.parent.persons {
		if p.ID == id {
			if p.DisplayName != "" {
				return p.DisplayName
			}
			return p.UserID
		}
	}
	return id
}

func (ws *Worksheet) threadPlain(cell string) string {
	var b strings.Builder
	for _, t := range ws.threads {
		if !strings.EqualFold(t.Cell, cell) {
			continue
		}
		if b.Len() > 0 {
			b.WriteString("\n")
		}
		name := ws.personName(t.PersonID)
		if name != "" {
			b.WriteString(name)
			b.WriteString(": ")
		}
		b.WriteString(t.Text)
	}
	return b.String()
}

func ensureThreadFallbacks(wb *Workbook) {
	if wb == nil {
		return
	}
	for _, ws := range wb.worksheets {
		seen := map[string]bool{}
		for _, t := range ws.threads {
			if seen[strings.ToUpper(t.Cell)] {
				continue
			}
			seen[strings.ToUpper(t.Cell)] = true
			ws.syncThreadNote(t.Cell)
		}
	}
}

func readPersons(wb *Workbook, pkg *ooxmlPackage) {
	if wb == nil || pkg == nil {
		return
	}
	data, ok := pkg.file("xl/persons/person.xml")
	if !ok {
		return
	}
	dec := xml.NewDecoder(bytes.NewReader(data))
	for {
		tok, err := dec.Token()
		if err != nil {
			return
		}
		se, ok := tok.(xml.StartElement)
		if !ok || se.Name.Local != "person" {
			continue
		}
		wb.persons = append(wb.persons, Person{
			ID:          attrLocal(se, "id"),
			DisplayName: attrLocal(se, "displayName"),
			UserID:      attrLocal(se, "userId"),
			Provider:    attrLocal(se, "providerId"),
		})
	}
}

func readThreadedComments(ws *Worksheet, pkg *ooxmlPackage) {
	rel, ok := findRel(ws.sheetRels, "threadedComment")
	if !ok {
		return
	}
	part := resolvePart(ws.partName, rel.Target)
	data, ok := pkg.file(part)
	if !ok {
		return
	}
	dec := xml.NewDecoder(bytes.NewReader(data))
	for {
		tok, err := dec.Token()
		if err != nil {
			return
		}
		se, ok := tok.(xml.StartElement)
		if !ok || se.Name.Local != "threadedComment" {
			continue
		}
		c := ThreadedComment{
			ID:       attrLocal(se, "id"),
			ParentID: attrLocal(se, "parentId"),
			Cell:     attrLocal(se, "ref"),
			PersonID: attrLocal(se, "personId"),
			Resolved: attrLocal(se, "done") == "1" || strings.EqualFold(attrLocal(se, "done"), "true"),
		}
		if ts := attrLocal(se, "dT"); ts != "" {
			if t, err := time.Parse("2006-01-02T15:04:05.000", ts); err == nil {
				c.Created = t
			} else if t, err := time.Parse(time.RFC3339, ts); err == nil {
				c.Created = t
			}
		}
		var text strings.Builder
		depth := 1
		for depth > 0 {
			tok, err := dec.Token()
			if err != nil {
				break
			}
			switch t := tok.(type) {
			case xml.StartElement:
				depth++
				if t.Name.Local == "text" {
					_ = readCharData(dec, &text, "text")
					depth--
				}
			case xml.EndElement:
				depth--
			}
		}
		c.Text = text.String()
		ws.threads = append(ws.threads, c)
	}
}

func (w *XLSXWriter) writeThreads(zw *zip.Writer, wb *Workbook) error {
	if wb == nil || !wb.hasThreads() {
		return nil
	}
	if err := writeZip(zw, "xl/persons/person.xml", personsXML(wb.persons)); err != nil {
		return err
	}
	for i, ws := range wb.worksheets {
		if len(ws.threads) == 0 {
			continue
		}
		name := fmt.Sprintf("xl/threadedComments/threadedComment%d.xml", i+1)
		if err := writeZip(zw, name, threadsXML(ws.threads)); err != nil {
			return err
		}
	}
	return nil
}

func personsXML(people []Person) string {
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	b.WriteString(`<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">`)
	for _, p := range people {
		fmt.Fprintf(&b, `<person displayName="%s" id="%s" userId="%s" providerId="%s"/>`,
			xmlText(p.DisplayName), xmlText(p.ID), xmlText(p.UserID), xmlText(p.Provider))
	}
	b.WriteString(`</personList>`)
	return b.String()
}

func threadsXML(list []ThreadedComment) string {
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	b.WriteString(`<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">`)
	for _, c := range list {
		fmt.Fprintf(&b, `<threadedComment ref="%s" dT="%s" personId="%s" id="%s"`,
			xmlText(c.Cell), c.Created.UTC().Format("2006-01-02T15:04:05.000"), xmlText(c.PersonID), xmlText(c.ID))
		if c.ParentID != "" {
			fmt.Fprintf(&b, ` parentId="%s"`, xmlText(c.ParentID))
		}
		if c.Resolved && c.ParentID == "" {
			b.WriteString(` done="1"`)
		}
		fmt.Fprintf(&b, `><text>%s</text></threadedComment>`, xmlText(c.Text))
	}
	b.WriteString(`</ThreadedComments>`)
	return b.String()
}

func skipThreadPart(name string) bool {
	l := strings.ToLower(strings.ReplaceAll(name, "\\", "/"))
	return strings.HasPrefix(l, "xl/threadedcomments/") || l == "xl/persons/person.xml"
}

func threadPartName(index int) string {
	return "xl/threadedComments/threadedComment" + strconv.Itoa(index+1) + ".xml"
}
