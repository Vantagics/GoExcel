package gospreadsheet

// PaperSize represents standard paper sizes.
type PaperSize int

const (
	PaperLetter      PaperSize = 1
	PaperLetterSmall PaperSize = 2
	PaperTabloid     PaperSize = 3
	PaperLedger      PaperSize = 4
	PaperLegal       PaperSize = 5
	PaperA3          PaperSize = 8
	PaperA4          PaperSize = 9
	PaperA4Small     PaperSize = 10
	PaperA5          PaperSize = 11
	PaperB4          PaperSize = 12
	PaperB5          PaperSize = 13
)

// Orientation represents page orientation.
type Orientation string

const (
	OrientationPortrait  Orientation = "portrait"
	OrientationLandscape Orientation = "landscape"
)

// PageMargins represents page margins in inches.
type PageMargins struct {
	Top    float64
	Bottom float64
	Left   float64
	Right  float64
	Header float64
	Footer float64
}

// DefaultPageMargins returns the default page margins.
func DefaultPageMargins() PageMargins {
	return PageMargins{
		Top:    0.75,
		Bottom: 0.75,
		Left:   0.7,
		Right:  0.7,
		Header: 0.3,
		Footer: 0.3,
	}
}

// HeaderFooter represents page header and footer content.
type HeaderFooter struct {
	OddHeader        string
	OddFooter        string
	EvenHeader       string
	EvenFooter       string
	DifferentOddEven bool
}

// PrintArea represents the print area of a worksheet.
type PrintArea struct {
	Range string // e.g., "A1:H50"
}

// PageSetup represents the page setup/print settings for a worksheet.
type PageSetup struct {
	PaperSize          PaperSize
	Orientation        Orientation
	Scale              int // percentage, 10-400
	FitToWidth         int // number of pages wide
	FitToHeight        int // number of pages tall
	Margins            PageMargins
	HeaderFooter       *HeaderFooter
	PrintArea          *PrintArea
	PrintGridlines     bool
	PrintHeadings      bool
	CenterHorizontally bool
	CenterVertically   bool
	RepeatRows         string // e.g., "1:2" for rows 1-2
	RepeatColumns      string // e.g., "A:B" for columns A-B
	rid                string // relationship id of printer settings, when present
}

// NewPageSetup creates a new page setup with defaults.
func NewPageSetup() *PageSetup {
	return &PageSetup{
		PaperSize:   PaperA4,
		Orientation: OrientationPortrait,
		Scale:       100,
		Margins:     DefaultPageMargins(),
	}
}

// SetPaperSize sets the paper size.
func (ps *PageSetup) SetPaperSize(size PaperSize) *PageSetup {
	ps.PaperSize = size
	return ps
}

// SetOrientation sets the page orientation.
func (ps *PageSetup) SetOrientation(o Orientation) *PageSetup {
	ps.Orientation = o
	return ps
}

// SetScale sets the print scale percentage.
func (ps *PageSetup) SetScale(scale int) *PageSetup {
	if scale < 10 {
		scale = 10
	}
	if scale > 400 {
		scale = 400
	}
	ps.Scale = scale
	return ps
}

// SetFitToPage sets fit-to-page printing.
func (ps *PageSetup) SetFitToPage(width, height int) *PageSetup {
	ps.FitToWidth = width
	ps.FitToHeight = height
	return ps
}

// SetMargins sets all page margins.
func (ps *PageSetup) SetMargins(m PageMargins) *PageSetup {
	ps.Margins = m
	return ps
}

// SetPrintArea sets the print area.
func (ps *PageSetup) SetPrintArea(rangeStr string) *PageSetup {
	ps.PrintArea = &PrintArea{Range: rangeStr}
	return ps
}

// SetRepeatRows sets the rows to repeat at top of each page.
func (ps *PageSetup) SetRepeatRows(rows string) *PageSetup {
	ps.RepeatRows = rows
	return ps
}

// SetRepeatColumns sets the columns to repeat at left of each page.
func (ps *PageSetup) SetRepeatColumns(cols string) *PageSetup {
	ps.RepeatColumns = cols
	return ps
}
