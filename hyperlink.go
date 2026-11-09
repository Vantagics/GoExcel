package gospreadsheet

// Hyperlink represents a cell hyperlink.
type Hyperlink struct {
	URL     string
	Tooltip string
	rid     string // workbook-local relationship id, set while writing
}

// NewHyperlink creates a new hyperlink.
func NewHyperlink(url string) *Hyperlink {
	return &Hyperlink{URL: url}
}

// SetTooltip sets the tooltip and returns the hyperlink for chaining.
func (h *Hyperlink) SetTooltip(tooltip string) *Hyperlink {
	h.Tooltip = tooltip
	return h
}
