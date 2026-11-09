package gospreadsheet

import (
	"fmt"
	"path/filepath"
	"strings"
)

// OpenFile opens a spreadsheet file, auto-detecting the format by extension.
func OpenFile(filename string) (*Workbook, error) {
	ext := strings.ToLower(filepath.Ext(filename))
	switch ext {
	case ".xlsx", ".xlsm", ".xltx", ".xltm":
		return NewXLSXReader().Open(filename)
	case ".csv":
		return NewCSVReader().Open(filename)
	default:
		return nil, fmt.Errorf("unsupported file format: %s", ext)
	}
}

// SaveFile saves a workbook to a file, auto-detecting the format by extension.
func SaveFile(wb *Workbook, filename string) error {
	ext := strings.ToLower(filepath.Ext(filename))
	switch ext {
	case ".xlsx", ".xlsm", ".xltx", ".xltm":
		return NewXLSXWriter().Save(wb, filename)
	case ".csv":
		return NewCSVWriter().Save(wb, filename)
	default:
		return fmt.Errorf("unsupported file format: %s", ext)
	}
}
