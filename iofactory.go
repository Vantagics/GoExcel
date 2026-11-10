package gospreadsheet

import (
	"fmt"
	"path/filepath"
	"strings"
)

// Option configures OpenFile.
type Option func(*fileOptions)

type fileOptions struct {
	password string
}

// WithPassword supplies the password for an ECMA-376 encrypted workbook
// (Agile or Standard). It is ignored for a workbook that is not encrypted.
func WithPassword(password string) Option {
	return func(o *fileOptions) { o.password = password }
}

// OpenFile opens a spreadsheet file, auto-detecting the format by extension.
// Pass WithPassword to open an encrypted xlsx/xlsm/xltx/xltm package.
func OpenFile(filename string, opts ...Option) (*Workbook, error) {
	var o fileOptions
	for _, opt := range opts {
		if opt != nil {
			opt(&o)
		}
	}
	ext := strings.ToLower(filepath.Ext(filename))
	switch ext {
	case ".xlsx", ".xlsm", ".xltx", ".xltm":
		if o.password != "" {
			return NewXLSXReader().OpenPassword(filename, o.password)
		}
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
