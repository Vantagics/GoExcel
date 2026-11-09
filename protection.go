package gospreadsheet

import (
	"crypto/rand"
	"crypto/sha256"
	"encoding/base64"
	"golang.org/x/crypto/pbkdf2"
)

// SheetProtection represents protection settings for a worksheet.
type SheetProtection struct {
	Sheet               bool
	Objects             bool
	Scenarios           bool
	FormatCells         bool
	FormatColumns       bool
	FormatRows          bool
	InsertColumns       bool
	InsertRows          bool
	InsertHyperlinks    bool
	DeleteColumns       bool
	DeleteRows          bool
	SelectLockedCells   bool
	Sort                bool
	AutoFilter          bool
	PivotTables         bool
	SelectUnlockedCells bool
	Password            string // hashed password
	algoName            string
	hashValue           string
	saltValue           string
	spinCount           string
}

// NewSheetProtection creates a new sheet protection with default settings.
func NewSheetProtection() *SheetProtection {
	return &SheetProtection{
		Sheet:     true,
		Objects:   true,
		Scenarios: true,
	}
}

// SetPassword sets the protection password (will be hashed).
func (sp *SheetProtection) SetPassword(password string) *SheetProtection {
	sp.Password = hashPassword(password)
	return sp
}

// AllowFormatCells allows formatting cells.
func (sp *SheetProtection) AllowFormatCells() *SheetProtection {
	sp.FormatCells = false
	return sp
}

// AllowInsertRows allows inserting rows.
func (sp *SheetProtection) AllowInsertRows() *SheetProtection {
	sp.InsertRows = false
	return sp
}

// AllowDeleteRows allows deleting rows.
func (sp *SheetProtection) AllowDeleteRows() *SheetProtection {
	sp.DeleteRows = false
	return sp
}

// AllowSort allows sorting.
func (sp *SheetProtection) AllowSort() *SheetProtection {
	sp.Sort = false
	return sp
}

// AllowAutoFilter allows auto filter.
func (sp *SheetProtection) AllowAutoFilter() *SheetProtection {
	sp.AutoFilter = false
	return sp
}

// hashPassword creates a salted, iterated hash of the password for XLSX protection.
// Uses PBKDF2 with SHA-256 and 100,000 iterations per OWASP recommendations.
func hashPassword(password string) string {
	salt := make([]byte, 16)
	if _, err := rand.Read(salt); err != nil {
		panic("crypto/rand failed: " + err.Error())
	}
	dk := pbkdf2.Key([]byte(password), salt, 100000, 32, sha256.New)
	// Encode as "salt:hash" so we can verify later
	return base64.StdEncoding.EncodeToString(salt) + ":" + base64.StdEncoding.EncodeToString(dk)
}

// WorkbookProtection represents protection settings for the workbook.
type WorkbookProtection struct {
	LockStructure bool
	LockWindows   bool
	Password      string
	revisions     bool
}

// NewWorkbookProtection creates a new workbook protection.
func NewWorkbookProtection() *WorkbookProtection {
	return &WorkbookProtection{
		LockStructure: true,
	}
}

// SetPassword sets the workbook protection password.
func (wp *WorkbookProtection) SetPassword(password string) *WorkbookProtection {
	wp.Password = hashPassword(password)
	return wp
}
