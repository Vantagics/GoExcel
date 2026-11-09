package gospreadsheet

// ConditionalType represents the type of conditional formatting rule.
type ConditionalType string

const (
	ConditionalCellIs       ConditionalType = "cellIs"
	ConditionalExpression   ConditionalType = "expression"
	ConditionalColorScale   ConditionalType = "colorScale"
	ConditionalDataBar      ConditionalType = "dataBar"
	ConditionalIconSet      ConditionalType = "iconSet"
	ConditionalTop10        ConditionalType = "top10"
	ConditionalAboveAvg     ConditionalType = "aboveAverage"
	ConditionalDuplicates   ConditionalType = "duplicateValues"
	ConditionalUniqueVals   ConditionalType = "uniqueValues"
	ConditionalContainsText ConditionalType = "containsText"
)

// ConditionalOperator represents the comparison operator for cellIs rules.
type ConditionalOperator string

const (
	OperatorEqual          ConditionalOperator = "equal"
	OperatorNotEqual       ConditionalOperator = "notEqual"
	OperatorGreaterThan    ConditionalOperator = "greaterThan"
	OperatorGreaterOrEqual ConditionalOperator = "greaterThanOrEqual"
	OperatorLessThan       ConditionalOperator = "lessThan"
	OperatorLessOrEqual    ConditionalOperator = "lessThanOrEqual"
	OperatorBetween        ConditionalOperator = "between"
	OperatorNotBetween     ConditionalOperator = "notBetween"
)

// ConditionalRule represents a single conditional formatting rule.
type ConditionalRule struct {
	Type       ConditionalType
	Operator   ConditionalOperator
	Formula    []string // one or two formulas depending on operator
	Style      *Style
	Priority   int
	StopIfTrue bool
}

// ConditionalFormatting represents conditional formatting applied to a range.
type ConditionalFormatting struct {
	Range    string // e.g., "A1:A10"
	Rules    []ConditionalRule
	fromFile bool
}

// NewConditionalFormatting creates a new conditional formatting for a range.
func NewConditionalFormatting(rangeStr string) *ConditionalFormatting {
	return &ConditionalFormatting{
		Range: rangeStr,
		Rules: make([]ConditionalRule, 0),
	}
}

// AddRule adds a conditional formatting rule.
func (cf *ConditionalFormatting) AddRule(rule ConditionalRule) *ConditionalFormatting {
	if rule.Priority == 0 {
		rule.Priority = len(cf.Rules) + 1
	}
	cf.Rules = append(cf.Rules, rule)
	return cf
}

// CellIsRule creates a convenience cellIs rule.
func CellIsRule(op ConditionalOperator, formula string, style *Style) ConditionalRule {
	return ConditionalRule{
		Type:     ConditionalCellIs,
		Operator: op,
		Formula:  []string{formula},
		Style:    style,
	}
}

// BetweenRule creates a convenience between rule.
func BetweenRule(formula1, formula2 string, style *Style) ConditionalRule {
	return ConditionalRule{
		Type:     ConditionalCellIs,
		Operator: OperatorBetween,
		Formula:  []string{formula1, formula2},
		Style:    style,
	}
}

// ExpressionRule creates a convenience expression rule.
func ExpressionRule(formula string, style *Style) ConditionalRule {
	return ConditionalRule{
		Type:    ConditionalExpression,
		Formula: []string{formula},
		Style:   style,
	}
}
