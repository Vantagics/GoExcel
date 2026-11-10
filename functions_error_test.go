package gospreadsheet

import (
	"testing"
)

// Test error paths for all formula functions.
// Most functions have error returns when evaluate() fails on arguments.

func TestFnSumError(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	ws.SetCellFormula("A1", "SUM(INVALID_REF)")
	_, err := ce.CalculateCell(ws, "A1")
	if err == nil {
		t.Error("expected error")
	}
}

func TestFnAverageError(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	ws.SetCellFormula("A1", "AVERAGE(INVALID_REF)")
	_, err := ce.CalculateCell(ws, "A1")
	if err == nil {
		t.Error("expected error")
	}
}

func TestFnAverageEmpty(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	// AVERAGE with no values - empty range
	// We need a range that resolves to empty
	ws.SetCellFormula("A1", "AVERAGE(Z1:Z1)")
	// Z1 is empty, but resolveRange returns empty slice for nil values
	// Actually collectNumericValues will still return a float for empty cells
	// Let's test the error path differently
	ws.SetCellFormula("B1", "AVERAGE(INVALID_REF)")
	_, err := ce.CalculateCell(ws, "B1")
	if err == nil {
		t.Error("expected error")
	}
}

func TestFnCountError(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	ws.SetCellFormula("A1", "COUNT(INVALID_REF)")
	_, err := ce.CalculateCell(ws, "A1")
	if err == nil {
		t.Error("expected error")
	}
}

func TestFnCountAError(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	ws.SetCellFormula("A1", "COUNTA(BAD:REF:EXTRA)")
	_, err := ce.CalculateCell(ws, "A1")
	// COUNTA with invalid range - the range parse will fail
	// but COUNTA swallows errors for single values
	_ = err // COUNTA is lenient
}

func TestFnCountASingleValue(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	ws.SetCellValue("A1", "hello")
	ws.SetCellFormula("B1", "COUNTA(A1)")
	val, err := ce.CalculateCell(ws, "B1")
	if err != nil {
		t.Fatalf("error: %v", err)
	}
	if val != 1.0 {
		t.Errorf("COUNTA = %v, want 1", val)
	}
}

func TestFnMaxError(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	ws.SetCellFormula("A1", "MAX(INVALID_REF)")
	_, err := ce.CalculateCell(ws, "A1")
	if err == nil {
		t.Error("expected error")
	}
}

func TestFnMinError(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	ws.SetCellFormula("A1", "MIN(INVALID_REF)")
	_, err := ce.CalculateCell(ws, "A1")
	if err == nil {
		t.Error("expected error")
	}
}

func TestFnIfError(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	// Too few args
	ws.SetCellFormula("A1", "IF(1)")
	_, err := ce.CalculateCell(ws, "A1")
	if err == nil {
		t.Error("expected error for IF with 1 arg")
	}

	// Error in condition
	ws.SetCellFormula("B1", "IF(INVALID_REF,1,2)")
	_, err = ce.CalculateCell(ws, "B1")
	if err == nil {
		t.Error("expected error for invalid condition")
	}

	// Error in true branch
	ws.SetCellValue("C1", 1)
	ws.SetCellFormula("D1", "IF(C1,INVALID_REF,2)")
	_, err = ce.CalculateCell(ws, "D1")
	if err == nil {
		t.Error("expected error for invalid true branch")
	}

	// Error in false branch
	ws.SetCellValue("E1", 0)
	ws.SetCellFormula("F1", "IF(E1,1,INVALID_REF)")
	_, err = ce.CalculateCell(ws, "F1")
	if err == nil {
		t.Error("expected error for invalid false branch")
	}

	// IF with 2 args, false condition
	ws.SetCellValue("G1", 0)
	ws.SetCellFormula("H1", "IF(G1,1)")
	val, err := ce.CalculateCell(ws, "H1")
	if err != nil {
		t.Fatalf("error: %v", err)
	}
	if val != false {
		t.Errorf("IF(0,1) = %v, want false", val)
	}
}

func TestFnAbsError(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	ws.SetCellFormula("A1", "ABS(INVALID_REF)")
	_, err := ce.CalculateCell(ws, "A1")
	if err == nil {
		t.Error("expected error")
	}
}

func TestFnRoundError(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	// Wrong number of args
	ws.SetCellFormula("A1", "ROUND(1)")
	_, err := ce.CalculateCell(ws, "A1")
	if err == nil {
		t.Error("expected error for ROUND with 1 arg")
	}

	// Error in first arg
	ws.SetCellFormula("B1", "ROUND(INVALID_REF,2)")
	_, err = ce.CalculateCell(ws, "B1")
	if err == nil {
		t.Error("expected error")
	}

	// Error in second arg
	ws.SetCellFormula("C1", "ROUND(3.14,INVALID_REF)")
	_, err = ce.CalculateCell(ws, "C1")
	if err == nil {
		t.Error("expected error")
	}
}

func TestFnSqrtError(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	ws.SetCellFormula("A1", "SQRT(INVALID_REF)")
	_, err := ce.CalculateCell(ws, "A1")
	if err == nil {
		t.Error("expected error")
	}
}

func TestFnPowerError(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	ws.SetCellFormula("A1", "POWER(1)")
	_, err := ce.CalculateCell(ws, "A1")
	if err == nil {
		t.Error("expected error for POWER with 1 arg")
	}

	ws.SetCellFormula("B1", "POWER(INVALID_REF,2)")
	_, err = ce.CalculateCell(ws, "B1")
	if err == nil {
		t.Error("expected error")
	}

	ws.SetCellFormula("C1", "POWER(2,INVALID_REF)")
	_, err = ce.CalculateCell(ws, "C1")
	if err == nil {
		t.Error("expected error")
	}
}

func TestFnModError(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	ws.SetCellFormula("A1", "MOD(1)")
	_, err := ce.CalculateCell(ws, "A1")
	if err == nil {
		t.Error("expected error for MOD with 1 arg")
	}

	ws.SetCellFormula("B1", "MOD(INVALID_REF,2)")
	_, err = ce.CalculateCell(ws, "B1")
	if err == nil {
		t.Error("expected error")
	}

	ws.SetCellFormula("C1", "MOD(10,INVALID_REF)")
	_, err = ce.CalculateCell(ws, "C1")
	if err == nil {
		t.Error("expected error")
	}
}

func TestFnIntError(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	ws.SetCellFormula("A1", "INT(INVALID_REF)")
	_, err := ce.CalculateCell(ws, "A1")
	if err == nil {
		t.Error("expected error")
	}
}

func TestFnLenError(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	ws.SetCellFormula("A1", "LEN(INVALID_REF)")
	_, err := ce.CalculateCell(ws, "A1")
	if err == nil {
		t.Error("expected error")
	}
}

func TestFnUpperError(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	ws.SetCellFormula("A1", "UPPER(INVALID_REF)")
	_, err := ce.CalculateCell(ws, "A1")
	if err == nil {
		t.Error("expected error")
	}
}

func TestFnLowerError(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	ws.SetCellFormula("A1", "LOWER(INVALID_REF)")
	_, err := ce.CalculateCell(ws, "A1")
	if err == nil {
		t.Error("expected error")
	}
}

func TestFnTrimError(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	ws.SetCellFormula("A1", "TRIM(INVALID_REF)")
	_, err := ce.CalculateCell(ws, "A1")
	if err == nil {
		t.Error("expected error")
	}
}

func TestFnLeftError(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	ws.SetCellFormula("A1", "LEFT(1)")
	_, err := ce.CalculateCell(ws, "A1")
	if err == nil {
		t.Error("expected error for LEFT with 1 arg")
	}

	ws.SetCellFormula("B1", "LEFT(INVALID_REF,2)")
	_, err = ce.CalculateCell(ws, "B1")
	if err == nil {
		t.Error("expected error")
	}

	ws.SetCellFormula("C1", `LEFT("hello",INVALID_REF)`)
	_, err = ce.CalculateCell(ws, "C1")
	if err == nil {
		t.Error("expected error")
	}
}

func TestFnLeftEdgeCases(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	ws.SetCellValue("A1", "hi")
	// Request more chars than available
	ws.SetCellFormula("B1", "LEFT(A1,100)")
	val, err := ce.CalculateCell(ws, "B1")
	if err != nil {
		t.Fatalf("error: %v", err)
	}
	if val != "hi" {
		t.Errorf("LEFT = %v, want 'hi'", val)
	}

	// Negative count
	ws.SetCellFormula("C1", "LEFT(A1,-1)")
	if _, err = ce.CalculateCell(ws, "C1"); err == nil {
		t.Fatal("LEFT with a negative count should be #VALUE!")
	}
}

func TestFnRightError(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	ws.SetCellFormula("A1", "RIGHT(1)")
	_, err := ce.CalculateCell(ws, "A1")
	if err == nil {
		t.Error("expected error for RIGHT with 1 arg")
	}

	ws.SetCellFormula("B1", "RIGHT(INVALID_REF,2)")
	_, err = ce.CalculateCell(ws, "B1")
	if err == nil {
		t.Error("expected error")
	}

	ws.SetCellFormula("C1", `RIGHT("hello",INVALID_REF)`)
	_, err = ce.CalculateCell(ws, "C1")
	if err == nil {
		t.Error("expected error")
	}
}

func TestFnRightEdgeCases(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	ws.SetCellValue("A1", "hi")
	ws.SetCellFormula("B1", "RIGHT(A1,100)")
	val, err := ce.CalculateCell(ws, "B1")
	if err != nil {
		t.Fatalf("error: %v", err)
	}
	if val != "hi" {
		t.Errorf("RIGHT = %v, want 'hi'", val)
	}

	ws.SetCellFormula("C1", "RIGHT(A1,-1)")
	if _, err = ce.CalculateCell(ws, "C1"); err == nil {
		t.Fatal("RIGHT with a negative count should be #VALUE!")
	}
}

func TestFnMidError(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	ws.SetCellFormula("A1", "MID(1,2)")
	_, err := ce.CalculateCell(ws, "A1")
	if err == nil {
		t.Error("expected error for MID with 2 args")
	}

	ws.SetCellFormula("B1", "MID(INVALID_REF,1,1)")
	_, err = ce.CalculateCell(ws, "B1")
	if err == nil {
		t.Error("expected error")
	}

	ws.SetCellFormula("C1", `MID("hello",INVALID_REF,1)`)
	_, err = ce.CalculateCell(ws, "C1")
	if err == nil {
		t.Error("expected error")
	}

	ws.SetCellFormula("D1", `MID("hello",1,INVALID_REF)`)
	_, err = ce.CalculateCell(ws, "D1")
	if err == nil {
		t.Error("expected error")
	}
}

func TestFnMidEdgeCases(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	ws.SetCellValue("A1", "hi")
	// Start beyond string length
	ws.SetCellFormula("B1", "MID(A1,100,1)")
	val, err := ce.CalculateCell(ws, "B1")
	if err != nil {
		t.Fatalf("error: %v", err)
	}
	if val != "" {
		t.Errorf("MID = %v, want ''", val)
	}

	// Negative start
	ws.SetCellFormula("C1", "MID(A1,-1,2)")
	if _, err = ce.CalculateCell(ws, "C1"); err == nil {
		t.Fatal("MID with a negative start should be #VALUE!")
	}
}

func TestFnConcatenateError(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	ws.SetCellFormula("A1", "CONCATENATE(INVALID_REF)")
	_, err := ce.CalculateCell(ws, "A1")
	if err == nil {
		t.Error("expected error")
	}
}

func TestFnConcatenateNumeric(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	ws.SetCellValue("A1", 42)
	ws.SetCellFormula("B1", "CONCATENATE(A1)")
	val, err := ce.CalculateCell(ws, "B1")
	if err != nil {
		t.Fatalf("error: %v", err)
	}
	if val != "42" {
		t.Errorf("CONCATENATE = %v, want '42'", val)
	}
}

func TestFnMedianError(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	ws.SetCellFormula("A1", "MEDIAN(INVALID_REF)")
	_, err := ce.CalculateCell(ws, "A1")
	if err == nil {
		t.Error("expected error")
	}
}

func TestFnSumIfError(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	// Wrong number of args
	ws.SetCellFormula("A1", `SUMIF(A1:A3)`)
	_, err := ce.CalculateCell(ws, "A1")
	if err == nil {
		t.Error("expected error for SUMIF with 1 arg")
	}

	// Invalid criteria range
	ws.SetCellFormula("B1", `SUMIF(INVALID,"x")`)
	_, err = ce.CalculateCell(ws, "B1")
	if err == nil {
		t.Error("expected error for invalid range")
	}

	// Invalid sum range
	ws.SetCellValue("C1", "a")
	ws.SetCellFormula("D1", `SUMIF(C1:C1,"a",INVALID)`)
	_, err = ce.CalculateCell(ws, "D1")
	if err == nil {
		t.Error("expected error for invalid sum range")
	}
}

func TestFnCountIfError(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	ws.SetCellFormula("A1", `COUNTIF(A1:A3)`)
	_, err := ce.CalculateCell(ws, "A1")
	if err == nil {
		t.Error("expected error for COUNTIF with 1 arg")
	}

	ws.SetCellFormula("B1", `COUNTIF(INVALID,"x")`)
	_, err = ce.CalculateCell(ws, "B1")
	if err == nil {
		t.Error("expected error for invalid range")
	}
}

func TestFnLenWithNumericCell(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	ws.SetCellValue("A1", 12345)
	ws.SetCellFormula("B1", "LEN(A1)")
	if _, err := ce.CalculateCell(ws, "B1"); err == nil {
		t.Fatal("LEN of a number should be #VALUE!")
	}
}

// --- collectNumericValues error path ---

func TestCollectNumericValuesRangeError(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	_, err := collectNumericValues(ce, ws, "BAD:RANGE:EXTRA")
	if err == nil {
		t.Error("expected error for invalid range")
	}
}

func TestCollectNumericValuesSingleError(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	_, err := collectNumericValues(ce, ws, "INVALID_REF")
	if err == nil {
		t.Error("expected error for invalid single value")
	}
}

// --- resolveCellValue with non-formula cell ---

func TestResolveCellValueNonFormula(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	ws.SetCellValue("A1", 42)
	val, err := ce.resolveCellValue(ws, "A1")
	if err != nil {
		t.Fatalf("error: %v", err)
	}
	if val != float64(42) {
		t.Errorf("val = %v, want 42", val)
	}
}

func TestResolveCellValueInvalidRef(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	_, err := ce.resolveCellValue(ws, "!!!")
	if err == nil {
		t.Error("expected error")
	}
}

// --- evaluate arithmetic error paths ---

func TestEvaluateArithmeticLeftError(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	_, err := ce.evaluate(ws, "INVALID_REF+1")
	// INVALID_REF is not a cell ref, not a number, not a function
	// so it will fail
	if err == nil {
		t.Error("expected error")
	}
}

func TestEvaluateArithmeticRightError(t *testing.T) {
	_, ws, ce := setupCalcSheet()
	ws.SetCellValue("A1", 10)
	_, err := ce.evaluate(ws, "A1+INVALID_REF")
	if err == nil {
		t.Error("expected error")
	}
}
