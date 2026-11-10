package gospreadsheet

import (
	"fmt"
	"strings"
	"time"
)

func holidayList(ce *CalculationEngine, ws *Worksheet, expr string) ([]time.Time, error) {
	rg, err := ce.evalRange(ws, expr)
	if err != nil {
		return nil, err
	}
	var out []time.Time
	for _, v := range flattenAreas(rg) {
		if v == nil || excelError(v) != nil {
			continue
		}
		if n, ok := asNumber(v, true); ok {
			out = append(out, excelSerialToTime(n))
			continue
		}
		if t, ok := v.(time.Time); ok {
			out = append(out, t)
		}
	}
	return out, nil
}

func sameDay(a, b time.Time) bool {
	ay, am, ad := a.Date()
	by, bm, bd := b.Date()
	return ay == by && am == bm && ad == bd
}

func isHoliday(d time.Time, holidays []time.Time) bool {
	for _, h := range holidays {
		if sameDay(d, h) {
			return true
		}
	}
	return false
}

func networkDayCount(start, end time.Time, weekend [7]bool, holidays []time.Time) int {
	sign := 1
	if end.Before(start) {
		start, end = end, start
		sign = -1
	}
	n := 0
	for d := start; !d.After(end); d = d.AddDate(0, 0, 1) {
		if weekend[d.Weekday()] || isHoliday(d, holidays) {
			continue
		}
		n++
	}
	return n * sign
}

func workdayShift(start time.Time, days int, weekend [7]bool, holidays []time.Time) time.Time {
	step := 1
	left := days
	if left < 0 {
		step = -1
		left = -left
	}
	d := start
	for left > 0 {
		d = d.AddDate(0, 0, step)
		if weekend[d.Weekday()] || isHoliday(d, holidays) {
			continue
		}
		left--
	}
	return d
}

func weekendPattern(v interface{}) ([7]bool, error) {
	var m [7]bool
	if s, ok := v.(string); ok && len([]rune(s)) == 7 {
		rs := []rune(s)
		for i, ch := range rs {
			if ch != '0' && ch != '1' {
				return m, fmt.Errorf("#VALUE!")
			}
			if ch == '1' {
				// String starts on Monday.
				m[time.Weekday((i+1)%7)] = true
			}
		}
		return m, nil
	}
	n, ok := asNumber(v, true)
	if !ok {
		return m, fmt.Errorf("#VALUE!")
	}
	switch int(n) {
	case 1:
		m[time.Saturday], m[time.Sunday] = true, true
	case 2:
		m[time.Sunday], m[time.Monday] = true, true
	case 3:
		m[time.Monday], m[time.Tuesday] = true, true
	case 4:
		m[time.Tuesday], m[time.Wednesday] = true, true
	case 5:
		m[time.Wednesday], m[time.Thursday] = true, true
	case 6:
		m[time.Thursday], m[time.Friday] = true, true
	case 7:
		m[time.Friday], m[time.Saturday] = true, true
	case 11:
		m[time.Sunday] = true
	case 12:
		m[time.Monday] = true
	case 13:
		m[time.Tuesday] = true
	case 14:
		m[time.Wednesday] = true
	case 15:
		m[time.Thursday] = true
	case 16:
		m[time.Friday] = true
	case 17:
		m[time.Saturday] = true
	default:
		return m, fmt.Errorf("#NUM!")
	}
	return m, nil
}

func fnNetWorkDaysIntl(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	start, err := evalDate(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	end, err := evalDate(ce, ws, parts[1])
	if err != nil {
		return nil, err
	}
	weekend := [7]bool{time.Sunday: true, time.Saturday: true}
	if len(parts) >= 3 && strings.TrimSpace(parts[2]) != "" {
		v, err := ce.evaluate(ws, parts[2])
		if err != nil {
			return nil, err
		}
		sv, err := ce.asScalar(v)
		if err != nil {
			return nil, err
		}
		weekend, err = weekendPattern(sv)
		if err != nil {
			return nil, err
		}
	}
	var holidays []time.Time
	if len(parts) >= 4 {
		holidays, err = holidayList(ce, ws, parts[3])
		if err != nil {
			return nil, err
		}
	}
	return float64(networkDayCount(start, end, weekend, holidays)), nil
}

func fnWorkdayIntl(ce *CalculationEngine, ws *Worksheet, args string) (interface{}, error) {
	parts := parseArgs(args)
	if len(parts) < 2 {
		return nil, fmt.Errorf("#VALUE!")
	}
	start, err := evalDate(ce, ws, parts[0])
	if err != nil {
		return nil, err
	}
	daysV, err := ce.evaluate(ws, parts[1])
	if err != nil {
		return nil, err
	}
	daysN, err := ce.mustNum(daysV)
	if err != nil {
		return nil, err
	}
	weekend := [7]bool{time.Sunday: true, time.Saturday: true}
	if len(parts) >= 3 && strings.TrimSpace(parts[2]) != "" {
		v, err := ce.evaluate(ws, parts[2])
		if err != nil {
			return nil, err
		}
		sv, err := ce.asScalar(v)
		if err != nil {
			return nil, err
		}
		weekend, err = weekendPattern(sv)
		if err != nil {
			return nil, err
		}
	}
	var holidays []time.Time
	if len(parts) >= 4 {
		holidays, err = holidayList(ce, ws, parts[3])
		if err != nil {
			return nil, err
		}
	}
	return dateToSerial(workdayShift(start, int(daysN), weekend, holidays)), nil
}
