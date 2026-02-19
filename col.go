package xls

import (
	"fmt"
	"math"
	"strconv"
	"strings"
	"time"
)

//content type
type contentHandler interface {
	String(*WorkBook) []string
	FirstCol() uint16
	LastCol() uint16
}

type Col struct {
	RowB      uint16
	FirstColB uint16
}

type Coler interface {
	Row() uint16
}

func (c *Col) Row() uint16 {
	return c.RowB
}

func (c *Col) FirstCol() uint16 {
	return c.FirstColB
}

func (c *Col) LastCol() uint16 {
	return c.FirstColB
}

func (c *Col) String(wb *WorkBook) []string {
	return []string{"default"}
}

// isDateFormat returns true if the format string looks like a date/time format.
// Checks for unambiguous date tokens (yy, dd, mmm) before number tokens to avoid
// misclassifying formats like "mm/dd/yyyy" as numeric due to containing "d/y".
func isDateFormat(fmtLower string) bool {
	return strings.Contains(fmtLower, "yy") ||
		strings.Contains(fmtLower, "dd") ||
		strings.Contains(fmtLower, "mmm")
}

// isNumberFormat returns true if the format string looks like a number format.
func isNumberFormat(fmtStr, fmtLower string) bool {
	return fmtLower == "general" ||
		strings.Contains(fmtStr, "#") ||
		strings.Contains(fmtStr, ".00") ||
		strings.Contains(fmtLower, "h:") ||
		strings.Contains(fmtLower, "д.г")
}

// isBuiltinDateFmt returns true if the built-in format number is a date/time format.
// See http://www.openoffice.org/sc/excelfileformat.pdf Page #174
func isBuiltinDateFmt(fNo uint16) bool {
	return 14 <= fNo && fNo <= 22 || 27 <= fNo && fNo <= 36 || 45 <= fNo && fNo <= 58
}

// countCI counts a case-insensitive run of ch starting at position start in s.
func countCI(s string, start int, ch byte) int {
	lower := ch | 0x20
	n := 0
	for i := start; i < len(s); i++ {
		if s[i]|0x20 == lower {
			n++
		} else {
			break
		}
	}
	return n
}

// excelDateFmtToGo converts an Excel date/time format string to a Go time.Format layout.
// Handles the context-dependent m/mm token (month after d/y, minutes after h).
func excelDateFmtToGo(excel string) string {
	// Use first section only (positive values)
	if idx := strings.IndexByte(excel, ';'); idx >= 0 {
		excel = excel[:idx]
	}

	upper := strings.ToUpper(excel)
	is12h := strings.Contains(upper, "AM/PM") || strings.Contains(upper, "A/P")

	var result strings.Builder
	i := 0
	afterH := false

	for i < len(excel) {
		c := excel[i]
		cl := c | 0x20

		switch {
		case cl == 'y':
			n := countCI(excel, i, 'y')
			if n >= 3 {
				result.WriteString("2006")
			} else {
				result.WriteString("06")
			}
			afterH = false
			i += n

		case cl == 'm':
			n := countCI(excel, i, 'm')
			if afterH {
				if n >= 2 {
					result.WriteString("04")
				} else {
					result.WriteString("4")
				}
			} else {
				switch {
				case n >= 4:
					result.WriteString("January")
				case n == 3:
					result.WriteString("Jan")
				case n == 2:
					result.WriteString("01")
				default:
					result.WriteString("1")
				}
			}
			i += n

		case cl == 'd':
			n := countCI(excel, i, 'd')
			switch {
			case n >= 4:
				result.WriteString("Monday")
			case n == 3:
				result.WriteString("Mon")
			case n == 2:
				result.WriteString("02")
			default:
				result.WriteString("2")
			}
			afterH = false
			i += n

		case cl == 'h':
			n := countCI(excel, i, 'h')
			if is12h {
				if n >= 2 {
					result.WriteString("03")
				} else {
					result.WriteString("3")
				}
			} else {
				result.WriteString("15")
			}
			afterH = true
			i += n

		case cl == 's':
			n := countCI(excel, i, 's')
			if n >= 2 {
				result.WriteString("05")
			} else {
				result.WriteString("5")
			}
			afterH = false
			i += n

		case cl == 'a':
			rem := strings.ToUpper(excel[i:])
			if strings.HasPrefix(rem, "AM/PM") {
				result.WriteString("PM")
				i += 5
			} else if strings.HasPrefix(rem, "A/P") {
				result.WriteString("PM")
				i += 3
			} else {
				result.WriteByte(c)
				i++
			}

		case c == '\\':
			if i+1 < len(excel) {
				result.WriteByte(excel[i+1])
				i += 2
			} else {
				i++
			}

		case c == '"':
			i++
			for i < len(excel) && excel[i] != '"' {
				result.WriteByte(excel[i])
				i++
			}
			if i < len(excel) {
				i++
			}

		case c == '[':
			end := strings.IndexByte(excel[i:], ']')
			if end >= 0 {
				i += end + 1
			} else {
				i++
			}

		default:
			result.WriteByte(c)
			i++
		}
	}
	return result.String()
}

// formatBuiltinNumber formats a float according to a built-in number format (1-13).
func formatBuiltinNumber(f float64, fNo uint16) string {
	switch fNo {
	case 1: // 0
		return strconv.FormatFloat(f, 'f', 0, 64)
	case 2: // 0.00
		return strconv.FormatFloat(f, 'f', 2, 64)
	case 3: // #,##0
		return strconv.FormatFloat(math.Round(f), 'f', 0, 64)
	case 4: // #,##0.00
		return strconv.FormatFloat(f, 'f', 2, 64)
	case 9: // 0%
		return strconv.FormatFloat(f*100, 'f', 0, 64) + "%"
	case 10: // 0.00%
		return strconv.FormatFloat(f*100, 'f', 2, 64) + "%"
	case 11: // 0.00E+00
		return strconv.FormatFloat(f, 'E', 2, 64)
	default:
		return strconv.FormatFloat(f, 'f', -1, 64)
	}
}

type XfRk struct {
	Index uint16
	Rk    RK
}

func (xf *XfRk) String(wb *WorkBook) string {
	idx := int(xf.Index)
	if len(wb.Xfs) > idx {
		fNo := wb.Xfs[idx].formatNo()
		if fNo >= 164 { // user defined format
			if formatter := wb.Formats[fNo]; formatter != nil {
				fmtLower := strings.ToLower(formatter.str)
				if isDateFormat(fmtLower) {
					i, f, isFloat := xf.Rk.number()
					if !isFloat {
						f = float64(i)
					}
					t := timeFromExcelTime(f, wb.dateMode == 1)
					return t.Format(excelDateFmtToGo(formatter.str))
				}
				if isNumberFormat(formatter.str, fmtLower) {
					return xf.Rk.String()
				}
				// unknown user format — try as date
				i, f, isFloat := xf.Rk.number()
				if !isFloat {
					f = float64(i)
				}
				t := timeFromExcelTime(f, wb.dateMode == 1)
				return t.Format(excelDateFmtToGo(formatter.str))
			}
		} else if isBuiltinDateFmt(fNo) {
			i, f, isFloat := xf.Rk.number()
			if !isFloat {
				f = float64(i)
			}
			t := timeFromExcelTime(f, wb.dateMode == 1)
			return t.Format(time.RFC3339)
		} else if fNo >= 1 && fNo <= 13 {
			i, f, isFloat := xf.Rk.number()
			if !isFloat {
				f = float64(i)
			}
			return formatBuiltinNumber(f, fNo)
		}
	}
	return xf.Rk.String()
}

type RK uint32

func (rk RK) number() (intNum int64, floatNum float64, isFloat bool) {
	multiplied := rk & 1
	isInt := rk & 2
	val := int32(rk) >> 2
	if isInt == 0 {
		isFloat = true
		floatNum = math.Float64frombits(uint64(val) << 34)
		if multiplied != 0 {
			floatNum = floatNum / 100
		}
		return
	}
	if multiplied != 0 {
		isFloat = true
		floatNum = float64(val) / 100
		return
	}
	return int64(val), 0, false
}

func (rk RK) String() string {
	i, f, isFloat := rk.number()
	if isFloat {
		return strconv.FormatFloat(f, 'f', -1, 64)
	}
	return strconv.FormatInt(i, 10)
}

var ErrIsInt = fmt.Errorf("is int")

func (rk RK) Float() (float64, error) {
	_, f, isFloat := rk.number()
	if !isFloat {
		return 0, ErrIsInt
	}
	return f, nil
}

type MulrkCol struct {
	Col
	Xfrks    []XfRk
	LastColB uint16
}

func (c *MulrkCol) LastCol() uint16 {
	return c.LastColB
}

func (c *MulrkCol) String(wb *WorkBook) []string {
	var res = make([]string, len(c.Xfrks))
	for i := 0; i < len(c.Xfrks); i++ {
		xfrk := c.Xfrks[i]
		res[i] = xfrk.String(wb)
	}
	return res
}

type MulBlankCol struct {
	Col
	Xfs      []uint16
	LastColB uint16
}

func (c *MulBlankCol) LastCol() uint16 {
	return c.LastColB
}

func (c *MulBlankCol) String(wb *WorkBook) []string {
	return make([]string, len(c.Xfs))
}

type NumberCol struct {
	Col
	Index uint16
	Float float64
}

func (c *NumberCol) String(wb *WorkBook) []string {
	idx := int(c.Index)
	if idx < len(wb.Xfs) {
		fNo := wb.Xfs[idx].formatNo()
		if fNo >= 164 { // user defined format
			if formatter := wb.Formats[fNo]; formatter != nil {
				fmtLower := strings.ToLower(formatter.str)
				if isDateFormat(fmtLower) {
					t := timeFromExcelTime(c.Float, wb.dateMode == 1)
					return []string{t.Format(excelDateFmtToGo(formatter.str))}
				}
				if isNumberFormat(formatter.str, fmtLower) {
					return []string{strconv.FormatFloat(c.Float, 'f', -1, 64)}
				}
				// unknown user format — try as date
				t := timeFromExcelTime(c.Float, wb.dateMode == 1)
				return []string{t.Format(excelDateFmtToGo(formatter.str))}
			}
		} else if isBuiltinDateFmt(fNo) {
			t := timeFromExcelTime(c.Float, wb.dateMode == 1)
			return []string{t.Format(time.RFC3339)}
		} else if fNo >= 1 && fNo <= 13 {
			return []string{formatBuiltinNumber(c.Float, fNo)}
		}
	}
	return []string{strconv.FormatFloat(c.Float, 'f', -1, 64)}
}

type FormulaStringCol struct {
	Col
	RenderedValue string
}

func (c *FormulaStringCol) String(wb *WorkBook) []string {
	return []string{c.RenderedValue}
}

//str, err = wb.get_string(buf_item, size)
//wb.sst[offset_pre] = wb.sst[offset_pre] + str

type FormulaCol struct {
	Header struct {
		Col
		IndexXf uint16
		Result  [8]byte
		Flags   uint16
		_       uint32
	}
	Bts []byte
}

func (c *FormulaCol) String(wb *WorkBook) []string {
	return []string{"FormulaCol"}
}

type RkCol struct {
	Col
	Xfrk XfRk
}

func (c *RkCol) String(wb *WorkBook) []string {
	return []string{c.Xfrk.String(wb)}
}

type LabelsstCol struct {
	Col
	Xf  uint16
	Sst uint32
}

func (c *LabelsstCol) String(wb *WorkBook) []string {
	idx := int(c.Sst)
	if idx < len(wb.sst) {
		return []string{wb.sst[idx]}
	}
	return []string{""}
}

type labelCol struct {
	BlankCol
	Str string
}

func (c *labelCol) String(wb *WorkBook) []string {
	return []string{c.Str}
}

type BlankCol struct {
	Col
	Xf uint16
}

func (c *BlankCol) String(wb *WorkBook) []string {
	return []string{""}
}
