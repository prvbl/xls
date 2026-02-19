package xls

import "fmt"

type rowInfo struct {
	Index    uint16
	Fcell    uint16
	Lcell    uint16
	Height   uint16
	Notused  uint16
	Notused2 uint16
	Flags    uint32
}

//Row the data of one row
type Row struct {
	wb   *WorkBook
	info *rowInfo
	cols map[uint16]contentHandler
}

//Col Get the Nth Col from the Row, if has not, return nil.
//Suggest use Has function to test it.
func (r *Row) Col(i int) string {
	serial := uint16(i)
	if ch, ok := r.cols[serial]; ok {
		strs := ch.String(r.wb)
		if len(strs) > 0 {
			return strs[0]
		}
		return ""
	} else {
		for _, v := range r.cols {
			if v.FirstCol() <= serial && v.LastCol() >= serial {
				strs := v.String(r.wb)
				idx := int(serial - v.FirstCol())
				if idx < len(strs) {
					return strs[idx]
				}
				return ""
			}
		}
	}
	return ""
}

//ColExact Get the Nth Col from the Row, if has not, return nil.
//For merged cells value is returned for first cell only
func (r *Row) ColExact(i int) string {
	serial := uint16(i)
	if ch, ok := r.cols[serial]; ok {
		strs := ch.String(r.wb)
		if len(strs) > 0 {
			return strs[0]
		}
		return ""
	}
	return ""
}

//LastCol Get the number of Last Col of the Row.
func (r *Row) LastCol() int {
	return int(r.info.Lcell)
}

//FirstCol Get the number of First Col of the Row.
func (r *Row) FirstCol() int {
	return int(r.info.Fcell)
}

//ColFormatInfo returns the cell type name, XF index, and format number for debugging.
func (r *Row) ColFormatInfo(i int) (typeName string, xfIdx int, fmtNo uint16) {
	serial := uint16(i)
	var ch contentHandler
	if c, ok := r.cols[serial]; ok {
		ch = c
	} else {
		for _, v := range r.cols {
			if v.FirstCol() <= serial && v.LastCol() >= serial {
				ch = v
				break
			}
		}
	}
	if ch == nil {
		return "nil", -1, 0
	}
	typeName = fmt.Sprintf("%T", ch)
	switch c := ch.(type) {
	case *NumberCol:
		xfIdx = int(c.Index)
	case *RkCol:
		xfIdx = int(c.Xfrk.Index)
	case *LabelsstCol:
		xfIdx = int(c.Xf)
	case *BlankCol:
		xfIdx = int(c.Xf)
	case *labelCol:
		xfIdx = int(c.Xf)
	default:
		return typeName, -1, 0
	}
	if xfIdx >= 0 && xfIdx < len(r.wb.Xfs) {
		fmtNo = r.wb.Xfs[xfIdx].formatNo()
	}
	return
}
