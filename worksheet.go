package xls

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"io"
	"unicode/utf16"
)

type TWorkSheetVisibility byte

const (
	WorkSheetVisible    TWorkSheetVisibility = 0
	WorkSheetHidden     TWorkSheetVisibility = 1
	WorkSheetVeryHidden TWorkSheetVisibility = 2
)

type boundsheet struct {
	Filepos uint32
	Visible byte
	Type    byte
	Name    byte
}

//WorkSheet in one WorkBook
type WorkSheet struct {
	bs         *boundsheet
	wb         *WorkBook
	Name       string
	Selected   bool
	Visibility TWorkSheetVisibility
	rows       map[uint16]*Row
	//NOTICE: this is the max row number of the sheet, so it should be count -1
	MaxRow      uint16
	parsed      bool
	rightToLeft bool
}

func (w *WorkSheet) Row(i int) *Row {
	row := w.rows[uint16(i)]
	if row != nil {
		row.wb = w.wb
	}
	return row
}

func (w *WorkSheet) parse(buf io.ReadSeeker) error {
	w.rows = make(map[uint16]*Row)
	b := new(bof)
	var bof_pre *bof
	var col_pre interface{}
	for {
		if err := binary.Read(buf, binary.LittleEndian, b); err != nil {
			if err == io.EOF || err == io.ErrUnexpectedEOF {
				break
			}
			return fmt.Errorf("reading worksheet bof header: %w", err)
		}
		var err error
		bof_pre, col_pre, err = w.parseBof(buf, b, bof_pre, col_pre)
		if err != nil {
			return fmt.Errorf("parsing worksheet record 0x%X: %w", b.Id, err)
		}
		if b.Id == 0xa {
			break
		}
	}
	w.parsed = true
	return nil
}

func (w *WorkSheet) parseBof(buf io.ReadSeeker, b *bof, pre *bof, col_pre interface{}) (*bof, interface{}, error) {
	var col interface{}
	var bts = make([]byte, b.Size)
	if err := binary.Read(buf, binary.LittleEndian, bts); err != nil {
		return b, col, fmt.Errorf("reading record body: %w", err)
	}
	buf = bytes.NewReader(bts)
	switch b.Id {
	// case 0x0E5: //MERGEDCELLS
	// ws.mergedCells(buf)
	case 0x23E: // WINDOW2
		var sheetOptions, firstVisibleRow, firstVisibleColumn uint16
		if err := binary.Read(buf, binary.LittleEndian, &sheetOptions); err != nil {
			return b, col, fmt.Errorf("reading WINDOW2 sheetOptions: %w", err)
		}
		if err := binary.Read(buf, binary.LittleEndian, &firstVisibleRow); err != nil {
			return b, col, fmt.Errorf("reading WINDOW2 firstVisibleRow: %w", err)
		}
		if err := binary.Read(buf, binary.LittleEndian, &firstVisibleColumn); err != nil {
			return b, col, fmt.Errorf("reading WINDOW2 firstVisibleColumn: %w", err)
		}
		//buf.Seek(int64(b.Size)-2*3, 1)
		w.rightToLeft = (sheetOptions & 0x40) != 0
		w.Selected = (sheetOptions & 0x400) != 0
	case 0x208: //ROW
		r := new(rowInfo)
		if err := binary.Read(buf, binary.LittleEndian, r); err != nil {
			return b, col, fmt.Errorf("reading ROW record: %w", err)
		}
		w.addRow(r)
	case 0x0BD: //MULRK
		mc := new(MulrkCol)
		size := (b.Size - 6) / 6
		if err := binary.Read(buf, binary.LittleEndian, &mc.Col); err != nil {
			return b, col, fmt.Errorf("reading MULRK col: %w", err)
		}
		mc.Xfrks = make([]XfRk, size)
		for i := uint16(0); i < size; i++ {
			if err := binary.Read(buf, binary.LittleEndian, &mc.Xfrks[i]); err != nil {
				return b, col, fmt.Errorf("reading MULRK xfrk %d: %w", i, err)
			}
		}
		if err := binary.Read(buf, binary.LittleEndian, &mc.LastColB); err != nil {
			return b, col, fmt.Errorf("reading MULRK lastcol: %w", err)
		}
		col = mc
	case 0x0BE: //MULBLANK
		mc := new(MulBlankCol)
		size := (b.Size - 6) / 2
		if err := binary.Read(buf, binary.LittleEndian, &mc.Col); err != nil {
			return b, col, fmt.Errorf("reading MULBLANK col: %w", err)
		}
		mc.Xfs = make([]uint16, size)
		for i := uint16(0); i < size; i++ {
			if err := binary.Read(buf, binary.LittleEndian, &mc.Xfs[i]); err != nil {
				return b, col, fmt.Errorf("reading MULBLANK xf %d: %w", i, err)
			}
		}
		if err := binary.Read(buf, binary.LittleEndian, &mc.LastColB); err != nil {
			return b, col, fmt.Errorf("reading MULBLANK lastcol: %w", err)
		}
		col = mc
	case 0x203: //NUMBER
		col = new(NumberCol)
		if err := binary.Read(buf, binary.LittleEndian, col); err != nil {
			return b, nil, fmt.Errorf("reading NUMBER record: %w", err)
		}
	case 0x06: //FORMULA
		c := new(FormulaCol)
		if err := binary.Read(buf, binary.LittleEndian, &c.Header); err != nil {
			return b, col, fmt.Errorf("reading FORMULA header: %w", err)
		}
		c.Bts = make([]byte, b.Size-20)
		if err := binary.Read(buf, binary.LittleEndian, &c.Bts); err != nil {
			return b, col, fmt.Errorf("reading FORMULA body: %w", err)
		}
		col = c
	case 0x207: //STRING = FORMULA-VALUE is expected right after FORMULA
		if ch, ok := col_pre.(*FormulaCol); ok {
			c := new(FormulaStringCol)
			c.Col = ch.Header.Col
			var cStringLen uint16
			if err := binary.Read(buf, binary.LittleEndian, &cStringLen); err != nil {
				return b, col, fmt.Errorf("reading STRING length: %w", err)
			}
			str, err := w.wb.get_string(buf, cStringLen)
			if err != nil && err != io.EOF {
				return b, col, fmt.Errorf("reading STRING value: %w", err)
			}
			c.RenderedValue = str
			col = c
		}
	case 0x27e: //RK
		col = new(RkCol)
		if err := binary.Read(buf, binary.LittleEndian, col); err != nil {
			return b, nil, fmt.Errorf("reading RK record: %w", err)
		}
	case 0xFD: //LABELSST
		col = new(LabelsstCol)
		if err := binary.Read(buf, binary.LittleEndian, col); err != nil {
			return b, nil, fmt.Errorf("reading LABELSST record: %w", err)
		}
	case 0x204:
		c := new(labelCol)
		if err := binary.Read(buf, binary.LittleEndian, &c.BlankCol); err != nil {
			return b, col, fmt.Errorf("reading LABEL blank: %w", err)
		}
		var count uint16
		if err := binary.Read(buf, binary.LittleEndian, &count); err != nil {
			return b, col, fmt.Errorf("reading LABEL count: %w", err)
		}
		var strErr error
		c.Str, strErr = w.wb.get_string(buf, count)
		if strErr != nil && strErr != io.EOF {
			return b, col, fmt.Errorf("reading LABEL string: %w", strErr)
		}
		col = c
	case 0x201: //BLANK
		col = new(BlankCol)
		if err := binary.Read(buf, binary.LittleEndian, col); err != nil {
			return b, nil, fmt.Errorf("reading BLANK record: %w", err)
		}
	case 0x1b8: //HYPERLINK
		var hy HyperLink
		if err := binary.Read(buf, binary.LittleEndian, &hy.CellRange); err != nil {
			return b, col, fmt.Errorf("reading HYPERLINK cell range: %w", err)
		}
		if _, err := buf.Seek(20, 1); err != nil {
			return b, col, fmt.Errorf("seeking in HYPERLINK: %w", err)
		}
		var flag uint32
		if err := binary.Read(buf, binary.LittleEndian, &flag); err != nil {
			return b, col, fmt.Errorf("reading HYPERLINK flag: %w", err)
		}
		var count uint32

		if flag&0x14 != 0 {
			if err := binary.Read(buf, binary.LittleEndian, &count); err != nil {
				return b, col, fmt.Errorf("reading HYPERLINK description count: %w", err)
			}
			var err error
			hy.Description, err = b.utf16String(buf, count)
			if err != nil {
				return b, col, fmt.Errorf("reading HYPERLINK description: %w", err)
			}
		}
		if flag&0x80 != 0 {
			if err := binary.Read(buf, binary.LittleEndian, &count); err != nil {
				return b, col, fmt.Errorf("reading HYPERLINK target frame count: %w", err)
			}
			var err error
			hy.TargetFrame, err = b.utf16String(buf, count)
			if err != nil {
				return b, col, fmt.Errorf("reading HYPERLINK target frame: %w", err)
			}
		}
		if flag&0x1 != 0 {
			var guid [2]uint64
			if err := binary.Read(buf, binary.BigEndian, &guid); err != nil {
				return b, col, fmt.Errorf("reading HYPERLINK GUID: %w", err)
			}
			if guid[0] == 0xE0C9EA79F9BACE11 && guid[1] == 0x8C8200AA004BA90B { //URL
				hy.IsUrl = true
				if err := binary.Read(buf, binary.LittleEndian, &count); err != nil {
					return b, col, fmt.Errorf("reading HYPERLINK URL count: %w", err)
				}
				var err error
				hy.Url, err = b.utf16String(buf, count/2)
				if err != nil {
					return b, col, fmt.Errorf("reading HYPERLINK URL: %w", err)
				}
			} else if guid[0] == 0x303000000000000 && guid[1] == 0xC000000000000046 { //URL{
				var upCount uint16
				if err := binary.Read(buf, binary.LittleEndian, &upCount); err != nil {
					return b, col, fmt.Errorf("reading HYPERLINK upCount: %w", err)
				}
				if err := binary.Read(buf, binary.LittleEndian, &count); err != nil {
					return b, col, fmt.Errorf("reading HYPERLINK file path count: %w", err)
				}
				bts := make([]byte, count)
				if err := binary.Read(buf, binary.LittleEndian, &bts); err != nil {
					return b, col, fmt.Errorf("reading HYPERLINK file path: %w", err)
				}
				hy.ShortedFilePath = string(bts)
				if _, err := buf.Seek(24, 1); err != nil {
					return b, col, fmt.Errorf("seeking in HYPERLINK file path: %w", err)
				}
				if err := binary.Read(buf, binary.LittleEndian, &count); err != nil {
					return b, col, fmt.Errorf("reading HYPERLINK extended count: %w", err)
				}
				if count > 0 {
					if err := binary.Read(buf, binary.LittleEndian, &count); err != nil {
						return b, col, fmt.Errorf("reading HYPERLINK extended inner count: %w", err)
					}
					if _, err := buf.Seek(2, 1); err != nil {
						return b, col, fmt.Errorf("seeking in HYPERLINK extended: %w", err)
					}
					var err error
					hy.ExtendedFilePath, err = b.utf16String(buf, count/2+1)
					if err != nil {
						return b, col, fmt.Errorf("reading HYPERLINK extended file path: %w", err)
					}
				}
			}
		}
		if flag&0x8 != 0 {
			if err := binary.Read(buf, binary.LittleEndian, &count); err != nil {
				return b, col, fmt.Errorf("reading HYPERLINK text mark count: %w", err)
			}
			var bts = make([]uint16, count)
			if err := binary.Read(buf, binary.LittleEndian, &bts); err != nil {
				return b, col, fmt.Errorf("reading HYPERLINK text mark: %w", err)
			}
			runes := utf16.Decode(bts[:len(bts)-1])
			hy.TextMark = string(runes)
		}

		w.addRange(&hy.CellRange, &hy)
	case 0x809:
		// already read into bts above
	case 0xa:
	default:
		// log.Printf("Unknow %X,%d\n", b.Id, b.Size)
	}
	if col != nil {
		w.add(col)
	}
	return b, col, nil
}

func (w *WorkSheet) add(content interface{}) {
	if ch, ok := content.(contentHandler); ok {
		if col, ok := content.(Coler); ok {
			w.addCell(col, ch)
		}
	}

}

func (w *WorkSheet) addCell(col Coler, ch contentHandler) {
	w.addContent(col.Row(), ch)
}

func (w *WorkSheet) addRange(rang Ranger, ch contentHandler) {

	for i := rang.FirstRow(); i <= rang.LastRow(); i++ {
		w.addContent(i, ch)
	}
}

func (w *WorkSheet) addContent(row_num uint16, ch contentHandler) {
	var row *Row
	var ok bool
	if row, ok = w.rows[row_num]; !ok {
		info := new(rowInfo)
		info.Index = row_num
		row = w.addRow(info)
	}
	if row.info.Lcell < ch.LastCol() {
		row.info.Lcell = ch.LastCol()
	}
	row.cols[ch.FirstCol()] = ch
}

func (w *WorkSheet) addRow(info *rowInfo) (row *Row) {
	if info.Index > w.MaxRow {
		w.MaxRow = info.Index
	}
	var ok bool
	if row, ok = w.rows[info.Index]; ok {
		row.info = info
	} else {
		row = &Row{info: info, cols: make(map[uint16]contentHandler)}
		w.rows[info.Index] = row
	}
	return
}
