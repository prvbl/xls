package xls

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"golang.org/x/text/encoding/charmap"
	"io"
	"unicode/utf16"
)

//xls workbook type
type WorkBook struct {
	Is5ver   bool
	Type     uint16
	Codepage uint16
	Xfs      []st_xf_data
	Fonts    []Font
	Formats  map[uint16]*Format
	//All the sheets from the workbook
	sheets         []*WorkSheet
	Author         string
	rs             io.ReadSeeker
	sst            []string
	continue_utf16        uint16
	continue_rich         uint16
	continue_apsb         uint32
	continue_rich_pending uint32 // remaining richtext bytes to skip in CONTINUE
	continue_apsb_pending uint32 // remaining phonetic bytes to skip in CONTINUE
	dateMode              uint16
}

//read workbook from ole2 file
func newWorkBookFromOle2(rs io.ReadSeeker) (*WorkBook, error) {
	wb := new(WorkBook)
	wb.Formats = make(map[uint16]*Format)
	wb.rs = rs
	wb.sheets = make([]*WorkSheet, 0)
	if err := wb.Parse(rs); err != nil {
		return nil, fmt.Errorf("parsing workbook: %w", err)
	}
	return wb, nil
}

func (w *WorkBook) Parse(buf io.ReadSeeker) error {
	b := new(bof)
	bof_pre := new(bof)
	offset := 0
	for {
		if err := binary.Read(buf, binary.LittleEndian, b); err != nil {
			if err == io.EOF || err == io.ErrUnexpectedEOF {
				return nil
			}
			return fmt.Errorf("reading bof header: %w", err)
		}
		var err error
		bof_pre, b, offset, err = w.parseBof(buf, b, bof_pre, offset)
		if err != nil {
			return fmt.Errorf("parsing record 0x%X: %w", b.Id, err)
		}
	}
}

func (w *WorkBook) addXf(xf st_xf_data) {
	w.Xfs = append(w.Xfs, xf)
}

func (w *WorkBook) addFont(font *FontInfo, buf io.ReadSeeker) error {
	name, err := w.get_string(buf, uint16(font.NameB))
	if err != nil && err != io.EOF {
		return fmt.Errorf("reading font name: %w", err)
	}
	w.Fonts = append(w.Fonts, Font{Info: font, Name: name})
	return nil
}

func (w *WorkBook) addFormat(format *Format) error {
	if w.Formats == nil {
		w.Formats = make(map[uint16]*Format)
	}
	w.Formats[format.Head.Index] = format
	return nil
}

func (wb *WorkBook) parseBof(buf io.ReadSeeker, b *bof, pre *bof, offset_pre int) (after *bof, after_using *bof, offset int, err error) {
	after = b
	after_using = pre
	offset = offset_pre
	var bts = make([]byte, b.Size)
	if err = binary.Read(buf, binary.LittleEndian, bts); err != nil {
		if err == io.EOF || err == io.ErrUnexpectedEOF {
			err = nil
			return
		}
		err = fmt.Errorf("reading record body: %w", err)
		return
	}
	buf_item := bytes.NewReader(bts)
	switch b.Id {
	case 0x809:
		bif := new(biffHeader)
		if err = binary.Read(buf_item, binary.LittleEndian, bif); err != nil {
			err = fmt.Errorf("reading BIFF header: %w", err)
			return
		}
		if bif.Ver != 0x600 {
			wb.Is5ver = true
		}
		wb.Type = bif.Type
	case 0x042: // CODEPAGE
		if err = binary.Read(buf_item, binary.LittleEndian, &wb.Codepage); err != nil {
			err = fmt.Errorf("reading codepage: %w", err)
			return
		}
	case 0x3c: // CONTINUE
		if pre.Id == 0xfc {
			// Skip pending richtext/phonetic formatting data that overflowed
			// from the previous record when characters were already complete.
			if wb.continue_utf16 == 0 && (wb.continue_rich_pending > 0 || wb.continue_apsb_pending > 0) {
				if wb.continue_rich_pending > 0 {
					available := int64(buf_item.Len())
					skip := int64(wb.continue_rich_pending)
					if skip <= available {
						if _, seekErr := buf_item.Seek(skip, io.SeekCurrent); seekErr != nil {
							err = fmt.Errorf("seeking past richtext in CONTINUE: %w", seekErr)
							return
						}
						wb.continue_rich_pending = 0
					} else {
						if _, seekErr := buf_item.Seek(0, io.SeekEnd); seekErr != nil {
							err = fmt.Errorf("seeking to end in CONTINUE richtext: %w", seekErr)
							return
						}
						wb.continue_rich_pending -= uint32(available)
					}
				}
				if wb.continue_apsb_pending > 0 {
					available := int64(buf_item.Len())
					skip := int64(wb.continue_apsb_pending)
					if skip <= available {
						if _, seekErr := buf_item.Seek(skip, io.SeekCurrent); seekErr != nil {
							err = fmt.Errorf("seeking past phonetic in CONTINUE: %w", seekErr)
							return
						}
						wb.continue_apsb_pending = 0
					} else {
						if _, seekErr := buf_item.Seek(0, io.SeekEnd); seekErr != nil {
							err = fmt.Errorf("seeking to end in CONTINUE phonetic: %w", seekErr)
							return
						}
						wb.continue_apsb_pending -= uint32(available)
					}
				}
				if wb.continue_rich_pending == 0 && wb.continue_apsb_pending == 0 {
					wb.continue_rich = 0
					wb.continue_apsb = 0
					offset_pre++
				}
			}

			var size uint16
			var readErr error
			if wb.continue_utf16 >= 1 {
				size = wb.continue_utf16
				wb.continue_utf16 = 0
			} else {
				readErr = binary.Read(buf_item, binary.LittleEndian, &size)
			}
			for readErr == nil && offset_pre < len(wb.sst) {
				var str string
				str, readErr = wb.get_string(buf_item, size)
				wb.sst[offset_pre] = wb.sst[offset_pre] + str

				if readErr == io.EOF {
					break
				}

				offset_pre++
				readErr = binary.Read(buf_item, binary.LittleEndian, &size)
			}
		}
		offset = offset_pre
		after = pre
		after_using = b
	case 0xfc: // SST
		info := new(SstInfo)
		if err = binary.Read(buf_item, binary.LittleEndian, info); err != nil {
			err = fmt.Errorf("reading SST info: %w", err)
			return
		}
		wb.sst = make([]string, info.Count)
		var size uint16
		var i = 0
		// dont forget to initialize offset
		offset = 0
		for ; i < int(info.Count); i++ {
			var readErr error
			readErr = binary.Read(buf_item, binary.LittleEndian, &size)
			if readErr == nil {
				var str string
				str, readErr = wb.get_string(buf_item, size)
				wb.sst[i] = wb.sst[i] + str
			}

			if readErr == io.EOF {
				break
			}
		}
		offset = i
	case 0x85: // boundsheet
		var bs = new(boundsheet)
		if err = binary.Read(buf_item, binary.LittleEndian, bs); err != nil {
			err = fmt.Errorf("reading boundsheet: %w", err)
			return
		}
		// different for BIFF5 and BIFF8
		if err = wb.addSheet(bs, buf_item); err != nil {
			return
		}
	case 0x0e0: // XF
		if wb.Is5ver {
			xf := new(Xf5)
			if err = binary.Read(buf_item, binary.LittleEndian, xf); err != nil {
				err = fmt.Errorf("reading XF5 record: %w", err)
				return
			}
			wb.addXf(xf)
		} else {
			xf := new(Xf8)
			if err = binary.Read(buf_item, binary.LittleEndian, xf); err != nil {
				err = fmt.Errorf("reading XF8 record: %w", err)
				return
			}
			wb.addXf(xf)
		}
	case 0x031: // FONT
		f := new(FontInfo)
		if err = binary.Read(buf_item, binary.LittleEndian, f); err != nil {
			err = fmt.Errorf("reading font info: %w", err)
			return
		}
		if err = wb.addFont(f, buf_item); err != nil {
			return
		}
	case 0x41E: //FORMAT
		font := new(Format)
		if err = binary.Read(buf_item, binary.LittleEndian, &font.Head); err != nil {
			err = fmt.Errorf("reading format header: %w", err)
			return
		}
		var fmtErr error
		font.str, fmtErr = wb.get_string(buf_item, font.Head.Size)
		if fmtErr != nil && fmtErr != io.EOF {
			err = fmt.Errorf("reading format string: %w", fmtErr)
			return
		}
		if err = wb.addFormat(font); err != nil {
			return
		}
	case 0x22: //DATEMODE
		if err = binary.Read(buf_item, binary.LittleEndian, &wb.dateMode); err != nil {
			err = fmt.Errorf("reading date mode: %w", err)
			return
		}
	}
	return
}
func decodeWindows1251(enc []byte) string {
	dec := charmap.Windows1251.NewDecoder()
	out, err := dec.Bytes(enc)
	if err != nil {
		return string(enc)
	}
	return string(out)
}
func (w *WorkBook) get_string(buf io.ReadSeeker, size uint16) (res string, err error) {
	if w.Is5ver {
		var bts = make([]byte, size)
		_, err = buf.Read(bts)
		res = decodeWindows1251(bts)
		//res = string(bts)
	} else {
		var richtext_num = uint16(0)
		var phonetic_size = uint32(0)
		var flag byte
		err = binary.Read(buf, binary.LittleEndian, &flag)
		if err != nil {
			return
		}
		if flag&0x8 != 0 {
			err = binary.Read(buf, binary.LittleEndian, &richtext_num)
			if err != nil {
				return
			}
		} else if w.continue_rich > 0 {
			richtext_num = w.continue_rich
			w.continue_rich = 0
		}
		if flag&0x4 != 0 {
			err = binary.Read(buf, binary.LittleEndian, &phonetic_size)
			if err != nil {
				return
			}
		} else if w.continue_apsb > 0 {
			phonetic_size = w.continue_apsb
			w.continue_apsb = 0
		}
		if flag&0x1 != 0 {
			var bts = make([]uint16, size)
			var i = uint16(0)
			for ; i < size && err == nil; i++ {
				err = binary.Read(buf, binary.LittleEndian, &bts[i])
			}

			// when eof found, we dont want to append last element
			var runes []rune
			if err == io.EOF {
				i = i - 1
			}
			runes = utf16.Decode(bts[:i])

			res = string(runes)
			if i < size {
				w.continue_utf16 = size - i
			}

		} else {
			var bts = make([]byte, size)
			var n int
			n, err = buf.Read(bts)
			if uint16(n) < size {
				w.continue_utf16 = size - uint16(n)
				err = io.EOF
			}

			var bts1 = make([]uint16, n)
			for k, v := range bts[:n] {
				bts1[k] = uint16(v)
			}
			runes := utf16.Decode(bts1)
			res = string(runes)
		}
		if richtext_num > 0 {
			var seek_size int64
			if w.Is5ver {
				seek_size = int64(2 * richtext_num)
			} else {
				seek_size = int64(4 * richtext_num)
			}
			bts := make([]byte, seek_size)
			n, readErr := io.ReadFull(buf, bts)
			if readErr != nil {
				w.continue_rich = richtext_num
				w.continue_rich_pending = uint32(seek_size) - uint32(n)
				err = io.EOF
			}
		}
		if phonetic_size > 0 {
			bts := make([]byte, phonetic_size)
			n, readErr := io.ReadFull(buf, bts)
			if readErr != nil {
				w.continue_apsb = phonetic_size
				w.continue_apsb_pending = phonetic_size - uint32(n)
				err = io.EOF
			}
		}
	}
	return
}

func (w *WorkBook) addSheet(sheet *boundsheet, buf io.ReadSeeker) error {
	name, err := w.get_string(buf, uint16(sheet.Name))
	if err != nil && err != io.EOF {
		return fmt.Errorf("reading sheet name: %w", err)
	}
	w.sheets = append(w.sheets, &WorkSheet{bs: sheet, Name: name, wb: w, Visibility: TWorkSheetVisibility(sheet.Visible)})
	return nil
}

//reading a sheet from the compress file to memory, you should call this before you try to get anything from sheet
func (w *WorkBook) prepareSheet(sheet *WorkSheet) error {
	if _, err := w.rs.Seek(int64(sheet.bs.Filepos), 0); err != nil {
		return fmt.Errorf("seeking to sheet position: %w", err)
	}
	if err := sheet.parse(w.rs); err != nil {
		return fmt.Errorf("parsing sheet %q: %w", sheet.Name, err)
	}
	return nil
}

//Get one sheet by its number
func (w *WorkBook) GetSheet(num int) (*WorkSheet, error) {
	if num < len(w.sheets) {
		s := w.sheets[num]
		if !s.parsed {
			if err := w.prepareSheet(s); err != nil {
				return nil, err
			}
		}
		return s, nil
	}
	return nil, nil
}

//Get the number of all sheets, look into example
func (w *WorkBook) NumSheets() int {
	return len(w.sheets)
}

//helper function to read all cells from file
//Notice: the max value is the limit of the max capacity of lines.
//Warning: the helper function will need big memeory if file is large.
func (w *WorkBook) ReadAllCells(max int) ([][]string, error) {
	res := make([][]string, 0)
	for _, sheet := range w.sheets {
		if len(res) < max {
			max = max - len(res)
			if err := w.prepareSheet(sheet); err != nil {
				return res, err
			}
			if sheet.MaxRow != 0 {
				leng := int(sheet.MaxRow) + 1
				if max < leng {
					leng = max
				}
				temp := make([][]string, leng)
				for k, row := range sheet.rows {
					if row == nil {
						continue
					}
					data := make([]string, 0)
					if len(row.cols) > 0 {
						for _, col := range row.cols {
							if uint16(len(data)) <= col.LastCol() {
								data = append(data, make([]string, col.LastCol()-uint16(len(data))+1)...)
							}
							str := col.String(w)

							for i := uint16(0); i < col.LastCol()-col.FirstCol()+1; i++ {
								if int(i) < len(str) {
									data[col.FirstCol()+i] = str[i]
								}
							}
						}
						if leng > int(k) {
							temp[k] = data
						}
					}
				}
				res = append(res, temp...)
			}
		}
	}
	return res, nil
}
