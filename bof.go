package xls

import (
	"encoding/binary"
	"fmt"
	"io"
	"unicode/utf16"
)

//the information unit in xls file
type bof struct {
	Id   uint16
	Size uint16
}

//read the utf16 string from reader
func (b *bof) utf16String(buf io.ReadSeeker, count uint32) (string, error) {
	var bts = make([]uint16, count)
	if err := binary.Read(buf, binary.LittleEndian, &bts); err != nil {
		return "", fmt.Errorf("reading utf16 string: %w", err)
	}
	runes := utf16.Decode(bts[:len(bts)-1])
	return string(runes), nil
}

type biffHeader struct {
	Ver     uint16
	Type    uint16
	Id_make uint16
	Year    uint16
	Flags   uint32
	Min_ver uint32
}
