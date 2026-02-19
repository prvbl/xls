package xls

type Format struct {
	Head struct {
		Index uint16
		Size  uint16
	}
	str string
}

func (f *Format) String() string {
	return f.str
}
