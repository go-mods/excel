package excel

import "github.com/xuri/excelize/v2"

type ReaderInfo struct {
	Sheet Sheet
	Axis  Axis

	file *excelize.File
}

// NewReader create the configuration used by the reader
func NewReader(file *excelize.File) (*Excel, error) {
	if file == nil {
		return nil, ErrFileIsNil
	}
	r := &ReaderInfo{
		file: file,
	}
	e := &Excel{
		File:       file,
		ReaderInfo: r,
	}
	return e, nil
}

func (r *ReaderInfo) Validate() error {
	if !r.isSheetValid() {
		return ErrSheetNotValid
	}
	if !r.isAxisValid() {
		return ErrAxisNotValid
	}
	return nil
}
