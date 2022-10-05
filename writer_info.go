package excel

import "github.com/xuri/excelize/v2"

type WriterInfo struct {
	Sheet Sheet
	Axis  Axis

	file *excelize.File
}

// NewWriter create the configuration used by the writer
func NewWriter(file *excelize.File) (*Excel, error) {
	if file == nil {
		return nil, ErrFileIsNil
	}
	w := &WriterInfo{
		file: file,
	}
	e := &Excel{
		File:       file,
		WriterInfo: w,
	}
	return e, nil
}

func (w *WriterInfo) Validate() error {
	if !w.isSheetValid() {
		return ErrSheetNotValid
	}
	if !w.isAxisValid() {
		return ErrAxisNotValid
	}
	return nil
}
