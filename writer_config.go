package excel

import "github.com/xuri/excelize/v2"

type WriterConfig struct {
	Sheet Sheet
	Axis  Axis

	file *excelize.File
}

// NewWriterConfig create the configuration used by the writer
func NewWriterConfig(file *excelize.File) (*Excel, error) {
	if file == nil {
		return nil, errFileIsNil
	}
	w := &WriterConfig{
		file: file,
	}
	e := &Excel{
		File:         file,
		WriterConfig: w,
	}
	return e, nil
}

func (w *WriterConfig) Validate() error {
	if !w.isSheetValid() {
		return errSheetNotValid
	}
	if !w.isAxisValid() {
		return errAxisNotValid
	}
	return nil
}
