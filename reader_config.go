package excel

import "github.com/xuri/excelize/v2"

type ReaderConfig struct {
	Sheet Sheet
	Axis  Axis

	file *excelize.File
}

// NewReaderConfig create the configuration used by the reader
func NewReaderConfig(file *excelize.File) (*Excel, error) {
	if file == nil {
		return nil, errFileIsNil
	}
	r := &ReaderConfig{
		file: file,
	}
	e := &Excel{
		File:         file,
		ReaderConfig: r,
	}
	return e, nil
}

func (r *ReaderConfig) Validate() error {
	if !r.isSheetValid() {
		return errSheetNotValid
	}
	if !r.isAxisValid() {
		return errAxisNotValid
	}
	return nil
}
