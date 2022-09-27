package excel

import (
	"github.com/xuri/excelize/v2"
)

type Excel struct {
	File         *excelize.File
	ReaderConfig *ReaderConfig
}

func (e *Excel) Unmarshal(container any) error {
	// Validate excel input
	err := e.Validate()
	if err != nil {
		return err
	}

	// Create the reader
	reader, err := newReader(e.ReaderConfig, container)
	if err != nil {
		return err
	}

	// unmarshall
	err = reader.Unmarshall()
	return err
}

func (e *Excel) SetSheetName(sheet string) {
	if e.ReaderConfig != nil {
		e.ReaderConfig.SetSheetName(sheet)
	}
}

func (e *Excel) SetAxis(axis string) {
	if e.ReaderConfig != nil {
		e.ReaderConfig.SetAxis(axis)
	}
}

func (e *Excel) Validate() error {
	if e.File == nil {
		return errFileIsNil
	}
	err := e.ReaderConfig.Validate()
	return err
}
