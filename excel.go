package excel

import (
	"github.com/xuri/excelize/v2"
)

type Excel struct {
	File         *excelize.File
	ReaderConfig *ReaderConfig
	WriterConfig *WriterConfig
}

func (e *Excel) Marshal(container any) error {
	// Validate excel input
	err := e.Validate()
	if err != nil {
		return err
	}

	// Create the writer
	writer, err := newWriter(e.WriterConfig, container)
	if err != nil {
		return err
	}

	// unmarshall
	err = writer.Marshall(container)
	return err
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
	if e.WriterConfig != nil {
		e.WriterConfig.SetSheetName(sheet)
	}
}

func (e *Excel) SetSheetIndex(index int) {
	if e.ReaderConfig != nil {
		e.ReaderConfig.SetSheetIndex(index)
	}
	if e.WriterConfig != nil {
		e.WriterConfig.SetSheetIndex(index)
	}
}

func (e *Excel) SetAxis(axis string) {
	if e.ReaderConfig != nil {
		e.ReaderConfig.SetAxis(axis)
	}
	if e.WriterConfig != nil {
		e.WriterConfig.SetAxis(axis)
	}
}

func (e *Excel) SetAxisCoordinates(col int, row int) {
	if e.ReaderConfig != nil {
		e.ReaderConfig.SetAxisCoordinates(col, row)
	}
	if e.WriterConfig != nil {
		e.WriterConfig.SetAxisCoordinates(col, row)
	}
}

func (e *Excel) Validate() error {
	if e.File == nil {
		return errFileIsNil
	}
	if e.ReaderConfig != nil {
		return e.ReaderConfig.Validate()
	} else if e.WriterConfig != nil {
		return e.WriterConfig.Validate()
	}
	return errConfigNotValid
}
