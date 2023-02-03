package excel

import (
	"github.com/xuri/excelize/v2"
)

type Excel struct {
	File       *excelize.File
	ReaderInfo *ReaderInfo
	WriterInfo *WriterInfo
}

func (e *Excel) Marshal(container any) error {
	// Validate excel input
	err := e.Validate()
	if err != nil {
		return err
	}

	// Create the writer
	writer, err := newWriter(e.WriterInfo, container)
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
	reader, err := newReader(e.ReaderInfo, container)
	if err != nil {
		return err
	}

	// unmarshall
	err = reader.Unmarshall()
	return err
}

func (e *Excel) SetSheetName(sheet string) {
	if e.ReaderInfo != nil {
		e.ReaderInfo.SetSheetName(sheet)
	}
	if e.WriterInfo != nil {
		e.WriterInfo.SetSheetName(sheet)
	}
}

func (e *Excel) GetSheetName() string {
	if e.ReaderInfo != nil {
		return e.ReaderInfo.GetSheetName()
	}
	if e.WriterInfo != nil {
		return e.WriterInfo.GetSheetName()
	}
	return ""
}

func (e *Excel) SetSheetIndex(index int) {
	if e.ReaderInfo != nil {
		e.ReaderInfo.SetSheetIndex(index)
	}
	if e.WriterInfo != nil {
		e.WriterInfo.SetSheetIndex(index)
	}
}

func (e *Excel) GetSheetIndex() int {
	if e.ReaderInfo != nil {
		i, err := e.ReaderInfo.GetSheetIndex()
		if err != nil {
			return 0
		}
		return i
	}
	if e.WriterInfo != nil {
		i, err := e.WriterInfo.GetSheetIndex()
		if err != nil {
			return 0
		}
		return i
	}
	return 0
}

func (e *Excel) SetAxis(axis string) {
	if e.ReaderInfo != nil {
		e.ReaderInfo.SetAxis(axis)
	}
	if e.WriterInfo != nil {
		e.WriterInfo.SetAxis(axis)
	}
}

func (e *Excel) SetAxisCoordinates(col int, row int) {
	if e.ReaderInfo != nil {
		e.ReaderInfo.SetAxisCoordinates(col, row)
	}
	if e.WriterInfo != nil {
		e.WriterInfo.SetAxisCoordinates(col, row)
	}
}

func (e *Excel) Validate() error {
	if e.File == nil {
		return ErrFileIsNil
	}
	if e.ReaderInfo != nil {
		return e.ReaderInfo.Validate()
	} else if e.WriterInfo != nil {
		return e.WriterInfo.Validate()
	}
	return ErrConfigNotValid
}
