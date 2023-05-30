package excel

import (
	"github.com/xuri/excelize/v2"
	"reflect"
)

type SliceWriter struct {
	container *Container
	Writer    *Writer
}

func newSliceWriter(writer *Writer, value reflect.Value) (*SliceWriter, error) {
	e := reflect.Indirect(value).Type().Elem()
	c := &Container{
		Value:   value,
		Type:    e,
		Pointer: e.Kind() == reflect.Pointer,
	}
	w := &SliceWriter{
		container: c,
		Writer:    writer,
	}
	return w, nil
}

func (w *SliceWriter) Marshall(data any) (*WriterResult, error) {

	// Write
	result, err := w.writeRows(data)
	if err != nil {
		return nil, err
	}

	return result, nil
}

func (w *SliceWriter) SetColumnsTags(_ map[string]*Tags) {
	panic(ErrNotImplemented.Error())
}

func (w *SliceWriter) writeRows(data any) (*WriterResult, error) {

	// Make sure 'data' is a Pointer to Slice
	s := reflect.ValueOf(data)
	if s.Kind() != reflect.Pointer || s.Elem().Kind() != reflect.Slice {
		return nil, ErrContainerInvalid
	}
	s = s.Elem()

	// Get default coordinates
	col, row, _ := excelize.CellNameToCoordinates(w.Writer.Axis.Axis)

	// prepare the result
	result := &WriterResult{}
	result.Rows = s.Len()

	// Write rows
	for i := 0; i < s.Len(); i++ {

		// data row
		values := s.Index(i)
		if values.Kind() == reflect.Pointer {
			values = values.Elem()
		}

		// loop over columns
		for j := 0; j < values.Len(); j++ {
			// value
			value := values.Index(j)

			// cell
			cell, _ := excelize.CoordinatesToCellName(col+j, row+i)
			if err := w.Writer.file.SetCellValue(w.Writer.Sheet.Name, cell, value); err != nil {
				return nil, err
			}
		}

		// update the result
		if result.Columns == 0 {
			result.Columns = values.Len()
		}
	}

	return result, nil
}
