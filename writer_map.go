package excel

import (
	"reflect"

	"github.com/xuri/excelize/v2"
)

type MapWriter struct {
	container *Container
	Writer    *Writer
}

func newMapWriter(writer *Writer, value reflect.Value) (*MapWriter, error) {
	e := reflect.Indirect(value).Type().Elem()
	c := &Container{
		Value:   value,
		Type:    e,
		Pointer: e.Kind() == reflect.Pointer,
	}
	w := &MapWriter{
		container: c,
		Writer:    writer,
	}
	return w, nil
}

func (w *MapWriter) Marshall(data any) (*WriterResult, error) {

	// Write
	result, err := w.writeRows(data)
	if err != nil {
		return nil, err
	}

	return result, nil
}

func (w *MapWriter) SetColumnsTags(_ map[string]*Tags) {
	panic(ErrNotImplemented.Error())
}

func (w *MapWriter) writeRows(slice any) (*WriterResult, error) {

	// Make sure 'slice' is a Pointer to Slice
	s := reflect.ValueOf(slice)
	if s.Kind() != reflect.Pointer || s.Elem().Kind() != reflect.Slice {
		return nil, ErrContainerInvalid
	}
	s = s.Elem()

	// Get default coordinates
	col, row, _ := excelize.CellNameToCoordinates(w.Writer.Axis.Axis)

	// Keys store the keys of the map*
	// (Columns headers)
	var keys []reflect.Value

	// Loop over slice rows
	for i := 0; i < s.Len(); i++ {

		// get values from map
		values := s.Index(i)

		// if pointer, get value
		if values.Kind() == reflect.Pointer {
			values = values.Elem()
		}

		// if first row, get the keys
		// (columns headers)
		if i == 0 {
			// title
			keys = values.MapKeys()
			// loop overs keys
			for j, key := range keys {
				// write key in cell
				cell, _ := excelize.CoordinatesToCellName(col+j, row)
				if err := w.Writer.file.SetCellValue(w.Writer.Sheet.Name, cell, key); err != nil {
					return nil, err
				}
			}
		}

		// loop over columns
		for j, key := range keys {
			// get value from key
			value := values.MapIndex(key)
			// cell
			cell, _ := excelize.CoordinatesToCellName(col+j, row+i+1)
			if err := w.Writer.file.SetCellValue(w.Writer.Sheet.Name, cell, value); err != nil {
				return nil, err
			}
		}
	}

	// prepare the result
	result := &WriterResult{}
	result.Rows = s.Len()
	result.Columns = len(keys)

	return result, nil
}
