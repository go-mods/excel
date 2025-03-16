package excel

import (
	"fmt"
	"reflect"

	"github.com/xuri/excelize/v2"
)

// StructWriter is the Excel writer for a struct
// It implements the IWriter interface
type StructWriter struct {
	container *Container
	Writer    *Writer
	Struct    *Struct
}

// newStructWriter create the appropriate writer
func newStructWriter(writer *Writer, value reflect.Value) (*StructWriter, error) {
	if writer == nil {
		return nil, fmt.Errorf("excel: writer is nil")
	}

	if !value.IsValid() {
		return nil, fmt.Errorf("excel: value is not valid")
	}

	e := reflect.Indirect(value).Type().Elem()
	if e.Kind() != reflect.Struct && !(e.Kind() == reflect.Pointer && e.Elem().Kind() == reflect.Struct) {
		return nil, fmt.Errorf("excel: expected struct type, got %v", e.Kind())
	}

	c := &Container{
		Value:   value,
		Type:    e,
		Pointer: e.Kind() == reflect.Pointer,
	}

	structInfo := getStruct(c)
	if structInfo == nil {
		return nil, fmt.Errorf("excel: failed to get struct information")
	}

	w := &StructWriter{
		container: c,
		Writer:    writer,
		Struct:    structInfo,
	}
	return w, nil
}

// Marshall writes the Excel file from the container
func (w *StructWriter) Marshall(data any) (*WriterResult, error) {
	if w == nil {
		return nil, fmt.Errorf("excel: struct writer is nil")
	}

	if w.Writer == nil || w.Writer.file == nil {
		return nil, fmt.Errorf("excel: writer or file is nil")
	}

	if w.Struct == nil || w.Struct.Fields == nil {
		return nil, fmt.Errorf("excel: struct or fields are nil")
	}

	if data == nil {
		return nil, fmt.Errorf("excel: data is nil")
	}

	// get excel rows to find titles if exists
	rows, err := w.Writer.file.Rows(w.Writer.Sheet.Name)
	if err != nil {
		return nil, fmt.Errorf("excel: failed to get rows from sheet '%s': %w", w.Writer.Sheet.Name, err)
	}

	// Get title row
	var titleRow []string
	for rows.Next() {
		row, err := rows.Columns()
		if err != nil {
			return nil, fmt.Errorf("excel: failed to get columns: %w", err)
		}
		titleRow = row
		if len(titleRow) > 0 {
			break
		}
	}

	// Close the rows reader to avoid resource leaks
	if err := rows.Close(); err != nil {
		return nil, fmt.Errorf("excel: failed to close rows reader: %w", err)
	}

	//
	w.updateColumnIndex(titleRow)

	// Write
	count, err := w.writeRows(data)
	if err != nil {
		return nil, fmt.Errorf("excel: failed to write rows: %w", err)
	}

	// prepare the result
	result := &WriterResult{}
	result.Rows = count
	result.Columns = w.Struct.Fields.Count() - w.Struct.Fields.CountWriteIgnored()

	return result, nil
}

func (w *StructWriter) SetColumnsTags(tags map[string]*Tags) {
	if w == nil || w.Struct == nil || w.Struct.Fields == nil {
		return
	}

	if tags == nil {
		return
	}

	// Loop throw all fields in Struct
	for _, field := range w.Struct.Fields {
		if field != nil && field.WriteTags != nil {
			w.Struct.freeze(tags[field.Name], field.WriteTags)
		}
	}
}

func (w *StructWriter) updateColumnIndex(row []string) {
	if w == nil || w.Struct == nil || w.Struct.Fields == nil {
		return
	}

	// Initialize all fields index
	for _, f := range w.Struct.Fields {
		if f == nil || f.GetWriteIgnore() {
			continue
		}

		// Loop throw all columns
		for colIndex, cell := range row {
			if f.GetWriteColumnName() == cell {
				f.WriteTags.index = colIndex
				break
			}
		}
	}

	// Get max column index
	var maxIndex int = 0
	for _, f := range w.Struct.Fields {
		if f == nil || f.GetWriteIgnore() {
			continue
		}

		if f.WriteTags.index > maxIndex {
			maxIndex = f.WriteTags.index
		}
	}

	// Update field column index
	for _, f := range w.Struct.Fields {
		if f == nil || f.GetWriteIgnore() {
			continue
		}

		if f.WriteTags.index == -1 {
			f.WriteTags.index = maxIndex
			maxIndex++
		}
	}
}

func (w *StructWriter) writeRows(slice any) (row int, err error) {
	if w == nil || w.Writer == nil || w.Writer.file == nil || w.Struct == nil {
		return 0, fmt.Errorf("excel: writer components are nil")
	}

	// Make sure 'slice' is a Pointer to Slice
	s := reflect.ValueOf(slice)
	if s.Kind() != reflect.Pointer || s.Elem().Kind() != reflect.Slice {
		return 0, fmt.Errorf("excel: expected pointer to slice, got %v", s.Kind())
	}

	if s.IsNil() {
		return 0, fmt.Errorf("excel: slice is nil")
	}

	s = s.Elem()

	// Get default coordinates
	col, row, err := excelize.CellNameToCoordinates(w.Writer.Axis.Axis)
	if err != nil {
		return 0, fmt.Errorf("excel: invalid axis '%s': %w", w.Writer.Axis.Axis, err)
	}

	// Write title
	// -----------
	for _, f := range w.Struct.Fields {
		if f == nil || f.GetWriteIgnore() {
			continue
		}

		cell, err := excelize.CoordinatesToCellName(col+f.WriteTags.index, row)
		if err != nil {
			return 0, fmt.Errorf("excel: failed to convert coordinates to cell name: %w", err)
		}

		if err := w.Writer.file.SetCellValue(w.Writer.Sheet.Name, cell, f.GetWriteColumnName()); err != nil {
			return 0, fmt.Errorf("excel: failed to set cell value for title at %s: %w", cell, err)
		}
	}
	row++

	// Write rows
	// ----------
	for i := 0; i < s.Len(); i++ {
		col, _, err = excelize.CellNameToCoordinates(w.Writer.Axis.Axis)
		if err != nil {
			return 0, fmt.Errorf("excel: invalid axis '%s': %w", w.Writer.Axis.Axis, err)
		}

		// data
		values := s.Index(i)
		if !values.IsValid() {
			continue
		}

		if values.Kind() == reflect.Pointer {
			if values.IsNil() {
				continue
			}
			values = values.Elem()
		}

		if values.Kind() != reflect.Struct {
			return 0, fmt.Errorf("excel: expected struct, got %v at index %d", values.Kind(), i)
		}

		// write
		for _, f := range w.Struct.Fields {
			if f == nil || f.GetWriteIgnore() {
				continue
			}

			// Get the field value using the container's findFieldByIndex
			fieldValue, err := w.container.findFieldByIndex(values, f.Index)
			if err != nil {
				return 0, fmt.Errorf("excel: failed to find field at index %d: %w", f.Index, err)
			}

			if fieldValue.Kind() == reflect.Pointer && fieldValue.IsNil() {
				continue
			}

			cell, err := excelize.CoordinatesToCellName(col+f.WriteTags.index, row)
			if err != nil {
				return 0, fmt.Errorf("excel: failed to convert coordinates to cell name: %w", err)
			}

			cellValue, err := f.toCellValue(fieldValue.Interface())
			if err != nil {
				return 0, fmt.Errorf("excel: failed to convert value for field '%s': %w", f.Name, err)
			}

			if err = w.Writer.file.SetCellValue(w.Writer.Sheet.Name, cell, cellValue); err != nil {
				return 0, fmt.Errorf("excel: failed to set cell value at %s: %w", cell, err)
			}
		}

		row++
	}

	return row - 1, nil
}
