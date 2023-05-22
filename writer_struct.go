package excel

import (
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
	e := reflect.Indirect(value).Type().Elem()
	c := &Container{
		Value:   value,
		Type:    e,
		Pointer: e.Kind() == reflect.Pointer,
	}
	w := &StructWriter{
		container: c,
		Writer:    writer,
		Struct:    getStruct(c),
	}
	return w, nil
}

// Marshall writes the Excel file from the container
func (w *StructWriter) Marshall(data any) error {

	// get excel rows to find titles if exists
	rows, err := w.Writer.file.Rows(w.Writer.Sheet.Name)
	if err != nil {
		return err
	}

	// Get title row
	var titleRow []string
	for rows.Next() {
		row, err := rows.Columns()
		if err != nil {
			break
		}
		titleRow = row
		if len(titleRow) > 0 {
			break
		}
	}

	//
	w.updateColumnIndex(titleRow)

	// Write
	err = w.writeRows(data)
	if err != nil {
		return err
	}

	return nil
}

func (w *StructWriter) SetColumnsTags(tags map[string]*Tags) {
	// Loop throw all fields in Struct
	for _, field := range w.Struct.Fields {
		w.Struct.freeze(tags[field.Name], field.WriteTags)
	}
}

func (w *StructWriter) updateColumnIndex(row []string) {
	// Initialize all fields index
	for _, f := range w.Struct.Fields {
		if !f.GetWriteIgnore() {
			// Loop throw all columns
			for colIndex, cell := range row {
				if f.GetWriteColumnName() == cell {
					f.WriteTags.index = colIndex
					break
				}
			}
		}
	}

	// Get max column index
	var maxIndex int = 0
	for _, f := range w.Struct.Fields {
		if !f.GetWriteIgnore() {
			if f.WriteTags.index > maxIndex {
				maxIndex = f.WriteTags.index
			}
		}
	}

	// Update field column index
	for _, f := range w.Struct.Fields {
		if !f.GetWriteIgnore() {
			if f.WriteTags.index == -1 {
				f.WriteTags.index = maxIndex
				maxIndex++
			}
		}
	}
}

func (w *StructWriter) writeRows(slice any) (err error) {

	// Make sure 'slice' is a Pointer to Slice
	s := reflect.ValueOf(slice)
	if s.Kind() != reflect.Pointer || s.Elem().Kind() != reflect.Slice {
		return ErrContainerInvalid
	}
	s = s.Elem()

	// Get default coordinates
	col, row, _ := excelize.CellNameToCoordinates(w.Writer.Axis.Axis)

	// Write title
	// -----------
	for _, f := range w.Struct.Fields {
		if !f.GetWriteIgnore() {
			cell, _ := excelize.CoordinatesToCellName(col+f.WriteTags.index, row)
			if err := w.Writer.file.SetCellValue(w.Writer.Sheet.Name, cell, f.GetWriteColumnName()); err != nil {
				return err
			}
		}
	}
	row++

	// Write rows
	// ----------
	for i := 0; i < s.Len(); i++ {

		col, _, _ = excelize.CellNameToCoordinates(w.Writer.Axis.Axis)

		// data
		values := s.Index(i)
		if values.Kind() == reflect.Pointer {
			values = values.Elem()
		}

		// write
		for j := 0; j < values.NumField(); j++ {
			value := values.Field(j)
			f := w.Struct.GetField(j)
			if !f.GetWriteIgnore() {
				cell, _ := excelize.CoordinatesToCellName(col+f.WriteTags.index, row)
				cellValue, err := f.toCellValue(value.Interface())
				if err != nil {
					return err
				}
				if err = w.Writer.file.SetCellValue(w.Writer.Sheet.Name, cell, cellValue); err != nil {
					return err
				}
			}
		}

		row++
	}

	return
}
