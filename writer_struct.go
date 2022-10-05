package excel

import (
	"github.com/xuri/excelize/v2"
	"reflect"
)

type structWriter struct {
	config *WriterConfig
	schema *schema
}

func newStructWriter(config *WriterConfig, containerElement reflect.Type) (*structWriter, error) {
	r := &structWriter{
		config: config,
		schema: newSchema(containerElement),
	}
	return r, nil
}

func (w *structWriter) Marshall(data any) error {

	// get excel rows to find titles if exists
	rows, err := w.config.file.Rows(w.config.Sheet.Name)
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
	err = w.WriteRows(data)
	if err != nil {
		return err
	}

	return nil
}

func (w *structWriter) updateColumnIndex(row []string) {
	// Initialize all fields index
	for _, f := range w.schema.Fields {
		if !f.IgnoreOut() {
			// Loop throw all columns
			for colIndex, cell := range row {
				if f.ColumnNameOut() == cell {
					f.TagsOut.columnIndex = colIndex
					break
				}
			}
		}
	}

	// Get max column index
	var maxIndex int = 0
	for _, f := range w.schema.Fields {
		if !f.IgnoreOut() {
			if f.TagsOut.columnIndex > maxIndex {
				maxIndex = f.TagsOut.columnIndex
			}
		}
	}

	// Update field column index
	for _, f := range w.schema.Fields {
		if !f.IgnoreOut() {
			if f.TagsOut.columnIndex == -1 {
				f.TagsOut.columnIndex = maxIndex
				maxIndex++
			}
		}
	}
}

func (w *structWriter) WriteRows(slice any) (err error) {

	// Make sure 'slice' is a Ptr to Slice
	s := reflect.ValueOf(slice)
	if s.Kind() != reflect.Ptr || s.Elem().Kind() != reflect.Slice {
		return errContainerInvalid
	}
	s = s.Elem()

	// Get default coordinates
	col, row, _ := excelize.CellNameToCoordinates(w.config.Axis.Axis)

	// Write title
	for _, f := range w.schema.Fields {
		if !f.IgnoreOut() {
			cell, _ := excelize.CoordinatesToCellName(col+f.TagsOut.columnIndex, row)
			if err := w.config.file.SetCellValue(w.config.Sheet.Name, cell, f.ColumnNameOut()); err != nil {
				return err
			}
		}
	}
	row++

	// Write rows
	for i := 0; i < s.Len(); i++ {

		col, _, _ = excelize.CellNameToCoordinates(w.config.Axis.Axis)

		// data
		values := s.Index(i)

		// write
		for j := 0; j < values.NumField(); j++ {
			value := values.Field(j)
			f := w.schema.GetFieldFromFieldIndex(j)
			if !f.IgnoreOut() {
				cell, _ := excelize.CoordinatesToCellName(col+f.TagsOut.columnIndex, row)
				cellValue := f.toCellValue(value.Interface())
				if err := w.config.file.SetCellValue(w.config.Sheet.Name, cell, cellValue); err != nil {
					return err
				}
			}
		}

		row++
	}

	return
}
