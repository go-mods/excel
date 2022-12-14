package excel

import (
	"reflect"

	"github.com/xuri/excelize/v2"
)

type structWriter struct {
	writerInfo *WriterInfo
	structInfo *StructInfo
}

func newStructWriter(writerInfo *WriterInfo, containerValue reflect.Value) (*structWriter, error) {
	containerTypeElem := reflect.Indirect(containerValue).Type().Elem()
	c := &ContainerInfo{
		value:     containerValue,
		typeElem:  containerTypeElem,
		isPointer: containerTypeElem.Kind() == reflect.Pointer,
	}
	r := &structWriter{
		writerInfo: writerInfo,
		structInfo: getStructInfo(c),
	}
	return r, nil
}

func (w *structWriter) Marshall(data any) error {

	// get excel rows to find titles if exists
	rows, err := w.writerInfo.file.Rows(w.writerInfo.Sheet.Name)
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
	for _, f := range w.structInfo.Fields {
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
	for _, f := range w.structInfo.Fields {
		if !f.IgnoreOut() {
			if f.TagsOut.columnIndex > maxIndex {
				maxIndex = f.TagsOut.columnIndex
			}
		}
	}

	// Update field column index
	for _, f := range w.structInfo.Fields {
		if !f.IgnoreOut() {
			if f.TagsOut.columnIndex == -1 {
				f.TagsOut.columnIndex = maxIndex
				maxIndex++
			}
		}
	}
}

func (w *structWriter) WriteRows(slice any) (err error) {

	// Make sure 'slice' is a Pointer to Slice
	s := reflect.ValueOf(slice)
	if s.Kind() != reflect.Pointer || s.Elem().Kind() != reflect.Slice {
		return ErrContainerInvalid
	}
	s = s.Elem()

	// Get default coordinates
	col, row, _ := excelize.CellNameToCoordinates(w.writerInfo.Axis.Axis)

	// Write title
	// -----------
	for _, f := range w.structInfo.Fields {
		if !f.IgnoreOut() {
			cell, _ := excelize.CoordinatesToCellName(col+f.TagsOut.columnIndex, row)
			if err := w.writerInfo.file.SetCellValue(w.writerInfo.Sheet.Name, cell, f.ColumnNameOut()); err != nil {
				return err
			}
		}
	}
	row++

	// Write rows
	// ----------
	for i := 0; i < s.Len(); i++ {

		col, _, _ = excelize.CellNameToCoordinates(w.writerInfo.Axis.Axis)

		// data
		values := s.Index(i)
		if values.Kind() == reflect.Pointer {
			values = values.Elem()
		}

		// write
		for j := 0; j < values.NumField(); j++ {
			value := values.Field(j)
			f := w.structInfo.GetFieldFromFieldIndex(j)
			if !f.IgnoreOut() {
				cell, _ := excelize.CoordinatesToCellName(col+f.TagsOut.columnIndex, row)
				cellValue, err := f.toCellValue(value.Interface())
				if err != nil {
					return err
				}
				if err = w.writerInfo.file.SetCellValue(w.writerInfo.Sheet.Name, cell, cellValue); err != nil {
					return err
				}
			}
		}

		row++
	}

	return
}
