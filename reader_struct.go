package excel

import (
	"reflect"
)

type StructReader struct {
	ReaderInfo *ReaderInfo
	container  *ContainerInfo
	StructInfo *StructInfo
}

func newStructReader(readerInfo *ReaderInfo, containerValue reflect.Value) (*StructReader, error) {
	containerTypeElem := reflect.Indirect(containerValue).Type().Elem()
	c := &ContainerInfo{
		value:     containerValue,
		typeElem:  containerTypeElem,
		isPointer: containerTypeElem.Kind() == reflect.Pointer,
	}
	r := &StructReader{
		ReaderInfo: readerInfo,
		container:  c,
		StructInfo: getStructInfo(c),
	}
	return r, nil
}

// Unmarshall must be called when reading an Excel file
func (r *StructReader) Unmarshall() error {

	// get excel row
	rows, err := r.ReaderInfo.file.Rows(r.ReaderInfo.Sheet.Name)
	if err != nil {
		return err
	}

	// prepare the slice ContainerInfo
	slice := reflect.MakeSlice(reflect.SliceOf(r.container.typeElem), 0, 0)

	// Loop throw all rows
	rowIndex := 0
	for rows.Next() {
		row, err := rows.Columns()
		if err != nil {
			break
		}
		if row == nil {
			break
		}

		// Title row
		if rowIndex == 0 {
			err := r.updateColumnIndex(row)
			if err != nil {
				return err
			}
		}

		// Data row
		if rowIndex > 0 {
			value, err := r.unmarshallRow(row)
			if err != nil {
				return err
			}

			if value.IsValid() {
				slice = reflect.Append(slice, value)
			}
		}
		rowIndex++
	}

	r.container.value.Elem().Set(slice)

	return rows.Close()
}

func (w *StructReader) SetColumnsOptions(options map[string]*FieldTags) {
	// Loop throw all fields in StructInfo
	for _, field := range w.StructInfo.Fields {
		w.StructInfo.freeze(options[field.Name], field.TagsIn)
	}
}

func (r *StructReader) updateColumnIndex(row []string) error {
	// Initialize all fields index
	for _, f := range r.StructInfo.Fields {
		// Loop throw all columns
		for colIndex, cell := range row {
			if f.ColumnNameIn() == cell && !f.IgnoreIn() {
				f.TagsIn.columnIndex = colIndex
				break
			}
		}
		// Required column
		if f.IsRequiredIn() && f.TagsIn.columnIndex == -1 {
			return ErrColumnRequired
		}
	}
	return nil
}

func (r *StructReader) unmarshallRow(row []string) (value reflect.Value, err error) {

	containerElement := r.container.create()

	// Loop throw all fields
	for _, fieldConfig := range r.StructInfo.Fields {
		if fieldConfig.TagsIn.columnIndex >= 0 {

			if len(row) >= fieldConfig.TagsIn.columnIndex+1 {
				value, err = fieldConfig.toValue(row[fieldConfig.TagsIn.columnIndex])
				if err != nil {
					return reflect.Value{}, nil
				}
			} else {
				value = reflect.ValueOf(fieldConfig.DefaultValueIn())
			}

			if value.IsValid() {
				r.container.setFieldValue(containerElement, fieldConfig.Index, value.Convert(fieldConfig.Type))
			}
		}
	}

	return containerElement, nil
}
