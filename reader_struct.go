package excel

import (
	"reflect"
)

type structReader struct {
	readerInfo *ReaderInfo
	container  *ContainerInfo
	structInfo *StructInfo
}

func newStructReader(readerInfo *ReaderInfo, containerValue reflect.Value) (*structReader, error) {
	containerTypeElem := reflect.Indirect(containerValue).Type().Elem()
	c := &ContainerInfo{
		value:     containerValue,
		typeElem:  containerTypeElem,
		isPointer: containerTypeElem.Kind() == reflect.Pointer,
	}
	r := &structReader{
		readerInfo: readerInfo,
		container:  c,
		structInfo: getStructInfo(c),
	}
	return r, nil
}

// Unmarshall must be called when reading an Excel file
func (r *structReader) Unmarshall() error {

	// get excel row
	rows, err := r.readerInfo.file.Rows(r.readerInfo.Sheet.Name)
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

func (w *structReader) SetColumnsOptions(options map[string]*FieldTags) {
	// Loop throw all fields in StructInfo
	for _, field := range w.structInfo.Fields {
		w.structInfo.freeze(options[field.Name], field.TagsIn)
	}
}

func (r *structReader) updateColumnIndex(row []string) error {
	// Initialize all fields index
	for _, f := range r.structInfo.Fields {
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

func (r *structReader) unmarshallRow(row []string) (value reflect.Value, err error) {

	containerElement := r.container.create()

	// Loop throw all fields
	for _, fieldConfig := range r.structInfo.Fields {
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
