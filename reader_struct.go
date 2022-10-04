package excel

import (
	"reflect"
)

type structReader struct {
	config *ReaderConfig
	schema *schema

	containerValue   reflect.Value
	containerElement reflect.Type
}

func newStructReader(config *ReaderConfig, containerValue reflect.Value, containerElement reflect.Type) (*structReader, error) {
	r := &structReader{
		config:           config,
		schema:           newSchema(containerElement),
		containerValue:   containerValue,
		containerElement: containerElement,
	}
	return r, nil
}

func (r *structReader) Unmarshall() error {

	// get excel row
	rows, err := r.config.file.Rows(r.config.Sheet.Name)
	if err != nil {
		return err
	}

	// prepare the slice container
	slice := reflect.MakeSlice(reflect.SliceOf(r.containerElement), 0, 0)

	// Loop throw all rows
	rowIndex := 0
	for rows.Next() {
		row, err := rows.Columns()
		if err != nil {
			break
		}

		// Title row
		if rowIndex == 0 {
			r.updateColumnIndex(row)
		}

		// Data row
		if rowIndex > 0 {
			containerElement, err := r.unmarshallRow(row)
			if err != nil {
				return err
			}

			if containerElement.IsValid() {
				slice = reflect.Append(slice, containerElement)
			}
		}
		rowIndex++
	}

	r.containerValue.Elem().Set(slice)

	return rows.Close()
}

func (r *structReader) updateColumnIndex(row []string) {
	// Initialize all fields index
	for _, f := range r.schema.Fields {
		// Loop throw all columns
		for colIndex, cell := range row {
			if f.ColumnName == cell {
				f.ColumnIndex = colIndex
				break
			}
		}
	}
}

func (r *structReader) unmarshallRow(row []string) (reflect.Value, error) {

	containerElement := reflect.New(r.containerElement).Elem()

	// Loop throw all fields
	for _, fieldConfig := range r.schema.Fields {
		if fieldConfig.ColumnIndex >= 0 {

			fieldValue := containerElement.Field(fieldConfig.FieldIndex)

			value, err := fieldConfig.toValue(row[fieldConfig.ColumnIndex])
			if err != nil {
				return reflect.Value{}, nil
			}

			fieldValue.Set(value.Convert(fieldConfig.FieldType))
		}
	}

	return containerElement, nil
}
