package excel

import (
	"fmt"
	"reflect"
)

// StructReader is the Excel reader for a struct
// It implements the IReader interface
type StructReader struct {
	container *Container
	Reader    *Reader
	Struct    *Struct
}

// newStructReader create the appropriate reader
func newStructReader(reader *Reader, value reflect.Value) (*StructReader, error) {
	if reader == nil {
		return nil, fmt.Errorf("excel: reader is nil")
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

	r := &StructReader{
		container: c,
		Reader:    reader,
		Struct:    structInfo,
	}
	return r, nil
}

// Unmarshall reads the excel file and fill the container
func (r *StructReader) Unmarshall() (*ReaderResult, error) {
	if r == nil {
		return nil, fmt.Errorf("excel: struct reader is nil")
	}

	if r.Reader == nil || r.Reader.file == nil {
		return nil, fmt.Errorf("excel: reader or file is nil")
	}

	if r.Struct == nil || r.Struct.Fields == nil {
		return nil, fmt.Errorf("excel: struct or fields are nil")
	}

	// get excel rows
	rows, err := r.Reader.file.Rows(r.Reader.Sheet.Name)
	if err != nil {
		return nil, fmt.Errorf("excel: failed to get rows from sheet '%s': %w", r.Reader.Sheet.Name, err)
	}

	// prepare the slice Container
	slice := reflect.MakeSlice(reflect.SliceOf(r.container.Type), 0, 0)

	// prepare the result
	result := &ReaderResult{}

	// Loop throw all rows
	rowIndex := 0
	for rows.Next() {
		row, err := rows.Columns()
		if err != nil {
			return nil, fmt.Errorf("excel: failed to get columns for row %d: %w", rowIndex, err)
		}
		if row == nil {
			break
		}

		// Title row
		if rowIndex == 0 {
			err := r.updateColumnIndex(row)
			if err != nil {
				if err == ErrColumnRequired {
					return nil, ErrColumnRequired
				}
				return nil, fmt.Errorf("excel: failed to update column index: %w", err)
			}
		}

		// Data row
		if rowIndex > 0 {
			value, err := r.unmarshallRow(row)
			if err != nil {
				return nil, fmt.Errorf("excel: failed to unmarshall row %d: %w", rowIndex, err)
			}

			if value.IsValid() {
				slice = reflect.Append(slice, value)
			}
		}

		// Set the result
		if result.Rows == 0 {
			result.Columns = len(row)
		}

		// Next row
		rowIndex++
	}

	// Set the result
	result.Rows = rowIndex

	// Set the slice to the container
	if !r.container.Value.Elem().CanSet() {
		return nil, fmt.Errorf("excel: container value cannot be set")
	}
	r.container.Value.Elem().Set(slice)

	return result, rows.Close()
}

func (r *StructReader) SetColumnsTags(tags map[string]*Tags) {
	if r == nil || r.Struct == nil || r.Struct.Fields == nil {
		return
	}

	if tags == nil {
		return
	}

	// Loop throw all fields in Struct
	for _, field := range r.Struct.Fields {
		if field != nil && field.ReadTags != nil {
			r.Struct.freeze(tags[field.Name], field.ReadTags)
		}
	}
}

func (r *StructReader) updateColumnIndex(row []string) error {
	if r == nil || r.Struct == nil || r.Struct.Fields == nil {
		return fmt.Errorf("excel: struct reader, struct or fields are nil")
	}

	if row == nil {
		return fmt.Errorf("excel: row is nil")
	}

	// Initialize all fields index
	for _, f := range r.Struct.Fields {
		if f == nil {
			continue
		}

		// Loop throw all columns
		for colIndex, cell := range row {
			if f.GetReadColumnName() == cell && !f.GetReadIgnore() {
				f.ReadTags.index = colIndex
				break
			}
		}
		// Required column
		if f.GetReadRequired() && f.ReadTags.index == -1 {
			return ErrColumnRequired
		}
	}
	return nil
}

func (r *StructReader) unmarshallRow(row []string) (value reflect.Value, err error) {
	if r == nil || r.container == nil || r.Struct == nil || r.Struct.Fields == nil {
		return reflect.Value{}, fmt.Errorf("excel: struct reader, container, struct or fields are nil")
	}

	if row == nil {
		return reflect.Value{}, fmt.Errorf("excel: row is nil")
	}

	containerValue := r.container.newValue()
	if !containerValue.IsValid() {
		return reflect.Value{}, fmt.Errorf("excel: failed to create new container value")
	}

	// Loop throw all fields
	for _, fieldConfig := range r.Struct.Fields {
		if fieldConfig == nil {
			continue
		}

		if fieldConfig.ReadTags.index >= 0 {
			var fieldValue reflect.Value

			if len(row) >= fieldConfig.ReadTags.index+1 {
				fieldValue, err = fieldConfig.toValue(row[fieldConfig.ReadTags.index])
				if err != nil {
					// Log the error but continue with other fields
					// This allows partial data to be read even if some fields fail
					continue
				}
			} else {
				// Use default value if the column is out of range
				fieldValue = reflect.ValueOf(fieldConfig.GetReadDefault())
			}

			// Assign the value to the containerValue
			if fieldValue.IsValid() {
				// Check if the field value can be converted to the target type
				if !fieldValue.Type().ConvertibleTo(fieldConfig.Type) {
					continue
				}

				err = r.container.assign(containerValue, fieldConfig.Index, fieldValue.Convert(fieldConfig.Type))
				if err != nil {
					return reflect.Value{}, fmt.Errorf("excel: failed to assign value to field '%s': %w", fieldConfig.Name, err)
				}
			}
		}
	}

	return containerValue, nil
}
