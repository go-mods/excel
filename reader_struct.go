package excel

import (
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
	e := reflect.Indirect(value).Type().Elem()
	c := &Container{
		Value:   value,
		Type:    e,
		Pointer: e.Kind() == reflect.Pointer,
	}
	r := &StructReader{
		container: c,
		Reader:    reader,
		Struct:    getStruct(c),
	}
	return r, nil
}

// Unmarshall reads the excel file and fill the container
func (r *StructReader) Unmarshall() error {

	// get excel rows
	rows, err := r.Reader.file.Rows(r.Reader.Sheet.Name)
	if err != nil {
		return err
	}

	// prepare the slice Container
	slice := reflect.MakeSlice(reflect.SliceOf(r.container.Type), 0, 0)

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

	r.container.Value.Elem().Set(slice)

	return rows.Close()
}

func (r *StructReader) SetColumnsTags(tags map[string]*Tags) {
	// Loop throw all fields in Struct
	for _, field := range r.Struct.Fields {
		r.Struct.freeze(tags[field.Name], field.ReadTags)
	}
}

func (r *StructReader) updateColumnIndex(row []string) error {
	// Initialize all fields index
	for _, f := range r.Struct.Fields {
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

	containerValue := r.container.newValue()

	// Loop throw all fields
	for _, fieldConfig := range r.Struct.Fields {
		if fieldConfig.ReadTags.index >= 0 {

			if len(row) >= fieldConfig.ReadTags.index+1 {
				value, err = fieldConfig.toValue(row[fieldConfig.ReadTags.index])
				if err != nil {
					value = reflect.Value{}
				}
			} else {
				value = reflect.ValueOf(fieldConfig.GetReadDefault())
			}

			// Assign the value to the containerValue
			if value.IsValid() {
				r.container.assign(containerValue, fieldConfig.Index, value.Convert(fieldConfig.Type))
			}
		}
	}

	return containerValue, nil
}
