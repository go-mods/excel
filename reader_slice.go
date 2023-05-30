package excel

import (
	"github.com/go-mods/convert"
	"reflect"
)

type SliceReader struct {
	container *Container
	Reader    *Reader
}

func newSliceReader(reader *Reader, value reflect.Value) (*SliceReader, error) {
	e := reflect.Indirect(value).Type().Elem()
	c := &Container{
		Value:   value,
		Type:    e,
		Pointer: e.Kind() == reflect.Pointer,
	}
	r := &SliceReader{
		container: c,
		Reader:    reader,
	}
	return r, nil
}

func (r *SliceReader) Unmarshall() (*ReaderResult, error) {

	// get excel rows
	rows, err := r.Reader.file.Rows(r.Reader.Sheet.Name)
	if err != nil {
		return nil, err
	}

	// prepare the slice Container
	slice := reflect.MakeSlice(reflect.SliceOf(r.container.Type), 0, 0)

	// prepare the result
	result := &ReaderResult{}

	// Loop throw all rows
	for rows.Next() {
		row, err := rows.Columns()
		if err != nil {
			break
		}
		if row == nil {
			break
		}

		value, err := r.unmarshallRow(row)
		if err != nil {
			return nil, err
		}

		if value.IsValid() {
			slice = reflect.Append(slice, value)
		}

		// Set the result
		if result.Rows == 0 {
			result.Columns = len(row)
		}
	}

	// Set the result
	result.Rows = slice.Len()

	// Set the slice to the container
	r.container.Value.Elem().Set(slice)

	return result, rows.Close()
}

func (r *SliceReader) SetColumnsTags(_ map[string]*Tags) {
	panic(ErrNotImplemented.Error())
}

func (r *SliceReader) unmarshallRow(row []string) (reflect.Value, error) {

	containerValue := r.container.newValue()
	containerValueType := containerValue.Type().Elem()

	// The containerValue must be of type Slice
	if containerValue.Kind() != reflect.Slice {
		return reflect.Value{}, ErrContainerNotSlice
	}

	// Resize the containerValue to the number of cells in the row
	if containerValue.IsNil() {
		containerValue.Set(reflect.MakeSlice(containerValue.Type(), len(row), len(row)))
	}

	// loop throw all cells of the row
	for index, cell := range row {
		sCell := convert.ToValidString(cell)
		value, err := convert.ToValue(sCell, containerValueType)
		if err != nil {
			// try to fins the type
			t := convert.GetConvertType(sCell)
			if t != nil {
				value, err = convert.ToValue(sCell, t)
				if err != nil {
					return reflect.Value{}, err
				}
			}
		}

		// Assign the value to the containerValue
		if value.IsValid() {
			r.container.assign(containerValue, index, value)
		}
	}

	return containerValue, nil
}
