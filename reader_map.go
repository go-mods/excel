package excel

import (
	"reflect"

	"github.com/go-mods/convert"
)

// mapReader is the Excel reader for a map
// It implements the IReader interface
type mapReader struct {
	container *Container
	Reader    *Reader
	Columns   columns
}

// column is a column of excel
type column struct {
	index int
	title string
}

// columns is a list of column
type columns []column

func newMapReader(reader *Reader, value reflect.Value) (*mapReader, error) {
	e := reflect.Indirect(value).Type().Elem()
	c := &Container{
		Value:   value,
		Type:    e,
		Pointer: e.Kind() == reflect.Pointer,
	}
	r := &mapReader{
		container: c,
		Reader:    reader,
	}
	return r, nil
}

func (r *mapReader) Unmarshall() (*ReaderResult, error) {
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
			err := r.getColumns(row)
			if err != nil {
				return nil, err
			}
		}

		// Data row
		if rowIndex > 0 {
			value, err := r.unmarshallRow(row)
			if err != nil {
				return nil, err
			}

			if value.IsValid() {
				slice = reflect.Append(slice, value)
			}
		}
		rowIndex++
	}

	// Set the result
	result.Rows = rowIndex
	result.Columns = len(r.Columns)

	// Set the slice to the container
	r.container.Value.Elem().Set(slice)

	return result, rows.Close()
}

func (r *mapReader) SetColumnsTags(_ map[string]*Tags) {
	panic(ErrNotImplemented.Error())
}

func (r *mapReader) getColumns(row []string) error {

	for index, title := range row {
		r.Columns = append(r.Columns, column{
			index: index,
			title: title,
		})
	}

	return nil
}

func (r *mapReader) unmarshallRow(row []string) (reflect.Value, error) {

	containerValue := r.container.newValue()
	containerValueType := containerValue.Type().Elem()

	// The containerValue must be of type Slice
	if containerValue.Kind() != reflect.Map {
		return reflect.Value{}, ErrContainerNotMap
	}

	// Define the map
	if containerValue.IsNil() {
		containerValue.Set(reflect.MakeMap(containerValue.Type()))
	}

	// loop throw all cells of the row
	for index, cell := range row {
		sCell := convert.ToString(cell)
		value, err := convert.ToValueE(sCell, containerValueType)
		if err != nil {
			// try to fins the type
			t := convert.GetConvertType(sCell)
			if t != nil {
				value, err = convert.ToValueE(sCell, t)
				if err != nil {
					return reflect.Value{}, err
				}
			}
		}

		// Assign the value to the containerValue
		if value.IsValid() {
			if r.container.Pointer {
				containerValue.Elem().SetMapIndex(reflect.ValueOf(r.Columns[index].title), value)
			} else {
				containerValue.SetMapIndex(reflect.ValueOf(r.Columns[index].title), value)
			}
		}
	}

	return containerValue, nil
}
