package excel

import (
	"fmt"
	"reflect"
	"sort"

	"github.com/xuri/excelize/v2"
)

type MapWriter struct {
	container *Container
	Writer    *Writer
}

func newMapWriter(writer *Writer, value reflect.Value) (*MapWriter, error) {
	if writer == nil {
		return nil, fmt.Errorf("excel: writer is nil")
	}

	if !value.IsValid() {
		return nil, fmt.Errorf("excel: value is not valid")
	}

	e := reflect.Indirect(value).Type().Elem()
	c := &Container{
		Value:   value,
		Type:    e,
		Pointer: e.Kind() == reflect.Pointer,
	}
	w := &MapWriter{
		container: c,
		Writer:    writer,
	}
	return w, nil
}

func (w *MapWriter) Marshall(data any) (*WriterResult, error) {
	if w == nil || w.Writer == nil || w.Writer.file == nil {
		return nil, fmt.Errorf("excel: writer components are nil")
	}

	if data == nil {
		return nil, fmt.Errorf("excel: data is nil")
	}

	// Write
	result, err := w.writeRows(data)
	if err != nil {
		return nil, fmt.Errorf("excel: failed to write rows: %w", err)
	}

	return result, nil
}

func (w *MapWriter) SetColumnsTags(_ map[string]*Tags) {
	panic(ErrNotImplemented.Error())
}

func (w *MapWriter) writeRows(slice any) (*WriterResult, error) {
	if w == nil || w.Writer == nil || w.Writer.file == nil {
		return nil, fmt.Errorf("excel: writer components are nil")
	}

	// Make sure 'slice' is a Pointer to Slice
	s := reflect.ValueOf(slice)
	if s.Kind() != reflect.Pointer || s.Elem().Kind() != reflect.Slice {
		return nil, fmt.Errorf("excel: expected pointer to slice, got %v", s.Kind())
	}

	if s.IsNil() {
		return nil, fmt.Errorf("excel: slice is nil")
	}

	s = s.Elem()

	if s.Len() == 0 {
		return &WriterResult{Rows: 0, Columns: 0}, nil
	}

	// Get default coordinates
	col, row, err := excelize.CellNameToCoordinates(w.Writer.Axis.Axis)
	if err != nil {
		return nil, fmt.Errorf("excel: invalid axis '%s': %w", w.Writer.Axis.Axis, err)
	}

	// Keys store the keys of the map
	// (Columns headers)
	var keys []reflect.Value
	var sortedKeys []string

	// Loop over slice rows
	for i := 0; i < s.Len(); i++ {
		// get values from map
		values := s.Index(i)
		if !values.IsValid() {
			continue
		}

		// if pointer, get value
		if values.Kind() == reflect.Pointer {
			if values.IsNil() {
				continue
			}
			values = values.Elem()
		}

		if values.Kind() != reflect.Map {
			return nil, fmt.Errorf("excel: expected map, got %v at index %d", values.Kind(), i)
		}

		// if first row, get the keys
		// (columns headers)
		if i == 0 {
			// Get the keys
			keys = values.MapKeys()

			// Sort the keys alphabetically for consistent order
			sortedKeys = make([]string, 0, len(keys))
			for _, key := range keys {
				if key.Kind() == reflect.String {
					sortedKeys = append(sortedKeys, key.String())
				} else {
					// If the key is not a string, convert it to a string
					sortedKeys = append(sortedKeys, fmt.Sprintf("%v", key.Interface()))
				}
			}

			// Sort the keys
			sort.Strings(sortedKeys)

			// Write the headers
			for j, keyStr := range sortedKeys {
				cell, err := excelize.CoordinatesToCellName(col+j, row)
				if err != nil {
					return nil, fmt.Errorf("excel: failed to convert coordinates to cell name: %w", err)
				}

				if err := w.Writer.file.SetCellValue(w.Writer.Sheet.Name, cell, keyStr); err != nil {
					return nil, fmt.Errorf("excel: failed to set cell value for header at %s: %w", cell, err)
				}
			}
		}

		// loop over columns
		for j, keyStr := range sortedKeys {
			// Convert the string key back to a reflect.Value
			var keyValue reflect.Value
			for _, k := range keys {
				if k.Kind() == reflect.String && k.String() == keyStr {
					keyValue = k
					break
				} else if fmt.Sprintf("%v", k.Interface()) == keyStr {
					keyValue = k
					break
				}
			}

			if !keyValue.IsValid() {
				continue
			}

			// get value from key
			value := values.MapIndex(keyValue)
			if !value.IsValid() {
				continue
			}

			// cell
			cell, err := excelize.CoordinatesToCellName(col+j, row+i+1)
			if err != nil {
				return nil, fmt.Errorf("excel: failed to convert coordinates to cell name: %w", err)
			}

			if err := w.Writer.file.SetCellValue(w.Writer.Sheet.Name, cell, value.Interface()); err != nil {
				return nil, fmt.Errorf("excel: failed to set cell value at %s: %w", cell, err)
			}
		}
	}

	// prepare the result
	result := &WriterResult{}
	result.Rows = s.Len()
	result.Columns = len(sortedKeys)

	return result, nil
}
