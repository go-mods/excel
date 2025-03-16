package excel

import (
	"reflect"

	"github.com/xuri/excelize/v2"
)

// IReader interface defines the contract for all Excel readers.
// All readers must implement this interface to provide consistent
// functionality for unmarshaling Excel data into Go structures.
type IReader interface {
	// Unmarshall reads Excel data and converts it into a Go structure.
	// Returns a ReaderResult containing information about the read operation
	// and an error if the operation fails.
	Unmarshall() (*ReaderResult, error)

	// SetColumnsTags sets custom tags for columns to control the unmarshaling process.
	SetColumnsTags(tags map[string]*Tags)
}

// Reader is the base Excel reader that provides common functionality
// for all specific reader implementations (struct, slice, map).
type Reader struct {
	file   *excelize.File
	Sheet  Sheet
	Axis   Axis
	Result *ReaderResult
}

// ReaderResult contains information about the result of a read operation,
// including the number of rows and columns processed.
type ReaderResult struct {
	Rows    int
	Columns int
}

// validate validates the reader configuration.
// It returns an error if:
// - the file is nil
// - the sheet is not valid
// - the axis is not valid
func (r *Reader) validate() error {
	if r.file == nil {
		return ErrFileIsNil
	}
	if !r.Sheet.IsValid() {
		return ErrSheetNotValid
	}
	if !r.isAxisValid() {
		return ErrAxisNotValid
	}
	return nil
}

// getRows returns the rows from the sheet starting from the defined axis
// If the axis is valid, it will return rows starting from the axis row
// Otherwise, it will return all rows from the sheet
// It also returns the starting column index if an axis is defined
func (r *Reader) getRows() (*excelize.Rows, int, error) {
	rows, err := r.file.Rows(r.Sheet.Name)
	if err != nil {
		return nil, 0, err
	}

	// If axis is valid, skip rows until we reach the axis row
	// and set the starting column index
	if r.isAxisValid() {
		// Get coordinates from the axis
		startCol, startRow, err := excelize.CellNameToCoordinates(r.Axis.Axis)
		if err != nil {
			return nil, 0, err
		}

		// Skip rows until we reach the axis row
		if startRow > 1 {
			rowIndex := 1
			for rows.Next() && rowIndex < startRow-1 {
				rowIndex++
			}
		}

		return rows, startCol - 1, nil
	}

	return rows, 0, nil
}

// newReader creates the appropriate reader implementation based on the container type.
// It analyzes the container's type and returns a reader that can handle that specific type.
// Supported container types are slices of structs, slices of slices, and slices of maps.
// Returns an error if no appropriate reader can be created for the container type.
func (r *Reader) newReader(container any) (IReader, error) {
	// The type of the reader depends on the Container
	v := reflect.ValueOf(container)
	t := reflect.Indirect(v).Type()

	// The container must be a pointer to a slice
	if v.Kind() != reflect.Pointer || t.Kind() != reflect.Slice {
		return nil, ErrContainerInvalid
	}

	// Get the element type of the container
	e := t.Elem()
	if e.Kind() == reflect.Pointer {
		e = e.Elem()
	}

	// Create the reader according to the type of the container
	// and the type of the elements
	switch e.Kind() {
	case reflect.Struct:
		reader, err := newStructReader(r, v)
		return reader, err
	case reflect.Slice:
		reader, err := newSliceReader(r, v)
		return reader, err
	case reflect.Map:
		reader, err := newMapReader(r, v)
		return reader, err
	default:
		return nil, ErrNoReaderFound
	}
}
