package excel

import (
	"github.com/xuri/excelize/v2"
	"reflect"
)

// IReader interface
// All readers must implement this interface
type IReader interface {
	Unmarshall() (*ReaderResult, error)
	SetColumnsTags(tags map[string]*Tags)
}

// Reader is the Excel reader
type Reader struct {
	file   *excelize.File
	Sheet  Sheet
	Axis   Axis
	Result *ReaderResult
}

// ReaderResult is a struct that contains the result of the reader
type ReaderResult struct {
	Rows    int
	Columns int
}

// validate validates the reader
// It returns an error if :
// - the sheet is not valid
// - the axis is not valid
func (r *Reader) validate() error {
	if r.file == nil {
		return ErrFileIsNil
	}
	if !r.isSheetValid() {
		return ErrSheetNotValid
	}
	if !r.isAxisValid() {
		return ErrAxisNotValid
	}
	return nil
}

// newReader create the appropriate reader
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
