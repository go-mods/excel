package excel

import (
	"reflect"

	"github.com/xuri/excelize/v2"
)

// IWriter interface defines the contract for all Excel writers.
// All writers must implement this interface to provide consistent
// functionality for marshaling Go structures into Excel data.
type IWriter interface {
	// Marshall converts a Go structure into Excel data and writes it to the file.
	// Returns a WriterResult containing information about the write operation
	// and an error if the operation fails.
	Marshall(data any) (*WriterResult, error)

	// SetColumnsTags sets custom tags for columns to control the marshaling process.
	SetColumnsTags(tags map[string]*Tags)
}

// Writer is the base Excel writer that provides common functionality
// for all specific writer implementations (struct, slice, map).
type Writer struct {
	file   *excelize.File
	Sheet  Sheet
	Axis   Axis
	Result *WriterResult
}

// WriterResult contains information about the result of a write operation,
// including the number of rows and columns processed.
type WriterResult struct {
	Rows    int
	Columns int
}

// validate validates the writer configuration.
// It returns an error if:
// - the file is nil
// - the sheet is not valid
// - the axis is not valid
func (w *Writer) validate() error {
	if w.file == nil {
		return ErrFileIsNil
	}
	if !w.Sheet.IsValid() {
		return ErrSheetNotValid
	}
	if !w.isAxisValid() {
		return ErrAxisNotValid
	}
	return nil
}

// newWriter creates the appropriate writer implementation based on the container type.
// It analyzes the container's type and returns a writer that can handle that specific type.
// Supported container types are slices of structs, slices of slices, and slices of maps.
// Returns an error if no appropriate writer can be created for the container type.
func (w *Writer) newWriter(container any) (IWriter, error) {
	// The type of the reader depends on the Container
	v := reflect.ValueOf(container)
	t := reflect.Indirect(v).Type()

	// The container must be a pointer to a slice
	if v.Kind() != reflect.Pointer || t.Kind() != reflect.Slice {
		return nil, ErrContainerInvalid
	}

	// Get element type of the container
	e := t.Elem()
	if e.Kind() == reflect.Pointer {
		e = e.Elem()
	}

	// Create the reader according to the type of the container
	// and the type of the elements
	switch e.Kind() {
	case reflect.Struct:
		writer, err := newStructWriter(w, v)
		return writer, err
	case reflect.Slice:
		writer, err := newSliceWriter(w, v)
		return writer, err
	case reflect.Map:
		writer, err := newMapWriter(w, v)
		return writer, err
	default:
		return nil, ErrNoWriterFound
	}
}
