package excel

import (
	"github.com/xuri/excelize/v2"
	"reflect"
)

// IWriter interface
// All writers must implement this interface
type IWriter interface {
	Marshall(data any) error
	SetColumnsTags(tags map[string]*Tags)
}

// Writer is the Excel writer
type Writer struct {
	file  *excelize.File
	Sheet Sheet
	Axis  Axis
}

// validate validates the writer
// It returns an error if :
// - the sheet is not valid
// - the axis is not valid
func (w *Writer) validate() error {
	if w.file == nil {
		return ErrFileIsNil
	}
	if !w.isSheetValid() {
		return ErrSheetNotValid
	}
	if !w.isAxisValid() {
		return ErrAxisNotValid
	}
	return nil
}

// newWriter create the appropriate writer
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
		return nil, ErrNoReaderFound
	}
}
