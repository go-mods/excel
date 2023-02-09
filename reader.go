package excel

import (
	"reflect"
)

// Reader interface
type Reader interface {
	Unmarshall() error
	SetColumnsOptions(options map[string]*FieldTags)
}

// newReader create the appropriate reader
func newReader(info *ReaderInfo, container any) (Reader, error) {
	// The type of the reader depends on the ContainerInfo
	containerValue := reflect.ValueOf(container)
	containerType := reflect.Indirect(containerValue).Type()

	// Validate ContainerInfo
	if containerValue.Kind() != reflect.Pointer && containerType.Kind() != reflect.Slice {
		return nil, ErrContainerInvalid
	}

	// Get element
	containerElement := containerType.Elem()
	if containerElement.Kind() == reflect.Pointer {
		containerElement = containerElement.Elem()
	}

	// create the reader according to the
	// type of element
	switch containerElement.Kind() {
	case reflect.Struct:
		reader, err := newStructReader(info, containerValue)
		return reader, err
	case reflect.Map:
		reader, err := newMapReader(info, containerValue)
		return reader, err
	case reflect.Slice, reflect.Array:
		reader, err := newSliceReader(info, containerValue)
		return reader, err
	}
	return nil, ErrNoReaderFound
}
