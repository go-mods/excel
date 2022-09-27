package excel

import (
	"reflect"
)

// Reader interface
type Reader interface {
	Unmarshall() error
}

// newReader create the appropriate reader
func newReader(config *ReaderConfig, container any) (Reader, error) {
	// The type of the reader depends on the container
	containerValue := reflect.ValueOf(container)
	containerType := reflect.Indirect(containerValue).Type()

	// Validate container
	if containerValue.Kind() != reflect.Pointer && containerType.Kind() != reflect.Slice {
		return nil, errContainerInvalid
	}

	// Get element
	containerElement := containerType.Elem()
	if containerElement.Kind() == reflect.Ptr {
		containerElement = containerElement.Elem()
	}

	// create the reader
	switch containerElement.Kind() {
	case reflect.Struct:
		reader, err := newStructReader(config, containerValue, containerElement)
		return reader, err
	case reflect.Map:
		reader, err := newMapReader(config, containerValue)
		return reader, err
	case reflect.Slice, reflect.Array:
		reader, err := newSliceReader(config, containerValue)
		return reader, err
	}
	return nil, errNoReaderFound
}
