package excel

import "reflect"

// Writer interface
type Writer interface {
	Marshall(data any) error
}

// newWriter create the appropriate writer
func newWriter(config *WriterConfig, container any) (Writer, error) {
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
		writer, err := newStructWriter(config, containerElement)
		return writer, err
	case reflect.Map:
		writer, err := newMapWriter(config)
		return writer, err
	case reflect.Slice, reflect.Array:
		writer, err := newSliceWriter(config)
		return writer, err
	}
	return nil, errNoWriterFound
}
