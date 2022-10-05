package excel

import "reflect"

// Writer interface
type Writer interface {
	Marshall(data any) error
}

// newWriter create the appropriate writer
func newWriter(info *WriterInfo, container any) (Writer, error) {
	// The type of the reader depends on the ContainerInfo
	containerValue := reflect.ValueOf(container)
	containerType := reflect.Indirect(containerValue).Type()

	// Validate ContainerInfo
	if containerValue.Kind() != reflect.Pointer && containerType.Kind() != reflect.Slice {
		return nil, ErrContainerInvalid
	}

	// Get element
	containerElement := containerType.Elem()
	if containerElement.Kind() == reflect.Ptr {
		containerElement = containerElement.Elem()
	}

	// create the reader
	switch containerElement.Kind() {
	case reflect.Struct:
		writer, err := newStructWriter(info, containerValue)
		return writer, err
	case reflect.Map:
		writer, err := newMapWriter(info, containerValue)
		return writer, err
	case reflect.Slice, reflect.Array:
		writer, err := newSliceWriter(info, containerValue)
		return writer, err
	}
	return nil, ErrNoWriterFound
}
