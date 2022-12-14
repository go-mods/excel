package excel

import "errors"

var (
	ErrFileIsNil      = errors.New("excel: the file is nil")
	ErrSheetNotValid  = errors.New("excel: the sheet name is not valid")
	ErrAxisNotValid   = errors.New("excel: the axis is not valid")
	ErrConfigNotValid = errors.New("excel: the configuration is not valid")

	ErrNoReaderFound = errors.New("excel: unable to create an appropriate reader")
	ErrNoWriterFound = errors.New("excel: unable to create an appropriate writer")

	ErrContainerInvalid = errors.New("excel: the ContainerInfo must be a slice or a pointer")
	ErrColumnRequired   = errors.New("excel: required colum")

	ErrNotImplemented = errors.New("excel: not implemented")
)
