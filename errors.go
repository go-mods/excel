package excel

import "errors"

var (
	ErrFileIsNil      = errors.New("excel: the file is nil")
	ErrAxisNotValid   = errors.New("excel: the axis is not valid")
	ErrConfigNotValid = errors.New("excel: the configuration is not valid")

	ErrSheetNotValid  = errors.New("excel: the sheet name is not valid")
	ErrSheetNotFound  = errors.New("excel: the sheet is not found")
	ErrSheetNameEmpty = errors.New("excel: the sheet name is empty")
	ErrSheetIndex     = errors.New("excel: the sheet index is not valid")

	ErrMapKeyNotString   = errors.New("excel: the map key must be a string")
	ErrNoReaderFound     = errors.New("excel: unable to create an appropriate reader")
	ErrNoWriterFound     = errors.New("excel: unable to create an appropriate writer")
	ErrContainerNotSlice = errors.New("excel: the Container must be a slice")
	ErrContainerNotMap   = errors.New("excel: the Container must be a map")

	ErrContainerInvalid = errors.New("excel: the Container must be a slice or a pointer")
	ErrColumnRequired   = errors.New("excel: required colum")

	ErrNotImplemented = errors.New("excel: not implemented")
)
