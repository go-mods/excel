package excel

import "errors"

var (
	errFileIsNil      = errors.New("excel: the file is nil")
	errSheetNotValid  = errors.New("excel: the sheet name is not valid")
	errAxisNotValid   = errors.New("excel: the axis is not valid")
	errConfigNotValid = errors.New("excel: the configuration is not valid")

	errNoReaderFound = errors.New("excel: unable to create an appropriate reader")
	errNoWriterFound = errors.New("excel: unable to create an appropriate writer")

	errContainerInvalid = errors.New("excel: the container must be a slice or a pointer")

	errNotImplemented = errors.New("excel: not implemented")
)
