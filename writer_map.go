package excel

import "reflect"

type mapWriter struct {
}

func newMapWriter(_ *WriterInfo, _ reflect.Value) (*mapWriter, error) {
	return nil, ErrNotImplemented
}

func (w *mapWriter) Marshall(data any) error {
	return ErrNotImplemented
}
