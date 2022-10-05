package excel

import "reflect"

type sliceWriter struct {
}

func newSliceWriter(_ *WriterInfo, _ reflect.Value) (*sliceWriter, error) {
	return nil, ErrNotImplemented
}

func (w *sliceWriter) Marshall(data any) error {
	return ErrNotImplemented
}
