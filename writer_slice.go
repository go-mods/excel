package excel

import "reflect"

type sliceWriter struct {
}

func newSliceWriter(_ *Writer, _ reflect.Value) (*sliceWriter, error) {
	return nil, ErrNotImplemented
}

func (w *sliceWriter) Marshall(_ any) error {
	return ErrNotImplemented
}

func (w *sliceWriter) SetColumnsTags(_ map[string]*Tags) {
	panic(ErrNotImplemented.Error())
}
