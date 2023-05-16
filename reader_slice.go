package excel

import "reflect"

type sliceReader struct {
}

func newSliceReader(_ *Reader, _ reflect.Value) (*sliceReader, error) {
	return nil, ErrNotImplemented
}

func (r *sliceReader) Unmarshall() error {
	return ErrNotImplemented
}

func (r *sliceReader) SetColumnsTags(_ map[string]*Tags) {
	panic(ErrNotImplemented.Error())
}
