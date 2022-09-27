package excel

import "reflect"

type sliceReader struct {
}

func newSliceReader(_ *ReaderConfig, _ reflect.Value) (*sliceReader, error) {
	return nil, errNotImplemented
}

func (r *sliceReader) Unmarshall() error {
	return errNotImplemented
}
