package excel

import "reflect"

type mapReader struct {
}

func newMapReader(_ *ReaderConfig, _ reflect.Value) (*mapReader, error) {
	return nil, errNotImplemented
}

func (r *mapReader) Unmarshall() error {
	return errNotImplemented
}
