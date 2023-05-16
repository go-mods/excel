package excel

import "reflect"

type mapReader struct {
}

func newMapReader(_ *Reader, _ reflect.Value) (*mapReader, error) {
	return nil, ErrNotImplemented
}

func (r *mapReader) Unmarshall() error {
	return ErrNotImplemented
}

func (r *mapReader) SetColumnsTags(_ map[string]*Tags) {
	panic(ErrNotImplemented.Error())
}
