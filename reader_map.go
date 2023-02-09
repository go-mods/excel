package excel

import "reflect"

type mapReader struct {
}

func newMapReader(_ *ReaderInfo, _ reflect.Value) (*mapReader, error) {
	return nil, ErrNotImplemented
}

func (r *mapReader) Unmarshall() error {
	return ErrNotImplemented
}

func (w *mapReader) SetColumnsOptions(options map[string]*FieldTags) {
	panic(ErrNotImplemented.Error())
}
