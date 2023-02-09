package excel

import "reflect"

type sliceReader struct {
}

func newSliceReader(_ *ReaderInfo, _ reflect.Value) (*sliceReader, error) {
	return nil, ErrNotImplemented
}

func (r *sliceReader) Unmarshall() error {
	return ErrNotImplemented
}

func (w *sliceReader) SetColumnsOptions(options map[string]*FieldTags) {
	panic(ErrNotImplemented.Error())
}
