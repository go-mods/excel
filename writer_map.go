package excel

import "reflect"

type mapWriter struct {
}

func newMapWriter(_ *Writer, _ reflect.Value) (*mapWriter, error) {
	return nil, ErrNotImplemented
}

func (w *mapWriter) Marshall(_ any) error {
	return ErrNotImplemented
}

func (w *mapWriter) SetColumnsTags(_ map[string]*Tags) {
	panic(ErrNotImplemented.Error())
}
