package excel

type mapWriter struct {
}

func newMapWriter(_ *WriterConfig) (*mapWriter, error) {
	return nil, errNotImplemented
}

func (w *mapWriter) Marshall(data any) error {
	return errNotImplemented
}
