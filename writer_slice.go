package excel

type sliceWriter struct {
}

func newSliceWriter(_ *WriterConfig) (*sliceWriter, error) {
	return nil, errNotImplemented
}

func (w *sliceWriter) Marshall(data any) error {
	return errNotImplemented
}
