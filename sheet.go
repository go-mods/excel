package excel

import "github.com/xuri/excelize/v2"

// Sheet represent the sheet in the
// Excel file where data will read or write
type Sheet struct {
	Name  string
	Index int
}

// SetSheetName sets the sheet name to be used by the reader or writer
func (e *Excel) SetSheetName(sheet string) {
	if e.Reader != nil {
		e.Reader.setSheetName(sheet)
	}
	if e.Writer != nil {
		e.Writer.setSheetName(sheet)
	}
}

// GetSheetName gets the sheet name used by the reader or writer
func (e *Excel) GetSheetName() string {
	if e.Reader != nil {
		return e.Reader.getSheetName()
	}
	if e.Writer != nil {
		return e.Writer.getSheetName()
	}
	return ""
}

// SetSheetIndex sets the sheet index to be used by the reader or writer
func (e *Excel) SetSheetIndex(index int) {
	if e.Reader != nil {
		e.Reader.setSheetIndex(index)
	}
	if e.Writer != nil {
		e.Writer.setSheetIndex(index)
	}
}

// GetSheetIndex gets the sheet index used by the reader or writer
func (e *Excel) GetSheetIndex() int {
	if e.Reader != nil {
		i, err := e.Reader.getSheetIndex()
		if err != nil {
			return 0
		}
		return i
	}
	if e.Writer != nil {
		i, err := e.Writer.getSheetIndex()
		if err != nil {
			return 0
		}
		return i
	}
	return 0
}

func (r *Reader) setSheetName(n string) {
	_ = setSheetName(r.file, &r.Sheet, n)
}

func (r *Reader) getSheetName() string {
	return getSheetName(r.file, &r.Sheet)
}

func (r *Reader) setSheetIndex(i int) {
	setSheetIndex(r.file, &r.Sheet, i)
}

func (r *Reader) getSheetIndex() (int, error) {
	return getSheetIndex(r.file, &r.Sheet)
}

func (r *Reader) isSheetValid() bool {
	return isSheetValid(&r.Sheet)
}

func (w *Writer) setSheetName(n string) {
	_ = setSheetName(w.file, &w.Sheet, n)
}

func (w *Writer) getSheetName() string {
	return getSheetName(w.file, &w.Sheet)
}

func (w *Writer) setSheetIndex(i int) {
	setSheetIndex(w.file, &w.Sheet, i)
}

func (w *Writer) getSheetIndex() (int, error) {
	return getSheetIndex(w.file, &w.Sheet)
}

func (w *Writer) isSheetValid() bool {
	return isSheetValid(&w.Sheet)
}

func setSheetName(file *excelize.File, sheet *Sheet, name string) error {
	index, err := file.GetSheetIndex(name)
	if err != nil {
		return err
	}
	sheet.Index = index
	sheet.Name = file.GetSheetName(sheet.Index)
	return nil
}

func setSheetIndex(file *excelize.File, sheet *Sheet, i int) {
	sheet.Name = file.GetSheetName(i)
	sheet.Index, _ = file.GetSheetIndex(sheet.Name)
}

func getSheetName(file *excelize.File, sheet *Sheet) string {
	return file.GetSheetName(sheet.Index)
}

func getSheetIndex(file *excelize.File, sheet *Sheet) (int, error) {
	return file.GetSheetIndex(sheet.Name)
}

func isSheetValid(sheet *Sheet) bool {
	if len(sheet.Name) > 0 && sheet.Index >= 0 {
		return true
	}
	return false
}
