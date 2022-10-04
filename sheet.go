package excel

import "github.com/xuri/excelize/v2"

type Sheet struct {
	Name  string
	Index int
}

func (r *ReaderConfig) SetSheetName(n string) {
	setSheetName(r.file, &r.Sheet, n)
}

func (r *ReaderConfig) SetSheetIndex(i int) {
	setSheetIndex(r.file, &r.Sheet, i)
}

func (r *ReaderConfig) isSheetValid() bool {
	return isSheetValid(&r.Sheet)
}

func (w *WriterConfig) SetSheetName(n string) {
	setSheetName(w.file, &w.Sheet, n)
}

func (w *WriterConfig) SetSheetIndex(i int) {
	setSheetIndex(w.file, &w.Sheet, i)
}

func (w *WriterConfig) isSheetValid() bool {
	return isSheetValid(&w.Sheet)
}

func setSheetName(file *excelize.File, sheet *Sheet, n string) {
	sheet.Index = file.GetSheetIndex(n)
	sheet.Name = file.GetSheetName(sheet.Index)
}

func setSheetIndex(file *excelize.File, sheet *Sheet, i int) {
	sheet.Name = file.GetSheetName(i)
	sheet.Index = file.GetSheetIndex(sheet.Name)
}

func isSheetValid(sheet *Sheet) bool {
	if len(sheet.Name) > 0 && sheet.Index >= 0 {
		return true
	}
	return false
}
