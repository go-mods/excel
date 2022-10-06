package excel

import "github.com/xuri/excelize/v2"

// Sheet represent the sheet in the
// Excel file where data will read or write
type Sheet struct {
	Name  string
	Index int
}

func (r *ReaderInfo) SetSheetName(n string) {
	setSheetName(r.file, &r.Sheet, n)
}

func (r *ReaderInfo) SetSheetIndex(i int) {
	setSheetIndex(r.file, &r.Sheet, i)
}

func (r *ReaderInfo) GetSheetName() string {
	return getSheetName(r.file, &r.Sheet)
}

func (r *ReaderInfo) GetSheetIndex() int {
	return getSheetIndex(r.file, &r.Sheet)
}

func (r *ReaderInfo) isSheetValid() bool {
	return isSheetValid(&r.Sheet)
}

func (w *WriterInfo) SetSheetName(n string) {
	setSheetName(w.file, &w.Sheet, n)
}

func (w *WriterInfo) SetSheetIndex(i int) {
	setSheetIndex(w.file, &w.Sheet, i)
}

func (w *WriterInfo) GetSheetName() string {
	return getSheetName(w.file, &w.Sheet)
}

func (w *WriterInfo) GetSheetIndex() int {
	return getSheetIndex(w.file, &w.Sheet)
}

func (w *WriterInfo) isSheetValid() bool {
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

func getSheetName(file *excelize.File, sheet *Sheet) string {
	return file.GetSheetName(sheet.Index)
}

func getSheetIndex(file *excelize.File, sheet *Sheet) int {
	return file.GetSheetIndex(sheet.Name)
}

func isSheetValid(sheet *Sheet) bool {
	if len(sheet.Name) > 0 && sheet.Index >= 0 {
		return true
	}
	return false
}
