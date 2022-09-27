package excel

type Sheet struct {
	Name  string
	Index int
}

func (r *ReaderConfig) SetSheetName(n string) {
	r.Sheet.Index = r.file.GetSheetIndex(n)
	r.Sheet.Name = r.file.GetSheetName(r.Sheet.Index)
}

func (r *ReaderConfig) SetSheetIndex(i int) {
	r.Sheet.Name = r.file.GetSheetName(i)
	r.Sheet.Index = r.file.GetSheetIndex(r.Sheet.Name)
}

func (r *ReaderConfig) isSheetValid() bool {
	if len(r.Sheet.Name) > 0 && r.Sheet.Index >= 0 {
		return true
	}
	return false
}
