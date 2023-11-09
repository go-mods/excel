package excel

import "github.com/xuri/excelize/v2"

// Sheet represent the sheet in the
// Excel file where data will read or write
type Sheet struct {
	file  *excelize.File
	Name  string
	Index int
}

// Sheet returns the sheet object used by the reader or writer
func (e *Excel) Sheet() *Sheet {
	if e.Reader != nil {
		return &e.Reader.Sheet
	}
	if e.Writer != nil {
		return &e.Writer.Sheet
	}
	return nil
}

// GetSheet returns the sheet object
func (e *Excel) GetSheet(name string) *Sheet {
	if e.File == nil {
		return nil
	}
	sheet := &Sheet{
		file: e.File,
	}
	// Get the sheet name
	sheet.Name = name
	// Get the sheet index
	index, err := e.File.GetSheetIndex(name)
	if err != nil {
		return nil
	}
	sheet.Index = index
	// Return the sheet
	return sheet
}

// GetSheetFromIndex returns the sheet object
func (e *Excel) GetSheetFromIndex(index int) *Sheet {
	if e.File == nil {
		return nil
	}
	sheet := &Sheet{
		file: e.File,
	}
	// Get the sheet name
	sheet.Name = e.File.GetSheetName(index)
	// Get the sheet index
	sheet.Index = index
	// Return the sheet
	return sheet
}

// GetActiveSheet returns the active sheet
func (e *Excel) GetActiveSheet() *Sheet {
	if e.File == nil {
		return nil
	}
	sheet := &Sheet{
		file: e.File,
	}
	// Get the sheet index
	sheet.Index = e.File.GetActiveSheetIndex()
	// Get the sheet name
	sheet.Name = e.File.GetSheetName(sheet.Index)
	// Return the sheet
	return sheet
}

// SetActiveSheet sets the active sheet
func (e *Excel) SetActiveSheet(sheet *Sheet) {
	if e.File == nil {
		return
	}
	// Set the active sheet
	e.File.SetActiveSheet(sheet.Index)
	// Set the sheet to be used by the reader or writer
	e.SetSheet(sheet)
}

// SetSheet sets the sheet to be used by the reader or writer
func (e *Excel) SetSheet(sheet *Sheet) {
	if e.Reader != nil {
		e.Reader.Sheet = *sheet
	}
	if e.Writer != nil {
		e.Writer.Sheet = *sheet
	}
}

// SetSheetFromName sets the sheet name to be used by the reader or writer
func (e *Excel) SetSheetFromName(name string) {
	e.SetSheet(e.GetSheet(name))
}

// SetSheetFromIndex sets the sheet index to be used by the reader or writer
func (e *Excel) SetSheetFromIndex(index int) {
	e.SetSheet(e.GetSheetFromIndex(index))
}

// IsValid returns true if the sheet is valid
func (s *Sheet) IsValid() bool {
	if s.file != nil && len(s.Name) > 0 && s.Index >= 0 {
		return true
	}
	return false
}

// GetComment returns the comment of the cell
func (s *Sheet) GetComment(cell string) *excelize.Comment {
	if !s.IsValid() {
		return nil
	}
	comments, _ := s.file.GetComments(s.Name)
	for _, c := range comments {
		if c.Cell == cell {
			return &c
		}
	}
	return nil
}
