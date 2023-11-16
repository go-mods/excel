package excel

import (
	"errors"
	"github.com/xuri/excelize/v2"
)

// Table represent the table in the Excel file
type Table struct {
	Sheet *Sheet
	*excelize.Table
}

// GetTables returns the tables in the Excel file
func (e *Excel) GetTables() ([]Table, error) {
	if e.File == nil {
		return nil, ErrFileIsNil
	}

	var result []Table

	// Loop through the sheets
	for _, sheetName := range e.File.GetSheetList() {
		// Get the tables in the sheet
		tables, err := e.File.GetTables(sheetName)
		if err != nil {
			return nil, err
		}
		// Convert the tables to Table
		for _, t := range tables {
			t := t
			result = append(result, Table{
				Sheet: e.GetSheet(sheetName),
				Table: &t,
			})
		}
	}

	return result, nil
}

// GetTable returns the table in the Excel file
func (e *Excel) GetTable(name string) (*Table, error) {
	if e.File == nil {
		return nil, ErrFileIsNil
	}
	if name == "" {
		return nil, ErrTableNameEmpty
	}

	// Loop through the sheets
	for _, sheetName := range e.File.GetSheetList() {

		// Get the tables in the sheet
		tables, err := e.File.GetTables(sheetName)
		if err != nil {
			return nil, err
		}
		// Convert the tables to Table
		for _, t := range tables {
			t := t
			if t.Name == name {
				return &Table{
					Sheet: e.GetSheet(sheetName),
					Table: &t,
				}, nil
			}
		}
	}

	return nil, errors.New("table not found")
}

// GetTableSheet returns the sheet where the table is located
func (e *Excel) GetTableSheet(name string) (*Sheet, error) {
	table, err := e.GetTable(name)
	if err != nil {
		return nil, err
	}
	return table.Sheet, nil
}

// AddTable adds a table to the Excel file
func (e *Excel) AddTable(table *Table) error {
	// File must be set
	if e.File == nil {
		return ErrFileIsNil
	}
	// Table must be set
	if table == nil || table.Table == nil {
		return errors.New("table is nil")
	}
	if err := table.IsValidError(); err != nil {
		return err
	}
	return e.File.AddTable(table.Sheet.Name, table.Table)
}

// DeleteTable deletes the table in the Excel file
func (e *Excel) DeleteTable(name string) error {
	if e.File == nil {
		return ErrFileIsNil
	}
	if name == "" {
		return ErrTableNameEmpty
	}
	return e.File.DeleteTable(name)
}

// DeleteTableContent deletes the content of the table in the Excel file
func (e *Excel) DeleteTableContent(name string) error {
	table, err := e.GetTable(name)
	if err != nil {
		return err
	}
	return table.DeleteContent()
}

// ResizeTable resize the table by changing the range
func (e *Excel) ResizeTable(table *Table, newRange string) error {
	return table.Resize(newRange)
}

// IsValidError returns an error if the table is not valid
func (t *Table) IsValidError() error {
	// Table must have a name
	if t.Name == "" {
		return ErrTableNameEmpty
	}
	// Table must have a range
	if t.Range == "" {
		return ErrTableRange
	}
	// Sheet must be set
	if t.Sheet == nil {
		return ErrSheetNotValid
	}
	// Sheet must have a name
	if t.Sheet.Name == "" {
		return ErrSheetNameEmpty
	}
	return nil
}

// IsValid returns true if the table is valid
func (t *Table) IsValid() bool {
	return t.IsValidError() == nil
}

// Delete the table
func (t *Table) Delete() error {
	if err := t.IsValidError(); err != nil {
		return err
	}
	return t.Sheet.file.DeleteTable(t.Name)
}

// Resize the table by changing the range
func (t *Table) Resize(newRange string) error {
	if err := t.IsValidError(); err != nil {
		return err
	}
	if newRange == "" {
		return ErrTableRange
	}

	// Resize the table
	if err := t.Sheet.file.ResizeTable(t.Name, newRange); err != nil {
		return err
	}

	// Update the range
	t.Range = newRange

	// Loop through all cells in the table
	tr, err := ToRange(t.Range)
	if err != nil {
		return err
	}

	dv, _ := t.Sheet.file.GetDataValidations(t.Sheet.Name)
	_ = dv

	for col := tr.StartColumn; col <= tr.EndColumn; col++ {
		firstCoord, _ := excelize.CoordinatesToCellName(col, tr.StartRow+1)
		// Comments
		comment := t.Sheet.GetComment(firstCoord)

		for row := tr.StartRow + 2; row <= tr.EndRow; row++ {
			curCoord, _ := excelize.CoordinatesToCellName(col, row)
			// Comments
			if comment != nil {
				comment.Cell = curCoord
				_ = t.Sheet.file.AddComment(t.Sheet.Name, *comment)
			}
		}
	}

	return nil
}

// DeleteContent deletes the content of the table
func (t *Table) DeleteContent() error {
	if err := t.IsValidError(); err != nil {
		return err
	}

	// The table range
	tRange, err := ToRange(t.Range)
	if err != nil {
		return err
	}

	// Removes cells values from the table
	// except the first row (the header)
	for col := tRange.StartColumn; col <= tRange.EndColumn; col++ {
		for row := tRange.StartRow + 1; row <= tRange.EndRow; row++ {
			cell, _ := excelize.CoordinatesToCellName(col, row)
			_ = t.Sheet.file.SetCellValue(t.Sheet.Name, cell, nil)
		}
	}

	return nil
}

// GetRange returns the range of the table
func (t *Table) GetRange() (*Range, error) {
	return ToRange(t.Range)
}

// GetHeaderRange returns the range of the header of the table
func (t *Table) GetHeaderRange() (*Range, error) {
	if err := t.IsValidError(); err != nil {
		return nil, err
	}
	// The table range
	tRange, err := ToRange(t.Range)
	if err != nil {
		return nil, err
	}
	// The header range
	hRange := tRange
	hRange.EndRow = hRange.StartRow
	return hRange, hRange.UpdateNames()
}

// GetDataRange returns the range of the data of the table
func (t *Table) GetDataRange() (*Range, error) {
	if err := t.IsValidError(); err != nil {
		return nil, err
	}
	// The table range
	tRange, err := ToRange(t.Range)
	if err != nil {
		return nil, err
	}
	// The data range
	dRange := tRange
	dRange.StartRow++
	return dRange, dRange.UpdateNames()
}

// GetColumn returns the column index of the title
func (t *Table) GetColumn(title string) (int, error) {
	hr, err := t.GetHeaderRange()
	if err != nil {
		return 0, err
	}
	for col := hr.StartColumn; col <= hr.EndColumn; col++ {
		cellRef, err := excelize.CoordinatesToCellName(col, hr.StartRow)
		if err != nil {
			return 0, err
		}
		cell, err := t.Sheet.file.GetCellValue(t.Sheet.Name, cellRef)
		if err != nil {
			return 0, err
		}
		if cell == title {
			return col, nil
		}
	}
	return 0, nil
}

// GetColumnAt returns the column name at the desired index
func (t *Table) GetColumnAt(index int) (string, error) {
	hr, err := t.GetHeaderRange()
	if err != nil {
		return "", err
	}
	ir, err := hr.ColumnAsRange(index)
	if err != nil {
		return "", err
	}
	return t.Sheet.file.GetCellValue(t.Sheet.Name, ir.StartName)
}
