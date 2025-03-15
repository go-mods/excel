package excel

import (
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"
)

// Range represent the range in the Excel file
// where data will read or write
type Range struct {
	// StartColumn is the start column of the range
	StartColumn int
	// StartRow is the start row of the range
	StartRow int
	// StartName is the start name of the range
	StartName string
	// EndColumn is the end column of the range
	EndColumn int
	// EndRow is the end row of the range
	EndRow int
	// EndName is the end name of the range
	EndName string
}

// ToRange converts a string to a Range
func ToRange(ref string) (*Range, error) {
	rng := strings.Split(strings.ReplaceAll(ref, "$", ""), ":")
	if len(rng) < 2 {
		return nil, excelize.ErrParameterInvalid
	}
	// Create the range
	r := &Range{
		StartName: rng[0],
		EndName:   rng[1],
	}
	// Get the start row and column
	startColumn, startRow, err := excelize.CellNameToCoordinates(r.StartName)
	if err != nil {
		return nil, err
	}
	r.StartRow = startRow
	r.StartColumn = startColumn
	// Get the end row and column
	endColumn, endRow, err := excelize.CellNameToCoordinates(r.EndName)
	if err != nil {
		return nil, err
	}
	r.EndRow = endRow
	r.EndColumn = endColumn

	return r, nil
}

// MinRange returns the minimum range
func MinRange(startName string) (*Range, error) {
	return ToRange(fmt.Sprintf("%s:%s", startName, startName))
}

// ToRef converts a Range to a string
func (r *Range) ToRef() string {
	return r.StartName + ":" + r.EndName
}

// UpdateNames updates the name of the range
func (r *Range) UpdateNames() error {
	startName, err := excelize.CoordinatesToCellName(r.StartColumn, r.StartRow)
	if err != nil {
		return err
	}
	r.StartName = startName
	endName, err := excelize.CoordinatesToCellName(r.EndColumn, r.EndRow)
	if err != nil {
		return err
	}
	r.EndName = endName
	return nil
}

// Rows returns the number of rows in the range
func (r *Range) Rows() int {
	return r.EndRow - r.StartRow + 1
}

// AddRows adds rows to the range
func (r *Range) AddRows(rows int) error {
	r.EndRow += rows
	return r.UpdateNames()
}

// RemoveRows removes rows from the range
func (r *Range) RemoveRows(rows int) error {
	r.EndRow -= rows
	return r.UpdateNames()
}

// SetRows sets the number of rows in the range
func (r *Range) SetRows(rows int) error {
	nbRows := r.EndRow - r.StartRow + 1
	if nbRows < rows {
		return r.AddRows(rows - nbRows)
	} else if nbRows > rows {
		return r.RemoveRows(nbRows - rows)
	}
	return nil
}

// RowAsRange returns the range of the row
func (r *Range) RowAsRange(row int) (*Range, error) {
	if row < r.StartRow || row > r.EndRow {
		return nil, excelize.ErrParameterInvalid
	}
	rRange := Range{
		StartColumn: r.StartColumn,
		StartRow:    row,
		EndColumn:   r.EndColumn,
		EndRow:      row,
	}

	return &rRange, rRange.UpdateNames()
}

// FirstRowAsRange returns the range of the first row
func (r *Range) FirstRowAsRange() (*Range, error) {
	return r.RowAsRange(r.StartRow)
}

// LastRowAsRange returns the range of the last row
func (r *Range) LastRowAsRange() (*Range, error) {
	return r.RowAsRange(r.EndRow)
}

// Columns returns the number of columns in the range
func (r *Range) Columns() int {
	return r.EndColumn - r.StartColumn + 1
}

// AddColumns adds columns to the range
func (r *Range) AddColumns(columns int) error {
	r.EndColumn += columns
	return r.UpdateNames()
}

// RemoveColumns removes columns from the range
func (r *Range) RemoveColumns(columns int) error {
	r.EndColumn -= columns
	return r.UpdateNames()
}

// SetColumns sets the number of columns in the range
func (r *Range) SetColumns(columns int) error {
	nbColumns := r.EndColumn - r.StartColumn + 1
	if nbColumns < columns {
		return r.AddColumns(columns - nbColumns)
	} else if nbColumns > columns {
		return r.RemoveColumns(nbColumns - columns)
	}
	return nil
}

// ColumnAsRange returns the range of the column
func (r *Range) ColumnAsRange(column int) (*Range, error) {
	if column < r.StartColumn || column > r.EndColumn {
		return nil, excelize.ErrParameterInvalid
	}
	rRange := Range{
		StartColumn: column,
		StartRow:    r.StartRow,
		EndColumn:   column,
		EndRow:      r.EndRow,
	}

	return &rRange, rRange.UpdateNames()
}

// FirstColumnAsRange returns the range of the first column
func (r *Range) FirstColumnAsRange() (*Range, error) {
	return r.ColumnAsRange(r.StartColumn)
}

// LastColumnAsRange returns the range of the last column
func (r *Range) LastColumnAsRange() (*Range, error) {
	return r.ColumnAsRange(r.EndColumn)
}
