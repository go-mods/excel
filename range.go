package excel

import (
	"github.com/xuri/excelize/v2"
	"strings"
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

// ToRef converts a Range to a string
func (r *Range) ToRef() string {
	return r.StartName + ":" + r.EndName
}

// Rows returns the number of rows in the range
func (r *Range) Rows() int {
	return r.EndRow - r.StartRow + 1
}

// AddRows adds rows to the range
func (r *Range) AddRows(rows int) error {
	r.EndRow += rows
	endName, err := excelize.CoordinatesToCellName(r.EndColumn, r.EndRow)
	if err != nil {
		return err
	}
	r.EndName = endName
	return nil
}

// RemoveRows removes rows from the range
func (r *Range) RemoveRows(rows int) error {
	r.EndRow -= rows
	endName, err := excelize.CoordinatesToCellName(r.EndColumn, r.EndRow)
	if err != nil {
		return err
	}
	r.EndName = endName
	return nil
}

// Columns returns the number of columns in the range
func (r *Range) Columns() int {
	return r.EndColumn - r.StartColumn + 1
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

// AddColumns adds columns to the range
func (r *Range) AddColumns(columns int) error {
	r.EndColumn += columns
	endName, err := excelize.CoordinatesToCellName(r.EndColumn, r.EndRow)
	if err != nil {
		return err
	}
	r.EndName = endName
	return nil
}

// RemoveColumns removes columns from the range
func (r *Range) RemoveColumns(columns int) error {
	r.EndColumn -= columns
	endName, err := excelize.CoordinatesToCellName(r.EndColumn, r.EndRow)
	if err != nil {
		return err
	}
	r.EndName = endName
	return nil
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
