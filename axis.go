package excel

import "github.com/xuri/excelize/v2"

type Axis struct {
	Axis string
	Col  int
	Row  int
}

func (r *ReaderConfig) SetAxis(axis string) {
	col, row, err := excelize.CellNameToCoordinates(axis)
	if err == nil {
		r.Axis.Axis = axis
		r.Axis.Col = col
		r.Axis.Row = row
	} else {
		r.Axis.Axis = ""
		r.Axis.Col = -1
		r.Axis.Row = -1
	}
}

func (r *ReaderConfig) SetAxisCoordinates(col int, row int) {
	axis, err := excelize.CoordinatesToCellName(col, row)
	if err == nil {
		r.Axis.Axis = axis
		r.Axis.Col = col
		r.Axis.Row = row
	} else {
		r.Axis.Axis = ""
		r.Axis.Col = -1
		r.Axis.Row = -1
	}
}

func (r *ReaderConfig) isAxisValid() bool {
	if len(r.Axis.Axis) > 0 && r.Axis.Col > 0 && r.Axis.Row > 0 {
		return true
	}
	return false
}
