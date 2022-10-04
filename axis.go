package excel

import "github.com/xuri/excelize/v2"

type Axis struct {
	Axis string
	Col  int
	Row  int
}

func (r *ReaderConfig) SetAxis(axis string) {
	setAxis(&r.Axis, axis)
}

func (r *ReaderConfig) SetAxisCoordinates(col int, row int) {
	setAxisCoordinates(&r.Axis, col, row)
}

func (r *ReaderConfig) isAxisValid() bool {
	return isAxisValid(&r.Axis)
}

func (w *WriterConfig) SetAxis(axis string) {
	setAxis(&w.Axis, axis)
}

func (w *WriterConfig) SetAxisCoordinates(col int, row int) {
	setAxisCoordinates(&w.Axis, col, row)
}

func (w *WriterConfig) isAxisValid() bool {
	return isAxisValid(&w.Axis)
}

func setAxis(axis *Axis, a string) {
	col, row, err := excelize.CellNameToCoordinates(a)
	if err == nil {
		axis.Axis = a
		axis.Col = col
		axis.Row = row
	} else {
		axis.Axis = ""
		axis.Col = -1
		axis.Row = -1
	}
}

func setAxisCoordinates(axis *Axis, col int, row int) {
	a, err := excelize.CoordinatesToCellName(col, row)
	if err == nil {
		axis.Axis = a
		axis.Col = col
		axis.Row = row
	} else {
		axis.Axis = ""
		axis.Col = -1
		axis.Row = -1
	}
}

func isAxisValid(axis *Axis) bool {
	if len(axis.Axis) > 0 && axis.Col > 0 && axis.Row > 0 {
		return true
	}
	return false
}
