package excel

import "github.com/xuri/excelize/v2"

// Axis represent the coordinates in the
// Excel file where data will read or write
type Axis struct {
	Axis string
	Col  int
	Row  int
}

// SetAxis sets the axis to be used by the reader or writer
func (e *Excel) SetAxis(axis string) {
	if e.Reader != nil {
		e.Reader.setAxis(axis)
	}
	if e.Writer != nil {
		e.Writer.setAxis(axis)
	}
}

// SetAxisCoordinates sets the axis coordinates to be used by the reader or writer
func (e *Excel) SetAxisCoordinates(col int, row int) {
	if e.Reader != nil {
		e.Reader.setAxisCoordinates(col, row)
	}
	if e.Writer != nil {
		e.Writer.setAxisCoordinates(col, row)
	}
}

func (r *Reader) setAxis(axis string) {
	setAxis(&r.Axis, axis)
}

func (r *Reader) setAxisCoordinates(col int, row int) {
	setAxisCoordinates(&r.Axis, col, row)
}

func (r *Reader) isAxisValid() bool {
	return isAxisValid(&r.Axis)
}

func (w *Writer) setAxis(axis string) {
	setAxis(&w.Axis, axis)
}

func (w *Writer) setAxisCoordinates(col int, row int) {
	setAxisCoordinates(&w.Axis, col, row)
}

func (w *Writer) isAxisValid() bool {
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
