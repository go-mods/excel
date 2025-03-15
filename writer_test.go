package excel

import (
	"reflect"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/xuri/excelize/v2"
)

func TestWriter_newWriter(t *testing.T) {

	w := &Writer{}

	tests := []struct {
		name    string
		args    any
		want    IWriter
		wantErr bool
	}{
		{
			name:    "struct writer",
			args:    []MyStruct{},
			want:    &StructWriter{},
			wantErr: true,
		},
		{
			name: "struct writer pointer",
			args: &[]MyStruct{},
			want: &StructWriter{},
		},
		{
			name:    "slice writer (string)",
			args:    MyStringSlice{},
			want:    &SliceWriter{},
			wantErr: true,
		},
		{
			name: "slice writer pointer (string)",
			args: &MyStringSlice{},
			want: &SliceWriter{},
		},
		{
			name:    "slice writer (int)",
			args:    MyIntSlice{},
			want:    &SliceWriter{},
			wantErr: true,
		},
		{
			name: "slice writer pointer (int)",
			args: &MyIntSlice{},
			want: &SliceWriter{},
		},
		{
			name:    "slice writer (any)",
			args:    MyAnySlice{},
			want:    &SliceWriter{},
			wantErr: true,
		},
		{
			name: "slice writer pointer (any)",
			args: &MyAnySlice{},
			want: &SliceWriter{},
		},
		{
			name:    "map writer (string)",
			args:    MyMapString{},
			want:    &MapWriter{},
			wantErr: true,
		},
		{
			name: "map writer pointer (string)",
			args: &MyMapString{},
			want: &MapWriter{},
		},
		{
			name:    "map writer (int)",
			args:    MyMapInt{},
			want:    &MapWriter{},
			wantErr: true,
		},
		{
			name: "map writer pointer (int)",
			args: &MyMapInt{},
			want: &MapWriter{},
		},
		{
			name:    "map writer (any)",
			args:    MyMapAny{},
			want:    &MapWriter{},
			wantErr: true,
		},
		{
			name: "map writer pointer (any)",
			args: &MyMapAny{},
			want: &MapWriter{},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := w.newWriter(tt.args)
			if (err != nil) != tt.wantErr {
				t.Errorf("newWriter() error = %v, wantErr %v", err, tt.wantErr)
				return
			}
			if err == nil {
				if got == nil {
					t.Errorf("newWriter() got = %v, want %v", got, tt.want)
				}
				if reflect.TypeOf(got) != reflect.TypeOf(tt.want) {
					t.Errorf("newWriter() got = %v, want %v", got, tt.want)
				}
			}
		})
	}

}

func TestSliceWriter_string(t *testing.T) {

	var stringSlice MyStringSlice
	stringSlice = append(stringSlice, []string{"ID", "Name"})
	stringSlice = append(stringSlice, []string{"1", "John Doe"})

	file := excelize.NewFile()
	defer func() { _ = file.Close() }()

	xl, _ := NewWriter(file)
	xl.SetSheet(xl.GetActiveSheet())
	xl.SetAxis("A1")

	err := xl.Marshal(&stringSlice)
	if err != nil {
		t.Error(err)
		return
	}

	sA1, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A1")
	sB1, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B1")
	sA2, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A2")
	sB2, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B2")

	assert.Equal(t, "ID", sA1, "they should be equal")
	assert.Equal(t, "Name", sB1, "they should be equal")
	assert.Equal(t, "1", sA2, "they should be equal")
	assert.Equal(t, "John Doe", sB2, "they should be equal")

	assert.Equal(t, xl.Writer.Result.Rows, 2, "they should be equal")
	assert.Equal(t, xl.Writer.Result.Columns, 2, "they should be equal")
}

func TestSliceWriter_int(t *testing.T) {

	var intSlice MyIntSlice
	intSlice = append(intSlice, []int{1, 2})
	intSlice = append(intSlice, []int{3, 4})

	file := excelize.NewFile()
	defer func() { _ = file.Close() }()

	xl, _ := NewWriter(file)
	xl.SetSheet(xl.GetActiveSheet())
	xl.SetAxis("A1")

	err := xl.Marshal(&intSlice)
	if err != nil {
		t.Error(err)
		return
	}

	sA1, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A1")
	sB1, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B1")
	sA2, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A2")
	sB2, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B2")

	assert.Equal(t, "1", sA1, "they should be equal")
	assert.Equal(t, "2", sB1, "they should be equal")
	assert.Equal(t, "3", sA2, "they should be equal")
	assert.Equal(t, "4", sB2, "they should be equal")

	assert.Equal(t, xl.Writer.Result.Rows, 2, "they should be equal")
	assert.Equal(t, xl.Writer.Result.Columns, 2, "they should be equal")
}

func TestSliceWriter_any(t *testing.T) {

	var anySlice MyAnySlice
	anySlice = append(anySlice, []interface{}{"ID", "Name"})
	anySlice = append(anySlice, []interface{}{1, "John Doe"})
	anySlice = append(anySlice, []interface{}{2, "Jane Doe"})

	file := excelize.NewFile()
	defer func() { _ = file.Close() }()

	xl, _ := NewWriter(file)
	xl.SetSheet(xl.GetActiveSheet())
	xl.SetAxis("A1")

	err := xl.Marshal(&anySlice)
	if err != nil {
		t.Error(err)
		return
	}

	sA1, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A1")
	sB1, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B1")
	sA2, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A2")
	sB2, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B2")

	assert.Equal(t, "ID", sA1, "they should be equal")
	assert.Equal(t, "Name", sB1, "they should be equal")
	assert.Equal(t, "1", sA2, "they should be equal")
	assert.Equal(t, "John Doe", sB2, "they should be equal")

	assert.Equal(t, xl.Writer.Result.Rows, 3, "they should be equal")
	assert.Equal(t, xl.Writer.Result.Columns, 2, "they should be equal")
}

func TestMapWrite_string(t *testing.T) {

	var mapString MyMapString
	mapString = append(mapString, map[string]string{"ID": "1", "Name": "John Doe"})
	mapString = append(mapString, map[string]string{"ID": "2", "Name": "Jane Doe"})

	file := excelize.NewFile()
	defer func() { _ = file.Close() }()

	xl, _ := NewWriter(file)
	xl.SetSheet(xl.GetActiveSheet())
	xl.SetAxis("A1")

	err := xl.Marshal(&mapString)
	if err != nil {
		t.Error(err)
		return
	}

	sA1, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A1")
	sB1, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B1")
	sA2, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A2")
	sB2, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B2")
	sA3, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A3")
	sB3, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B3")

	assert.Equal(t, sA1, "ID", "they should be equal")
	assert.Equal(t, sB1, "Name", "they should be equal")
	assert.Equal(t, sA2, "1", "they should be equal")
	assert.Equal(t, sB2, "John Doe", "they should be equal")
	assert.Equal(t, sA3, "2", "they should be equal")
	assert.Equal(t, sB3, "Jane Doe", "they should be equal")
}

func TestMapWrite_int(t *testing.T) {

	var mapInt MyMapInt
	mapInt = append(mapInt, map[string]int{"ID1": 1, "ID2": 2})
	mapInt = append(mapInt, map[string]int{"ID1": 3, "ID2": 4})

	file := excelize.NewFile()
	defer func() { _ = file.Close() }()

	xl, _ := NewWriter(file)
	xl.SetSheet(xl.GetActiveSheet())
	xl.SetAxis("A1")

	err := xl.Marshal(&mapInt)
	if err != nil {
		t.Error(err)
		return
	}

	sA1, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A1")
	sB1, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B1")
	sA2, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A2")
	sB2, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B2")
	sA3, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A3")
	sB3, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B3")

	assert.Equal(t, sA1, "ID1", "they should be equal")
	assert.Equal(t, sB1, "ID2", "they should be equal")
	assert.Equal(t, sA2, "1", "they should be equal")
	assert.Equal(t, sB2, "2", "they should be equal")
	assert.Equal(t, sA3, "3", "they should be equal")
	assert.Equal(t, sB3, "4", "they should be equal")
}

func TestMapWrite_any(t *testing.T) {

	var mapAny MyMapAny
	mapAny = append(mapAny, map[string]interface{}{"ID": int64(1), "Name": "John Doe"})
	mapAny = append(mapAny, map[string]interface{}{"ID": int64(2), "Name": "Jane Doe"})

	file := excelize.NewFile()
	defer func() { _ = file.Close() }()

	xl, _ := NewWriter(file)
	xl.SetSheet(xl.GetActiveSheet())
	xl.SetAxis("A1")

	err := xl.Marshal(&mapAny)
	if err != nil {
		t.Error(err)
		return
	}

	sA1, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A1")
	sB1, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B1")
	sA2, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A2")
	sB2, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B2")
	sA3, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A3")
	sB3, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B3")

	assert.Equal(t, sA1, "ID", "they should be equal")
	assert.Equal(t, sB1, "Name", "they should be equal")
	assert.Equal(t, sA2, "1", "they should be equal")
	assert.Equal(t, sB2, "John Doe", "they should be equal")
	assert.Equal(t, sA3, "2", "they should be equal")
	assert.Equal(t, sB3, "Jane Doe", "they should be equal")
}

func TestMapWrite_string_interface(t *testing.T) {

	var mapAny []map[string]interface{}

	m1 := make(map[string]interface{})
	m1["ID"] = int64(1)
	m1["Name"] = "John Doe"

	m2 := make(map[string]interface{})
	m2["ID"] = int64(2)
	m2["Name"] = "Jane Doe"

	mapAny = append(mapAny, m1)
	mapAny = append(mapAny, m2)

	file := excelize.NewFile()
	defer func() { _ = file.Close() }()

	xl, _ := NewWriter(file)
	xl.SetSheet(xl.GetActiveSheet())
	xl.SetAxis("A1")

	err := xl.Marshal(&mapAny)
	if err != nil {
		t.Error(err)
		return
	}

	sA1, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A1")
	assert.Equal(t, sA1, "ID", "they should be equal")

	sB1, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B1")
	assert.Equal(t, sB1, "Name", "they should be equal")

	sA2, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A2")
	assert.Equal(t, sA2, "1", "they should be equal")

	sB2, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B2")
	assert.Equal(t, sB2, "John Doe", "they should be equal")

	sA3, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A3")
	assert.Equal(t, sA3, "2", "they should be equal")

	sB3, _ := file.GetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B3")
	assert.Equal(t, sB3, "Jane Doe", "they should be equal")
}
