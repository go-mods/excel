package excel

import (
	"log"
	"reflect"
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
	"github.com/xuri/excelize/v2"
)

// TestNewWriter verifies that the appropriate writer type is created based on the input data type.
// It tests writer creation for structures, slices, and maps, both for pointers and direct values.
func TestNewWriter(t *testing.T) {
	w := &Writer{}

	tests := []struct {
		name    string
		args    any
		want    IWriter
		wantErr bool
	}{
		{
			name:    "struct writer",
			args:    []Named{},
			want:    &StructWriter{},
			wantErr: true,
		},
		{
			name: "struct writer pointer",
			args: &[]Named{},
			want: &StructWriter{},
		},
		{
			name:    "slice writer (string)",
			args:    StringMatrix{},
			want:    &SliceWriter{},
			wantErr: true,
		},
		{
			name: "slice writer pointer (string)",
			args: &StringMatrix{},
			want: &SliceWriter{},
		},
		{
			name:    "slice writer (int)",
			args:    IntMatrix{},
			want:    &SliceWriter{},
			wantErr: true,
		},
		{
			name: "slice writer pointer (int)",
			args: &IntMatrix{},
			want: &SliceWriter{},
		},
		{
			name:    "slice writer (any)",
			args:    AnyMatrix{},
			want:    &SliceWriter{},
			wantErr: true,
		},
		{
			name: "slice writer pointer (any)",
			args: &AnyMatrix{},
			want: &SliceWriter{},
		},
		{
			name:    "map writer (string)",
			args:    StringMap{},
			want:    &MapWriter{},
			wantErr: true,
		},
		{
			name: "map writer pointer (string)",
			args: &StringMap{},
			want: &MapWriter{},
		},
		{
			name:    "map writer (int)",
			args:    IntMap{},
			want:    &MapWriter{},
			wantErr: true,
		},
		{
			name: "map writer pointer (int)",
			args: &IntMap{},
			want: &MapWriter{},
		},
		{
			name:    "map writer (any)",
			args:    AnyMap{},
			want:    &MapWriter{},
			wantErr: true,
		},
		{
			name: "map writer pointer (any)",
			args: &AnyMap{},
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

// TestStringMatrixWrite verifies writing a string matrix to an Excel file.
// It tests:
// - Writing headers and data
// - Exact matching of written values
// - Correct number of rows and columns
func TestStringMatrixWrite(t *testing.T) {
	var stringMatrix StringMatrix
	stringMatrix = append(stringMatrix, []string{"ID", "Name"})
	stringMatrix = append(stringMatrix, []string{"1", "John Doe"})

	file := excelize.NewFile()
	defer func() { _ = file.Close() }()

	xl, _ := NewWriter(file)
	xl.SetSheet(xl.GetActiveSheet())
	xl.SetAxis("A1")

	err := xl.Marshal(&stringMatrix)
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

// TestIntMatrixWrite verifies writing an integer matrix to an Excel file.
// It tests:
// - Correct conversion of integers to strings
// - Writing to appropriate cells
// - Correct number of rows and columns
func TestIntMatrixWrite(t *testing.T) {
	var intMatrix IntMatrix
	intMatrix = append(intMatrix, []int{1, 2})
	intMatrix = append(intMatrix, []int{3, 4})

	file := excelize.NewFile()
	defer func() { _ = file.Close() }()

	xl, _ := NewWriter(file)
	xl.SetSheet(xl.GetActiveSheet())
	xl.SetAxis("A1")

	err := xl.Marshal(&intMatrix)
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

// TestAnyMatrixWrite verifies writing a matrix of mixed types to an Excel file.
// It tests:
// - Writing different data types (strings, integers)
// - Appropriate type conversion
// - Correct structure of the resulting file
func TestAnyMatrixWrite(t *testing.T) {
	var anyMatrix AnyMatrix
	anyMatrix = append(anyMatrix, []interface{}{"ID", "Name"})
	anyMatrix = append(anyMatrix, []interface{}{1, "John Doe"})
	anyMatrix = append(anyMatrix, []interface{}{2, "Jane Doe"})

	file := excelize.NewFile()
	defer func() { _ = file.Close() }()

	xl, _ := NewWriter(file)
	xl.SetSheet(xl.GetActiveSheet())
	xl.SetAxis("A1")

	err := xl.Marshal(&anyMatrix)
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

// TestStringMapWrite verifies writing a map of strings to an Excel file.
// It tests:
// - Correct writing of keys as headers
// - Writing values to appropriate cells
// - Handling multiple rows of data
func TestStringMapWrite(t *testing.T) {
	var mapString StringMap
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

// TestIntMapWrite verifies writing a map of integers to an Excel file.
// It tests:
// - Converting keys to column headers
// - Converting integer values to strings
// - Correct formatting of the resulting Excel file
func TestIntMapWrite(t *testing.T) {
	var mapInt IntMap
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

// TestAnyMapWrite verifies writing a map of mixed types to an Excel file.
// It tests:
// - Handling different data types
// - Appropriate type conversion to strings
// - Correct data organization in the file
func TestAnyMapWrite(t *testing.T) {
	var mapAny AnyMap
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

// TestMapStringInterfaceWrite verifies writing a slice of interface{} maps to an Excel file.
// It tests:
// - Handling interface types
// - Complex type conversion
// - Data organization in columns and rows
func TestMapStringInterfaceWrite(t *testing.T) {
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

// TestStructNamedUserWrite verifies writing a NamedUser structure to an Excel file.
// It tests:
// - Writing structure fields
// - Respecting field tags
// - Converting complex types (dates, arrays)
// - Excluding ignored fields
func TestStructNamedUserWrite(t *testing.T) {
	// Create Excel file for testing
	outFile := excelize.NewFile()
	defer func() { _ = outFile.Close() }()

	// Create test data
	createdDate, _ := time.Parse("02/01/2006", "01/01/2023")
	namedUsers := []NamedUser{
		{
			Named: Named{
				ID:   1,
				Name: "Test User",
			},
			Ignored:     "This should be ignored",
			EncodedName: Encoded{Name: "encoded name"},
			Created:     createdDate,
			AnArray:     []int{1, 2, 3},
		},
	}

	// Configure Excel writer
	outExcel, _ := NewWriter(outFile)
	outExcel.SetSheet(outExcel.GetActiveSheet())
	outExcel.SetAxis("A1")

	// Serialize data
	err := outExcel.Marshal(&namedUsers)
	if err != nil {
		t.Error(err)
		return
	}

	// Check headers
	headerA1, _ := outFile.GetCellValue(outFile.GetSheetName(outFile.GetActiveSheetIndex()), "A1")
	headerB1, _ := outFile.GetCellValue(outFile.GetSheetName(outFile.GetActiveSheetIndex()), "B1")
	headerC1, _ := outFile.GetCellValue(outFile.GetSheetName(outFile.GetActiveSheetIndex()), "C1")
	headerD1, _ := outFile.GetCellValue(outFile.GetSheetName(outFile.GetActiveSheetIndex()), "D1")
	headerE1, _ := outFile.GetCellValue(outFile.GetSheetName(outFile.GetActiveSheetIndex()), "E1")

	assert.Equal(t, "Id", headerA1, "Header A1 should be 'Id'")
	assert.Equal(t, "Name", headerB1, "Header B1 should be 'Name'")
	assert.Equal(t, "Encoded_Name", headerC1, "Header C1 should be 'Encoded_Name'")
	assert.Equal(t, "created", headerD1, "Header D1 should be 'created'")
	assert.Equal(t, "array", headerE1, "Header E1 should be 'array'")

	// Check values
	valueA2, _ := outFile.GetCellValue(outFile.GetSheetName(outFile.GetActiveSheetIndex()), "A2")
	valueB2, _ := outFile.GetCellValue(outFile.GetSheetName(outFile.GetActiveSheetIndex()), "B2")
	valueC2, _ := outFile.GetCellValue(outFile.GetSheetName(outFile.GetActiveSheetIndex()), "C2")
	valueD2, _ := outFile.GetCellValue(outFile.GetSheetName(outFile.GetActiveSheetIndex()), "D2")
	valueE2, _ := outFile.GetCellValue(outFile.GetSheetName(outFile.GetActiveSheetIndex()), "E2")

	assert.Equal(t, "1", valueA2, "Value A2 should be '1'")
	assert.Equal(t, "Test User", valueB2, "Value B2 should be 'Test User'")
	assert.Equal(t, "{\"name\":\"encoded name\"}", valueC2, "Value C2 should be '{\"name\":\"encoded name\"}'")
	assert.Equal(t, "01/01/2023", valueD2, "Value D2 should be '01/01/2023'")
	assert.Equal(t, "1|2|3", valueE2, "Value E2 should be '1|2|3'")
}

// TestStructNamedUserReadWrite verifies writing and then reading a NamedUser structure.
// It tests:
// - Complete data serialization
// - File persistence
// - Data deserialization
// - Exact data matching before/after
func TestStructNamedUserReadWrite(t *testing.T) {
	// Create Excel file for write testing
	outFile := excelize.NewFile()

	// Create test data
	createdDate, _ := time.Parse("02/01/2006", "01/01/2023")
	originalUsers := []NamedUser{
		{
			Named: Named{
				ID:   1,
				Name: "Test User",
			},
			Ignored:     "This should be ignored",
			EncodedName: Encoded{Name: "encoded name"},
			Created:     createdDate,
			AnArray:     []int{1, 2, 3},
		},
	}

	// Configure Excel writer
	outExcel, _ := NewWriter(outFile)
	outExcel.SetSheet(outExcel.GetActiveSheet())
	outExcel.SetAxis("A1")

	// Serialize data
	err := outExcel.Marshal(&originalUsers)
	if err != nil {
		t.Error(err)
		return
	}
	if err := outFile.SaveAs("employees.gen.xlsx"); err != nil {
		log.Fatalf("Failed to save Excel file: %v", err)
	}

	// Now read data from the same file
	var readUsers []NamedUser

	// Configure Excel reader
	inExcel, _ := NewReader(outFile)
	inExcel.SetSheet(inExcel.GetActiveSheet())
	inExcel.SetAxis("A1")

	// Deserialize data
	err = inExcel.Unmarshal(&readUsers)
	if err != nil {
		t.Error(err)
		return
	}

	defer func() { _ = outFile.Close() }()

	// Verify that read data matches written data
	assert.Equal(t, 1, len(readUsers), "Should have one user")
	assert.Equal(t, originalUsers[0].ID, readUsers[0].ID, "IDs should match")
	assert.Equal(t, originalUsers[0].Name, readUsers[0].Name, "Names should match")
	assert.Equal(t, originalUsers[0].EncodedName.Name, readUsers[0].EncodedName.Name, "Encoded names should match")
	assert.Equal(t, len(originalUsers[0].AnArray), len(readUsers[0].AnArray), "Arrays should have the same length")
	for i := 0; i < len(originalUsers[0].AnArray); i++ {
		assert.Equal(t, originalUsers[0].AnArray[i], readUsers[0].AnArray[i], "Array elements should match")
	}
}
