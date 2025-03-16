package excel

import (
	"reflect"
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
	"github.com/xuri/excelize/v2"
)

// Named represents a structure with ID and Name fields
// The ID field is mapped to "Id" column for both import and export
// The Name field has different default values:
// - "error" for general excel tag
// - "anonymous" when importing
// - "not_used" when exporting
type Named struct {
	ID   int    `excel:"Id" excel-in:"Id" excel-out:"Id"`
	Name string `excel:"Name,default:error" excel-in:"default:anonymous" excel-out:"default:not_used"`
}

// User represents a structure with various field types and excel tags:
// - Id: mapped to different column names for import/export
// - Name: has different default values for import/export
// - Ignored: excluded from excel processing
// - EncodedName: JSON encoded/decoded
// - Created: formatted date field
// - AnArray: array split by different delimiters for import/export
type User struct {
	Id          int       `excel:"Id" excel-in:"ID" excel-out:"id"`
	Name        string    `excel:"Name,default:error" excel-in:"default:anonymous" excel-out:"default:not_used"`
	Ignored     string    `excel:"-"`
	EncodedName Encoded   `excel:"Encoded_Name,encoding:json"`
	Created     time.Time `excel:"created,format:d/m/Y"`
	AnArray     []int     `excel:"array,split:;" excel-out:"split:|"`
}

// NamedUser represents a structure that embeds Named and adds additional fields:
// - Ignored: excluded from excel processing
// - EncodedName: JSON encoded/decoded name field
// - Created: formatted date field
// - AnArray: array split by | delimiter
type NamedUser struct {
	Named
	Ignored     string    `excel:"-"`
	EncodedName Encoded   `excel:"Encoded_Name,encoding:json"`
	Created     time.Time `excel:"created,format:d/m/Y"`
	AnArray     []int     `excel:"array,split:|"`
}

// users represents a slice of User structs used for testing
var users []User

// Encoded represents a structure that can be encoded/decoded to/from JSON
// It contains a Name field that is tagged with json:"name"
type Encoded struct {
	Name string `json:"name"`
}

// DateTime wraps time.Time to provide custom marshalling/unmarshalling
type DateTime struct {
	time.Time
}

// Marshall formats the DateTime as YYYYMMDD string
func (date *DateTime) Marshall() (interface{}, error) {
	return date.Time.Format("20060201"), nil
}

// Unmarshall parses a YYYYMMDD string into the DateTime
func (date *DateTime) Unmarshall(s string) (err error) {
	date.Time, err = time.Parse("20060201", s)
	return err
}

type StructSlice []Named
type StringMatrix [][]string
type IntMatrix [][]int
type AnyMatrix [][]any
type StringMap []map[string]string
type IntMap []map[string]int
type AnyMap []map[string]any

// TestNewReader verifies that the correct reader type is created based on the input data type.
// It tests reader creation for structures, slices, and maps, both for pointers and direct values.
func TestNewReader(t *testing.T) {

	r := &Reader{}

	tests := []struct {
		name    string
		args    any
		want    IReader
		wantErr bool
	}{
		{
			name:    "struct reader",
			args:    StructSlice{},
			want:    &StructReader{},
			wantErr: true,
		},
		{
			name: "struct reader pointer",
			args: &StructSlice{},
			want: &StructReader{},
		},
		{
			name:    "slice reader (string)",
			args:    StringMatrix{},
			want:    &SliceReader{},
			wantErr: true,
		},
		{
			name: "slice reader pointer (string)",
			args: &StringMatrix{},
			want: &SliceReader{},
		},
		{
			name:    "slice reader (int)",
			args:    IntMatrix{},
			want:    &SliceReader{},
			wantErr: true,
		},
		{
			name: "slice reader pointer (int)",
			args: &IntMatrix{},
			want: &SliceReader{},
		},
		{
			name:    "slice reader (any)",
			args:    AnyMatrix{},
			want:    &SliceReader{},
			wantErr: true,
		},
		{
			name: "slice reader pointer (any)",
			args: &AnyMatrix{},
			want: &SliceReader{},
		},
		{
			name:    "map reader (string)",
			args:    StringMap{},
			want:    &mapReader{},
			wantErr: true,
		},
		{
			name: "map reader pointer (string)",
			args: &StringMap{},
			want: &mapReader{},
		},
		{
			name:    "map reader (int)",
			args:    IntMap{},
			want:    &mapReader{},
			wantErr: true,
		},
		{
			name: "map reader pointer (int)",
			args: &IntMap{},
			want: &mapReader{},
		},
		{
			name:    "map reader (any)",
			args:    AnyMap{},
			want:    &mapReader{},
			wantErr: true,
		},
		{
			name: "map reader pointer (any)",
			args: &AnyMap{},
			want: &mapReader{},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := r.newReader(tt.args)
			if (err != nil) != tt.wantErr {
				t.Errorf("newReader() error = %v, wantErr %v", err, tt.wantErr)
				return
			}
			if err == nil {
				if got == nil {
					t.Errorf("newReader() got = %v, want %v", got, tt.want)
				}
				if reflect.TypeOf(got) != reflect.TypeOf(tt.want) {
					t.Errorf("newReader() got = %v, want %v", got, tt.want)
				}
			}
		})
	}
}

// TestExcelColumnNameMapping verifies that column names are correctly mapped between Excel and struct.
// It tests:
// - Reading column headers
// - Mapping to struct fields
// - Writing back to Excel with correct headers
func TestExcelColumnNameMapping(t *testing.T) {

	inFile := excelize.NewFile()
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "A1", "ID")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "A2", 1)
	defer func() { _ = inFile.Close() }()

	inExcel, _ := NewReader(inFile)
	inExcel.SetSheet(inExcel.GetActiveSheet())
	inExcel.SetAxis("A1")

	_ = inExcel.Unmarshal(&users)

	if len(users) != 1 {
		t.Errorf("Unmarshal() failed: expected users length to be 1, got %d", len(users))
		return
	}
	if users[0].Id != 1 {
		t.Errorf("Unmarshal() failed: expected Id to be 1, got %d", users[0].Id)
		return
	}

	outFile := excelize.NewFile()
	defer func() { _ = outFile.Close() }()

	outExcel, _ := NewWriter(outFile)
	outExcel.SetSheet(outExcel.GetActiveSheet())
	outExcel.SetAxis("A1")

	_ = outExcel.Marshal(&users)

	var title string
	var value string
	title, _ = outFile.GetCellValue(outFile.GetSheetName(outFile.GetActiveSheetIndex()), "A1")
	value, _ = outFile.GetCellValue(outFile.GetSheetName(outFile.GetActiveSheetIndex()), "A2")

	if title != "id" {
		t.Error("Marshal() error")
		return
	}

	if value != "1" {
		t.Error("Marshal() error")
		return
	}
}

// TestDefaultValueHandling verifies that default values are correctly applied during reading/writing.
// It tests:
// - Empty cell handling
// - Default value application
// - Value persistence during read/write cycles
func TestDefaultValueHandling(t *testing.T) {

	inFile := excelize.NewFile()
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "A1", "ID")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "A2", 1)
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "B1", "Name")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "B2", "")
	defer func() { _ = inFile.Close() }()

	inExcel, _ := NewReader(inFile)
	inExcel.SetSheetFromIndex(0)
	inExcel.SetAxisCoordinates(1, 1)

	_ = inExcel.Unmarshal(&users)

	if len(users) != 1 {
		t.Error("Unmarshal() error")
		return
	}
	if users[0].Name != "anonymous" {
		t.Error("Unmarshal() error")
		return
	}

	outFile := excelize.NewFile()
	defer func() { _ = outFile.Close() }()

	outExcel, _ := NewWriter(outFile)
	outExcel.SetSheetFromIndex(0)
	outExcel.SetAxisCoordinates(1, 1)

	_ = outExcel.Marshal(&users)

	value, _ := outFile.GetCellValue(outFile.GetSheetName(outFile.GetActiveSheetIndex()), "B2")

	if value != "anonymous" {
		t.Error("Marshal() error")
		return
	}
}

// TestIgnoredFieldHandling verifies that fields marked as ignored are properly handled.
// It tests:
// - Ignored field exclusion
// - Data integrity for non-ignored fields
// - JSON encoding of remaining fields
func TestIgnoredFieldHandling(t *testing.T) {

	inFile := excelize.NewFile()
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "A1", "ID")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "A2", 1)
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "B1", "Name")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "B2", "")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "C1", "Ignored")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "C2", "this is ignored")
	defer func() { _ = inFile.Close() }()

	inExcel, _ := NewReader(inFile)
	inExcel.SetSheet(inExcel.GetActiveSheet())
	inExcel.SetAxis("A1")

	_ = inExcel.Unmarshal(&users)

	if len(users) != 1 {
		t.Error("Unmarshal() error")
		return
	}
	if users[0].Ignored != "" {
		t.Error("Unmarshal() error")
		return
	}

	outFile := excelize.NewFile()
	defer func() { _ = outFile.Close() }()

	outExcel, _ := NewWriter(outFile)
	outExcel.SetSheet(outExcel.GetActiveSheet())
	outExcel.SetAxis("A1")

	_ = outExcel.Marshal(&users)

	value, _ := outFile.GetCellValue(outFile.GetSheetName(outFile.GetActiveSheetIndex()), "C2")

	if value != "{\"name\":\"\"}" {
		t.Error("Marshal() error")
		return
	}
}

// TestJSONEncoding verifies JSON encoding/decoding of fields.
// It tests:
// - JSON string parsing
// - Struct field mapping
// - JSON serialization
func TestJSONEncoding(t *testing.T) {

	inFile := excelize.NewFile()
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "A1", "ID")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "A2", 1)
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "B1", "Encoded_Name")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "B2", "{ \"name\":\"a string\" }")
	defer func() { _ = inFile.Close() }()

	inExcel, _ := NewReader(inFile)
	inExcel.SetSheet(inExcel.GetActiveSheet())
	inExcel.SetAxis("A1")

	_ = inExcel.Unmarshal(&users)

	if len(users) != 1 {
		t.Error("Unmarshal() error")
		return
	}
	if users[0].EncodedName.Name != "a string" {
		t.Error("Unmarshal() error")
		return
	}

	outFile := excelize.NewFile()
	defer func() { _ = outFile.Close() }()

	outExcel, _ := NewWriter(outFile)
	outExcel.SetSheet(outExcel.GetActiveSheet())
	outExcel.SetAxis("A1")

	_ = outExcel.Marshal(&users)

	var value string
	value, _ = outFile.GetCellValue(outFile.GetSheetName(outFile.GetActiveSheetIndex()), "C2")

	if value != "{\"name\":\"a string\"}" {
		t.Error("Marshal() error")
		return
	}
}

// TestDateFormatting verifies date formatting during reading/writing.
// It tests:
// - Date string parsing
// - Time.Time conversion
// - Date format output
func TestDateFormatting(t *testing.T) {

	inFile := excelize.NewFile()
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "A1", "ID")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "A2", 1)
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "B1", "created")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "B2", "31/12/2022")
	defer func() { _ = inFile.Close() }()

	inExcel, _ := NewReader(inFile)
	inExcel.SetSheet(inExcel.GetActiveSheet())
	inExcel.SetAxis("A1")

	_ = inExcel.Unmarshal(&users)

	if len(users) != 1 {
		t.Error("Unmarshal() error")
		return
	}
	if users[0].Created.Year() != 2022 {
		t.Error("Unmarshal() error")
		return
	}
	if users[0].Created.Month() != 12 {
		t.Error("Unmarshal() error")
		return
	}
	if users[0].Created.Day() != 31 {
		t.Error("Unmarshal() error")
		return
	}

	outFile := excelize.NewFile()
	defer func() { _ = outFile.Close() }()

	outExcel, _ := NewWriter(outFile)
	outExcel.SetSheet(outExcel.GetActiveSheet())
	outExcel.SetAxis("A1")

	_ = outExcel.Marshal(&users)

	var value string
	value, _ = outFile.GetCellValue(outFile.GetSheetName(outFile.GetActiveSheetIndex()), "D2")

	if value != "31/12/2022" {
		t.Error("Marshal() error")
		return
	}
}

// TestArraySplitting verifies array splitting and joining during reading/writing.
// It tests:
// - Array string splitting
// - Type conversion
// - Array joining with delimiters
func TestArraySplitting(t *testing.T) {

	inFile := excelize.NewFile()
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "A1", "ID")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "A2", 1)
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "B1", "array")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "B2", "1;2;3")
	defer func() { _ = inFile.Close() }()

	inExcel, _ := NewReader(inFile)
	inExcel.SetSheet(inExcel.GetActiveSheet())
	inExcel.SetAxis("A1")

	_ = inExcel.Unmarshal(&users)

	if len(users) != 1 {
		t.Error("Unmarshal() error")
		return
	}
	if len(users[0].AnArray) != 3 {
		t.Error("Unmarshal() error")
		return
	}
	if users[0].AnArray[0] != 1 {
		t.Error("Unmarshal() error")
		return
	}
	if users[0].AnArray[1] != 2 {
		t.Error("Unmarshal() error")
		return
	}
	if users[0].AnArray[2] != 3 {
		t.Error("Unmarshal() error")
		return
	}

	outFile := excelize.NewFile()
	defer func() { _ = outFile.Close() }()

	outExcel, _ := NewWriter(outFile)
	outExcel.SetSheet(outExcel.GetActiveSheet())
	outExcel.SetAxis("A1")

	_ = outExcel.Marshal(&users)

	var value string
	value, _ = outFile.GetCellValue(outFile.GetSheetName(outFile.GetActiveSheetIndex()), "E2")

	if value != "1|2|3" {
		t.Error("Marshal() error")
		return
	}
}

// TestRequiredFields verifies that required fields are properly validated.
// It tests:
// - Required field presence
// - Error handling for missing fields
// - Validation process
func TestRequiredFields(t *testing.T) {

	type SimpleUser struct {
		Id   int    `excel:"Id,required"`
		Name string `excel:"Name,required"`
	}

	var simpleUsers []SimpleUser

	inFile := excelize.NewFile()
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "A1", "Id")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "A2", 1)
	defer func() { _ = inFile.Close() }()

	inExcel, _ := NewReader(inFile)
	inExcel.SetSheet(inExcel.GetActiveSheet())
	inExcel.SetAxis("A1")

	err := inExcel.Unmarshal(&simpleUsers)
	if err != ErrColumnRequired {
		t.Error("Required column error")
		return
	}
}

// TestCustomTypeConverter verifies custom type conversion functionality.
// It tests:
// - Custom type marshalling
// - Custom type unmarshalling
// - Data integrity through conversion
func TestCustomTypeConverter(t *testing.T) {

	type SimpleUser struct {
		Id       int       `excel:"Id"`
		Name     string    `excel:"Name"`
		Created  DateTime  `excel:"Created"`
		Modified *DateTime `excel:"Modified"`
	}

	now := time.Date(2022, 12, 31, 0, 0, 0, 0, time.UTC)
	later := now.AddDate(1, 0, 0)

	var simpleUsers []*SimpleUser
	user1 := &SimpleUser{Id: 1, Name: "One", Created: DateTime{now}, Modified: &DateTime{later}}
	user2 := &SimpleUser{Id: 2, Name: "two", Created: DateTime{now}, Modified: &DateTime{later}}
	user3 := &SimpleUser{Id: 3, Name: "three", Created: DateTime{now}, Modified: &DateTime{later}}
	simpleUsers = append(simpleUsers, user1, user2, user3)

	file := excelize.NewFile()
	defer func() { _ = file.Close() }()

	outExcel, _ := NewWriter(file)
	outExcel.SetSheet(outExcel.GetActiveSheet())
	outExcel.SetAxis("A1")

	err := outExcel.Marshal(&simpleUsers)
	if err != nil {
		t.Error("Marshal error")
		return
	}

	dateValue, _ := file.GetCellValue(outExcel.Sheet().Name, "C2")

	if dateValue != "20223112" {
		t.Error("Marshal() error")
		return
	}

	var usersOut []*SimpleUser

	inExcel, _ := NewReader(file)
	inExcel.SetSheet(outExcel.Sheet())
	inExcel.SetAxis("A1")

	_ = inExcel.Unmarshal(&usersOut)

	if len(usersOut) != 3 {
		t.Error("Unmarshal() error")
		return
	}
	if usersOut[0].Id != simpleUsers[0].Id {
		t.Error("Unmarshal() error")
		return
	}
	if usersOut[0].Name != simpleUsers[0].Name {
		t.Error("Unmarshal() error")
		return
	}
	if usersOut[0].Created.Time.Year() != simpleUsers[0].Created.Time.Year() || usersOut[0].Created.Time.Month() != simpleUsers[0].Created.Time.Month() || usersOut[0].Created.Time.Day() != simpleUsers[0].Created.Time.Day() {
		t.Error("Unmarshal() error")
		return
	}
	if usersOut[0].Modified.Time.Year() != simpleUsers[0].Modified.Time.Year() || usersOut[0].Modified.Time.Month() != simpleUsers[0].Modified.Time.Month() || usersOut[0].Modified.Time.Day() != simpleUsers[0].Modified.Time.Day() {
		t.Error("Unmarshal() error")
		return
	}
}

// TestStringMatrixRead verifies reading a string matrix from Excel.
// It tests:
// - Matrix structure reading
// - String value handling
// - Dimensional accuracy
func TestStringMatrixRead(t *testing.T) {

	file := excelize.NewFile()
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A1", "ID")
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B1", "Name")
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A2", 1)
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B2", "John Doe")
	defer func() { _ = file.Close() }()

	xl, _ := NewReader(file)
	xl.SetSheet(xl.GetActiveSheet())
	xl.SetAxis("A1")

	var stringSlice StringMatrix
	err := xl.Unmarshal(&stringSlice)
	if err != nil {
		t.Error(err)
		return
	}

	assert.Equal(t, stringSlice[0][0], "ID", "they should be equal")
	assert.Equal(t, stringSlice[0][1], "Name", "they should be equal")
	assert.Equal(t, stringSlice[1][0], "1", "they should be equal")
	assert.Equal(t, stringSlice[1][1], "John Doe", "they should be equal")
}

// TestIntMatrixRead verifies reading a matrix of integers from Excel.
// It tests:
// - Integer parsing
// - Matrix structure
// - Value conversion accuracy
func TestIntMatrixRead(t *testing.T) {

	file := excelize.NewFile()
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A1", 1)
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B1", 2)
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A2", 3)
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B2", 4)
	defer func() { _ = file.Close() }()

	xl, _ := NewReader(file)
	xl.SetSheet(xl.GetActiveSheet())
	xl.SetAxis("A1")

	var intSlice IntMatrix
	err := xl.Unmarshal(&intSlice)
	if err != nil {
		t.Error(err)
		return
	}

	assert.Equal(t, intSlice[0][0], 1, "they should be equal")
	assert.Equal(t, intSlice[0][1], 2, "they should be equal")
	assert.Equal(t, intSlice[1][0], 3, "they should be equal")
	assert.Equal(t, intSlice[1][1], 4, "they should be equal")
}

// TestAnyMatrixRead verifies reading a matrix of mixed types from Excel.
// It tests:
// - Mixed type handling
// - Type inference
// - Matrix structure preservation
func TestAnyMatrixRead(t *testing.T) {

	file := excelize.NewFile()
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A1", "ID")
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B1", "Name")
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A2", 1)
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B2", "John Doe")
	defer func() { _ = file.Close() }()

	xl, _ := NewReader(file)
	xl.SetSheet(xl.GetActiveSheet())
	xl.SetAxis("A1")

	var anySlice AnyMatrix
	err := xl.Unmarshal(&anySlice)
	if err != nil {
		t.Error(err)
		return
	}

	assert.Equal(t, anySlice[0][0], "ID", "they should be equal")
	assert.Equal(t, anySlice[0][1], "Name", "they should be equal")
	assert.Equal(t, anySlice[1][0], int64(1), "they should be equal")
	assert.Equal(t, anySlice[1][1], "John Doe", "they should be equal")

	assert.Equal(t, xl.Reader.Result.Rows, 2, "they should be equal")
	assert.Equal(t, xl.Reader.Result.Columns, 2, "they should be equal")
}

// TestStringMapRead verifies reading a map of strings from Excel.
// It tests:
// - Header row processing
// - Map structure creation
// - String value mapping
func TestStringMapRead(t *testing.T) {

	file := excelize.NewFile()
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A1", "ID")
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B1", "Name")
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A2", 1)
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B2", "John Doe")
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A3", 2)
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B3", "Jane Doe")
	defer func() { _ = file.Close() }()

	xl, _ := NewReader(file)
	xl.SetSheet(xl.GetActiveSheet())
	xl.SetAxis("A1")

	var mapString StringMap
	err := xl.Unmarshal(&mapString)
	if err != nil {
		t.Error(err)
		return
	}

	assert.Equal(t, mapString[0]["ID"], "1", "they should be equal")
	assert.Equal(t, mapString[0]["Name"], "John Doe", "they should be equal")
	assert.Equal(t, mapString[1]["ID"], "2", "they should be equal")
	assert.Equal(t, mapString[1]["Name"], "Jane Doe", "they should be equal")

	assert.Equal(t, xl.Reader.Result.Rows, 3, "they should be equal")
	assert.Equal(t, xl.Reader.Result.Columns, 2, "they should be equal")
}

// TestIntMapRead verifies reading a map of integers from Excel.
// It tests:
// - Header processing
// - Integer parsing
// - Map structure integrity
func TestIntMapRead(t *testing.T) {

	file := excelize.NewFile()
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A1", "ID1")
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B1", "ID2")
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A2", 1)
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B2", 2)
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A3", 3)
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B3", 4)
	defer func() { _ = file.Close() }()

	xl, _ := NewReader(file)
	xl.SetSheet(xl.GetActiveSheet())
	xl.SetAxis("A1")

	var mapInt IntMap
	err := xl.Unmarshal(&mapInt)
	if err != nil {
		t.Error(err)
		return
	}

	assert.Equal(t, mapInt[0]["ID1"], 1, "they should be equal")
	assert.Equal(t, mapInt[0]["ID2"], 2, "they should be equal")
	assert.Equal(t, mapInt[1]["ID1"], 3, "they should be equal")
	assert.Equal(t, mapInt[1]["ID2"], 4, "they should be equal")

	assert.Equal(t, xl.Reader.Result.Rows, 3, "they should be equal")
	assert.Equal(t, xl.Reader.Result.Columns, 2, "they should be equal")
}

// TestAnyMapRead verifies reading a map of mixed types from Excel.
// It tests:
// - Mixed type parsing
// - Map structure creation
// - Type inference and conversion
func TestAnyMapRead(t *testing.T) {

	file := excelize.NewFile()
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A1", "ID")
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B1", "Name")
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A2", 1)
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B2", "John Doe")
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A3", 2)
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B3", "Jane Doe")
	defer func() { _ = file.Close() }()

	xl, _ := NewReader(file)
	xl.SetSheet(xl.GetActiveSheet())
	xl.SetAxis("A1")

	var mapAny AnyMap
	err := xl.Unmarshal(&mapAny)
	if err != nil {
		t.Error(err)
		return
	}

	assert.Equal(t, mapAny[0]["ID"], int64(1), "they should be equal")
	assert.Equal(t, mapAny[0]["Name"], "John Doe", "they should be equal")
	assert.Equal(t, mapAny[1]["ID"], int64(2), "they should be equal")
	assert.Equal(t, mapAny[1]["Name"], "Jane Doe", "they should be equal")

	assert.Equal(t, xl.Reader.Result.Rows, 3, "they should be equal")
	assert.Equal(t, xl.Reader.Result.Columns, 2, "they should be equal")
}

// TestStructNamedUserRead verifies reading a NamedUser struct from Excel.
// It tests:
// - Structure field mapping
// - Complex type handling
// - Field tag processing
func TestStructNamedUserRead(t *testing.T) {
	// Create Excel file for testing
	inFile := excelize.NewFile()

	// Set column headers
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "A1", "Id")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "B1", "Name")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "C1", "Encoded_Name")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "D1", "created")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "E1", "array")

	// Set values
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "A2", 1)
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "B2", "Test User")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "C2", "{\"name\":\"encoded name\"}")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "D2", "01/01/2023")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "E2", "1|2|3")

	defer func() { _ = inFile.Close() }()

	// Create container for users
	var namedUsers []NamedUser

	// Configure Excel reader
	inExcel, _ := NewReader(inFile)
	inExcel.SetSheet(inExcel.GetActiveSheet())
	inExcel.SetAxis("A1")

	// Deserialize data
	err := inExcel.Unmarshal(&namedUsers)
	if err != nil {
		t.Error(err)
		return
	}

	// Verify results
	assert.Equal(t, 1, len(namedUsers), "Should have one user")
	assert.Equal(t, 1, namedUsers[0].ID, "ID should be 1")
	assert.Equal(t, "Test User", namedUsers[0].Name, "Name should be 'Test User'")
	assert.Equal(t, "encoded name", namedUsers[0].EncodedName.Name, "Encoded name should be 'encoded name'")
	assert.Equal(t, 3, len(namedUsers[0].AnArray), "Array should contain 3 elements")
	assert.Equal(t, 1, namedUsers[0].AnArray[0], "First array element should be 1")
	assert.Equal(t, 2, namedUsers[0].AnArray[1], "Second array element should be 2")
	assert.Equal(t, 3, namedUsers[0].AnArray[2], "Third array element should be 3")
}
