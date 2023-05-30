package excel

import (
	"github.com/stretchr/testify/assert"
	"github.com/xuri/excelize/v2"
	"reflect"
	"testing"
	"time"
)

type MyStruct struct {
	ID   int
	Name string
}

type User struct {
	Id          int       `excel:"Id" excel-in:"ID" excel-out:"id"`
	Name        string    `excel:"Name,default:error" excel-in:"default:anonymous" excel-out:"default:not_used"`
	Ignored     string    `excel:"-"`
	EncodedName Encoded   `excel:"Encoded_Name,encoding:json"`
	Created     time.Time `excel:"created,format:d/m/Y"`
	AnArray     []int     `excel:"array,split:;" excel-out:"split:|"`
}

var users []User

type Encoded struct {
	Name string `json:"name"`
}

type DateTime struct {
	time.Time
}

func (date *DateTime) Marshall() (interface{}, error) {
	return date.Time.Format("20060201"), nil
}

func (date *DateTime) Unmarshall(s string) (err error) {
	date.Time, err = time.Parse("20060201", s)
	return err
}

type MyStructSlice []MyStruct
type MyStringSlice [][]string
type MyIntSlice [][]int
type MyAnySlice [][]any
type MyMapString []map[string]string
type MyMapInt []map[string]int
type MyMapAny []map[string]any

func TestReader_newReader(t *testing.T) {

	r := &Reader{}

	tests := []struct {
		name    string
		args    any
		want    IReader
		wantErr bool
	}{
		{
			name:    "struct reader",
			args:    MyStructSlice{},
			want:    &StructReader{},
			wantErr: true,
		},
		{
			name: "struct reader pointer",
			args: &MyStructSlice{},
			want: &StructReader{},
		},
		{
			name:    "slice reader (string)",
			args:    MyStringSlice{},
			want:    &SliceReader{},
			wantErr: true,
		},
		{
			name: "slice reader pointer (string)",
			args: &MyStringSlice{},
			want: &SliceReader{},
		},
		{
			name:    "slice reader (int)",
			args:    MyIntSlice{},
			want:    &SliceReader{},
			wantErr: true,
		},
		{
			name: "slice reader pointer (int)",
			args: &MyIntSlice{},
			want: &SliceReader{},
		},
		{
			name:    "slice reader (any)",
			args:    MyAnySlice{},
			want:    &SliceReader{},
			wantErr: true,
		},
		{
			name: "slice reader pointer (any)",
			args: &MyAnySlice{},
			want: &SliceReader{},
		},
		{
			name:    "map reader (string)",
			args:    MyMapString{},
			want:    &mapReader{},
			wantErr: true,
		},
		{
			name: "map reader pointer (string)",
			args: &MyMapString{},
			want: &mapReader{},
		},
		{
			name:    "map reader (int)",
			args:    MyMapInt{},
			want:    &mapReader{},
			wantErr: true,
		},
		{
			name: "map reader pointer (int)",
			args: &MyMapInt{},
			want: &mapReader{},
		},
		{
			name:    "map reader (any)",
			args:    MyMapAny{},
			want:    &mapReader{},
			wantErr: true,
		},
		{
			name: "map reader pointer (any)",
			args: &MyMapAny{},
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

func TestActiveFieldTags_ColumnName(t *testing.T) {

	inFile := excelize.NewFile()
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "A1", "ID")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "A2", 1)
	defer func() { _ = inFile.Close() }()

	inExcel, _ := NewReader(inFile)
	inExcel.SetSheetName(inFile.GetSheetName(inFile.GetActiveSheetIndex()))
	inExcel.SetAxis("A1")

	_ = inExcel.Unmarshal(&users)

	if len(users) != 1 {
		t.Error("Unmarshal() error")
		return
	}
	if users[0].Id != 1 {
		t.Error("Unmarshal() error")
		return
	}

	outFile := excelize.NewFile()
	defer func() { _ = outFile.Close() }()

	outExcel, _ := NewWriter(outFile)
	outExcel.SetSheetName(outFile.GetSheetName(outFile.GetActiveSheetIndex()))
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

func TestActiveFieldTags_DefaultValue(t *testing.T) {

	inFile := excelize.NewFile()
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "A1", "ID")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "A2", 1)
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "B1", "Name")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "B2", "")
	defer func() { _ = inFile.Close() }()

	inExcel, _ := NewReader(inFile)
	inExcel.SetSheetIndex(0)
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
	outExcel.SetSheetIndex(0)
	outExcel.SetAxisCoordinates(1, 1)

	_ = outExcel.Marshal(&users)

	value, _ := outFile.GetCellValue(outFile.GetSheetName(outFile.GetActiveSheetIndex()), "B2")

	if value != "anonymous" {
		t.Error("Marshal() error")
		return
	}
}

func TestActiveFieldTags_Ignored(t *testing.T) {

	inFile := excelize.NewFile()
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "A1", "ID")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "A2", 1)
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "B1", "Name")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "B2", "")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "C1", "Ignored")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "C2", "this is ignored")
	defer func() { _ = inFile.Close() }()

	inExcel, _ := NewReader(inFile)
	inExcel.SetSheetName(inFile.GetSheetName(inFile.GetActiveSheetIndex()))
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
	outExcel.SetSheetName(outFile.GetSheetName(outFile.GetActiveSheetIndex()))
	outExcel.SetAxis("A1")

	_ = outExcel.Marshal(&users)

	value, _ := outFile.GetCellValue(outFile.GetSheetName(outFile.GetActiveSheetIndex()), "C2")

	if value != "{\"name\":\"\"}" {
		t.Error("Marshal() error")
		return
	}
}

func TestActiveFieldTags_Encoding(t *testing.T) {

	inFile := excelize.NewFile()
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "A1", "ID")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "A2", 1)
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "B1", "Encoded_Name")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "B2", "{ \"name\":\"a string\" }")
	defer func() { _ = inFile.Close() }()

	inExcel, _ := NewReader(inFile)
	inExcel.SetSheetName(inFile.GetSheetName(inFile.GetActiveSheetIndex()))
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
	outExcel.SetSheetName(outFile.GetSheetName(outFile.GetActiveSheetIndex()))
	outExcel.SetAxis("A1")

	_ = outExcel.Marshal(&users)

	var value string
	value, _ = outFile.GetCellValue(outFile.GetSheetName(outFile.GetActiveSheetIndex()), "C2")

	if value != "{\"name\":\"a string\"}" {
		t.Error("Marshal() error")
		return
	}
}

func TestActiveFieldTags_Format(t *testing.T) {

	inFile := excelize.NewFile()
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "A1", "ID")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "A2", 1)
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "B1", "created")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "B2", "31/12/2022")
	defer func() { _ = inFile.Close() }()

	inExcel, _ := NewReader(inFile)
	inExcel.SetSheetName(inFile.GetSheetName(inFile.GetActiveSheetIndex()))
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
	outExcel.SetSheetName(outFile.GetSheetName(outFile.GetActiveSheetIndex()))
	outExcel.SetAxis("A1")

	_ = outExcel.Marshal(&users)

	var value string
	value, _ = outFile.GetCellValue(outFile.GetSheetName(outFile.GetActiveSheetIndex()), "D2")

	if value != "31/12/2022" {
		t.Error("Marshal() error")
		return
	}
}

func TestActiveFieldTags_Split(t *testing.T) {

	inFile := excelize.NewFile()
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "A1", "ID")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "A2", 1)
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "B1", "array")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "B2", "1;2;3")
	defer func() { _ = inFile.Close() }()

	inExcel, _ := NewReader(inFile)
	inExcel.SetSheetName(inFile.GetSheetName(inFile.GetActiveSheetIndex()))
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
	outExcel.SetSheetName(outFile.GetSheetName(outFile.GetActiveSheetIndex()))
	outExcel.SetAxis("A1")

	_ = outExcel.Marshal(&users)

	var value string
	value, _ = outFile.GetCellValue(outFile.GetSheetName(outFile.GetActiveSheetIndex()), "E2")

	if value != "1|2|3" {
		t.Error("Marshal() error")
		return
	}
}

func TestActiveFieldTags_Required(t *testing.T) {

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
	inExcel.SetSheetName(inFile.GetSheetName(inFile.GetActiveSheetIndex()))
	inExcel.SetAxis("A1")

	err := inExcel.Unmarshal(&simpleUsers)
	if err != ErrColumnRequired {
		t.Error("Required column error")
		return
	}
}

func TestConverter(t *testing.T) {

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
	outExcel.SetSheetName(file.GetSheetName(file.GetActiveSheetIndex()))
	outExcel.SetAxis("A1")

	err := outExcel.Marshal(&simpleUsers)
	if err != nil {
		t.Error("Marshal error")
		return
	}

	dateValue, _ := file.GetCellValue(outExcel.GetSheetName(), "C2")

	if dateValue != "20223112" {
		t.Error("Marshal() error")
		return
	}

	var usersOut []*SimpleUser

	inExcel, _ := NewReader(file)
	inExcel.SetSheetName(outExcel.GetSheetName())
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

func TestSliceRead_string(t *testing.T) {

	file := excelize.NewFile()
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A1", "ID")
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B1", "Name")
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A2", 1)
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B2", "John Doe")
	defer func() { _ = file.Close() }()

	xl, _ := NewReader(file)
	xl.SetSheetName(file.GetSheetName(file.GetActiveSheetIndex()))
	xl.SetAxis("A1")

	var stringSlice MyStringSlice
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

func TestSliceRead_int(t *testing.T) {

	file := excelize.NewFile()
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A1", 1)
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B1", 2)
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A2", 3)
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B2", 4)
	defer func() { _ = file.Close() }()

	xl, _ := NewReader(file)
	xl.SetSheetName(file.GetSheetName(file.GetActiveSheetIndex()))
	xl.SetAxis("A1")

	var intSlice MyIntSlice
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

func TestSliceRead_any(t *testing.T) {

	file := excelize.NewFile()
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A1", "ID")
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B1", "Name")
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A2", 1)
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B2", "John Doe")
	defer func() { _ = file.Close() }()

	xl, _ := NewReader(file)
	xl.SetSheetName(file.GetSheetName(file.GetActiveSheetIndex()))
	xl.SetAxis("A1")

	var anySlice MyAnySlice
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

func TestMapRead_string(t *testing.T) {

	file := excelize.NewFile()
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A1", "ID")
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B1", "Name")
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A2", 1)
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B2", "John Doe")
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A3", 2)
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B3", "Jane Doe")
	defer func() { _ = file.Close() }()

	xl, _ := NewReader(file)
	xl.SetSheetName(file.GetSheetName(file.GetActiveSheetIndex()))
	xl.SetAxis("A1")

	var mapString MyMapString
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

func TestMapRead_int(t *testing.T) {

	file := excelize.NewFile()
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A1", "ID1")
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B1", "ID2")
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A2", 1)
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B2", 2)
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A3", 3)
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B3", 4)
	defer func() { _ = file.Close() }()

	xl, _ := NewReader(file)
	xl.SetSheetName(file.GetSheetName(file.GetActiveSheetIndex()))
	xl.SetAxis("A1")

	var mapInt MyMapInt
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

func TestMapRead_any(t *testing.T) {

	file := excelize.NewFile()
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A1", "ID")
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B1", "Name")
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A2", 1)
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B2", "John Doe")
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "A3", 2)
	_ = file.SetCellValue(file.GetSheetName(file.GetActiveSheetIndex()), "B3", "Jane Doe")
	defer func() { _ = file.Close() }()

	xl, _ := NewReader(file)
	xl.SetSheetName(file.GetSheetName(file.GetActiveSheetIndex()))
	xl.SetAxis("A1")

	var mapAny MyMapAny
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
