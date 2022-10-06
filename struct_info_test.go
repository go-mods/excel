package excel_test

import (
	"github.com/go-mods/excel"
	"testing"
	"time"

	"github.com/xuri/excelize/v2"
)

type User struct {
	Id          int       `excel:"Id" excel-in:"ID" excel-out:"id"`
	Name        string    `excel:"Name,default:error" excel-in:"default:anonymous" excel-out:"default:not_used"`
	Ignored     string    `excel:"-"`
	EncodedName Encoded   `excel:"Encoded_Name,encoding:json"`
	Created     time.Time `excel:"created,format:d/m/Y"`
	AnArray     []int     `excel:"array,split:;" excel-out:"split:|"`
}

type Encoded struct {
	Name string `json:"name"`
}

var users []User

func TestActiveFieldTags_ColumnName(t *testing.T) {

	inFile := excelize.NewFile()
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "A1", "ID")
	_ = inFile.SetCellValue(inFile.GetSheetName(inFile.GetActiveSheetIndex()), "A2", 1)
	defer func() { _ = inFile.Close() }()

	inExcel, _ := excel.NewReader(inFile)
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

	outExcel, _ := excel.NewWriter(outFile)
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

	inExcel, _ := excel.NewReader(inFile)
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

	outExcel, _ := excel.NewWriter(outFile)
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

	inExcel, _ := excel.NewReader(inFile)
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

	outExcel, _ := excel.NewWriter(outFile)
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

	inExcel, _ := excel.NewReader(inFile)
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

	outExcel, _ := excel.NewWriter(outFile)
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

	inExcel, _ := excel.NewReader(inFile)
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

	outExcel, _ := excel.NewWriter(outFile)
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

	inExcel, _ := excel.NewReader(inFile)
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

	outExcel, _ := excel.NewWriter(outFile)
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

	inExcel, _ := excel.NewReader(inFile)
	inExcel.SetSheetName(inFile.GetSheetName(inFile.GetActiveSheetIndex()))
	inExcel.SetAxis("A1")

	err := inExcel.Unmarshal(&simpleUsers)
	if err != excel.ErrColumnRequired {
		t.Error("Required column error")
		return
	}
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
	user13 := &SimpleUser{Id: 3, Name: "three", Created: DateTime{now}, Modified: &DateTime{later}}
	simpleUsers = append(simpleUsers, user1, user2, user13)

	file := excelize.NewFile()
	defer func() { _ = file.Close() }()

	outExcel, _ := excel.NewWriter(file)
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

	var users []*SimpleUser

	inExcel, _ := excel.NewReader(file)
	inExcel.SetSheetName(outExcel.GetSheetName())
	inExcel.SetAxis("A1")

	_ = inExcel.Unmarshal(&users)

	if len(users) != 3 {
		t.Error("Unmarshal() error")
		return
	}
	if users[0].Id != simpleUsers[0].Id {
		t.Error("Unmarshal() error")
		return
	}
	if users[0].Name != simpleUsers[0].Name {
		t.Error("Unmarshal() error")
		return
	}
	if users[0].Created.Time.Year() != simpleUsers[0].Created.Time.Year() || users[0].Created.Time.Month() != simpleUsers[0].Created.Time.Month() || users[0].Created.Time.Day() != simpleUsers[0].Created.Time.Day() {
		t.Error("Unmarshal() error")
		return
	}
	if users[0].Modified.Time.Year() != simpleUsers[0].Modified.Time.Year() || users[0].Modified.Time.Month() != simpleUsers[0].Modified.Time.Month() || users[0].Modified.Time.Day() != simpleUsers[0].Modified.Time.Day() {
		t.Error("Unmarshal() error")
		return
	}
}
