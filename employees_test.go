package excel_test

import (
	"github.com/go-mods/excel"
	"github.com/xuri/excelize/v2"
	"testing"
	"time"
)

type Employee struct {
	// use field name as default column name
	ID int
	// map the column 'firstName'
	First string `excel:"firstName"`
	// column can also be used to set the column name
	FirstPtr *string `excel:"column:firstName"`
	// map the column 'lastName'
	Last string `excel:"lastName"`
	// 'column' can be omitted when only mapping to column name
	// this is equal to 'column(email)'
	Email string `excel:"email"`
	// map the column 'contactNumber'
	ContactNumber string `excel:"contactNumber"`
	// map the column 'age'
	Age int `excel:"age"`
	// map the column 'dob'
	BirthDate time.Time `excel:"dob,format:02/01/2006"`
	// map the column 'salary'
	Salary int `excel:"salary"`
	// Slice of staff ID's
	// split will split the string into slice using `|` separator
	Staff []int `excel:"staff,split:|"`
	// Slice of pointer of staff ID's
	StaffPtr []*int `excel:"staff,split:|"`
	// 'FullName' column contains a json string
	EncodedName EncodedName `excel:"encodedName,encoding:json"`
	// 'FullName' column contains a json string
	EncodedNamePtr *EncodedName `excel:"encodedName,encoding:json"`
	// use '-' to ignore.
	Ignored string `excel:"-"`
}

type EncodedName struct {
	FirstName string `json:"first_name"`
	LastName  string `json:"last_name"`
	FullName  string `json:"full_name"`
}

func TestReadEmployees(t *testing.T) {
	// Open the employees test file
	file, _ := excelize.OpenFile(employeesTestFile)
	defer func() { _ = file.Close() }()

	// Employees container
	var employees []Employee

	// Configure what to read in the Excel file
	excel, _ := excel.NewReaderConfig(file)
	excel.SetSheetName(employeesSheet)
	excel.SetAxis(employeesAxis)

	// Unmarshal employees
	err := excel.Unmarshal(&employees)
	if err != nil {
		t.Error(err)
		return
	}
}
