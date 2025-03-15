package excel_test

import (
	"path/filepath"
	"testing"
	"time"

	"github.com/go-mods/excel"
	"github.com/xuri/excelize/v2"
)

var (
	employeesTestFile   = filepath.Join("test", "employees.xlsx")
	employeesExportFile = filepath.Join("test", "employees.gen.xlsx")
	employeesSheet      = "Employees"
	employeesAxis       = "A1"
)

type Employee struct {
	// use field name as default column name
	ID int
	// map the column 'firstName'
	First string `excel:"firstName,default:anonymous"`
	// column can also be used to set the column name
	FirstPtr *string `excel:"column:firstName,export:false" excel-out:"-"`
	// map the column 'lastName'
	Last string `excel:"lastName,default:anonymous"`
	// 'column' can be omitted when only mapping to column name
	// this is equal to 'column(email)'
	Email string `excel:"email"`
	// map the column 'contactNumber'
	ContactNumber string `excel:"contactNumber"`
	// map the column 'age'
	Age int `excel:"age"`
	// map the column 'dob'
	BirthDate time.Time `excel:"dob,format:d/m/Y,default:"`
	// map the column 'salary'
	Salary int `excel:"salary"`
	// Slice of staff ID's
	// split will split the string into slice using `|` separator
	Staff []int `excel:"staff,split:|" excel-out:"split:;"`
	// Slice of pointer of staff ID's
	StaffPtr []*int `excel:"staff,split:|,export:false" excel-out:"-"`
	// 'FullName' column contains a json string
	EncodedName EncodedName `excel:"encodedName,encoding:json"`
	// 'FullName' column contains a json string
	EncodedNamePtr *EncodedName `excel:"encodedName,encoding:json,export:false" excel-out:"-"`
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
	var employees []*Employee

	// Configure what to read in the Excel file
	xl, _ := excel.NewReader(file)
	xl.SetSheet(xl.GetSheet(employeesSheet))
	xl.SetAxis(employeesAxis)

	// Unmarshal employees
	err := xl.Unmarshal(&employees)
	if err != nil {
		t.Error(err)
		return
	}
}

func TestWriteEmployeesCustomOutput(t *testing.T) {
	// Create a new Excel file
	file := excelize.NewFile()
	_ = file.SetSheetName(file.GetSheetName(file.GetActiveSheetIndex()), employeesSheet)
	defer func() { _ = file.Close() }()

	// Employees container
	var employees []*Employee
	employee1 := &Employee{ID: 1, First: "First", Last: "last", Email: "test@test.com", BirthDate: time.Now()}
	employee2 := &Employee{ID: 2, First: "Second", Last: "last", Salary: 100, EncodedName: EncodedName{FirstName: "Second", LastName: "last", FullName: "Second last"}}
	employee3 := &Employee{ID: 3, BirthDate: time.Now(), Age: 32, Staff: []int{1, 2}}
	employees = append(employees, employee1, employee2, employee3)

	// Configure how to write in the Excel file
	xl, _ := excel.NewWriter(file)
	xl.SetSheet(xl.GetSheet(employeesSheet))
	xl.SetAxis(employeesAxis)

	// Set custom output
	var customOutput = map[string]*excel.Tags{}
	customOutput["ID"] = &excel.Tags{Column: "id"}
	customOutput["First"] = &excel.Tags{Ignore: true}
	customOutput["Last"] = &excel.Tags{Ignore: true}
	customOutput["ContactNumber"] = &excel.Tags{Ignore: true}
	customOutput["Age"] = &excel.Tags{Ignore: true}
	customOutput["BirthDate"] = &excel.Tags{Ignore: true}
	customOutput["Salary"] = &excel.Tags{Ignore: true}
	customOutput["Staff"] = &excel.Tags{Ignore: true}
	customOutput["EncodedName"] = &excel.Tags{Ignore: true}

	// Unmarshal employees
	err := xl.Marshal(&employees, customOutput)
	if err != nil {
		t.Error(err)
		return
	}

	// Save file
	_ = file.SaveAs(employeesExportFile)

	// Configure what to read in the Excel file
	fileRead, _ := excelize.OpenFile(employeesExportFile)
	defer func() { _ = fileRead.Close() }()
	xl, _ = excel.NewReader(fileRead)
	xl.SetSheet(xl.GetSheet(employeesSheet))
	xl.SetAxis(employeesAxis)

	// Read the file
	var employeesRead []*Employee
	var customInput = customOutput
	err = xl.Unmarshal(&employeesRead, customInput)
	if err != nil {
		t.Error(err)
		return
	}
}

func TestWriteEmployees(t *testing.T) {
	// Create a new Excel file
	file := excelize.NewFile()
	_ = file.SetSheetName(file.GetSheetName(file.GetActiveSheetIndex()), employeesSheet)
	defer func() { _ = file.Close() }()

	// Employees container
	var employees []*Employee
	employee1 := &Employee{ID: 1, First: "First", Last: "last", Email: "test@test.com", BirthDate: time.Now()}
	employee2 := &Employee{ID: 2, First: "Second", Last: "last", Salary: 100, EncodedName: EncodedName{FirstName: "Second", LastName: "last", FullName: "Second last"}}
	employee3 := &Employee{ID: 3, BirthDate: time.Now(), Age: 32, Staff: []int{1, 2}}
	employees = append(employees, employee1, employee2, employee3)

	// Configure how to write in the Excel file
	xl, _ := excel.NewWriter(file)
	xl.SetSheet(xl.GetSheet(employeesSheet))
	xl.SetAxis(employeesAxis)

	// Unmarshal employees
	err := xl.Marshal(&employees)
	if err != nil {
		t.Error(err)
		return
	}

	// Save file
	_ = file.SaveAs(employeesExportFile)
}
