# excel

Marshal and Unmarshal Excel file with the help of [excelize](https://github.com/qax-os/excelize).

## Installation

```shell
go get github.com/go-mods/excel
```

## Usage

Consider the following structs

```go
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
```


### Unmarshal Excel file to struct

```go
package main

import (
	"github.com/go-mods/excel"
	"github.com/xuri/excelize/v2"
)

func main() {
    // Open the employees test file
    file, _ := excelize.OpenFile(employeesTestFile)
    defer func() { _ = file.Close() }()
    
    // Employees container
    var employees []*Employee

    // Configure what to read in the Excel file
    excel, _ := excel.NewReader(file)
    excel.SetSheetName(employeesSheet)
    excel.SetAxis(employeesAxis)

	// Unmarshal employees
    err := excel.Unmarshal(&employees)
    if err != nil {
        t.Error(err)
        return
    }
}
```

### Marshal Excel file from struct

```go
package main

import (
	"github.com/go-mods/excel"
	"github.com/xuri/excelize/v2"
)

func main() {
    // Create a new Excel file
    file := excelize.NewFile()
    file.SetSheetName(file.GetSheetName(file.GetActiveSheetIndex()), employeesSheet)
    defer func() { _ = file.Close() }()

    // Employees container
    var employees []*Employee
    employee1 := &Employee{ID: 1, First: "First", Last: "last", Email: "test@test.com", BirthDate: time.Now()}
    employee2 := &Employee{ID: 2, First: "Second", Last: "last", Salary: 100, EncodedName: EncodedName{FirstName: "Second", LastName: "last", FullName: "Second last"}}
    employee3 := &Employee{ID: 3, BirthDate: time.Now(), Age: 32, Staff: []int{1, 2}}
    employees = append(employees, employee1, employee2, employee3)

    // Configure how to write in the Excel file
    excel, _ := excel.NewWriter(file)
    excel.SetSheetName(employeesSheet)
    excel.SetAxis(employeesAxis)
    
    // Unmarshal employees
    err := excel.Marshal(&employees)
        if err != nil {
        t.Error(err)
        return
    }
    
    // Save file
    _ = file.SaveAs(employeesExportFile)
}
```

## Customizable Converters
```go
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

type User struct {
    Id       int       `excel:"Id"`
    Name     string    `excel:"Name"`
    Created  DateTime  `excel:"Created"`
    Modified *DateTime `excel:"Modified"`
}
```

## Tags

This is the list of tags that can be used.
`excel-in` and `excel-out` always have precedence on `excel`

| Tags       | description                                                                                                  | excel | excel-in | excel-out | 
|------------|--------------------------------------------------------------------------------------------------------------|:-----:|:--------:|:---------:|
| column     | Field name in the title row.<br/>`by default the field name will be used`<br/>`in and out can be differents` | **X** |  **X**   |   **X**   |
| default    | Default value to use when none is defined in the cell.                                                       | **X** |  **X**   |           |
| format     | Format to apply                                                                                              | **X** |  **X**   |   **X**   |
| encoding   | Encode or decode to the specified format<br/>`only json encoding is supported at the moment`                 | **X** |  **X**   |   **X**   |
| split      | Define the split separator to use for array or slice field.                                                  | **X** |  **X**   |   **X**   |
| required   | Will return ann error if the column is not present                                                           | **X** |  **X**   |           |
| -          | Do not map the field to a column                                                                             | **X** |  **X**   |   **X**   |

