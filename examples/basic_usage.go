package main

import (
	"fmt"
	"log"
	"time"

	"github.com/go-mods/excel"
	"github.com/xuri/excelize/v2"
)

// Employee represents an employee record that can be marshalled to/from Excel
type Employee struct {
	ID            int       `excel:"id"`
	First         string    `excel:"firstName,default:anonymous"`
	Last          string    `excel:"lastName,default:anonymous"`
	Email         string    `excel:"email"`
	ContactNumber string    `excel:"contactNumber"`
	Age           int       `excel:"age"`
	BirthDate     time.Time `excel:"dob,format:2006-01-02,default:"`
	Salary        int       `excel:"salary"`
	Staff         []int     `excel:"staff,split:|" excel-out:"split:;"`
	Department    string    `excel:"department"`
}

func main() {
	// Example 1: Writing data to Excel
	writeExample()

	// Example 2: Reading data from Excel
	readExample()

	// Example 3: Using maps
	mapExample()
}

// writeExample demonstrates how to write Go structs to an Excel file
func writeExample() {
	// Create a new Excel file
	f := excelize.NewFile()
	defer func() { _ = f.Close() }()

	// Create a new Excel writer
	excelWriter, err := excel.NewWriter(f)
	if err != nil {
		log.Fatalf("Failed to create Excel writer: %v", err)
	}
	if excelWriter == nil {
		log.Fatalf("Excel writer is nil")
	}

	// Rename the default sheet
	err = excelWriter.SetActiveSheetName("Employees")
	if err != nil {
		log.Fatalf("Failed to set sheet name: %v", err)
	}

	// Set the axis
	excelWriter.SetAxis("D8")

	// Create sample data
	employees := []Employee{
		{
			ID:            1,
			First:         "John",
			Last:          "Doe",
			Email:         "john.doe@example.com",
			ContactNumber: "123-456-7890",
			Age:           30,
			BirthDate:     time.Date(1993, 5, 15, 0, 0, 0, 0, time.UTC),
			Salary:        50000,
			Staff:         []int{101, 102, 103},
			Department:    "Engineering",
		},
		{
			ID:            2,
			First:         "Jane",
			Last:          "Smith",
			Email:         "jane.smith@example.com",
			ContactNumber: "987-654-3210",
			Age:           28,
			BirthDate:     time.Date(1995, 8, 22, 0, 0, 0, 0, time.UTC),
			Salary:        55000,
			Staff:         []int{104, 105},
			Department:    "Marketing",
		},
	}

	// Marshal the data to Excel
	err = excelWriter.Marshal(&employees)
	if err != nil {
		log.Fatalf("Failed to marshal data: %v", err)
	}

	fmt.Printf("Successfully wrote data to Excel\n")

	// Save the Excel file
	if err := f.SaveAs("employees.gen.xlsx"); err != nil {
		log.Fatalf("Failed to save Excel file: %v", err)
	}

	fmt.Println("Successfully saved employees.xlsx")
}

// readExample demonstrates how to read data from an Excel file into Go structs
func readExample() {
	// Open the Excel file
	f, err := excelize.OpenFile("employees.gen.xlsx")
	if err != nil {
		log.Fatalf("Failed to open Excel file: %v", err)
	}
	defer func() { _ = f.Close() }()

	// Create a new Excel reader
	excelReader, err := excel.NewReader(f)
	if err != nil {
		log.Fatalf("Failed to create Excel reader: %v", err)
	}

	// Set the sheet to read from
	excelReader.SetSheetFromName("Employees")
	excelReader.SetAxis("D8")

	// Create a container for the data
	var employees []Employee

	// Unmarshal the Excel data into the container
	err = excelReader.Unmarshal(&employees)
	if err != nil {
		log.Fatalf("Failed to unmarshal data: %v", err)
	}

	// Print the data
	fmt.Println("Read employees from Excel:")
	for _, emp := range employees {
		fmt.Printf("ID: %d, Name: %s %s, Email: %s, Department: %s\n",
			emp.ID, emp.First, emp.Last, emp.Email, emp.Department)
	}
}

// mapExample demonstrates how to use maps with the Excel library
func mapExample() {
	// Create a new Excel file
	f := excelize.NewFile()
	defer func() { _ = f.Close() }()

	// Create a new Excel writer
	excelWriter, err := excel.NewWriter(f)
	if err != nil {
		log.Fatalf("Failed to create Excel writer: %v", err)
	}
	if excelWriter == nil {
		log.Fatalf("Excel writer is nil")
	}

	// Rename the default sheet
	sheetName := "Data"
	err = excelWriter.SetActiveSheetName(sheetName)
	if err != nil {
		log.Fatalf("Failed to set sheet name: %v", err)
	}

	// Set the axis
	excelWriter.SetAxis("A1")

	// Create sample data using a map
	data := []map[string]interface{}{
		{
			"Name":       "Product A",
			"Price":      19.99,
			"Quantity":   100,
			"InStock":    true,
			"UpdateDate": time.Now(),
		},
		{
			"Name":       "Product B",
			"Price":      29.99,
			"Quantity":   50,
			"InStock":    true,
			"UpdateDate": time.Now(),
		},
		{
			"Name":       "Product C",
			"Price":      9.99,
			"Quantity":   0,
			"InStock":    false,
			"UpdateDate": time.Now(),
		},
	}

	// Marshal the data to Excel
	err = excelWriter.Marshal(&data)
	if err != nil {
		log.Fatalf("Failed to marshal map data: %v", err)
	}

	fmt.Printf("Successfully wrote map data to Excel\n")

	// Save the Excel file
	if err := f.SaveAs("products.gen.xlsx"); err != nil {
		log.Fatalf("Failed to save Excel file: %v", err)
	}

	fmt.Println("Successfully saved products.xlsx")

	// Reading map data
	f2, err := excelize.OpenFile("products.gen.xlsx")
	if err != nil {
		log.Fatalf("Failed to open Excel file: %v", err)
	}
	defer func() { _ = f2.Close() }()

	excelReader, err := excel.NewReader(f2)
	if err != nil {
		log.Fatalf("Failed to create Excel reader: %v", err)
	}

	// Set the sheet to read from
	excelReader.SetSheetFromName("Data")
	excelReader.SetAxis("A1")

	var products []map[string]interface{}
	err = excelReader.Unmarshal(&products)
	if err != nil {
		log.Fatalf("Failed to unmarshal map data: %v", err)
	}

	fmt.Println("Read products from Excel:")
	for i, product := range products {
		fmt.Printf("Product %d: %s, Price: %.2f, In Stock: %v\n",
			i+1, product["Name"], product["Price"], product["InStock"])
	}
}
