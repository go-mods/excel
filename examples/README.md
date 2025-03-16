# Excel Library Usage Examples

This directory contains examples of how to use the [go-mods/excel](https://github.com/go-mods/excel) library to read and write Excel files in Go.

## Basic Example

The `basic/basic_usage.go` file demonstrates how to:

1. Write Go structures to an Excel file
2. Read data from an Excel file into Go structures
3. Use maps to read and write Excel data

### Running the Example

```bash
cd basic
go run basic_usage.go
```

This will create two Excel files:
- `employees.xlsx`: Contains employee data
- `products.xlsx`: Contains product data


## Features Demonstrated

- Creating new Excel files
- Reading from existing Excel files
- Using tags to customize column mapping
- Automatic type conversion
- Default value handling
- Date formatting
- Using slices and maps
- Handling separators for slice fields

## API Structure

The excel library API is organized as follows:

1. `excel.NewWriter(file)`: Creates a new Excel writer
2. `excel.NewReader(file)`: Creates a new Excel reader
3. `excelWriter.Marshal(data)`: Writes data to an Excel file
4. `excelReader.Unmarshal(&container)`: Reads data from an Excel file into a container


## Additional Examples

For more advanced examples, check the tests in the main project directory. 
