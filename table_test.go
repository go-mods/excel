package excel

import (
	"github.com/stretchr/testify/assert"
	"github.com/xuri/excelize/v2"
	"testing"
)

func TestGetTable(t *testing.T) {
	xls := Excel{File: excelize.NewFile()}

	t.Run("NoTable", func(t *testing.T) {
		_, err := xls.GetTable("Table1")
		assert.Error(t, err)
	})

	assert.NoError(t, xls.File.AddTable("Sheet1", &excelize.Table{
		Range: "A1:B4",
		Name:  "Table1",
	}))

	t.Run("GetTables", func(t *testing.T) {
		tables, err := xls.GetTables()
		assert.NoError(t, err)
		assert.Len(t, tables, 1)
		assert.Equal(t, "Table1", tables[0].Name)
	})

	t.Run("GetTable", func(t *testing.T) {
		table, err := xls.GetTable("Table1")
		assert.NoError(t, err)
		assert.Equal(t, "Table1", table.Name)
		assert.Equal(t, "Sheet1", table.Sheet.Name)
	})

	t.Run("GetTableSheet", func(t *testing.T) {
		sheet, err := xls.GetTableSheet("Table1")
		assert.NoError(t, err)
		assert.Equal(t, "Sheet1", sheet.Name)
	})
}

func TestTables(t *testing.T) {
	xls := Excel{}
	assert.Error(t, xls.AddTable(nil))
	assert.Error(t, xls.DeleteTable(""))
	assert.Error(t, xls.DeleteTable("Table1"))

	xls.File = excelize.NewFile()
	assert.Error(t, xls.AddTable(nil))
	assert.Error(t, xls.AddTable(&Table{}))
	assert.Error(t, xls.AddTable(&Table{Table: &excelize.Table{}}))
	assert.Error(t, xls.AddTable(&Table{Table: &excelize.Table{Name: "Table1"}}))
	assert.Error(t, xls.AddTable(&Table{Table: &excelize.Table{Name: "Table1", Range: "A1:B2"}}))
	assert.Error(t, xls.AddTable(&Table{Table: &excelize.Table{Name: "Table1", Range: "A1:B2"}, Sheet: &Sheet{}}))
	assert.NoError(t, xls.AddTable(&Table{Table: &excelize.Table{Name: "Table1", Range: "A1:B2"}, Sheet: &Sheet{Name: "Sheet1"}}))

	assert.NoError(t, xls.DeleteTable("Table1"))
}

func TestRange(t *testing.T) {
	xls := Excel{File: excelize.NewFile()}
	assert.NoError(t, xls.File.AddTable("Sheet1", &excelize.Table{
		Range: "A1:B4",
		Name:  "Table1",
	}))
	var table *Table

	t.Run("GetTable", func(t *testing.T) {
		tbl, err := xls.GetTable("Table1")
		assert.NoError(t, err)
		table = tbl
	})

	t.Run("GetTableRange", func(t *testing.T) {
		tRange, err := table.GetRange()
		assert.NoError(t, err)
		assert.Equal(t, "A1:B4", tRange.ToRef())
	})

	t.Run("GetHeaderRange", func(t *testing.T) {
		hRange, err := table.GetHeaderRange()
		assert.NoError(t, err)
		assert.Equal(t, "A1:B1", hRange.ToRef())
	})

	t.Run("GetDataRange", func(t *testing.T) {
		dRange, err := table.GetDataRange()
		assert.NoError(t, err)
		assert.Equal(t, "A2:B4", dRange.ToRef())
	})

}
