package excel

const (
	TagKeyMain = "excel"
	TagKeyIn   = TagKeyMain + "-in"
	TagKeyOut  = TagKeyMain + "-out"

	TagColumn   = "column"
	TagDefault  = "default"
	TagFormat   = "format"
	TagEncoding = "encoding"
	TagSplit    = "split"
	TagRequired = "required"
	TagIgnore   = "-"
)

// Tags is used to store the mainTags parameters of a field.
//
// The mainTags parameters are defined in the struct definition and are prefixed by "excel"
// and are used to configure the import and export of an Excel file.
//
// Example:
//
//	type Named struct {
//		Column1 string `excel:"column=MyColumn1"`
//		Column2 string `excel:"column=MyColumn2;required"`
//		Column3 string `excel:"column=MyColumn3;default=Hello World"`
//	}
//
// In this example:
// the Column1 field will be mapped to the "MyColumn1" column of the Excel file.
// The Column2 field will be mapped to the "MyColumn2" column of the Excel file and it will be required.
// The Column3 field will be mapped to the "MyColumn3" column of the Excel file and it will have a default value of "Hello World".
type Tags struct {
	Column   string
	Default  interface{}
	Format   string
	Encoding string
	Split    string
	Required bool
	Ignore   bool

	// internal
	index int // The index of the column in the Excel file.
}

// The ITags interface can be used as a replacement of the mainTags parameters.
// The GetTags method must return a map of the mainTags.
// The key of the map is the name of the field and the value is a Tags structure.
//
// Example:
//
//	type Named struct {
//		Column1 string `excel:"column=MyColumn1"`
//		Column2 string `excel:"column=MyColumn2;required"`
//		Column3 string `excel:"column=MyColumn3;default=Hello World"`
//	}
//
//	func (s *Named) GetTags() map[string]excel.MainTags {
//		return map[string]excel.MainTags{
//			"Column1": excel.MainTags{column: "MyColumn1"},
//			"Column2": excel.MainTags{column: "MyColumn2", Required: true},
//			"Column3": excel.MainTags{column: "MyColumn3", Default: "Hello World"},
//		}
//	}
//
// In this example:
// the Column1 field will be mapped to the "MyColumn1" column of the Excel file.
// The Column2 field will be mapped to the "MyColumn2" column of the Excel file and it will be required.
// The Column3 field will be mapped to the "MyColumn3" column of the Excel file and it will have a default value of "Hello World".
type ITags interface {
	GetTags() map[string]*Tags
}

// The IReadTags interface can be used as a replacement of the mainTags parameters when importing an Excel file.
//
// GetTagsIn is used when importing an Excel file and will be used if ITags is not implemented.
type IReadTags interface {
	GetReadTags() map[string]*Tags
}

// The IWriteTags interface can be used as a replacement of the mainTags parameters when exporting an Excel file.
//
// GetTagsOut is used when exporting an Excel file and will be used if ITags is not implemented.
type IWriteTags interface {
	GetWriteTags() map[string]*Tags
}

// newTag returns a Tags structure with default values.
func newTag() *Tags {
	tag := &Tags{
		index:    -1,
		Required: false,
		Ignore:   false,
	}
	return tag
}
