package excel

const (
	mainKey = "excel"
	inKey   = mainKey + "-in"
	outKey  = mainKey + "-out"

	columnTag   = "column"
	defaultTag  = "default"
	formatTag   = "format"
	encodingTag = "encoding"
	splitTag    = "split"
	requiredTag = "required"
	ignoreTag   = "-"
)

// FieldTags which can be used when reading or writing
type FieldTags struct {
	ColumnName   string
	columnIndex  int
	DefaultValue interface{}
	Format       string
	Encoding     string
	Split        string
	IsRequired   bool
	Ignore       bool
}

// The FieldsTags interface can be used as a replacement of the tags parameters.
type FieldsTags interface {
	GetFieldsTags() map[string]*FieldTags
}

// The FieldsTagsIn interface can be used as a replacement of the tags parameters when importing an Excel file.
type FieldsTagsIn interface {
	GetFieldsTagsIn() map[string]*FieldTags
}

// The FieldsTagsOut interface can be used as a replacement of the tags parameters when exporting an Excel file.
type FieldsTagsOut interface {
	GetFieldsTagsOut() map[string]*FieldTags
}

func newTags() *FieldTags {
	tags := &FieldTags{
		columnIndex: -1,
		IsRequired:  false,
		Ignore:      false,
	}
	return tags
}

func (f *FieldInfo) ColumnNameIn() string {
	if len(f.TagsIn.ColumnName) > 0 {
		return f.TagsIn.ColumnName
	}
	return f.Tags.ColumnName
}

func (f *FieldInfo) ColumnNameOut() string {
	if len(f.TagsOut.ColumnName) > 0 {
		return f.TagsOut.ColumnName
	}
	return f.Tags.ColumnName
}

func (f *FieldInfo) DefaultValueIn() interface{} {
	if f.TagsIn.DefaultValue != nil {
		return f.TagsIn.DefaultValue
	}
	return f.Tags.DefaultValue
}

func (f *FieldInfo) DefaultValueOut() interface{} {
	if f.TagsOut.DefaultValue != nil {
		return f.TagsOut.DefaultValue
	}
	return f.Tags.DefaultValue
}

func (f *FieldInfo) FormatIn() string {
	if len(f.TagsIn.Format) > 0 {
		return f.TagsIn.Format
	}
	return f.Tags.Format
}

func (f *FieldInfo) FormatOut() string {
	if len(f.TagsOut.Format) > 0 {
		return f.TagsOut.Format
	}
	return f.Tags.Format
}

func (f *FieldInfo) EncodingIn() string {
	if len(f.TagsIn.Encoding) > 0 {
		return f.TagsIn.Encoding
	}
	return f.Tags.Encoding
}

func (f *FieldInfo) EncodingOut() string {
	if len(f.TagsOut.Encoding) > 0 {
		return f.TagsOut.Encoding
	}
	return f.Tags.Encoding
}

func (f *FieldInfo) SplitIn() string {
	if len(f.TagsIn.Split) > 0 {
		return f.TagsIn.Split
	}
	return f.Tags.Split
}

func (f *FieldInfo) SplitOut() string {
	if len(f.TagsOut.Split) > 0 {
		return f.TagsOut.Split
	}
	return f.Tags.Split
}

func (f *FieldInfo) IsRequiredIn() bool {
	if f.TagsIn.IsRequired {
		return f.TagsIn.IsRequired
	}
	return f.Tags.IsRequired
}

func (f *FieldInfo) IsRequiredOut() bool {
	if f.TagsOut.IsRequired {
		return f.TagsOut.IsRequired
	}
	return f.Tags.IsRequired
}

func (f *FieldInfo) IgnoreIn() bool {
	if f.TagsIn.Ignore {
		return f.TagsIn.Ignore
	}
	return f.Tags.Ignore
}

func (f *FieldInfo) IgnoreOut() bool {
	if f.TagsOut.Ignore {
		return f.TagsOut.Ignore
	}
	return f.Tags.Ignore
}
