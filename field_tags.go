package excel

// GetReadColumnName returns the column name to read from the excel file
func (f *Field) GetReadColumnName() string {
	if len(f.ReadTags.Column) > 0 {
		return f.ReadTags.Column
	}
	return f.MainTags.Column
}

// GetReadDefault returns the default value to use if the cell is empty
func (f *Field) GetReadDefault() interface{} {
	if f.ReadTags.Default != nil {
		return f.ReadTags.Default
	}
	return f.MainTags.Default
}

// GetReadFormat returns the format to use when reading the cell
func (f *Field) GetReadFormat() string {
	if len(f.ReadTags.Format) > 0 {
		return f.ReadTags.Format
	}
	return f.MainTags.Format
}

// GetReadEncoding returns the encoding to use when reading the cell
func (f *Field) GetReadEncoding() string {
	if len(f.ReadTags.Encoding) > 0 {
		return f.ReadTags.Encoding
	}
	return f.MainTags.Encoding
}

// GetReadSplit returns the split to use when reading the cell
func (f *Field) GetReadSplit() string {
	if len(f.ReadTags.Split) > 0 {
		return f.ReadTags.Split
	}
	return f.MainTags.Split
}

// GetReadRequired returns whether the field is required when reading the cell
func (f *Field) GetReadRequired() bool {
	if f.ReadTags.Required {
		return f.ReadTags.Required
	}
	return f.MainTags.Required
}

// GetReadIgnore returns whether the field should be ignored when reading the cell
func (f *Field) GetReadIgnore() bool {
	if f.ReadTags.Ignore {
		return f.ReadTags.Ignore
	}
	return f.MainTags.Ignore
}

// GetWriteColumnName returns the column name to write to the excel file
func (f *Field) GetWriteColumnName() string {
	if len(f.WriteTags.Column) > 0 {
		return f.WriteTags.Column
	}
	return f.MainTags.Column
}

// GetWriteDefault returns the default value to use if the cell is empty
func (f *Field) GetWriteDefault() interface{} {
	if f.WriteTags.Default != nil {
		return f.WriteTags.Default
	}
	return f.MainTags.Default
}

// GetWriteFormat returns the format to use when writing the cell
func (f *Field) GetWriteFormat() string {
	if len(f.WriteTags.Format) > 0 {
		return f.WriteTags.Format
	}
	return f.MainTags.Format
}

// GetWriteEncoding returns the encoding to use when writing the cell
func (f *Field) GetWriteEncoding() string {
	if len(f.WriteTags.Encoding) > 0 {
		return f.WriteTags.Encoding
	}
	return f.MainTags.Encoding
}

// GetWriteSplit returns the split to use when writing the cell
func (f *Field) GetWriteSplit() string {
	if len(f.WriteTags.Split) > 0 {
		return f.WriteTags.Split
	}
	return f.MainTags.Split
}

// GetWriteRequired returns whether the field is required when writing the cell
func (f *Field) GetWriteRequired() bool {
	if f.WriteTags.Required {
		return f.WriteTags.Required
	}
	return f.MainTags.Required
}

// GetWriteIgnore returns whether the field should be ignored when writing the cell
func (f *Field) GetWriteIgnore() bool {
	if f.WriteTags.Ignore {
		return f.WriteTags.Ignore
	}
	return f.MainTags.Ignore
}
