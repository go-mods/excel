package excel

// Fields is a list of Field
type Fields []*Field

// Count returns the number of fields
func (f *Fields) Count() int {
	return len(*f)
}

// CountReadIgnored returns the number of ignored fields
func (f *Fields) CountReadIgnored() int {
	var count int
	for _, field := range *f {
		if field.GetReadIgnore() {
			count++
		}
	}
	return count
}

// CountWriteIgnored returns the number of ignored fields
func (f *Fields) CountWriteIgnored() int {
	var count int
	for _, field := range *f {
		if field.GetWriteIgnore() {
			count++
		}
	}
	return count
}
