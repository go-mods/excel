package excel

import (
	"reflect"
)

// Field is a struct used to store the information of a field of a struct
type Field struct {
	Name  string
	Index int
	Type  reflect.Type

	MainTags  *Tags // mainTags used by default
	ReadTags  *Tags // mainTags for reading
	WriteTags *Tags // mainTags for writing
}

// Marshaller can be implemented by any Value that has a Marshal method
// This converter is used to convert the Value to the desired representation
type Marshaller interface {
	Marshall() (interface{}, error)
}

// Unmarshaller can be implemented by any Value that has an Unmarshall method
// This converter is used to convert the Value to the desired representation
type Unmarshaller interface {
	Unmarshall(s string) error
}

// getFields returns a list of Field from the Struct
func getFields(s *Struct) Fields {

	fieldsCount := s.Type.NumField()
	fields := make(Fields, 0, fieldsCount)

	// Check if the Container implement ITags, IReadTags or IWriteTags interface
	// -------------------------------------------------------------------------
	type DefaultTags struct {
		mainTags  map[string]*Tags
		readTags  map[string]*Tags
		writeTags map[string]*Tags
	}
	var defaultTags DefaultTags

	v := reflect.New(s.Type)

	if v.CanInterface() {
		if i, ok := v.Interface().(ITags); ok {
			defaultTags.mainTags = i.GetTags()
		}
		if i, ok := v.Interface().(IReadTags); ok {
			defaultTags.readTags = i.GetReadTags()
		}
		if i, ok := v.Interface().(IWriteTags); ok {
			defaultTags.writeTags = i.GetWriteTags()
		}
	} else if vElem := v.Elem(); vElem.CanInterface() {
		if i, ok := vElem.Interface().(ITags); ok {
			defaultTags.mainTags = i.GetTags()
		}
		if i, ok := vElem.Interface().(IReadTags); ok {
			defaultTags.readTags = i.GetReadTags()
		}
		if i, ok := vElem.Interface().(IWriteTags); ok {
			defaultTags.writeTags = i.GetWriteTags()
		}
	}

	// Loop throw each field of the Container to get each field configuration
	// ----------------------------------------------------------------------
	for i := 0; i < fieldsCount; i++ {

		f := s.Type.Field(i)

		// Get Tags from struct
		field := &Field{
			Index:     i,
			Type:      f.Type,
			Name:      f.Name,
			MainTags:  s.getTags(f, mainKey),
			ReadTags:  s.getTags(f, inKey),
			WriteTags: s.getTags(f, outKey),
		}

		// Overwrite default tags
		s.freeze(defaultTags.mainTags[f.Name], field.MainTags)
		s.freeze(defaultTags.readTags[f.Name], field.ReadTags)
		s.freeze(defaultTags.writeTags[f.Name], field.WriteTags)

		// Add the field to the list
		fields = append(fields, field)
	}

	return fields
}
