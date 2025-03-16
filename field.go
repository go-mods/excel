package excel

import (
	"reflect"
)

// Field is a struct used to store the information of a field of a struct
type Field struct {
	Name  string
	Index int
	Type  reflect.Type

	MainTags  *Tags // Tags used by default
	ReadTags  *Tags // Tags for reading
	WriteTags *Tags // Tags for writing
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
	fields := make(Fields, 0)

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

	// Call the recursive function to collect all fields
	return collectFields(s, s.Type, 0, fields, defaultTags)
}

// collectFields recursively traverses all fields, including those of embedded structures
func collectFields(s *Struct, t reflect.Type, startIndex int, fields Fields, defaultTags struct {
	mainTags  map[string]*Tags
	readTags  map[string]*Tags
	writeTags map[string]*Tags
}) Fields {
	fieldsCount := t.NumField()

	for i := 0; i < fieldsCount; i++ {
		f := t.Field(i)

		// If it's an embedded structure (Anonymous), process its fields
		if f.Anonymous {
			fieldType := f.Type
			if fieldType.Kind() == reflect.Ptr {
				fieldType = fieldType.Elem()
			}

			// If it's a structure, process its fields recursively
			if fieldType.Kind() == reflect.Struct {
				// Get the list of fields from the structure
				anonymousFields := collectFields(s, fieldType, startIndex, fields, defaultTags)

				// Add the field to the list
				for j := 0; j < len(anonymousFields); j++ {
					field := anonymousFields[j]
					if field != nil {
						field.Index = startIndex
						fields = append(fields, field)
						startIndex++
					}
				}

				continue
			}
		}

		// Create the field
		field := &Field{
			Index:     startIndex,
			Type:      f.Type,
			Name:      f.Name,
			MainTags:  s.getTags(f, TagKeyMain),
			ReadTags:  s.getTags(f, TagKeyIn),
			WriteTags: s.getTags(f, TagKeyOut),
		}

		// Overwrite default tags
		s.freeze(defaultTags.mainTags[f.Name], field.MainTags)
		s.freeze(defaultTags.readTags[f.Name], field.ReadTags)
		s.freeze(defaultTags.writeTags[f.Name], field.WriteTags)

		// Add the field to the list
		fields = append(fields, field)
		startIndex++
	}

	return fields
}
