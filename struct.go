package excel

import (
	"github.com/go-mods/tags"
	"reflect"
)

// Struct is a struct used to store information about a struct
type Struct struct {
	Type   reflect.Type
	Fields Fields
}

// getStruct returns a Struct with the information of the struct
// contained in the container
func getStruct(container *Container) *Struct {

	// Get struct type
	t := container.Type
	if container.Pointer {
		t = t.Elem()
	}

	// Create struct and get fields
	s := &Struct{Type: t}
	s.Fields = getFields(s)

	return s
}

// getFields get the tags of the struct fields
func (s *Struct) getTags(field reflect.StructField, key string) (t *Tags) {
	if ts := tags.Lookup(field, key); ts != nil {
		t = s.parseTag(ts)
	} else {
		t = newTag()
		if key == mainKey {
			t.Column = field.Name
		}
	}
	return
}

// parseTag parse the tag and return a Tags
func (s *Struct) parseTag(tag *tags.Tag) (t *Tags) {
	t = newTag()

	if tag.Name == ignoreTag {
		t.Ignore = true
		return
	}

	if len(tag.Name) > 0 {
		t.Column = tag.Name
	}

	if o := tag.GetOption(columnTag); o != nil {
		t.Column = o.Value
	}
	if o := tag.GetOption(defaultTag); o != nil {
		t.Default = o.Value
	}
	if o := tag.GetOption(formatTag); o != nil {
		t.Format = o.Value
	}
	if o := tag.GetOption(encodingTag); o != nil {
		t.Encoding = o.Value
	}
	if o := tag.GetOption(splitTag); o != nil {
		if len(o.Value) != 0 {
			t.Split = o.Value
		}
	}
	if o := tag.GetOption(requiredTag); o != nil {
		t.Required = true
	}

	return t
}

// freeze copy the tags from one Tags to another
func (s *Struct) freeze(from *Tags, to *Tags) {
	if from != nil && to == nil {
		to = newTag()
	}
	if from != nil && to != nil {
		to.Column = from.Column
		to.Default = from.Default
		to.Format = from.Format
		to.Encoding = from.Encoding
		to.Split = from.Split
		to.Required = from.Required
		to.Ignore = from.Ignore
	}
}

// GetField returns the field from the index
func (s *Struct) GetField(index int) *Field {
	for _, f := range s.Fields {
		if f.Index == index {
			return f
		}
	}
	return nil
}
