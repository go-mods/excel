package excel

import (
	"reflect"

	"github.com/go-mods/convert"
	"github.com/go-mods/tags"
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

// getTags gets the tags of the struct fields
func (s *Struct) getTags(field reflect.StructField, key string) (t *Tags) {
	if ts := tags.Lookup(field, key); ts != nil {
		t = s.parseTag(ts)
	} else {
		t = newTag()
		if key == TagKeyMain {
			t.Column = field.Name
		}
	}
	return
}

// parseTag parses the tag and returns a Tags
func (s *Struct) parseTag(tag *tags.Tag) (t *Tags) {
	t = newTag()

	if tag.Value == TagIgnore {
		t.Ignore = true
		return
	}

	if len(tag.Name) > 0 {
		t.Column = tag.Name
	}

	if o := tag.GetOption(TagColumn); o != nil {
		t.Column = convert.ToString(o.Value)
	}
	if o := tag.GetOption(TagDefault); o != nil {
		t.Default = o.Value
	}
	if o := tag.GetOption(TagFormat); o != nil {
		t.Format = convert.ToString(o.Value)
	}
	if o := tag.GetOption(TagEncoding); o != nil {
		t.Encoding = convert.ToString(o.Value)
	}
	if o := tag.GetOption(TagSplit); o != nil {
		if o.Value != nil {
			t.Split = convert.ToString(o.Value)
		}
	}
	if o := tag.GetOption(TagRequired); o != nil {
		t.Required = true
	}

	return t
}

// freeze copies the tags from one Tags to another
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
