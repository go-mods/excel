package excel

import (
	"github.com/go-mods/convert"
	"github.com/go-mods/tags"
	"reflect"
	"strings"
	"time"
)

const (
	tagIdentify = "excel"

	columnTag   = "column"
	defaultTag  = "default"
	encodingTag = "encoding"
	formatTag   = "format"
	ignoreTag   = "-"
	nilTag      = "nil"
	reqTag      = "required"
	splitTag    = "split"

	defaultSplitSeparator = ";"
)

// The FieldConfig struct hold all possibles tags for a field
type FieldConfig struct {
	// The config equals to tag: column
	ColumnName string
	// The config equals to tag: format
	Format string
	// The config equals to tag: default
	DefaultValue string
	// The config equals to tag: split
	Split string
	// The config equals to tag: decode
	Encoding string
	// The config equals to tag: nil
	// if cell.value == NilValue, will skip fc scan
	NilValue string
	// The config equals to tag: req
	// panic if required fc column but not set
	IsRequired bool
	// The config equals to tag: -
	Ignore bool
}

// The FieldsConfig interface can be used as a replacement
// of the tags parameters.
type FieldsConfig interface {
	GetFieldsTag() map[string]FieldConfig
}

type currentFieldConfig struct {
	FieldIndex int
	FieldType  reflect.Type

	ColumnName  string
	ColumnIndex int

	Format       string
	DefaultValue string
	Split        string
	Encoding     string
	NilValue     string
	IsRequired   bool
}

type schema struct {
	Type   reflect.Type
	Fields []*currentFieldConfig
}

var timeType = reflect.TypeOf((*time.Time)(nil)).Elem()

func newSchema(t reflect.Type) *schema {
	s := &schema{
		Fields: make([]*currentFieldConfig, 0, t.NumField()),
	}

	// Check if the container implement FieldsConfig interface
	var defaultFieldsConfig map[string]FieldConfig
	v := reflect.New(t)
	if v.CanInterface() {
		if i, ok := v.Interface().(FieldsConfig); ok {
			defaultFieldsConfig = i.GetFieldsTag()
		}
	} else if vElem := v.Elem(); vElem.CanInterface() {
		if i, ok := vElem.Interface().(FieldsConfig); ok {
			defaultFieldsConfig = i.GetFieldsTag()
		}
	}

	// Loop throw each field of the container to get each field configuration
	for i := 0; i < t.NumField(); i++ {
		field := t.Field(i)
		if selectedFieldConfig, ok := defaultFieldsConfig[field.Name]; ok {
			if !selectedFieldConfig.Ignore {
				frozenConfig := selectedFieldConfig.freeze(i)
				if frozenConfig.ColumnName == "" {
					frozenConfig.ColumnName = field.Name
				}
				frozenConfig.FieldType = field.Type
				s.Fields = append(s.Fields, frozenConfig)
			}
		} else if tag := tags.Lookup(field, tagIdentify); tag != nil {
			if tag.Key != ignoreTag {
				fieldConfig := s.parseTag(tag)
				fieldConfig.FieldIndex = i
				if fieldConfig.ColumnName == "" {
					fieldConfig.ColumnName = field.Name
				}
				fieldConfig.FieldType = field.Type
				s.Fields = append(s.Fields, fieldConfig)
			}
		} else {
			fieldConfig := &currentFieldConfig{
				FieldIndex: i,
				FieldType:  field.Type,
				ColumnName: field.Name,
			}
			s.Fields = append(s.Fields, fieldConfig)
		}
	}
	s.Type = t
	return s
}

func (s *schema) parseTag(t *tags.Tag) *currentFieldConfig {
	c := &currentFieldConfig{}
	c.Split = defaultSplitSeparator

	if len(t.Name) > 0 {
		c.ColumnName = t.Name
	}

	if o := t.GetOption(columnTag); o != nil {
		c.ColumnName = o.Value
	}
	if o := t.GetOption(defaultTag); o != nil {
		c.DefaultValue = o.Value
	}
	if o := t.GetOption(formatTag); o != nil {
		c.Format = o.Value
	}
	if o := t.GetOption(splitTag); o != nil {
		c.Split = o.Value
	}
	if o := t.GetOption(encodingTag); o != nil {
		c.Encoding = o.Value
	}
	if o := t.GetOption(nilTag); o != nil {
		c.NilValue = o.Value
	}
	if o := t.GetOption(reqTag); o != nil {
		c.IsRequired = true
	}

	return c
}

func (s *schema) GetFieldFromFieldIndex(index int) *currentFieldConfig {
	for _, f := range s.Fields {
		if f.FieldIndex == index {
			return f
		}
	}
	return nil
}

func (s *schema) GetFieldFromColumnName(name string) *currentFieldConfig {
	for _, f := range s.Fields {
		if f.ColumnName == name {
			return f
		}
	}
	return nil
}

func (s *schema) GetFieldFromColumnIndex(index int) *currentFieldConfig {
	for _, f := range s.Fields {
		if f.ColumnIndex == index {
			return f
		}
	}
	return nil
}

func (f *FieldConfig) freeze(fieldIdx int) *currentFieldConfig {
	return &currentFieldConfig{
		FieldIndex:   fieldIdx,
		ColumnName:   f.ColumnName,
		Format:       f.Format,
		DefaultValue: f.DefaultValue,
		Split:        f.Split,
		Encoding:     f.Encoding,
		NilValue:     f.NilValue,
		IsRequired:   f.IsRequired,
	}
}

func (f *currentFieldConfig) toValue(from string) (value reflect.Value, err error) {

	// Field of type Time
	if f.FieldType == timeType {
		dt, err := convert.ToLayoutTime(f.Format, from)
		if err != nil {
			return reflect.Value{}, nil
		}
		value = reflect.ValueOf(dt)
		return value, err
	}

	// Field of type Slice or Array
	if f.FieldType.Kind() == reflect.Slice || f.FieldType.Kind() == reflect.Array {
		values := strings.Split(from, f.Split)
		value = reflect.MakeSlice(reflect.SliceOf(f.FieldType.Elem()), 0, len(values))
		for _, vs := range values {
			v, err := f.decode(vs, f.FieldType.Elem())
			if err != nil {
				return reflect.Value{}, nil
			}

			value = reflect.Append(value, v)
		}
		return
	}

	// Field of type Pointer
	if f.FieldType.Kind() == reflect.Pointer {
		value, err = f.decode(from, f.FieldType)
		if err != nil {
			return reflect.Value{}, nil
		}
		return
	}

	value, err = f.decode(from, f.FieldType)
	if err != nil {
		return reflect.Value{}, nil
	}

	return
}

func (f *currentFieldConfig) decode(from string, to reflect.Type) (value reflect.Value, err error) {
	switch f.Encoding {
	case "json":
		value, err = convert.ToJsonValue(from, to)
	default:
		value, err = convert.ToValue(from, to)
	}
	return
}
