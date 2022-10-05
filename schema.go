package excel

import (
	"reflect"
	"strings"
	"time"

	"github.com/go-mods/convert"
	"github.com/go-mods/tags"
)

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

// The FieldTags struct hold all possibles tags for a field
type FieldTags struct {
	Tags    Tags
	TagsIn  Tags
	TagsOut Tags
}

type Tags struct {
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
	GetFieldsTags() map[string]*Tags
}

// The FieldsTagsIn interface can be used as a replacement of the tags parameters when importing an Excel file.
type FieldsTagsIn interface {
	GetFieldsTagsIn() map[string]*Tags
}

// The FieldsTagsOut interface can be used as a replacement of the tags parameters when exporting an Excel file.
type FieldsTagsOut interface {
	GetFieldsTagsOut() map[string]*Tags
}

type activeFieldTags struct {
	FieldIndex int
	FieldType  reflect.Type

	Tags    *Tags
	TagsIn  *Tags
	TagsOut *Tags
}

func (f *activeFieldTags) ColumnNameIn() string {
	if len(f.TagsIn.ColumnName) > 0 {
		return f.TagsIn.ColumnName
	}
	return f.Tags.ColumnName
}

func (f *activeFieldTags) ColumnNameOut() string {
	if len(f.TagsOut.ColumnName) > 0 {
		return f.TagsOut.ColumnName
	}
	return f.Tags.ColumnName
}

func (f *activeFieldTags) DefaultValueIn() interface{} {
	if f.TagsIn.DefaultValue != nil {
		return f.TagsIn.DefaultValue
	}
	return f.Tags.DefaultValue
}

func (f *activeFieldTags) DefaultValueOut() interface{} {
	if f.TagsOut.DefaultValue != nil {
		return f.TagsOut.DefaultValue
	}
	return f.Tags.DefaultValue
}

func (f *activeFieldTags) FormatIn() string {
	if len(f.TagsIn.Format) > 0 {
		return f.TagsIn.Format
	}
	return f.Tags.Format
}

func (f *activeFieldTags) FormatOut() string {
	if len(f.TagsOut.Format) > 0 {
		return f.TagsOut.Format
	}
	return f.Tags.Format
}

func (f *activeFieldTags) EncodingIn() string {
	if len(f.TagsIn.Encoding) > 0 {
		return f.TagsIn.Encoding
	}
	return f.Tags.Encoding
}

func (f *activeFieldTags) EncodingOut() string {
	if len(f.TagsOut.Encoding) > 0 {
		return f.TagsOut.Encoding
	}
	return f.Tags.Encoding
}

func (f *activeFieldTags) SplitIn() string {
	if len(f.TagsIn.Split) > 0 {
		return f.TagsIn.Split
	}
	return f.Tags.Split
}

func (f *activeFieldTags) SplitOut() string {
	if len(f.TagsOut.Split) > 0 {
		return f.TagsOut.Split
	}
	return f.Tags.Split
}

func (f *activeFieldTags) IsRequiredIn() bool {
	if f.TagsIn.IsRequired {
		return f.TagsIn.IsRequired
	}
	return f.Tags.IsRequired
}

func (f *activeFieldTags) IsRequiredOut() bool {
	if f.TagsOut.IsRequired {
		return f.TagsOut.IsRequired
	}
	return f.Tags.IsRequired
}

func (f *activeFieldTags) IgnoreIn() bool {
	if f.TagsIn.Ignore {
		return f.TagsIn.Ignore
	}
	return f.Tags.Ignore
}

func (f *activeFieldTags) IgnoreOut() bool {
	if f.TagsOut.Ignore {
		return f.TagsOut.Ignore
	}
	return f.Tags.Ignore
}

func newTags() *Tags {
	tags := &Tags{
		columnIndex: -1,
		IsRequired:  false,
		Ignore:      false,
	}
	return tags
}

type schema struct {
	Type   reflect.Type
	Fields []*activeFieldTags
}

func newSchema(t reflect.Type) *schema {
	s := &schema{
		Fields: make([]*activeFieldTags, 0, t.NumField()),
	}

	// Check if the container implement FieldsTags, FieldsTagsIn or FieldsTagsOut interface
	// ------------------------------------------------------------------------------------
	type defaultFieldsTags struct {
		tags    map[string]*Tags
		tagsIn  map[string]*Tags
		tagsOut map[string]*Tags
	}
	var defaultFieldsConfig defaultFieldsTags

	v := reflect.New(t)

	if v.CanInterface() {
		if i, ok := v.Interface().(FieldsTags); ok {
			defaultFieldsConfig.tags = i.GetFieldsTags()
		}
		if i, ok := v.Interface().(FieldsTagsIn); ok {
			defaultFieldsConfig.tagsIn = i.GetFieldsTagsIn()
		}
		if i, ok := v.Interface().(FieldsTagsOut); ok {
			defaultFieldsConfig.tagsOut = i.GetFieldsTagsOut()
		}
	} else if vElem := v.Elem(); vElem.CanInterface() {
		if i, ok := vElem.Interface().(FieldsTags); ok {
			defaultFieldsConfig.tags = i.GetFieldsTags()
		}
		if i, ok := vElem.Interface().(FieldsTagsIn); ok {
			defaultFieldsConfig.tagsIn = i.GetFieldsTagsIn()
		}
		if i, ok := vElem.Interface().(FieldsTagsOut); ok {
			defaultFieldsConfig.tagsOut = i.GetFieldsTagsOut()
		}
	}

	// Loop throw each field of the container to get each field configuration
	// ----------------------------------------------------------------------
	for i := 0; i < t.NumField(); i++ {

		field := t.Field(i)

		// Get tags from struct
		fieldTags := &activeFieldTags{
			FieldIndex: i,
			FieldType:  field.Type,
		}

		if tag := tags.Lookup(field, mainKey); tag != nil {
			if tag.Name != ignoreTag {
				fieldTags.Tags = s.parseTag(tag)
			} else {
				fieldTags.Tags = newTags()
				fieldTags.Tags.Ignore = true
			}
		}

		if tag := tags.Lookup(field, inKey); tag != nil {
			if tag.Name != ignoreTag {
				fieldTags.TagsIn = s.parseTag(tag)
			} else {
				fieldTags.TagsIn = newTags()
				fieldTags.TagsIn.Ignore = true
			}
		}

		if tag := tags.Lookup(field, outKey); tag != nil {
			if tag.Name != ignoreTag {
				fieldTags.TagsOut = s.parseTag(tag)
			} else {
				fieldTags.TagsOut = newTags()
				fieldTags.TagsOut.Ignore = true
			}
		}

		// Overwrite with tags from interfaces
		selectedFieldConfigMain := defaultFieldsConfig.tags[field.Name]
		selectedFieldConfigIn := defaultFieldsConfig.tagsIn[field.Name]
		selectedFieldConfigOut := defaultFieldsConfig.tagsOut[field.Name]

		if selectedFieldConfigMain != nil {
			s.freeze(selectedFieldConfigMain, fieldTags.Tags)
		}

		if selectedFieldConfigIn != nil {
			s.freeze(selectedFieldConfigIn, fieldTags.TagsIn)
		}

		if selectedFieldConfigOut != nil {
			s.freeze(selectedFieldConfigOut, fieldTags.TagsOut)
		}

		// default values if no tags are set
		if fieldTags.Tags == nil {
			fieldTags.Tags = newTags()
		}
		if fieldTags.Tags.ColumnName == "" {
			fieldTags.Tags.ColumnName = field.Name
		}
		if fieldTags.TagsIn == nil {
			fieldTags.TagsIn = newTags()
		}
		if fieldTags.TagsOut == nil {
			fieldTags.TagsOut = newTags()
		}

		//
		s.Fields = append(s.Fields, fieldTags)
	}
	s.Type = t
	return s
}

func (s *schema) parseTag(tg *tags.Tag) *Tags {
	tgs := newTags()

	if len(tg.Name) > 0 {
		tgs.ColumnName = tg.Name
	}

	if o := tg.GetOption(columnTag); o != nil {
		tgs.ColumnName = o.Value
	}
	if o := tg.GetOption(defaultTag); o != nil {
		tgs.DefaultValue = o.Value
	}
	if o := tg.GetOption(formatTag); o != nil {
		tgs.Format = o.Value
	}
	if o := tg.GetOption(encodingTag); o != nil {
		tgs.Encoding = o.Value
	}
	if o := tg.GetOption(splitTag); o != nil {
		if len(o.Value) != 0 {
			tgs.Split = o.Value
		}
	}
	if o := tg.GetOption(requiredTag); o != nil {
		tgs.IsRequired = true
	}

	return tgs
}

func (s *schema) freeze(from *Tags, to *Tags) {
	to.ColumnName = from.ColumnName
	to.DefaultValue = from.DefaultValue
	to.Format = from.Format
	to.Encoding = from.Encoding
	to.Split = from.Split
	to.IsRequired = from.IsRequired
}

func (s *schema) GetFieldFromFieldIndex(index int) *activeFieldTags {
	for _, f := range s.Fields {
		if f.FieldIndex == index {
			return f
		}
	}
	return nil
}

var timeType = reflect.TypeOf((*time.Time)(nil)).Elem()

// toValue is called when reading an Excel file
func (f *activeFieldTags) toValue(from string) (value reflect.Value, err error) {

	// Field of type Slice or Array
	if f.FieldType.Kind() == reflect.Slice || f.FieldType.Kind() == reflect.Array {
		if len(from) > 0 {
			values := strings.Split(convert.ToValidString(from), f.SplitIn())
			value = reflect.MakeSlice(reflect.SliceOf(f.FieldType.Elem()), 0, len(values))
			for _, vs := range values {
				v, err := f.decode(vs, f.FieldType.Elem())
				if err != nil {
					return reflect.Value{}, err
				}
				value = reflect.Append(value, v)
			}
		} else {
			return reflect.MakeSlice(reflect.SliceOf(f.FieldType.Elem()), 0, 0), nil
		}
		return
	}

	// Field of type Pointer
	if f.FieldType.Kind() == reflect.Pointer {
		value, err = f.decode(from, f.FieldType)
		if err != nil {
			return reflect.Value{}, err
		}
		return
	}

	// Decode the string
	value, err = f.decode(from, f.FieldType)
	if err != nil {
		return reflect.Value{}, err
	}

	return
}

// decode is called when reading an Excel file
func (f *activeFieldTags) decode(from string, to reflect.Type) (value reflect.Value, err error) {
	switch f.EncodingIn() {
	case "json":
		value, err = convert.ToJsonValue(from, to)
	default:
		if f.FieldType == timeType {
			dt, err := convert.ToLayoutTime(f.FormatIn(), from)
			if err != nil {
				return reflect.Value{}, nil
			}
			return reflect.ValueOf(dt), err
		} else {
			if len(from) == 0 {
				value = reflect.ValueOf(f.DefaultValueIn())
			} else {
				value, err = convert.ToValue(from, to)
			}
		}
	}
	return
}

// toCellValue is called when writing an Excel file
func (f *activeFieldTags) toCellValue(from interface{}) interface{} {

	// Field of type Slice or Array
	if f.FieldType.Kind() == reflect.Slice || f.FieldType.Kind() == reflect.Array {
		slice := reflect.ValueOf(from)
		var values []string
		for i := 0; i < slice.Len(); i++ {
			es, err := f.encode(slice.Index(i).Interface(), reflect.TypeOf(""))
			if err != nil {
				return nil
			}
			values = append(values, convert.ToValidString(es))
		}
		return strings.Join(values, f.SplitOut())
	}

	// Field of type Pointer
	if f.FieldType.Kind() == reflect.Pointer {
		return from
	}

	// Encode the value
	encoded, err := f.encode(from, f.FieldType)
	if err != nil {
		return nil
	}

	if len(convert.ToValidString(from)) == 0 {
		return f.DefaultValueOut()
	} else {
		return encoded.Interface()
	}
}

// toCellValue is called when writing an Excel file
func (f *activeFieldTags) encode(from interface{}, fieldType reflect.Type) (value reflect.Value, err error) {
	switch f.EncodingOut() {
	case "json":
		j, err := convert.ToJsonString(from)
		if err != nil {
			return reflect.Value{}, err
		}
		return reflect.ValueOf(j), nil
	default:
		if f.FieldType == timeType {
			dt, err := convert.ToTime(from)
			if err != nil {
				return reflect.Value{}, err
			}
			if dt.Year() == 1 {
				return reflect.ValueOf(""), nil
			}
			if len(f.FormatOut()) > 0 {
				s, err := convert.ToTimeString(dt, f.FormatOut())
				if err != nil {
					return reflect.Value{}, err
				}
				return reflect.ValueOf(s), nil
			}
			s, err := convert.ToTimeString(dt)
			if err != nil {
				return reflect.Value{}, err
			}
			return reflect.ValueOf(s), nil
		} else {
			value, err = convert.ToValue(from, fieldType)
		}
	}
	return
}
