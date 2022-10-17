package excel

import (
	"reflect"
	"strings"
	"time"

	"github.com/go-mods/convert"
)

type FieldInfo struct {
	FieldIndex int
	FieldType  reflect.Type

	Tags    *FieldTags
	TagsIn  *FieldTags
	TagsOut *FieldTags
}

// Marshaller can be implemented by any value that has a Marshal method
// This converter is used to convert the value to the desired representation
type Marshaller interface {
	Marshall() (interface{}, error)
}

// Unmarshaller can be implemented by any value that has an Unmarshall method
// This converter is used to convert the value to the desired representation
type Unmarshaller interface {
	Unmarshall(s string) error
}

func getFieldsInfos(s *StructInfo) []*FieldInfo {

	fieldsCount := s.StructType.NumField()
	fieldsInfos := make([]*FieldInfo, 0, fieldsCount)

	// Check if the ContainerInfo implement FieldsTags, FieldsTagsIn or FieldsTagsOut interface
	// ------------------------------------------------------------------------------------
	type DefaultFieldsTags struct {
		tags    map[string]*FieldTags
		tagsIn  map[string]*FieldTags
		tagsOut map[string]*FieldTags
	}
	var defaultFieldsTags DefaultFieldsTags

	v := reflect.New(s.StructType)

	if v.CanInterface() {
		if i, ok := v.Interface().(FieldsTags); ok {
			defaultFieldsTags.tags = i.GetFieldsTags()
		}
		if i, ok := v.Interface().(FieldsTagsIn); ok {
			defaultFieldsTags.tagsIn = i.GetFieldsTagsIn()
		}
		if i, ok := v.Interface().(FieldsTagsOut); ok {
			defaultFieldsTags.tagsOut = i.GetFieldsTagsOut()
		}
	} else if vElem := v.Elem(); vElem.CanInterface() {
		if i, ok := vElem.Interface().(FieldsTags); ok {
			defaultFieldsTags.tags = i.GetFieldsTags()
		}
		if i, ok := vElem.Interface().(FieldsTagsIn); ok {
			defaultFieldsTags.tagsIn = i.GetFieldsTagsIn()
		}
		if i, ok := vElem.Interface().(FieldsTagsOut); ok {
			defaultFieldsTags.tagsOut = i.GetFieldsTagsOut()
		}
	}

	// Loop throw each field of the ContainerInfo to get each field configuration
	// ----------------------------------------------------------------------
	for i := 0; i < fieldsCount; i++ {

		field := s.StructType.Field(i)

		// Get tags from struct
		fieldInfo := &FieldInfo{
			FieldIndex: i,
			FieldType:  field.Type,
			Tags:       s.filterTags(field, mainKey),
			TagsIn:     s.filterTags(field, inKey),
			TagsOut:    s.filterTags(field, outKey),
		}

		// Overwrite with tags from interfaces
		s.freeze(defaultFieldsTags.tags[field.Name], fieldInfo.Tags)
		s.freeze(defaultFieldsTags.tagsIn[field.Name], fieldInfo.TagsIn)
		s.freeze(defaultFieldsTags.tagsOut[field.Name], fieldInfo.TagsOut)

		//
		fieldsInfos = append(fieldsInfos, fieldInfo)
	}

	return fieldsInfos
}

var timeType = reflect.TypeOf((*time.Time)(nil)).Elem()

// toValue is called when reading an Excel file
func (f *FieldInfo) toValue(from string) (value reflect.Value, err error) {

	// Converter call
	if f.FieldType.Kind() == reflect.Pointer {
		vp := reflect.New(f.FieldType.Elem())
		if unmarshall, ok := vp.Interface().(Unmarshaller); ok {
			err = unmarshall.Unmarshall(from)
			return reflect.ValueOf(vp.Interface()), err
		}
	}

	// Converter call
	if f.FieldType.Kind() == reflect.Struct {
		vp := reflect.New(f.FieldType)
		if unmarshall, ok := vp.Interface().(Unmarshaller); ok {
			err = unmarshall.Unmarshall(from)
			return reflect.ValueOf(vp.Elem().Interface()), err
		}
	}

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

	// Field of type AsPointer
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
func (f *FieldInfo) decode(from string, to reflect.Type) (value reflect.Value, err error) {
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

// toCellValue is called when writing to an Excel file
func (f *FieldInfo) toCellValue(from interface{}) (interface{}, error) {

	// Converter call for pointer to struct
	if f.FieldType.Kind() == reflect.Pointer {
		vp := reflect.New(f.FieldType).Elem()
		vp.Set(reflect.ValueOf(from))
		if marshall, ok := vp.Interface().(Marshaller); ok {
			vi, err := marshall.Marshall()
			return reflect.ValueOf(vi), err
		}
	}
	// Converter call for struct
	if f.FieldType.Kind() == reflect.Struct {
		vp := reflect.New(f.FieldType)
		vp.Elem().Set(reflect.ValueOf(from))
		if marshall, ok := vp.Interface().(Marshaller); ok {
			vi, err := marshall.Marshall()
			return reflect.ValueOf(vi), err
		}
	}

	// Field of type Slice or Array
	if f.FieldType.Kind() == reflect.Slice || f.FieldType.Kind() == reflect.Array {
		slice := reflect.ValueOf(from)
		var values []string
		for i := 0; i < slice.Len(); i++ {
			es, err := f.encode(slice.Index(i).Interface(), reflect.TypeOf(""))
			if err != nil {
				return nil, err
			}
			values = append(values, convert.ToValidString(es))
		}
		return strings.Join(values, f.SplitOut()), nil
	}

	// Field of type AsPointer
	if f.FieldType.Kind() == reflect.Pointer {
		return from, nil
	}

	// Encode the value
	encoded, err := f.encode(from, f.FieldType)
	if err != nil {
		return nil, err
	}

	if len(convert.ToValidString(from)) == 0 {
		return f.DefaultValueOut(), nil
	} else {
		return encoded.Interface(), nil
	}
}

// toCellValue is called when writing an Excel file
func (f *FieldInfo) encode(from interface{}, fieldType reflect.Type) (value reflect.Value, err error) {
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
