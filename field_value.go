package excel

import (
	"github.com/go-mods/convert"
	"reflect"
	"strings"
	"time"
)

var timeType = reflect.TypeOf((*time.Time)(nil)).Elem()

// toValue is called when reading an Excel file to get the value of a field
func (f *Field) toValue(from string) (value reflect.Value, err error) {

	// Get the value of the field if it is a pointer
	// and the pointer implements the Unmarshaller interface
	if f.Type.Kind() == reflect.Pointer {
		vp := reflect.New(f.Type.Elem())
		if unmarshall, ok := vp.Interface().(Unmarshaller); ok {
			err = unmarshall.Unmarshall(from)
			return reflect.ValueOf(vp.Interface()), err
		}
	}

	// Get the value of the field if it is a struct
	// and the struct implements the Unmarshaller interface
	if f.Type.Kind() == reflect.Struct {
		vp := reflect.New(f.Type)
		if unmarshall, ok := vp.Interface().(Unmarshaller); ok {
			err = unmarshall.Unmarshall(from)
			return reflect.ValueOf(vp.Elem().Interface()), err
		}
	}

	// Decode the value of the field if it is a slice or array
	if f.Type.Kind() == reflect.Slice || f.Type.Kind() == reflect.Array {
		if len(from) > 0 {
			values := strings.Split(convert.ToValidString(from), f.GetReadSplit())
			value = reflect.MakeSlice(reflect.SliceOf(f.Type.Elem()), 0, len(values))
			for _, vs := range values {
				v, err := f.decode(vs, f.Type.Elem())
				if err != nil {
					return reflect.Value{}, err
				}
				value = reflect.Append(value, v)
			}
		} else {
			return reflect.MakeSlice(reflect.SliceOf(f.Type.Elem()), 0, 0), nil
		}
		return
	}

	// Decode the value of the field if it is a pointer
	if f.Type.Kind() == reflect.Pointer {
		value, err = f.decode(from, f.Type)
		if err != nil {
			return reflect.Value{}, err
		}
		return
	}

	// Decode the value of the field
	value, err = f.decode(from, f.Type)
	if err != nil {
		return reflect.Value{}, err
	}

	return
}

// decode is called when reading an Excel file to get the value of a field
func (f *Field) decode(from string, to reflect.Type) (value reflect.Value, err error) {
	switch f.GetReadEncoding() {
	case "json":
		value, err = convert.ToJsonValue(from, to)
	default:
		if f.Type == timeType {
			dt, err := convert.ToLayoutTime(f.GetReadFormat(), from)
			if err != nil {
				return reflect.Value{}, nil
			}
			return reflect.ValueOf(dt), err
		} else {
			if len(from) == 0 {
				value = reflect.ValueOf(f.GetReadDefault())
			} else {
				value, err = convert.ToValue(from, to)
			}
		}
	}
	return
}

// toCellValue is called when writing to an Excel file to set the value of a cell
func (f *Field) toCellValue(from interface{}) (interface{}, error) {

	// Set the value of the field if it is a pointer
	// and the pointer implements the Marshaller interface
	if f.Type.Kind() == reflect.Pointer {
		vp := reflect.New(f.Type).Elem()
		vp.Set(reflect.ValueOf(from))
		if marshall, ok := vp.Interface().(Marshaller); ok {
			vi, err := marshall.Marshall()
			return reflect.ValueOf(vi), err
		}
	}

	// Set the value of the field if it is a struct
	// and the struct implements the Marshaller interface
	if f.Type.Kind() == reflect.Struct {
		vp := reflect.New(f.Type)
		vp.Elem().Set(reflect.ValueOf(from))
		if marshall, ok := vp.Interface().(Marshaller); ok {
			vi, err := marshall.Marshall()
			return reflect.ValueOf(vi), err
		}
	}

	// Encode the Value if it is a slice or array
	if f.Type.Kind() == reflect.Slice || f.Type.Kind() == reflect.Array {
		slice := reflect.ValueOf(from)
		var values []string
		for i := 0; i < slice.Len(); i++ {
			es, err := f.encode(slice.Index(i).Interface(), reflect.TypeOf(""))
			if err != nil {
				return nil, err
			}
			values = append(values, convert.ToValidString(es))
		}
		return strings.Join(values, f.GetWriteSplit()), nil
	}

	// Encode the Value if it is a pointer
	if f.Type.Kind() == reflect.Pointer {
		return from, nil
	}

	// Encode the value
	encoded, err := f.encode(from, f.Type)
	if err != nil {
		return nil, err
	}

	// Return the default value if the value is empty
	if len(convert.ToValidString(from)) == 0 {
		return f.GetWriteDefault(), nil
	} else {
		return encoded.Interface(), nil
	}
}

// toCellValue is called when writing an Excel file
func (f *Field) encode(from interface{}, fieldType reflect.Type) (value reflect.Value, err error) {
	switch f.GetWriteEncoding() {
	case "json":
		j, err := convert.ToJsonString(from)
		if err != nil {
			return reflect.Value{}, err
		}
		return reflect.ValueOf(j), nil
	default:
		if f.Type == timeType {
			dt, err := convert.ToTime(from)
			if err != nil {
				return reflect.Value{}, err
			}
			if dt.Year() == 1 {
				return reflect.ValueOf(""), nil
			}
			if len(f.GetWriteFormat()) > 0 {
				s, err := convert.ToTimeString(dt, f.GetWriteFormat())
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
