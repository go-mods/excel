package excel

import (
	"fmt"
	"reflect"
	"strings"
	"time"

	"github.com/go-mods/convert"
)

var timeType = reflect.TypeOf((*time.Time)(nil)).Elem()

// toValue is called when reading an Excel file to get the value of a field
func (f *Field) toValue(from string) (value reflect.Value, err error) {

	// Validate field type is not nil
	if f.Type == nil {
		return reflect.Value{}, fmt.Errorf("excel: field type is nil")
	}

	// Get the value of the field if it is a pointer
	// and the pointer implements the Unmarshaller interface
	if f.Type.Kind() == reflect.Pointer {
		vp := reflect.New(f.Type.Elem())
		if unmarshall, ok := vp.Interface().(Unmarshaller); ok {
			err = unmarshall.Unmarshall(from)
			if err != nil {
				return reflect.Value{}, fmt.Errorf("excel: failed to unmarshall pointer: %w", err)
			}
			return reflect.ValueOf(vp.Interface()), nil
		}
	}

	// Get the value of the field if it is a struct
	// and the struct implements the Unmarshaller interface
	if f.Type.Kind() == reflect.Struct {
		vp := reflect.New(f.Type)
		if unmarshall, ok := vp.Interface().(Unmarshaller); ok {
			err = unmarshall.Unmarshall(from)
			if err != nil {
				return reflect.Value{}, fmt.Errorf("excel: failed to unmarshall struct: %w", err)
			}
			return reflect.ValueOf(vp.Elem().Interface()), nil
		}
	}

	// Decode the value of the field if it is a slice or array
	if f.Type.Kind() == reflect.Slice || f.Type.Kind() == reflect.Array {
		if len(from) > 0 {
			// Validate split character is not empty
			splitChar := f.GetReadSplit()
			if splitChar == "" {
				splitChar = "," // Default split character
			}

			values := strings.Split(convert.ToString(from), splitChar)
			value = reflect.MakeSlice(reflect.SliceOf(f.Type.Elem()), 0, len(values))
			for i, vs := range values {
				v, err := f.decode(vs, f.Type.Elem())
				if err != nil {
					return reflect.Value{}, fmt.Errorf("excel: failed to decode slice element %d: %w", i, err)
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
			return reflect.Value{}, fmt.Errorf("excel: failed to decode pointer: %w", err)
		}
		return
	}

	// Decode the value of the field
	value, err = f.decode(from, f.Type)
	if err != nil {
		return reflect.Value{}, fmt.Errorf("excel: failed to decode value: %w", err)
	}

	return
}

// decode is called when reading an Excel file to get the value of a field
func (f *Field) decode(from string, to reflect.Type) (value reflect.Value, err error) {
	// Validate to type is not nil
	if to == nil {
		return reflect.Value{}, fmt.Errorf("excel: target type is nil")
	}

	switch f.GetReadEncoding() {
	case "json":
		value, err = convert.ToJsonValueE(from, to)
		if err != nil {
			return reflect.Value{}, fmt.Errorf("excel: failed to decode JSON: %w", err)
		}
	default:
		if f.Type == timeType {
			// Validate format is not empty for time
			format := f.GetReadFormat()
			if format == "" {
				format = "2006-01-02" // Default date format
			}

			dt, err := convert.ToLayoutTimeE(format, from)
			if err != nil {
				return reflect.Value{}, fmt.Errorf("excel: failed to parse time with format '%s': %w", format, err)
			}
			return reflect.ValueOf(dt), nil
		} else {
			if len(from) == 0 {
				defaultValue := f.GetReadDefault()
				// Validate default value is compatible with target type
				if defaultValue != nil {
					defaultType := reflect.TypeOf(defaultValue)
					if !defaultType.ConvertibleTo(to) {
						return reflect.Value{}, fmt.Errorf("excel: default value type %v is not convertible to target type %v", defaultType, to)
					}
				}
				value = reflect.ValueOf(defaultValue)
			} else {
				value, err = convert.ToValueE(from, to)
				if err != nil {
					return reflect.Value{}, fmt.Errorf("excel: failed to convert '%s' to type %v: %w", from, to, err)
				}
			}
		}
	}
	return
}

// toCellValue is called when writing to an Excel file to set the value of a cell
func (f *Field) toCellValue(from interface{}) (interface{}, error) {

	// Validate field type is not nil
	if f.Type == nil {
		return nil, fmt.Errorf("excel: field type is nil")
	}

	// Validate from is not nil (unless the field type is a pointer)
	if from == nil && f.Type.Kind() != reflect.Pointer {
		return f.GetWriteDefault(), nil
	}

	// Set the value of the field if it is a pointer
	// and the pointer implements the Marshaller interface
	if f.Type.Kind() == reflect.Pointer {
		vp := reflect.New(f.Type).Elem()
		if from != nil {
			vp.Set(reflect.ValueOf(from))
		}
		if marshall, ok := vp.Interface().(Marshaller); ok {
			vi, err := marshall.Marshall()
			if err != nil {
				return nil, fmt.Errorf("excel: failed to marshall pointer: %w", err)
			}
			return vi, nil
		}
	}

	// Set the value of the field if it is a struct
	// and the struct implements the Marshaller interface
	if f.Type.Kind() == reflect.Struct {
		vp := reflect.New(f.Type)
		vp.Elem().Set(reflect.ValueOf(from))
		if marshall, ok := vp.Interface().(Marshaller); ok {
			vi, err := marshall.Marshall()
			if err != nil {
				return nil, fmt.Errorf("excel: failed to marshall struct: %w", err)
			}
			return vi, nil
		}
	}

	// Encode the Value if it is a slice or array
	if f.Type.Kind() == reflect.Slice || f.Type.Kind() == reflect.Array {
		slice := reflect.ValueOf(from)
		var values []string

		// Validate split character is not empty
		splitChar := f.GetWriteSplit()
		if splitChar == "" {
			splitChar = "," // Default split character
		}

		for i := 0; i < slice.Len(); i++ {
			es, err := f.encode(slice.Index(i).Interface(), reflect.TypeOf(""))
			if err != nil {
				return nil, fmt.Errorf("excel: failed to encode slice element %d: %w", i, err)
			}
			values = append(values, convert.ToString(es))
		}
		return strings.Join(values, splitChar), nil
	}

	// Encode the Value if it is a pointer
	if f.Type.Kind() == reflect.Pointer {
		if from == nil {
			return f.GetWriteDefault(), nil
		}
		return from, nil
	}

	// Encode the value
	encoded, err := f.encode(from, f.Type)
	if err != nil {
		return nil, fmt.Errorf("excel: failed to encode value: %w", err)
	}

	// Return the default value if the value is empty
	if len(convert.ToString(from)) == 0 {
		return f.GetWriteDefault(), nil
	} else {
		return encoded.Interface(), nil
	}
}

// encode is called when writing an Excel file
func (f *Field) encode(from interface{}, fieldType reflect.Type) (value reflect.Value, err error) {
	// Validate fieldType is not nil
	if fieldType == nil {
		return reflect.Value{}, fmt.Errorf("excel: field type is nil")
	}

	switch f.GetWriteEncoding() {
	case "json":
		j, err := convert.ToJsonStringE(from)
		if err != nil {
			return reflect.Value{}, fmt.Errorf("excel: failed to encode to JSON: %w", err)
		}
		return reflect.ValueOf(j), nil
	default:
		if f.Type == timeType {
			dt, err := convert.ToTimeE(from)
			if err != nil {
				return reflect.Value{}, fmt.Errorf("excel: failed to convert to time: %w", err)
			}
			if dt.Year() == 1 {
				return reflect.ValueOf(""), nil
			}
			if len(f.GetWriteFormat()) > 0 {
				format := f.GetWriteFormat()
				s, err := convert.ToTimeStringE(dt, format)
				if err != nil {
					return reflect.Value{}, fmt.Errorf("excel: failed to format time with format '%s': %w", format, err)
				}
				return reflect.ValueOf(s), nil
			}
			return reflect.ValueOf(dt), nil
		} else {
			value, err = convert.ToValueE(from, fieldType)
			if err != nil {
				return reflect.Value{}, fmt.Errorf("excel: failed to convert value to type %v: %w", fieldType, err)
			}
		}
	}
	return
}
