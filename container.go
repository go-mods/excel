package excel

import (
	"fmt"
	"reflect"
)

// Container is a struct that contains the Value and type of the container
// it is used to create the appropriate reader or writer
type Container struct {
	Value   reflect.Value
	Type    reflect.Type
	Pointer bool
}

// newValue returns a Value of the appropriate type
// It creates a new instance of the container's type, handling both pointer and non-pointer types.
func (c *Container) newValue() reflect.Value {
	if c == nil || c.Type == nil {
		return reflect.Value{}
	}

	if c.Pointer {
		return reflect.New(c.Type.Elem())
	}
	return reflect.New(c.Type).Elem()
}

// assign sets a value to a field of the container
// It handles different container types (struct, slice, map) and respects the pointer status.
// Returns an error if the container or value is invalid, or if the index is out of bounds.
func (c *Container) assign(container reflect.Value, index int, value reflect.Value) error {
	if c == nil {
		return fmt.Errorf("excel: container is nil")
	}

	if !container.IsValid() {
		return fmt.Errorf("excel: container value is invalid")
	}

	if !value.IsValid() {
		return fmt.Errorf("excel: value is invalid")
	}

	// If the container has a type Struct, we need to assign the value to the field
	// If the container has a type Slice, we need to assign the value to the element of the slice
	// If the container has a type Map, we need to assign the value to the element of the map

	kind := container.Kind()
	if c.Pointer {
		if container.IsNil() {
			return fmt.Errorf("excel: container pointer is nil")
		}
		kind = container.Elem().Kind()
	}

	switch kind {
	case reflect.Struct:
		// For structures, we need to find the field corresponding to the index
		// taking into account embedded structures
		field, err := c.findFieldByIndex(container, index)
		if err != nil {
			return err
		}

		if !field.CanSet() {
			return fmt.Errorf("excel: cannot set field at index %d (possibly unexported)", index)
		}

		// Check if the value type is compatible with the field type
		if !value.Type().AssignableTo(field.Type()) {
			// If types are not directly assignable, try to convert
			if value.Type().ConvertibleTo(field.Type()) {
				field.Set(value.Convert(field.Type()))
			} else {
				return fmt.Errorf("excel: value of type %v is not assignable to field of type %v", value.Type(), field.Type())
			}
		} else {
			field.Set(value)
		}
	case reflect.Slice:
		target := container
		if c.Pointer {
			target = container.Elem()
		}

		if index < 0 || index >= target.Len() {
			return fmt.Errorf("excel: slice index %d out of bounds for slice with length %d", index, target.Len())
		}
		target.Index(index).Set(value)
	case reflect.Map:
		target := container
		if c.Pointer {
			target = container.Elem()
		}

		if value.NumField() < 2 {
			return fmt.Errorf("excel: map value must have at least 2 fields (key and value)")
		}
		target.SetMapIndex(value.Field(0), value.Field(1))
	default:
		return fmt.Errorf("excel: unsupported container kind: %v", kind)
	}

	return nil
}

// findFieldByIndex finds the field corresponding to the index in the structure
// This function handles embedded structures
func (c *Container) findFieldByIndex(container reflect.Value, index int) (reflect.Value, error) {
	target := container

	if container.Kind() == reflect.Pointer {
		target = container.Elem()
	}

	// Traverse all fields to find the one that corresponds to the index
	return c.findFieldRecursive(target, index, 0)
}

// findFieldRecursive recursively traverses the structure to find the field corresponding to the index
func (c *Container) findFieldRecursive(structValue reflect.Value, targetIndex, currentFieldIndex int) (reflect.Value, error) {
	if structValue.Kind() != reflect.Struct {
		return reflect.Value{}, fmt.Errorf("excel: expected struct, got %structValue", structValue.Kind())
	}

	for i := 0; i < structValue.NumField(); i++ {
		field := structValue.Field(i)
		fieldType := structValue.Type().Field(i)

		// If it's an embedded structure, traverse its fields
		if fieldType.Anonymous {
			if field.Kind() == reflect.Ptr {
				if field.IsNil() {
					// Initialize the pointer if necessary
					field.Set(reflect.New(field.Type().Elem()))
				}
				field = field.Elem()
			}

			if field.Kind() == reflect.Struct {
				// Recursively traverse the embedded structure
				result, err := c.findFieldRecursive(field, targetIndex, currentFieldIndex)
				if err == nil {
					return result, nil
				}

				// If the field was not found, continue with the updated counter
				currentFieldIndex += c.countFields(field)
				continue
			}
		}

		// If we found the target index
		if currentFieldIndex == targetIndex {
			return field, nil
		}

		currentFieldIndex++
	}

	return reflect.Value{}, fmt.Errorf("excel: field index %d not found", targetIndex)
}

// countFields counts the number of fields in a structure, including fields of embedded structures
func (c *Container) countFields(structValue reflect.Value) int {
	if structValue.Kind() != reflect.Struct {
		return 0
	}

	count := 0
	for i := 0; i < structValue.NumField(); i++ {
		field := structValue.Field(i)
		fieldType := structValue.Type().Field(i)

		// If it's an embedded structure, count its fields
		if fieldType.Anonymous {
			if field.Kind() == reflect.Ptr {
				if field.IsNil() {
					// Initialize the pointer if necessary
					field.Set(reflect.New(field.Type().Elem()))
				}
				field = field.Elem()
			}

			if field.Kind() == reflect.Struct {
				count += c.countFields(field)
				continue
			}
		}

		count++
	}

	return count
}
