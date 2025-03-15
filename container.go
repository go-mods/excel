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
		if c.Pointer {
			elem := container.Elem()
			if index < 0 || index >= elem.NumField() {
				return fmt.Errorf("excel: field index %d out of bounds for struct with %d fields", index, elem.NumField())
			}
			field := elem.Field(index)
			if !field.CanSet() {
				return fmt.Errorf("excel: cannot set field at index %d (possibly unexported)", index)
			}
			field.Set(value)
		} else {
			if index < 0 || index >= container.NumField() {
				return fmt.Errorf("excel: field index %d out of bounds for struct with %d fields", index, container.NumField())
			}
			field := container.Field(index)
			if !field.CanSet() {
				return fmt.Errorf("excel: cannot set field at index %d (possibly unexported)", index)
			}
			field.Set(value)
		}
	case reflect.Slice:
		if c.Pointer {
			elem := container.Elem()
			if index < 0 || index >= elem.Len() {
				return fmt.Errorf("excel: slice index %d out of bounds for slice with length %d", index, elem.Len())
			}
			elem.Index(index).Set(value)
		} else {
			if index < 0 || index >= container.Len() {
				return fmt.Errorf("excel: slice index %d out of bounds for slice with length %d", index, container.Len())
			}
			container.Index(index).Set(value)
		}
	case reflect.Map:
		if c.Pointer {
			elem := container.Elem()
			if value.NumField() < 2 {
				return fmt.Errorf("excel: map value must have at least 2 fields (key and value)")
			}
			elem.SetMapIndex(value.Field(0), value.Field(1))
		} else {
			if value.NumField() < 2 {
				return fmt.Errorf("excel: map value must have at least 2 fields (key and value)")
			}
			container.SetMapIndex(value.Field(0), value.Field(1))
		}
	default:
		return fmt.Errorf("excel: unsupported container kind: %v", kind)
	}

	return nil
}
