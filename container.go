package excel

import "reflect"

// Container is a struct that contains the Value and type of the container
// it is used to create the appropriate reader or writer
type Container struct {
	Value   reflect.Value
	Type    reflect.Type
	Pointer bool
}

// newValue returns a Value of the appropriate type
func (c *Container) newValue() reflect.Value {
	if c.Pointer {
		return reflect.New(c.Type.Elem())
	}
	return reflect.New(c.Type).Elem()
}

// assign a Value to a field of the container
func (c *Container) assign(container reflect.Value, index int, value reflect.Value) {

	// If the container has a type Struct, we need to assign the value to the field
	// If the container has a type Slice, we need to assign the value to the element of the slice
	// If the container has a type Map, we need to assign the value to the element of the map

	kind := container.Kind()
	if c.Pointer {
		kind = container.Elem().Kind()
	}

	switch kind {
	case reflect.Struct:
		if c.Pointer {
			container.Elem().Field(index).Set(value)
		} else {
			container.Field(index).Set(value)
		}
	case reflect.Slice:
		if c.Pointer {
			container.Elem().Index(index).Set(value)
		} else {
			container.Index(index).Set(value)
		}
	case reflect.Map:
		if c.Pointer {
			container.Elem().SetMapIndex(value.Field(0), value.Field(1))
		} else {
			container.SetMapIndex(value.Field(0), value.Field(1))
		}
	}
}
