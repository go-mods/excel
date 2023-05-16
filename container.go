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
	if c.Pointer {
		container.Elem().Field(index).Set(value)
	} else {
		container.Field(index).Set(value)
	}
}
