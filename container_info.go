package excel

import "reflect"

type ContainerInfo struct {
	value     reflect.Value
	typeElem  reflect.Type
	isPointer bool
}

func (c *ContainerInfo) create() reflect.Value {
	if c.isPointer {
		return reflect.New(c.typeElem.Elem())
	}
	return reflect.New(c.typeElem).Elem()
}

func (c *ContainerInfo) setFieldValue(container reflect.Value, index int, value reflect.Value) {
	if c.isPointer {
		container.Elem().Field(index).Set(value)
	} else {
		container.Field(index).Set(value)
	}
}
