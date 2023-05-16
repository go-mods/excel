package excel

import (
	"reflect"

	"github.com/go-mods/tags"
)

type StructInfo struct {
	StructType reflect.Type
	Fields     FieldInfoList
}

func getStructInfo(container *ContainerInfo) *StructInfo {

	structType := container.typeElem
	if container.isPointer {
		structType = structType.Elem()
	}

	structInfo := &StructInfo{StructType: structType}
	structInfo.Fields = getFieldsInfos(structInfo)

	return structInfo
}

func (s *StructInfo) filterTags(field reflect.StructField, key string) (tgs *FieldTags) {
	if tag := tags.Lookup(field, key); tag != nil {
		tgs = s.parseTag(tag)
	} else {
		tgs = newTags()
		if key == mainKey {
			tgs.ColumnName = field.Name
		}
	}
	return
}

func (s *StructInfo) parseTag(tag *tags.Tag) (tgs *FieldTags) {
	tgs = newTags()

	if tag.Name == ignoreTag {
		tgs.Ignore = true
		return
	}

	if len(tag.Name) > 0 {
		tgs.ColumnName = tag.Name
	}

	if o := tag.GetOption(columnTag); o != nil {
		tgs.ColumnName = o.Value
	}
	if o := tag.GetOption(defaultTag); o != nil {
		tgs.DefaultValue = o.Value
	}
	if o := tag.GetOption(formatTag); o != nil {
		tgs.Format = o.Value
	}
	if o := tag.GetOption(encodingTag); o != nil {
		tgs.Encoding = o.Value
	}
	if o := tag.GetOption(splitTag); o != nil {
		if len(o.Value) != 0 {
			tgs.Split = o.Value
		}
	}
	if o := tag.GetOption(requiredTag); o != nil {
		tgs.IsRequired = true
	}

	return tgs
}

func (s *StructInfo) freeze(from *FieldTags, to *FieldTags) {
	if from != nil && to != nil {
		to.ColumnName = from.ColumnName
		to.DefaultValue = from.DefaultValue
		to.Format = from.Format
		to.Encoding = from.Encoding
		to.Split = from.Split
		to.IsRequired = from.IsRequired
		to.Ignore = from.Ignore
	}
}

func (s *StructInfo) GetFieldFromFieldIndex(index int) *FieldInfo {
	for _, f := range s.Fields {
		if f.Index == index {
			return f
		}
	}
	return nil
}
