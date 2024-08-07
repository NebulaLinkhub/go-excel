package go_excel

import (
	"errors"
	"reflect"
	"strings"
)

type Parser interface {
	Convert() map[string]any
}

type Column struct {
	Field       string
	FieldType   reflect.Type
	NaturalName string
	Index       int // 索引
	ExportFunc  Parser
	ImportFunc  Parser
}

func getField(data any) (fieldsMap map[string]*Column, err error) {
	val, is := refType(data)
	if !is {
		return nil, errors.New("model type err")
	}
	rt := reflect.TypeOf(val)
	fieldsMap = make(map[string]*Column)
	index := 0
	for i := 0; i < rt.NumField(); i++ {
		tagName := rt.Field(i).Tag.Get("excel")
		if tagName == "" {
			continue
		}

		tags := strings.Split(tagName, " ")
		if len(tags) == 0 {
			continue
		}
		filed := new(Column)
		filed.NaturalName = tags[0]
		filed.Field = rt.Field(i).Name
		filed.FieldType = rt.Field(i).Type
		filed.Index = index
		fieldsMap[rt.Field(i).Name] = filed
		index++
	}
	return fieldsMap, nil
}

func getFiledMap(tags []string) map[string]struct{} {
	filedMap := make(map[string]struct{})
	for _, k := range tags {
		if _, ok := filedMap[k]; !ok {
			filedMap[k] = struct{}{}
		}
	}
	return filedMap
}
