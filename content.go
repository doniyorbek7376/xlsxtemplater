package xlsxtemplater

import (
	"reflect"
	"strings"
)

func extractSlice(content any, path string) []any {
	extracted := extractContent(content, path)

	rv := reflect.ValueOf(extracted)
	if rv.Kind() != reflect.Slice && rv.Kind() != reflect.Array {
		return []any{extracted}
	}

	items := []any{}
	for i := 0; i < rv.Len(); i++ {
		item := rv.Index(i).Interface()

		items = append(items, item)
	}

	return items
}

func extractContent(content any, path string) any {
	path = strings.TrimSpace(path)
	path = strings.TrimPrefix(path, ".")
	val := reflect.ValueOf(content)

	for part := range strings.SplitSeq(path, ".") {
		if val.Kind() == reflect.Pointer {
			val = val.Elem()
		}

		if val.Kind() == reflect.Struct {
			val = val.FieldByName(part)
		} else if val.Kind() == reflect.Map {
			val = val.MapIndex(reflect.ValueOf(part))
		} else {
			return nil
		}
	}

	if val.IsValid() {
		return val.Interface()
	}

	return nil
}
