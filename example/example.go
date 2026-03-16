package main

import (
	"math/rand"
	"strings"

	"github.com/doniyorbek7376/xlsxtemplater"
)

func main() {
	itemsCount := 10
	items := make([]map[string]any, 0, itemsCount)
	for range itemsCount {
		items = append(items, map[string]any{
			"name":     randomName(),
			"price":    rand.Float64() * 10000,
			"quantity": rand.Float64() * 10,
		})
	}

	content := map[string]any{
		"order": map[string]any{
			"number": "121231",
		},
		"items": items,
	}

	err := xlsxtemplater.Generate(
		"template.xlsx",
		content,
		"generated.xlsx",
	)
	if err != nil {
		panic(err)
	}
}

func randomName() string {
	length := rand.Intn(5) + 5

	var sb strings.Builder
	for range length {
		c := 'A' + rune(rand.Intn(26))

		sb.WriteRune(c)
	}

	return sb.String()
}
