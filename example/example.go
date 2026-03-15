package main

import "github.com/doniyorbek7376/xlsxtemplater"

func main() {
	content := map[string]any{
		"order": map[string]any{
			"number": "121231",
		},
		"items": []map[string]any{
			{
				"name":     "Foo",
				"price":    1000.25,
				"quantity": 10,
			},
			{
				"name":     "Bar",
				"price":    200.00,
				"quantity": 121.01,
			},
		},
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
