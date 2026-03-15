# xlsxtemplater

Generate Excel (.xlsx) files from templates using Go's `text/template` syntax.

## Installation

```bash
go get github.com/doniyorbek7376/xlsxtemplater
```

## Usage

```go
err := xlsxtemplater.Generate("template.xlsx", data, "output.xlsx")
```

### Example

```go
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

err := xlsxtemplater.Generate("template.xlsx", content, "generated.xlsx")
```

## Template Syntax

Use Go's `text/template` syntax directly in Excel cells:

| Syntax | Description |
|--------|-------------|
| `{{.Field}}` | Access field |
| `{{.Nested.Field}}` | Nested field access |
| `{{range .Items}}...{{end}}` | Loop over slice |
| `{{if .Condition}}...{{end}}` | Conditional |
| `{{if .Condition}}...{{else}}...{{end}}` | If-else |

### Rows as Directives

The `{{range}}`, `{{if}}`, `{{else}}`, and `{{end}}` directives are placed in a single cell in a row. The entire row becomes the directive and is not rendered directly.

## Template Functions

| Function | Description | Example |
|----------|-------------|---------|
| `upper` | Uppercase string | `{{upper .Name}}` |
| `lower` | Lowercase string | `{{lower .Name}}` |
| `title` | Title case | `{{title .Name}}` |
| `round` | Round number | `{{round .Price 2}}` |
| `multiply` | Multiply numbers | `{{multiply .Price .Quantity}}` |
| `format_date` | Reformat date | `{{format_date .Date "2006-01-02" "02.01.2006"}}` |
| `to_number` | Convert to number | `{{to_number .Value}}` |
| `format_number` | Format with locale | `{{format_number 33928.57}}` → `"33 928,57"` |

## Custom Functions

Pass custom template functions via options:

```go
xlsxtemplater.GenerateWithOptions("template.xlsx", data, "output.xlsx", &xlsxtemplater.Options{
    CustomFuncMap: template.FuncMap{
        "myfunc": func(s string) string { return "custom: " + s },
    },
})
```

## License

MIT