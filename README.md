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
| `format_number` | Number cell with `#,##0.00` (see Numbers) | `{{format_number 33928.57}}` → numeric `33928.57` |

## Numbers

A cell whose whole content is a single action is written as a numeric cell when
its value is a number; otherwise it is text, as before. Text around the action
(`Total: {{.Sum}}`) always gives a text cell.

- `{{.Price | format_number}}` / `{{format_number .Price}}`: numeric when the
  argument of `format_number` is a Go number or any string `strconv.ParseFloat`
  reads as a finite number (`"0012"` → `12`). The value is rounded to 2
  decimals. The cell keeps its own number format, or gets `#,##0.00` when it is
  General. A non-number is text, formatted as `"33 928,57"`.
- `{{.Quantity}}`, `{{.Quantity | f}}`: numeric when the pipeline's value is a
  Go number, or a decimal string without leading zeros of at most 11 digits
  (`"1234"`, `"-12.5"`). `"0012"`, barcodes and other long digit strings stay
  text. The value is not rounded and the cell keeps its own number format.

A `format_number` replaced through `CustomFuncMap` is treated as any other
function.

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