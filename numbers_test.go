package xlsxtemplater

import (
	"fmt"
	"math"
	"path/filepath"
	"testing"
	"text/template"

	"codeberg.org/tealeg/xlsx/v4"
)

type namedFloat float64

func TestPlainNumberValue(t *testing.T) {
	tests := []struct {
		name  string
		value any
		want  float64
		ok    bool
	}{
		{"integer", 1500000, 1500000, true},
		{"decimal", 12345.67, 12345.67, true},
		{"negative", -12.5, -12.5, true},
		{"zero", 0, 0, true},
		{"uint32", uint32(1234), 1234, true},
		{"named float type", namedFloat(2.5), 2.5, true},
		{"NaN", math.NaN(), 0, false},
		{"Inf", math.Inf(1), 0, false},
		{"string integer", "1234", 1234, true},
		{"string negative", "-12.5", -12.5, true},
		{"string zero", "0", 0, true},
		{"string fraction", "0.5", 0.5, true},
		{"string leading zeros", "0012", 12, false},
		{"string 11 digits", "12345678901", 12345678901, true},
		{"string 12 digits", "123456789012", 123456789012, false},
		{"string localized", "1 234,50", 0, false},
		{"string NaN", "NaN", 0, false},
		{"string dash", "-", 0, false},
		{"empty string", "", 0, false},
		{"nil", nil, 0, false},
		{"bool", true, 0, false},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, ok := plainNumberValue(tt.value)
			if ok != tt.ok || (ok && got != tt.want) {
				t.Errorf("plainNumberValue(%#v) = %v, %v; want %v, %v", tt.value, got, ok, tt.want, tt.ok)
			}
		})
	}
}

func TestGenerateNumericCells(t *testing.T) {
	type templateCell struct {
		raw    string
		numFmt string
	}

	tests := []struct {
		name       string
		cell       templateCell
		content    map[string]any
		options    *Options
		wantType   xlsx.CellType
		wantValue  string
		wantNumFmt string
	}{
		{
			name:       "format_number with an integer",
			cell:       templateCell{raw: "{{ .x | format_number }}"},
			content:    map[string]any{"x": 1500000},
			wantType:   xlsx.CellTypeNumeric,
			wantValue:  "1500000",
			wantNumFmt: "#,##0.00",
		},
		{
			name:       "format_number rounds to 2 decimals",
			cell:       templateCell{raw: "{{ format_number .x }}"},
			content:    map[string]any{"x": 12345.678},
			wantType:   xlsx.CellTypeNumeric,
			wantValue:  "12345.68",
			wantNumFmt: "#,##0.00",
		},
		{
			name:       "format_number keeps the template's numeric format",
			cell:       templateCell{raw: "{{ .x | format_number }}", numFmt: "0.000"},
			content:    map[string]any{"x": 12.5},
			wantType:   xlsx.CellTypeNumeric,
			wantValue:  "12.5",
			wantNumFmt: "0.000",
		},
		{
			name:       "format_number with a numeric string",
			cell:       templateCell{raw: "{{ .x | format_number }}"},
			content:    map[string]any{"x": "1234.5"},
			wantType:   xlsx.CellTypeNumeric,
			wantValue:  "1234.5",
			wantNumFmt: "#,##0.00",
		},
		{
			name:      "format_number with leading zeros stays text",
			cell:      templateCell{raw: "{{ .x | format_number }}"},
			content:   map[string]any{"x": "0012"},
			wantType:  xlsx.CellTypeString,
			wantValue: "0012",
		},
		{
			name:      "format_number piped on prints formatted",
			cell:      templateCell{raw: `{{ .x | format_number | printf "%v UZS" }}`},
			content:   map[string]any{"x": 1234.5},
			wantType:  xlsx.CellTypeString,
			wantValue: "1 234,50 UZS",
		},
		{
			name:      "format_number with a non-number stays text",
			cell:      templateCell{raw: "{{ .x | format_number }}"},
			content:   map[string]any{"x": "-"},
			wantType:  xlsx.CellTypeString,
			wantValue: "-",
		},
		{
			name:       "plain uint32 in a General cell",
			cell:       templateCell{raw: "{{ .x }}"},
			content:    map[string]any{"x": uint32(1234)},
			wantType:   xlsx.CellTypeNumeric,
			wantValue:  "1234",
			wantNumFmt: "general",
		},
		{
			name:       "plain numeric string in a General cell",
			cell:       templateCell{raw: "{{ .x }}"},
			content:    map[string]any{"x": "1234"},
			wantType:   xlsx.CellTypeNumeric,
			wantValue:  "1234",
			wantNumFmt: "general",
		},
		{
			name:       "plain float keeps the template's numeric format",
			cell:       templateCell{raw: "{{ .x }}", numFmt: "#,##0.00"},
			content:    map[string]any{"x": 12345.678},
			wantType:   xlsx.CellTypeNumeric,
			wantValue:  "12345.678",
			wantNumFmt: "#,##0.00",
		},
		{
			name:      "plain string with leading zeros stays text",
			cell:      templateCell{raw: "{{ .x }}"},
			content:   map[string]any{"x": "0012"},
			wantType:  xlsx.CellTypeString,
			wantValue: "0012",
		},
		{
			name:      "string function with a non-number stays text",
			cell:      templateCell{raw: "{{ .x | upper }}"},
			content:   map[string]any{"x": "abc"},
			wantType:  xlsx.CellTypeString,
			wantValue: "ABC",
		},
		{
			name:      "text around format_number stays text",
			cell:      templateCell{raw: "Total: {{ .x | format_number }}"},
			content:   map[string]any{"x": 1234.5},
			wantType:  xlsx.CellTypeString,
			wantValue: "Total: 1 234,50",
		},
		{
			name:      "missing value stays text",
			cell:      templateCell{raw: "{{ .missing }}"},
			content:   map[string]any{"x": 12},
			wantType:  xlsx.CellTypeString,
			wantValue: "<no value>",
		},
		{
			name:    "custom format_number is used as is",
			cell:    templateCell{raw: "{{ .x | format_number }}"},
			content: map[string]any{"x": 12.5},
			options: &Options{CustomFuncMap: template.FuncMap{
				"format_number": func(v any) string { return fmt.Sprintf("%v UZS", v) },
			}},
			wantType:  xlsx.CellTypeString,
			wantValue: "12.5 UZS",
		},
		{
			name:      "text around a plain number stays text",
			cell:      templateCell{raw: "{{ .x }} pcs"},
			content:   map[string]any{"x": 12},
			wantType:  xlsx.CellTypeString,
			wantValue: "12 pcs",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			dir := t.TempDir()
			templatePath := filepath.Join(dir, "template.xlsx")
			generatedPath := filepath.Join(dir, "generated.xlsx")

			file := xlsx.NewFile()
			sheet, err := file.AddSheet("Sheet1")
			if err != nil {
				t.Fatal(err)
			}

			cell := sheet.AddRow().AddCell()
			cell.SetString(tt.cell.raw)
			if tt.cell.numFmt != "" {
				cell.SetFormat(tt.cell.numFmt)
			}

			if err := file.Save(templatePath); err != nil {
				t.Fatal(err)
			}

			if err := GenerateWithOptions(templatePath, tt.content, generatedPath, tt.options); err != nil {
				t.Fatal(err)
			}

			generated, err := xlsx.OpenFile(generatedPath)
			if err != nil {
				t.Fatal(err)
			}

			got, err := generated.Sheets[0].Cell(0, 0)
			if err != nil {
				t.Fatal(err)
			}

			if got.Type() != tt.wantType {
				t.Errorf("type = %v; want %v", got.Type(), tt.wantType)
			}

			if got.Value != tt.wantValue {
				t.Errorf("value = %q; want %q", got.Value, tt.wantValue)
			}

			if tt.wantType == xlsx.CellTypeNumeric && got.NumFmt != tt.wantNumFmt {
				t.Errorf("number format = %q; want %q", got.NumFmt, tt.wantNumFmt)
			}
		})
	}
}
