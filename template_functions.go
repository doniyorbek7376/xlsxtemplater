package xlsxtemplater

import (
	"fmt"
	"math"
	"reflect"
	"strconv"
	"strings"
	"text/template"
	"time"
)

func getDefaultTemplateFunctions() template.FuncMap {
	return template.FuncMap{
		"upper": strings.ToUpper,
		"lower": strings.ToLower,
		"title": strings.ToTitle,
		"round": func(a any, precision int) any {
			return math.Round(parseFloat(a)*math.Pow10(precision)) / math.Pow10(precision)
		},
		"multiply": func(a any, b any) float64 {
			return parseFloat(a) * parseFloat(b)
		},
		"format_date": func(a, oldLayout, newLayout string) string {
			if len(a) > len(oldLayout) {
				a = a[:len(oldLayout)]
			}

			time, _ := time.Parse(oldLayout, a)

			return time.Format(newLayout)
		},
		"to_number":     parseFloat,
		"format_number": FormatNumberForExcel,
	}
}

func parseFloat(value any) float64 {
	v := reflect.ValueOf(value)

	switch v.Kind() {
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		return float64(v.Int())
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64, reflect.Uintptr:
		return float64(v.Uint())
	case reflect.Float32, reflect.Float64:
		return v.Float()
	case reflect.String:
		parsed, _ := strconv.ParseFloat(v.String(), 64)

		return parsed
	}

	return 0
}

// FormatNumberForExcel formats a number to a localized string format suitable for Excel export.
// Formats numbers like: 33928.57 → "33 928,57" (space as thousands separator, comma as decimal separator)
// This format is commonly used in European and Central Asian locales.
func FormatNumberForExcel(value any) string {
	var floatValue float64

	// Type assertion to handle various numeric types
	switch v := value.(type) {
	case float64:
		floatValue = v
	case float32:
		floatValue = float64(v)
	case int:
		floatValue = float64(v)
	case int32:
		floatValue = float64(v)
	case int64:
		floatValue = float64(v)
	case uint32:
		floatValue = float64(v)
	case uint64:
		floatValue = float64(v)
	default:
		return fmt.Sprintf("%v", value)
	}

	// Format with 2 decimal places
	formatted := fmt.Sprintf("%.2f", floatValue)

	// Split into integer and decimal parts
	parts := strings.Split(formatted, ".")
	intPart := parts[0]

	decimalPart := ""
	if len(parts) > 1 {
		decimalPart = parts[1]
	}

	// Add thousands separators (spaces) to the integer part
	intPartFormatted := addThousandsSeparators(intPart)

	// Combine with comma as decimal separator
	return intPartFormatted + "," + decimalPart
}

// FormatDate reformats string into 31.12.1999 format
func FormatDate(value string) string {
	dateFormat := "02.01.2006"
	if value == "" {
		return ""
	}

	if len(value) > len(dateFormat) {
		value = value[:len(dateFormat)]
	}

	tm, err := time.Parse("2006-01-02", value)
	if err != nil {
		return value
	}

	return tm.Format(dateFormat)
}

// addThousandsSeparators adds space as thousands separator to a number string
// Example: "33928" → "33 928".
func addThousandsSeparators(numStr string) string {
	// Handle negative sign
	negative := false
	if strings.HasPrefix(numStr, "-") {
		negative = true
		numStr = numStr[1:]
	}

	// Reverse the string for easier processing from right to left
	runes := []rune(numStr)
	for i, j := 0, len(runes)-1; i < j; i, j = i+1, j-1 {
		runes[i], runes[j] = runes[j], runes[i]
	}

	// Add separators every 3 digits
	var result strings.Builder

	for i, r := range runes {
		if i > 0 && i%3 == 0 {
			result.WriteRune(' ')
		}

		result.WriteRune(r)
	}

	// Reverse back to original direction
	resultRunes := []rune(result.String())
	for i, j := 0, len(resultRunes)-1; i < j; i, j = i+1, j-1 {
		resultRunes[i], resultRunes[j] = resultRunes[j], resultRunes[i]
	}

	formatted := string(resultRunes)
	if negative {
		formatted = "-" + formatted
	}

	return formatted
}
