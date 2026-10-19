package xlsxtemplater

import (
	"math"
	"reflect"
	"regexp"
	"strconv"
	"strings"
	"text/template"
	tparse "text/template/parse"
)

type numberCell int

const (
	numberCellNone numberCell = iota
	// numberCellPlain is a cell holding a single action, e.g. {{ .x }}.
	numberCellPlain
	// numberCellFormatNumber is a cell holding a single action ending in format_number.
	numberCellFormatNumber
)

const (
	plainNumberFunc  = "xlsxtemplater_plain_number"
	formatNumberFunc = "xlsxtemplater_format_number"

	// numberMarker prefixes the output of a number cell's template when the value
	// is a number, so render can write a numeric cell. NUL never appears in xlsx text.
	numberMarker = "\x00number:"

	// formatNumberNumFmt is what format_number displays: grouped, 2 decimals.
	formatNumberNumFmt = "#,##0.00"

	// maxPlainNumberDigits keeps barcodes and other long codes as text: Excel shows
	// General numbers of 12+ digits in scientific notation.
	maxPlainNumberDigits = 11
)

var plainNumberRgx = regexp.MustCompile(`^-?(0|[1-9][0-9]*)(\.[0-9]+)?$`)

var numberFunctions = template.FuncMap{
	plainNumberFunc:  capturePlainNumber,
	formatNumberFunc: captureFormatNumber,
}

// markNumberCell detects a cell whose whole content is a single action and
// rewrites its pipeline so the template reports a number to render.
func markNumberCell(tmpl *template.Template, templateFunctions template.FuncMap) numberCell {
	nodes := tmpl.Tree.Root.Nodes
	if len(nodes) != 1 {
		return numberCellNone
	}

	action, ok := nodes[0].(*tparse.ActionNode)
	if !ok || len(action.Pipe.Decl) > 0 {
		return numberCellNone
	}

	last := action.Pipe.Cmds[len(action.Pipe.Cmds)-1]
	if ident, ok := last.Args[0].(*tparse.IdentifierNode); ok &&
		ident.Ident == "format_number" && isDefaultFormatNumber(templateFunctions["format_number"]) {
		ident.Ident = formatNumberFunc

		return numberCellFormatNumber
	}

	capture := tparse.NewIdentifier(plainNumberFunc).SetTree(tmpl.Tree).SetPos(last.Pos)
	action.Pipe.Cmds = append(action.Pipe.Cmds, &tparse.CommandNode{
		NodeType: tparse.NodeCommand,
		Pos:      last.Pos,
		Args:     []tparse.Node{capture},
	})

	return numberCellPlain
}

// isDefaultFormatNumber reports whether format_number was not replaced through
// Options.CustomFuncMap: a custom one is treated as any other function.
func isDefaultFormatNumber(fn any) bool {
	return fn != nil && reflect.ValueOf(fn).Pointer() == reflect.ValueOf(FormatNumberForExcel).Pointer()
}

func capturePlainNumber(value any) any {
	if number, ok := plainNumberValue(value); ok {
		return encodeNumber(number)
	}

	return value
}

func captureFormatNumber(value any) string {
	if number, ok := formatNumberValue(value); ok {
		return encodeNumber(math.Round(number*100) / 100)
	}

	return FormatNumberForExcel(value)
}

func encodeNumber(number float64) string {
	return numberMarker + strconv.FormatFloat(number, 'g', -1, 64)
}

func decodeNumber(text string) (float64, bool) {
	encoded, ok := strings.CutPrefix(text, numberMarker)
	if !ok {
		return 0, false
	}

	number, err := strconv.ParseFloat(encoded, 64)

	return number, err == nil
}

// formatNumberValue reads value as a number for a format_number cell: a Go
// number, or any string strconv.ParseFloat reads, finite.
func formatNumberValue(value any) (float64, bool) {
	v := reflect.ValueOf(value)
	if v.Kind() != reflect.String {
		return numericValue(v)
	}

	number, err := strconv.ParseFloat(v.String(), 64)

	return number, err == nil && isFinite(number)
}

// plainNumberValue reads value as a number for a plain cell: a Go number, or a
// decimal string without leading zeros of at most maxPlainNumberDigits digits.
func plainNumberValue(value any) (float64, bool) {
	v := reflect.ValueOf(value)
	if v.Kind() != reflect.String {
		return numericValue(v)
	}

	s := v.String()
	if !plainNumberRgx.MatchString(s) {
		return 0, false
	}

	digits := len(strings.TrimPrefix(strings.Replace(s, ".", "", 1), "-"))
	if digits > maxPlainNumberDigits {
		return 0, false
	}

	number, err := strconv.ParseFloat(s, 64)

	return number, err == nil
}

func numericValue(v reflect.Value) (float64, bool) {
	var number float64

	switch v.Kind() {
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		number = float64(v.Int())
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64, reflect.Uintptr:
		number = float64(v.Uint())
	case reflect.Float32, reflect.Float64:
		number = v.Float()
	default:
		return 0, false
	}

	return number, isFinite(number)
}

func isFinite(number float64) bool {
	return !math.IsNaN(number) && !math.IsInf(number, 0)
}

// numberFormat is the number format of a rendered number: the template cell's
// own, or what format_number displays when a format_number cell is General.
func (c Cell) numberFormat() string {
	numFmt := c.cell.NumFmt
	if c.number == numberCellFormatNumber && (numFmt == "" || strings.EqualFold(numFmt, "general")) {
		return formatNumberNumFmt
	}

	return numFmt
}
