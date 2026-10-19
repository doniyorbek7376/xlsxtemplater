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

const (
	captureNumberFunc = "xlsxtemplater_capture_number"

	// numberMarker and formattedNumberMarker prefix the output of a cell's
	// template when its value is a number, so render can write a numeric cell.
	// NUL never appears in xlsx text.
	numberMarker          = "\x00number:"
	formattedNumberMarker = "\x00formatted_number:"

	// formatNumberNumFmt is what format_number displays: grouped, 2 decimals.
	formatNumberNumFmt = "#,##0.00"

	// maxPlainNumberDigits keeps barcodes and other long codes as text: Excel shows
	// General numbers of 12+ digits in scientific notation.
	maxPlainNumberDigits = 11
)

var plainNumberRgx = regexp.MustCompile(`^-?(0|[1-9][0-9]*)(\.[0-9]+)?$`)

var numberFunctions = template.FuncMap{
	captureNumberFunc: captureNumber,
}

// formattedNumber is what format_number returns for a number. It prints as
// "33 928,57"; a cell holding only it is a numeric cell shown with #,##0.00.
type formattedNumber float64

func (n formattedNumber) String() string {
	return FormatNumberForExcel(float64(n))
}

// formatNumber is the format_number template function: a number, rounded to
// 2 decimals, or the value printed as text when it is not a number.
func formatNumber(value any) any {
	if number, ok := plainNumberValue(value); ok {
		return formattedNumber(math.Round(number*100) / 100)
	}

	return FormatNumberForExcel(value)
}

// markNumberCell appends captureNumber to the pipeline of a cell whose whole
// content is a single action, so the template reports a number to render.
func markNumberCell(tmpl *template.Template) {
	nodes := tmpl.Tree.Root.Nodes
	if len(nodes) != 1 {
		return
	}

	action, ok := nodes[0].(*tparse.ActionNode)
	if !ok || len(action.Pipe.Decl) > 0 {
		return
	}

	last := action.Pipe.Cmds[len(action.Pipe.Cmds)-1]
	capture := tparse.NewIdentifier(captureNumberFunc).SetTree(tmpl.Tree).SetPos(last.Pos)
	action.Pipe.Cmds = append(action.Pipe.Cmds, &tparse.CommandNode{
		NodeType: tparse.NodeCommand,
		Pos:      last.Pos,
		Args:     []tparse.Node{capture},
	})
}

func captureNumber(value any) any {
	if number, ok := value.(formattedNumber); ok {
		return formattedNumberMarker + strconv.FormatFloat(float64(number), 'g', -1, 64)
	}

	if number, ok := plainNumberValue(value); ok {
		return numberMarker + strconv.FormatFloat(number, 'g', -1, 64)
	}

	return value
}

// decodeNumber reads the output of a cell's template back as a number;
// formatted tells a format_number value.
func decodeNumber(text string) (number float64, formatted, ok bool) {
	encoded, formatted := strings.CutPrefix(text, formattedNumberMarker)
	if !formatted {
		if encoded, ok = strings.CutPrefix(text, numberMarker); !ok {
			return 0, false, false
		}
	}

	number, err := strconv.ParseFloat(encoded, 64)

	return number, formatted, err == nil
}

// plainNumberValue reads value as a number: a Go number, or a decimal string
// without leading zeros of at most maxPlainNumberDigits digits.
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
// own, or what format_number displays when the cell is General.
func (c Cell) numberFormat(formatted bool) string {
	numFmt := c.cell.NumFmt
	if formatted && (numFmt == "" || strings.EqualFold(numFmt, "general")) {
		return formatNumberNumFmt
	}

	return numFmt
}
