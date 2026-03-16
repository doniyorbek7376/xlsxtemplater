package xlsxtemplater

import (
	"fmt"
	"regexp"
	"strings"
	"text/template"

	"codeberg.org/tealeg/xlsx/v4"
)

var (
	rangeRgx = regexp.MustCompile(`\{\{\s*range\s+(.+)\s*\}\}`)
	ifRgx    = regexp.MustCompile(`\{\{\s*if\s+(.+)\s*\}\}`)
	endRgx   = regexp.MustCompile(`\{\{\s*end\s*\}\}`)
	elseRgx  = regexp.MustCompile(`\{\{\s*else\s*\}\}`)
)

func parse(file *xlsx.File, templateFunctions template.FuncMap) (*File, error) {
	sheetNodes := []*Sheet{}
	for _, sheet := range file.Sheets {
		sheetNode, err := parseSheet(sheet, templateFunctions)
		if err != nil {
			return nil, err
		}

		sheetNodes = append(sheetNodes, sheetNode)
	}

	return &File{Sheets: sheetNodes}, nil
}

func parseSheet(sheet *xlsx.Sheet, templateFunctions template.FuncMap) (*Sheet, error) {
	sheetNode := &Sheet{Name: sheet.Name, sheet: sheet}
	parentStack := []ParentNode{sheetNode}

	emptyRows := []*Row{}

	rowIndex := 0
	sheet.ForEachRow(func(row *xlsx.Row) error {
		rowIndex++

		rowNode := parseRow(row, rowIndex, templateFunctions)

		if rowNode.IsEmpty {
			emptyRows = append(emptyRows, rowNode)

			return nil
		}

		currentParent := lastParent(parentStack)

		for _, emptyRow := range emptyRows {
			currentParent.AddChild(emptyRow)
		}

		emptyRows = emptyRows[:0]

		if rowNode.IsRange {
			rangeNode := &Range{
				Expr: rowNode.Expr,
			}

			currentParent.AddChild(rangeNode)

			parentStack = append(parentStack, rangeNode)

			return nil
		}

		if rowNode.IsIf {
			conditionNode := &Condition{
				Expr: rowNode.Expr,
			}

			currentParent.AddChild(conditionNode)

			parentStack = append(parentStack, conditionNode)

			return nil
		}

		if rowNode.IsElse {
			if currentParent, ok := currentParent.(*Condition); ok {
				currentParent.elseFound = true
			}

			return nil

		}

		if rowNode.IsEnd {
			for _, emptyRow := range emptyRows {
				currentParent.AddChild(emptyRow)
			}

			emptyRows = emptyRows[:0]
			parentStack = parentStack[:len(parentStack)-1]

			return nil
		}

		currentParent.AddChild(rowNode)

		return nil
	})

	return sheetNode, nil
}

func parseRow(row *xlsx.Row, rowIndex int, templateFunctions template.FuncMap) *Row {
	rowNode := &Row{
		Index: rowIndex,
		row:   row,
	}

	if expr, ok := getRangeExpr(row); ok {
		rowNode.Expr = expr
		rowNode.IsRange = true

		return rowNode
	}

	if expr, ok := getIfExpr(row); ok {
		rowNode.Expr = expr
		rowNode.IsIf = true

		return rowNode
	}

	if isElse(row) {
		rowNode.IsElse = true

		return rowNode
	}

	if isEnd(row) {
		rowNode.IsEnd = true

		return rowNode
	}

	columnIndex := 1
	row.ForEachCell(func(cell *xlsx.Cell) error {
		cellName := getCellName(columnIndex, rowIndex)

		cellNode := &Cell{
			CellName: cellName,
			Col:      columnIndex,
			Raw:      cell.Value,
			cell:     cell,
		}

		if strings.Contains(cell.Value, "{{") {
			tmpl, err := template.New(cellName).
				Funcs(templateFunctions).
				Parse(cell.Value)
			if err != nil {
				println("warning: cannot parse template: " + cellName + " " + err.Error())
			}

			cellNode.Template = tmpl
		}

		if cell.Value != "" {
			rowNode.Cells = append(rowNode.Cells, cellNode)
		}

		columnIndex += 1 + cell.HMerge
		return nil
	})

	if len(rowNode.Cells) == 0 {
		rowNode.IsEmpty = true
	}

	return rowNode
}

func lastParent(stack []ParentNode) ParentNode {
	if len(stack) == 0 {
		return nil
	}

	return stack[len(stack)-1]
}

func getRangeExpr(row *xlsx.Row) (string, bool) {
	out := ""
	row.ForEachCell(func(cell *xlsx.Cell) error {
		value := cell.Value

		ok := rangeRgx.MatchString(value)
		if ok {
			out = rangeRgx.FindStringSubmatch(value)[1]
		}

		return nil
	})

	return out, out != ""
}

func getIfExpr(row *xlsx.Row) (string, bool) {
	out := ""
	row.ForEachCell(func(cell *xlsx.Cell) error {
		value := cell.Value

		ok := ifRgx.MatchString(value)
		if ok {
			out = ifRgx.FindStringSubmatch(value)[1]
		}

		return nil
	})

	return out, out != ""
}

func isEnd(row *xlsx.Row) bool {
	out := false
	row.ForEachCell(func(cell *xlsx.Cell) error {
		ok := endRgx.MatchString(cell.Value)
		if ok {
			out = ok
		}

		return nil
	})

	return out
}

func isElse(row *xlsx.Row) bool {
	out := false
	row.ForEachCell(func(cell *xlsx.Cell) error {
		ok := elseRgx.MatchString(cell.Value)
		if ok {
			out = ok
		}

		return nil
	})

	return out
}

func getCellName(column, row int) string {
	colName := ""
	for column > 0 {
		column--

		colName = string([]rune{'A' + rune(column%26)}) + colName
		column /= 26
	}

	return fmt.Sprintf("%s%d", colName, row)
}
