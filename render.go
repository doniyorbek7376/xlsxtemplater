package xlsxtemplater

import (
	"bytes"
	"fmt"
	"sync"
	"text/template"

	"codeberg.org/tealeg/xlsx/v4"
)

func (f *File) render(file *xlsx.File, content any, templateFunctions template.FuncMap) error {
	for _, sheetNode := range f.Sheets {
		sheet, err := file.AddSheet(sheetNode.Name)
		if err != nil {
			return err
		}

		err = sheetNode.render(sheet, content, templateFunctions)
		if err != nil {
			return err
		}
	}

	return nil
}

func (n *Sheet) render(sheet *xlsx.Sheet, content any, templateFunctions template.FuncMap) error {
	// copy old sheet column
	n.sheet.Cols.ForEach(func(idx int, col *xlsx.Col) {
		newCol := xlsx.Col{}
		style := col.GetStyle()
		newCol.SetStyle(style)
		newCol.Width = col.Width
		newCol.CustomWidth = col.CustomWidth
		newCol.Hidden = col.Hidden
		newCol.Collapsed = col.Collapsed
		newCol.Min = col.Min
		newCol.Max = col.Max

		sheet.Cols.Add(&newCol)
	})

	sheet.SheetViews = n.sheet.SheetViews

	for _, node := range n.Nodes {
		render(sheet, node, content, templateFunctions)
	}

	return nil
}

func render(sheet *xlsx.Sheet, node Node, content any, templateFunctions template.FuncMap) {
	switch node := node.(type) {
	case *Row:
		renderRow(sheet, node, content)
	case *Range:
		renderRange(sheet, node, content, templateFunctions)
	case *Condition:
		renderCondition(sheet, node, content, templateFunctions)
	}
}

func renderRange(sheet *xlsx.Sheet, node *Range, content any, templateFunctions template.FuncMap) {
	items := extractSlice(content, node.Expr)

	for _, item := range items {
		for _, child := range node.Body {
			render(sheet, child, item, templateFunctions)
		}
	}
}

func renderCondition(sheet *xlsx.Sheet, node *Condition, content any, templateFunctions template.FuncMap) {
	var nodes []Node
	if checkCondition(node.Expr, content, templateFunctions) {
		nodes = node.Body
	} else {
		nodes = node.Else
	}

	for _, child := range nodes {
		render(sheet, child, content, templateFunctions)
	}
}

var (
	conditionCache = make(map[string]*template.Template)
	conditionMu    sync.Mutex
)

func checkCondition(expr string, content any, templateFunctions template.FuncMap) bool {
	checkerTemplate := fmt.Sprintf("{{ %s }}", expr)

	conditionMu.Lock()
	tmpl, ok := conditionCache[checkerTemplate]
	if !ok {
		tmpl, _ = template.New("checker").
			Funcs(templateFunctions).
			Parse(checkerTemplate)
		conditionCache[checkerTemplate] = tmpl
	}
	conditionMu.Unlock()

	if tmpl == nil {
		return false
	}

	buf := bytes.NewBuffer(nil)

	_ = tmpl.Execute(buf, content)

	return buf.String() == "true"
}

func renderRow(sheet *xlsx.Sheet, row *Row, content any) {
	newRow := sheet.AddRow()
	height := row.row.GetHeight()
	if height != 0 {
		newRow.SetHeight(height)
	}

	columnIndex := 1
	for _, cell := range row.Cells {
		for cell.Col > columnIndex {
			newRow.AddCell()
			columnIndex++
		}

		newCell := newRow.AddCell()
		columnIndex++

		cloneCell(newCell, cell.cell)
		for i := 0; i < newCell.HMerge; i++ {
			row.row.AddCell()
			columnIndex++
		}

		newCell.Merge(newCell.HMerge, newCell.VMerge)

		newCell.SetValue(cell.GetValue(content))
	}
}

func cloneCell(to, from *xlsx.Cell) {
	to.HMerge = from.HMerge
	to.VMerge = from.VMerge
	to.Hidden = from.Hidden
	style := from.GetStyle()
	if style != nil {
		to.SetStyle(style)
	}
	to.Value = from.Value
	to.SetFormula(from.Formula())
	to.NumFmt = from.NumFmt

	to.SetHyperlink(from.Hyperlink.Link, from.Hyperlink.DisplayString, from.Hyperlink.Tooltip)
}
