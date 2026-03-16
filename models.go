package xlsxtemplater

import (
	"bytes"
	"fmt"
	"strings"
	"sync"
	"text/template"

	"codeberg.org/tealeg/xlsx/v4"
)

var bufferPool = sync.Pool{
	New: func() interface{} {
		return &bytes.Buffer{}
	},
}

type NodeType int

const (
	NodeTypeSheet NodeType = iota
	NodeTypeRow
	NodeTypeCell
	NodeTypeRange
	NodeTypeCondition
)

type Node interface {
	Repr() string
}

type ParentNode interface {
	Node
	AddChild(Node)
}

type File struct {
	Sheets []*Sheet
}

type Sheet struct {
	Name  string
	Nodes []Node

	sheet *xlsx.Sheet
}

func (n *Sheet) AddChild(child Node) {
	n.Nodes = append(n.Nodes, child)
}

type Row struct {
	Index int
	Cells []*Cell

	row *xlsx.Row
}

type Cell struct {
	CellName string
	Col      int
	Raw      string

	Template *template.Template
	cell     *xlsx.Cell
}

func (c Cell) GetValue(content any) string {
	if c.Template == nil {
		return c.Raw
	}

	// Skip template execution for static content (no template delimiters)
	if !strings.Contains(c.Raw, "{{") {
		return c.Raw
	}

	buf := bufferPool.Get().(*bytes.Buffer)
	buf.Reset()
	defer bufferPool.Put(buf)

	err := c.Template.Execute(buf, content)
	if err != nil {
		return c.Raw
	}

	return buf.String()
}

type Range struct {
	Expr string
	Body []Node
}

func (n *Range) AddChild(child Node) {
	n.Body = append(n.Body, child)
}

type Condition struct {
	elseFound bool
	Expr      string
	Body      []Node
	Else      []Node
}

func (n *Condition) AddChild(child Node) {
	if n.elseFound {
		n.Else = append(n.Else, child)
	} else {
		n.Body = append(n.Body, child)
	}
}

func (n Sheet) Type() NodeType {
	return NodeTypeSheet
}

func (n Sheet) Repr() string {
	var sb strings.Builder

	sb.WriteString("Sheet " + n.Name + "\n")
	for _, child := range n.Nodes {
		childRepr := child.Repr()
		for line := range strings.SplitSeq(childRepr, "\n") {
			sb.WriteString("----" + line)
			sb.WriteString("\n")
		}
	}

	return sb.String()
}

func (n Row) Type() NodeType {
	return NodeTypeRow
}

func (n Row) Repr() string {
	var sb strings.Builder

	sb.WriteString("Row " + fmt.Sprint(n.Index) + "\n")

	for _, cell := range n.Cells {
		sb.WriteString("----" + cell.Repr() + "\n")
	}

	return sb.String()
}

func (n Cell) Type() NodeType {
	return NodeTypeCell
}

func (n Cell) Repr() string {
	return "Cell " + n.CellName + " " + n.Raw + " " + fmt.Sprint(n.Template)
}

func (n Range) Type() NodeType {
	return NodeTypeRange
}

func (n Range) Repr() string {
	var sb strings.Builder
	sb.WriteString("Range " + n.Expr + "\n")

	for _, child := range n.Body {
		childRepr := child.Repr()
		for line := range strings.SplitSeq(childRepr, "\n") {
			sb.WriteString("----" + line)
			sb.WriteString("\n")
		}
	}

	return sb.String()
}

func (n Condition) Type() NodeType {
	return NodeTypeCondition
}

func (n Condition) Repr() string {
	var sb strings.Builder
	sb.WriteString("If " + n.Expr + "\n")
	sb.WriteString("Then\n")

	for _, child := range n.Body {
		childRepr := child.Repr()
		for line := range strings.SplitSeq(childRepr, "\n") {
			sb.WriteString("----" + line)
			sb.WriteString("\n")
		}
	}

	if n.elseFound || len(n.Else) > 0 {
		sb.WriteString("ELSE\n")
	}

	for _, child := range n.Else {
		childRepr := child.Repr()
		for line := range strings.SplitSeq(childRepr, "\n") {
			sb.WriteString("----" + line)
			sb.WriteString("\n")
		}
	}

	return sb.String()
}
