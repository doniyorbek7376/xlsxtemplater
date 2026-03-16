package xlsxtemplater

import (
	"strings"
	"text/template"

	"codeberg.org/tealeg/xlsx/v4"
)

// Template is a parsed tree from given xlsx file.
type Template struct {
	parsedFile        *File
	templateFunctions template.FuncMap
}

func (t Template) String() string {
	var sb strings.Builder

	for _, sheet := range t.parsedFile.Sheets {
		sb.WriteString(sheet.Repr())
		sb.WriteRune('\n')
	}

	return sb.String()
}

// Render generates xlsx file from the content and saves to the disk.
func (t Template) Render(content any, saveAs string) error {
	file := xlsx.NewFile()

	err := t.parsedFile.render(file, content, t.templateFunctions)
	if err != nil {
		return err
	}

	return file.Save(saveAs)
}
