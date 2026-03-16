package xlsxtemplater

import (
	"fmt"
	"maps"
	"text/template"

	"codeberg.org/tealeg/xlsx/v4"
)

type Options struct {
	CustomFuncMap template.FuncMap
}

// Generate opens xlsx template from templatePath, renders with given content
// and writes a generated file to generatedFilePath.
func Generate(templatePath string, content any, generatedFilePath string) error {
	return GenerateWithOptions(templatePath, content, generatedFilePath, nil)
}

// GenerateWithOptions accepts Options with CustomFuncMap for template functions.
func GenerateWithOptions(templatePath string, content any, generatedFilePath string, options *Options) error {
	templateFunctions := getDefaultTemplateFunctions()

	if options != nil {
		maps.Copy(templateFunctions, options.CustomFuncMap)
	}

	file, err := xlsx.OpenFile(templatePath)
	if err != nil {
		return err
	}

	parsed, err := parse(file, templateFunctions)
	if err != nil {
		return err
	}

	fmt.Println(parsed.Sheets[0].Repr())

	renderedFile := xlsx.NewFile()

	err = parsed.render(renderedFile, content, templateFunctions)
	if err != nil {
		return err
	}

	return renderedFile.Save(generatedFilePath)
}
