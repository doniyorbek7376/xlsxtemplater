package xlsxtemplater

import (
	"maps"
	"text/template"

	"codeberg.org/tealeg/xlsx/v4"
)

type Options struct {
	CustomFuncMap template.FuncMap
}

// ParseTemplate parses the xlsx file in templatePath, constructs AST and returns Template pointer.
func ParseTemplate(templatePath string) (*Template, error) {
	return ParseTemplateWithOptions(templatePath, nil)
}

// ParseTemplateWithOptions accepts options with CustomFuncMap for template functions.
func ParseTemplateWithOptions(templatePath string, options *Options) (*Template, error) {
	templateFunctions := getDefaultTemplateFunctions()

	if options != nil {
		maps.Copy(templateFunctions, options.CustomFuncMap)
	}

	file, err := xlsx.OpenFile(templatePath)
	if err != nil {
		return nil, err
	}

	parsed, err := parse(file, templateFunctions)
	if err != nil {
		return nil, err
	}

	return &Template{parsedFile: parsed}, nil
}

// Generate opens xlsx template from templatePath, renders with given content
// and writes a generated file to generatedFilePath.
func Generate(templatePath string, content any, generatedFilePath string) error {
	return GenerateWithOptions(templatePath, content, generatedFilePath, nil)
}

// GenerateWithOptions accepts Options with CustomFuncMap for template functions.
func GenerateWithOptions(templatePath string, content any, generatedFilePath string, options *Options) error {
	template, err := ParseTemplateWithOptions(templatePath, options)
	if err != nil {
		return err
	}

	return template.Render(content, generatedFilePath)
}
