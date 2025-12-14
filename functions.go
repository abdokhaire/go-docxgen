package docxtpl

import (
	"fmt"
	"maps"
	"text/template"

	"github.com/abdokhaire/go-docxgen/internal/functions"
)

// Register a function which can then be used within your template
//
//	d.RegisterFunction("sayHello", func(text string) string {
//		return "Hello " + text
//	})
func (d *DocxTmpl) RegisterFunction(name string, fn any) error {
	if !functions.FunctionNameValid(name) {
		return fmt.Errorf("function name %q is not a valid identifier", name)
	}
	// Go's text/template handles function signature validation at execution time
	d.funcMap[name] = fn
	return nil
}

// Get a pointer to the documents function map. This will include built-in functions.
func (d *DocxTmpl) GetRegisteredFunctions() *template.FuncMap {
	copiedFuncMap := make(template.FuncMap)
	maps.Copy(copiedFuncMap, d.funcMap)
	return &copiedFuncMap
}

// RegisterFuncMap registers all functions from a template.FuncMap.
// This is useful for adding external function libraries like Sprig.
//
//	doc.RegisterFuncMap(sprig.FuncMap())
func (d *DocxTmpl) RegisterFuncMap(funcs template.FuncMap) {
	maps.Copy(d.funcMap, funcs)
}
