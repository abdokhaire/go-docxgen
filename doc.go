// Package docxtpl provides a library for generating DOCX documents from templates
// using Go's text/template syntax.
//
// This package wraps the go-docx library (github.com/fumiama/go-docx) and adds
// template rendering capabilities, making it easy to create dynamic Word documents.
//
// # Key Features
//
//   - Go template syntax ({{.Field}}, {{range}}, {{if}}) in Word documents
//   - Automatic handling of fragmented XML tags (common in Word's editing)
//   - Inline image insertion with automatic sizing and DPI detection
//   - 60+ built-in template functions for text, math, dates, and more
//   - Support for headers, footers, footnotes, endnotes, and watermarks
//   - Custom function registration
//
// # Basic Usage
//
// Create a Word document with template placeholders like {{.Name}}, then:
//
//	doc, err := docxtpl.ParseFromFilename("template.docx")
//	if err != nil {
//	    log.Fatal(err)
//	}
//
//	data := map[string]any{
//	    "Name":    "John Doe",
//	    "Company": "Acme Corp",
//	}
//
//	if err := doc.Render(data); err != nil {
//	    log.Fatal(err)
//	}
//
//	if err := doc.SaveToFile("output.docx"); err != nil {
//	    log.Fatal(err)
//	}
//
// # Template Syntax
//
// The package supports standard Go text/template syntax:
//
//   - Simple fields: {{.Name}}
//   - Nested fields: {{.Person.Address.City}}
//   - Conditionals: {{if .Active}}...{{end}}
//   - Loops: {{range .Items}}{{.Name}}{{end}}
//   - Functions: {{.Name | upper}} or {{formatMoney .Price "$"}}
//
// # Built-in Functions
//
// The package includes many built-in functions:
//
// Text: upper, lower, title, bold, italic, underline, color, highlight
// Numbers: formatNumber, formatMoney, formatPercent, add, sub, mul, div
// Dates: now, formatDate, parseDate, addDays, addMonths, addYears
// Collections: len, first, last, join, contains, index, slice
// Logic: eq, ne, lt, le, gt, ge, and, or, not
// Utilities: default, ternary, trim, replace, truncate, pluralize
//
// # Inline Images
//
// Images can be inserted using file paths (auto-detected) or InlineImage:
//
//	// Auto-detect from file path
//	data := map[string]any{
//	    "Logo": "/path/to/logo.png",
//	}
//
//	// Or use InlineImage for size control
//	logo, _ := docxtpl.CreateInlineImage("/path/to/logo.png")
//	logo.Resize(200, 100)
//	data := map[string]any{
//	    "Logo": logo,
//	}
//
// # Custom Functions
//
// Register custom template functions before rendering:
//
//	doc.RegisterFunction("greet", func(name string) string {
//	    return "Hello, " + name + "!"
//	})
//
// Then use in template: {{greet .Name}}
//
// # Data Types
//
// The Render method accepts structs or maps with various field types:
//
//   - Primitives: string, int, float64, bool, etc.
//   - Structs: regular and nested structs
//   - Pointers: automatically dereferenced
//   - Slices: []string, []int, []Struct, []*Struct
//   - Maps: map[string]any, map[string]string, etc.
package docxtpl
