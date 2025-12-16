# go-docxgen

A Go library for generating DOCX documents from templates using Go's `text/template` syntax.

[![Go Reference](https://pkg.go.dev/badge/github.com/abdokhaire/go-docxgen.svg)](https://pkg.go.dev/github.com/abdokhaire/go-docxgen)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## Features

- **Go Template Syntax** - Use familiar `{{.Field}}`, `{{range}}`, `{{if}}` syntax directly in Word documents
- **Programmatic Document Creation** - Build documents from scratch with fluent API
- **Handles Fragmented Tags** - Automatically merges tags split across XML runs by Word's editing
- **Inline Images** - Insert images dynamically with automatic sizing and DPI detection
- **Extensible Functions** - Register custom functions or use community libraries like Sprig/Sprout
- **Full Document Support** - Headers, footers, footnotes, endnotes, and document properties
- **Document Properties** - Get/set title, author, subject, keywords, and more
- **Watermarks** - Extract and replace watermark text (supports template syntax)
- **Flexible Data Types** - Structs, maps, slices, pointers, nested structures, and graceful nil handling

## Installation

```sh
go get github.com/abdokhaire/go-docxgen@latest
```

## Quick Start

Create a Word document with template placeholders like `{{.Name}}`, then:

```go
package main

import (
    "os"
    docxtpl "github.com/abdokhaire/go-docxgen"
)

func main() {
    // Parse the template
    doc, err := docxtpl.ParseFromFilename("template.docx")
    if err != nil {
        panic(err)
    }

    // Render with data
    err = doc.Render(map[string]any{
        "Name":    "John Doe",
        "Company": "Acme Corp",
        "Items":   []string{"Item 1", "Item 2", "Item 3"},
    })
    if err != nil {
        panic(err)
    }

    // Save the result
    doc.SaveToFile("output.docx")
}
```

## Programmatic Document Creation

Create documents without templates using the builder API:

```go
package main

import docxtpl "github.com/abdokhaire/go-docxgen"

func main() {
    // Create a new document
    doc := docxtpl.New()

    // Set document properties
    doc.SetTitle("My Report")
    doc.SetAuthor("Go Developer")

    // Add heading and paragraphs
    doc.AddHeading("Introduction", 1)
    doc.AddParagraph("This document was created programmatically.")

    // Add formatted text
    para := doc.AddParagraph("This is ")
    para.AddText("bold").Bold()
    para.AddText(" and this is ").Then().AddText("italic").Italic()

    // Add a table
    doc.AddHeading("Data Table", 2)
    table := doc.AddTable(3, 3)

    // Header row with styling
    table.SetCell(0, 0, "Name").Bold().Background("CCCCCC")
    table.SetCell(0, 1, "Age").Bold().Background("CCCCCC")
    table.SetCell(0, 2, "City").Bold().Background("CCCCCC")

    // Data rows
    table.SetCell(1, 0, "Alice")
    table.SetCell(1, 1, "30")
    table.SetCell(1, 2, "New York")

    doc.SaveToFile("generated.docx")
}
```

### Paragraph Formatting

```go
doc.AddParagraph("Centered text").Center()
doc.AddParagraph("Right aligned").Right()
doc.AddParagraph("Red bold text").Bold().Color("FF0000")
doc.AddParagraph("Large text").SizePoints(16)
doc.AddParagraph("Custom font").Font("Arial")
doc.AddParagraph("Highlighted").Highlight("yellow")
doc.AddParagraph("With background").Background("FFFF00")
```

### Table Builder

```go
// Create table with custom column widths (in twips: 1 inch = 1440 twips)
table := doc.AddTableWithWidths(5, []int{2880, 1440, 1440}) // 2", 1", 1"

// Access and format cells
cell := table.Cell(0, 0)
cell.SetText("Header")
cell.Bold()
cell.Center()
cell.Background("E0E0E0")

// Format entire row
row := table.Row(0)
row.Justify("center")
```

### Document Properties

```go
// Set properties individually
doc.SetTitle("Annual Report 2025")
doc.SetAuthor("Finance Team")
doc.SetSubject("Q4 Financial Summary")
doc.SetKeywords("finance, report, quarterly")
doc.SetContentStatus("Final")

// Or set all at once
props := doc.GetProperties()
props.Title = "My Document"
props.Creator = "Go Application"
doc.SetProperties(props)
```

## Template Syntax

Use standard Go `text/template` syntax in your DOCX files:

| Syntax | Description | Example |
|--------|-------------|---------|
| `{{.Field}}` | Simple field | `{{.Name}}` |
| `{{.Nested.Field}}` | Nested struct/map | `{{.Person.Address.City}}` |
| `{{if .Condition}}...{{end}}` | Conditional | `{{if .Active}}Active{{end}}` |
| `{{range .Items}}...{{end}}` | Loop | `{{range .Products}}{{.Name}}{{end}}` |
| `{{.Field \| function}}` | Pipe to function | `{{.Name \| upper}}` (requires registered function) |
| `{{function .Args}}` | Function call | `{{greet .Name}}` (requires registered function) |

## Template Functions

This library provides a flexible function system. You can register your own custom functions or use popular community function libraries.

### Built-in Function

The library includes one built-in function:

| Function | Example | Description |
|----------|---------|-------------|
| `link` | `{{link "https://example.com" "Click here"}}` | Create a clickable hyperlink |

### Registering Custom Functions

Register your own functions before rendering:

```go
doc, _ := docxtpl.ParseFromFilename("template.docx")

// Simple function
doc.RegisterFunction("greet", func(name string) string {
    return "Hello, " + name + "!"
})

// Function with multiple parameters
doc.RegisterFunction("formatPrice", func(price float64, currency string) string {
    return fmt.Sprintf("%s%.2f", currency, price)
})

// Function for calculations
doc.RegisterFunction("multiply", func(a, b float64) float64 {
    return a * b
})

doc.Render(data)
```

**Template usage:**
```
{{greet .Name}}
Price: {{formatPrice .Amount "$"}}
Total: {{multiply .Price .Quantity}}
```

### Using Community Function Libraries

For a rich set of template functions, use community libraries like **Sprig** or **Sprout**.

#### Using Sprig (100+ functions)

[Sprig](https://github.com/Masterminds/sprig) provides 100+ template functions for strings, math, dates, and more.

```bash
go get github.com/Masterminds/sprig/v3
```

```go
import (
    "github.com/Masterminds/sprig/v3"
    docxtpl "github.com/abdokhaire/go-docxgen"
)

func main() {
    doc, _ := docxtpl.ParseFromFilename("template.docx")

    // Register all Sprig functions
    doc.RegisterFuncMap(sprig.FuncMap())

    doc.Render(data)
    doc.SaveToFile("output.docx")
}
```

**Template with Sprig functions:**
```
Name: {{.Name | upper}}
Email: {{.Email | lower}}
Date: {{now | date "2006-01-02"}}
Total: {{.Items | len}} items
Status: {{.IsActive | ternary "Active" "Inactive"}}
```

#### Using Sprout (Modern Sprig Alternative)

[Sprout](https://github.com/go-sprout/sprout) is a modern, modular alternative to Sprig with better performance.

```bash
go get github.com/go-sprout/sprout
```

```go
import (
    "github.com/go-sprout/sprout"
    docxtpl "github.com/abdokhaire/go-docxgen"
)

func main() {
    doc, _ := docxtpl.ParseFromFilename("template.docx")

    // Register Sprout functions
    handler := sprout.New()
    doc.RegisterFuncMap(handler.Build())

    doc.Render(data)
    doc.SaveToFile("output.docx")
}
```

### Common Functions (via Sprig/Sprout)

Once you register Sprig or Sprout, these functions become available:

| Category | Functions |
|----------|-----------|
| **Strings** | `upper`, `lower`, `title`, `trim`, `replace`, `contains`, `hasPrefix`, `hasSuffix` |
| **Math** | `add`, `sub`, `mul`, `div`, `mod`, `max`, `min`, `ceil`, `floor`, `round` |
| **Dates** | `now`, `date`, `dateModify`, `toDate`, `dateInZone` |
| **Lists** | `list`, `first`, `last`, `append`, `prepend`, `concat`, `join` |
| **Logic** | `eq`, `ne`, `lt`, `le`, `gt`, `ge`, `and`, `or`, `not`, `default`, `ternary` |
| **Encoding** | `b64enc`, `b64dec`, `toJson`, `fromJson` |

See [Sprig documentation](http://masterminds.github.io/sprig/) or [Sprout documentation](https://docs.atom.codes/sprout) for the complete function reference.

## Inline Images

Insert images dynamically using file paths or the `InlineImage` type:

```go
// Using file path (auto-detected)
data := map[string]any{
    "Logo": "/path/to/logo.png",
}

// Using InlineImage for more control
logo := docxtpl.CreateInlineImage("/path/to/logo.png")
logo.Resize(200, 100) // width, height in pixels

data := map[string]any{
    "Logo": logo,
}
```

Supported formats: JPEG (.jpg, .jpeg) and PNG (.png)

## Working with Document Parts

### Get All Placeholders

```go
placeholders := doc.GetPlaceholders()
// Returns: []string{"Name", "Company", "Items", ...}
```

### Watermarks

```go
// Get watermarks from headers
watermarks := doc.GetWatermarks()

// Replace watermark text (supports template syntax)
doc.ReplaceWatermark("DRAFT", "FINAL")

// Or use template syntax in the watermark itself: {{.Status}}
doc.Render(map[string]any{"Status": "APPROVED"})
```

## The Fragmentation Problem

Word processors often split text across multiple XML elements based on editing history, spell-check, or formatting. A placeholder like `{{.FirstName}}` might appear in the XML as:

```xml
<w:r><w:t>{{.First</w:t></w:r>
<w:r><w:t>Name}}</w:t></w:r>
```

This library automatically detects and merges these fragmented tags before template processing, ensuring reliable placeholder replacement.

## API Reference

### Parsing
- `Parse(reader io.ReaderAt, size int64)` - Parse from reader
- `ParseFromFilename(filename string)` - Parse from file path
- `ParseFromBytes(data []byte)` - Parse from byte slice

### Rendering
- `Render(data any)` - Replace placeholders with data
- `RegisterFunction(name string, fn any)` - Add custom function

### Saving
- `Save(writer io.Writer)` - Save to writer
- `SaveToFile(filename string)` - Save to file path

### Inspection
- `GetPlaceholders()` - Get all unique placeholders
- `GetWatermarks()` - Get watermark texts
- `ReplaceWatermark(old, new string)` - Replace watermark text

## Examples

### Basic Document with Struct Data

**Template (template.docx):**
```
Project: {{.ProjectNumber}}
Client: {{.Client}}
Status: {{.Status}}
```

**Code:**
```go
package main

import docxtpl "github.com/abdokhaire/go-docxgen"

func main() {
    doc, _ := docxtpl.ParseFromFilename("template.docx")

    data := struct {
        ProjectNumber string
        Client        string
        Status        string
    }{
        ProjectNumber: "PRJ-2025-001",
        Client:        "Acme Corporation",
        Status:        "In Progress",
    }

    doc.Render(data)
    doc.SaveToFile("output.docx")
}
```

### Tables with Loops

**Template (invoice.docx):**
```
Invoice #{{.InvoiceNumber}}
Date: {{.Date}}

| Item | Quantity | Price |
|------|----------|-------|
{{range .Items}}| | | |
|------|----------|-------|
{{.Name}} | {{.Qty}} | {{.PriceFormatted}} |
|------|----------|-------|
{{end}}
|------|----------|-------|

Total: {{.TotalFormatted}}
```

**Code:**
```go
type Item struct {
    Name           string
    Qty            int
    Price          float64
    PriceFormatted string
}

data := map[string]any{
    "InvoiceNumber": "INV-001",
    "Date":          "December 5, 2025",
    "Items": []Item{
        {Name: "Widget A", Qty: 10, Price: 29.99, PriceFormatted: "$29.99"},
        {Name: "Widget B", Qty: 5, Price: 49.99, PriceFormatted: "$49.99"},
        {Name: "Service Fee", Qty: 1, Price: 100.00, PriceFormatted: "$100.00"},
    },
    "TotalFormatted": "$549.85",
}

doc.Render(data)
```

### Conditionals

**Template:**
```
{{if .IsApproved}}
✓ This document has been approved by {{.ApprovedBy}}
{{else}}
⏳ Pending approval
{{end}}

{{if gt .Amount 1000}}
⚠️ Large transaction - requires manager approval
{{end}}
```

**Code:**
```go
data := map[string]any{
    "IsApproved": true,
    "ApprovedBy": "John Smith",
    "Amount":     1500.00,
}
```

### Images in Documents

**Template:**
```
Company Logo: {{.Logo}}

Team Members:
{{range .Team}}
  Name: {{.Name}}
  Photo: {{.Photo}}
{{end}}
```

**Code:**
```go
// Method 1: Auto-detect from file path
data := map[string]any{
    "Logo": "/path/to/logo.png",
    "Team": []map[string]any{
        {"Name": "Alice", "Photo": "/path/to/alice.jpg"},
        {"Name": "Bob", "Photo": "/path/to/bob.jpg"},
    },
}

// Method 2: Using InlineImage for size control
logo, _ := docxtpl.CreateInlineImage("/path/to/logo.png")
logo.Resize(150, 50) // width x height in pixels

profile, _ := docxtpl.CreateInlineImage("/path/to/profile.jpg")
profile.Resize(100, 100)

data := map[string]any{
    "Logo":    logo,
    "Profile": profile,
}
```

### Using Sprig/Sprout Functions

This example requires registering Sprig or Sprout functions first (see [Template Functions](#template-functions)).

**Template:**
```
Name: {{.Name | upper}}
Email: {{.Email | lower}}
Title: {{.Title | title}}

Joined: {{.JoinDate | date "January 2, 2006"}}
Salary: ${{.Salary | printf "%.2f"}}
Bonus: {{.BonusRate | mul 100 | printf "%.1f"}}%

Status: {{ternary "Active" "Inactive" .IsActive}}
Department: {{default "Unassigned" .Department}}

Tags: {{.Tags | join ", "}}
```

**Code:**
```go
import (
    "time"
    "github.com/Masterminds/sprig/v3"
    docxtpl "github.com/abdokhaire/go-docxgen"
)

func main() {
    doc, _ := docxtpl.ParseFromFilename("template.docx")
    doc.RegisterFuncMap(sprig.FuncMap()) // Register Sprig functions

    data := map[string]any{
        "Name":       "john doe",
        "Email":      "JOHN@EXAMPLE.COM",
        "Title":      "senior developer",
        "JoinDate":   time.Now(),
        "Salary":     85000.00,
        "BonusRate":  0.15,
        "IsActive":   true,
        "Department": "", // Will show "Unassigned"
        "Tags":       []string{"golang", "backend", "api"},
    }

    doc.Render(data)
    doc.SaveToFile("output.docx")
}
```

### Custom Functions

**Template:**
```
{{greet .Name}}
Order Total: {{calculateTotal .Items}}
```

**Code:**
```go
doc, _ := docxtpl.ParseFromFilename("template.docx")

// Register custom functions before rendering
doc.RegisterFunction("greet", func(name string) string {
    return "Welcome, " + name + "!"
})

doc.RegisterFunction("calculateTotal", func(items []map[string]any) string {
    total := 0.0
    for _, item := range items {
        price := item["price"].(float64)
        qty := item["qty"].(int)
        total += price * float64(qty)
    }
    return fmt.Sprintf("$%.2f", total)
})

data := map[string]any{
    "Name": "Alice",
    "Items": []map[string]any{
        {"name": "Item 1", "price": 10.0, "qty": 2},
        {"name": "Item 2", "price": 25.0, "qty": 1},
    },
}

doc.Render(data)
```

### Multi-line Text

Newlines in your data are automatically converted to line breaks:

```go
data := map[string]any{
    "Address": "123 Main Street\nSuite 456\nNew York, NY 10001",
    "Notes":   "Line 1\nLine 2\nLine 3",
}
```

### Loading from Different Sources

```go
// From file path
doc, _ := docxtpl.ParseFromFilename("template.docx")

// From bytes (useful for embedded templates or HTTP responses)
templateBytes, _ := os.ReadFile("template.docx")
doc, _ := docxtpl.ParseFromBytes(templateBytes)

// From io.ReaderAt (useful for http.Response.Body after reading)
file, _ := os.Open("template.docx")
stat, _ := file.Stat()
doc, _ := docxtpl.Parse(file, stat.Size())
```

### Saving to Different Destinations

```go
// To file path
doc.SaveToFile("output.docx")

// To io.Writer (useful for HTTP responses)
func handleDownload(w http.ResponseWriter, r *http.Request) {
    doc, _ := docxtpl.ParseFromFilename("template.docx")
    doc.Render(data)

    w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    w.Header().Set("Content-Disposition", "attachment; filename=document.docx")
    doc.Save(w)
}
```

### Working with JSON Data

```go
import "encoding/json"

// Load JSON data
jsonData := `{
    "name": "John Doe",
    "company": "Acme Corp",
    "items": [
        {"product": "Widget", "price": 29.99},
        {"product": "Gadget", "price": 49.99}
    ]
}`

var data map[string]any
json.Unmarshal([]byte(jsonData), &data)

doc, _ := docxtpl.ParseFromFilename("template.docx")
doc.Render(data)
doc.SaveToFile("output.docx")
```

### Inspecting Template Placeholders

```go
doc, _ := docxtpl.ParseFromFilename("template.docx")

// Get all placeholders to know what data is needed
placeholders, _ := doc.GetPlaceholders()
fmt.Println("Required fields:", placeholders)
// Output: [{{.Name}} {{.Email}} {{.Items}} ...]

// Useful for validation or building dynamic forms
```

### Complete Example: Generate Report

```go
package main

import (
    "time"
    docxtpl "github.com/abdokhaire/go-docxgen"
)

type Employee struct {
    Name       string
    Department string
    Salary     float64
    Photo      string
}

type Report struct {
    Title       string
    GeneratedAt time.Time
    Author      string
    Employees   []Employee
    TotalBudget float64
    IsFinalized bool
}

func main() {
    doc, err := docxtpl.ParseFromFilename("report_template.docx")
    if err != nil {
        panic(err)
    }

    report := Report{
        Title:       "Q4 2025 Staff Report",
        GeneratedAt: time.Now(),
        Author:      "HR Department",
        Employees: []Employee{
            {Name: "Alice Johnson", Department: "Engineering", Salary: 95000, Photo: "photos/alice.jpg"},
            {Name: "Bob Smith", Department: "Marketing", Salary: 75000, Photo: "photos/bob.jpg"},
            {Name: "Carol White", Department: "Engineering", Salary: 105000, Photo: "photos/carol.jpg"},
        },
        TotalBudget: 275000,
        IsFinalized: true,
    }

    if err := doc.Render(report); err != nil {
        panic(err)
    }

    if err := doc.SaveToFile("Q4_2025_Staff_Report.docx"); err != nil {
        panic(err)
    }
}
```

**Corresponding Template (report_template.docx):**

*Note: This template uses Sprig functions. Register them with `doc.RegisterFuncMap(sprig.FuncMap())`*

```
{{.Title | upper}}
Generated: {{.GeneratedAt | date "January 2, 2006"}}
Author: {{.Author}}

{{if .IsFinalized}}✓ FINALIZED{{else}}DRAFT{{end}}

EMPLOYEE ROSTER
{{range .Employees}}
━━━━━━━━━━━━━━━━━━━━
Name: {{.Name}}
Department: {{.Department}}
Salary: ${{.Salary | printf "%.2f"}}
{{.Photo}}
{{end}}

Total Budget: ${{.TotalBudget | printf "%.2f"}}
Employee Count: {{len .Employees}}
```

---

Example template files are also available in the [test/testdata/templates](https://github.com/abdokhaire/go-docxgen/tree/main/test/testdata/templates) directory.

## Acknowledgements

- [go-docx](https://github.com/fumiama/go-docx) by [fumiama](https://github.com/fumiama) - XML parsing and DOCX structure handling
- [go-docx-template](https://github.com/tomwatkins1994/go-docx-template) by [tomwatkins1994](https://github.com/tomwatkins1994) - Template merging approach and tag fragmentation handling
- [go-docx](https://github.com/lukasjarosch/go-docx) by [lukasjarosch](https://github.com/lukasjarosch) - Document manipulation concepts
- [python-docx-template](https://github.com/elapouya/python-docx-template) by [elapouya](https://github.com/elapouya) - Inspiration for the template approach

## License

MIT License - see [LICENSE](LICENSE) for details.
