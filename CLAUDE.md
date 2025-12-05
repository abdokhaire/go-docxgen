# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

go-docxgen is a Go library for merging DOCX files with data using Go's `text/template` syntax. It wraps the [go-docx](https://github.com/fumiama/go-docx) library (by fumiama), adding template rendering capabilities. Licensed under MIT.

## Common Commands

```bash
# Run all tests
make test

# Run a single test
go test -run TestName ./...

# Run tests in a specific package
go test ./internal/tags/...

# Run tests with coverage (opens HTML report in browser)
make test-coverage

# Run benchmarks
make benchmark
```

## The Fragmentation Problem

DOCX files are ZIP archives containing XML. WordprocessingML often fragments text across multiple XML `<w:r>` (run) elements based on editing history, spell-check, or formatting. A placeholder like `{{.FirstName}}` may appear in the XML as:

```xml
<w:r><w:t>{{.First</w:t></w:r>
<w:r><w:t>Name}}</w:t></w:r>
```

Simple string replacement fails here. This library solves this by **merging fragmented tags** before template processing.

## Architecture

### Core API (`docxtpl.go`)
- `DocxTmpl` struct embeds `*docx.Docx` from go-docx and adds template functionality
- `Parse`/`ParseFromFilename`/`ParseFromBytes` - Parse DOCX files into memory
- `Render(data)` - Replace template placeholders with provided data (struct or map)
- `Save(writer)`/`SaveToFile(filename)` - Write rendered document to output
- `GetPlaceholders()` - Returns all unique placeholders found in the document
- `GetWatermarks()` - Returns all watermark texts from document headers
- `ReplaceWatermark(old, new)` - Replace watermark text before rendering

### Rendering Pipeline
1. **Tag Merging** (`internal/tags/tag_merging.go`): Scans paragraphs and tables for incomplete tags (unmatched `{{` or `}}`), accumulates text across runs until complete, then writes merged text back. Supports whitespace-trimming syntax (`{{- ... -}}`)
2. **Data Processing** (`internal/templatedata/`): Converts structs to maps, handles pointers, nested structs, slices of any type (`[]string`, `[]int`, `[]*Struct`), maps with various key types, XML-escapes string values, detects image file paths
3. **Tag Replacement** (`internal/tags/tag_replacement.go`): Uses Go's `text/template` to execute template against prepared XML
4. **Processable Files** (`internal/headerfooter/`): Extracts and processes headers (`word/header*.xml`), footers (`word/footer*.xml`), footnotes (`word/footnotes.xml`), endnotes (`word/endnotes.xml`), and document properties (`docProps/core.xml`, `docProps/app.xml`). Also handles watermark extraction/replacement via VML `<v:textpath string="...">` elements. Watermarks support Go template syntax (e.g., `{{.Status}}`)
5. **XML Utilities** (`internal/xmlutils/`): Pre-processes XML for template compatibility, `MergeFragmentedTagsInXml` handles fragmented tags in raw XML strings, converts newlines to `<w:br/>` line breaks, fixes issues post-replacement

### Supported Data Types
The library supports all common Go data types in templates:
- **Primitives**: `string`, `int`, `int64`, `float64`, `bool`, etc.
- **Structs**: Regular structs and nested structs
- **Pointers**: `*string`, `*Struct` - automatically dereferenced
- **Slices**: `[]string`, `[]int`, `[]Struct`, `[]*Struct`
- **Maps**: `map[string]any`, `map[string]string`, etc.
- **Nil values**: Nil pointers are handled gracefully

### Inline Images (`inline_image.go`)
- `CreateInlineImage(filepath)` - Load image from file path
- String values that are valid image paths are auto-converted to inline images
- `InlineImage.Resize(width, height)` - Resize before rendering
- Supports JPEG (.jpg, .jpeg) and PNG formats
- Reads EXIF data for proper DPI/sizing (defaults to 72 DPI)

### Custom Functions (`functions.go`)
- Register via `doc.RegisterFunction(name, fn)` - validates function signature

**Built-in Functions:**

*Text Formatting:*
| Function | Usage | Description |
|----------|-------|-------------|
| `upper` | `{{.Name \| upper}}` | Uppercase text |
| `lower` | `{{.Name \| lower}}` | Lowercase text |
| `title` | `{{.Name \| title}}` | Title case text |
| `bold` | `{{bold .Name}}` | Bold formatting |
| `italic` | `{{italic .Name}}` | Italic formatting |
| `underline` | `{{underline .Name}}` | Underline formatting |
| `strikethrough` | `{{strikethrough .Name}}` | Strikethrough formatting |
| `color` | `{{color "FF0000" .Name}}` | Colored text (hex) |
| `highlight` | `{{highlight "yellow" .Name}}` | Highlighted text |
| `link` | `{{link "https://..." "Click"}}` | Hyperlink (basic) |
| `br` | `{{br}}` | Line break |
| `tab` | `{{tab}}` | Tab character |

*Comparison:*
| Function | Usage | Description |
|----------|-------|-------------|
| `eq` | `{{if eq .Status "active"}}` | Equal |
| `ne` | `{{if ne .Status "deleted"}}` | Not equal |
| `lt` | `{{if lt .Count 10}}` | Less than |
| `le` | `{{if le .Count 10}}` | Less than or equal |
| `gt` | `{{if gt .Count 0}}` | Greater than |
| `ge` | `{{if ge .Count 1}}` | Greater than or equal |

*Logical:*
| Function | Usage | Description |
|----------|-------|-------------|
| `and` | `{{if and .A .B}}` | All args truthy |
| `or` | `{{if or .A .B}}` | Any arg truthy |
| `not` | `{{if not .Deleted}}` | Negation |

*Collections:*
| Function | Usage | Description |
|----------|-------|-------------|
| `len` | `{{len .Items}}` | Length of slice/map/string |
| `first` | `{{first .Items}}` | First element |
| `last` | `{{last .Items}}` | Last element |
| `index` | `{{index .Items 0}}` | Element at index |
| `slice` | `{{slice .Items 1 3}}` | Sub-slice |
| `join` | `{{join .Names ", "}}` | Join with separator |
| `contains` | `{{if contains .Roles "admin"}}` | Check membership |

*Utilities:*
| Function | Usage | Description |
|----------|-------|-------------|
| `default` | `{{default "N/A" .Name}}` | Default if empty |
| `coalesce` | `{{coalesce .Nick .Name "Anon"}}` | First non-empty |
| `ternary` | `{{ternary "Yes" "No" .Active}}` | Conditional value |
| `split` | `{{range split .Tags ","}}` | Split string |
| `concat` | `{{concat .First " " .Last}}` | Concatenate |
| `trim` | `{{trim .Text}}` | Trim whitespace |
| `replace` | `{{replace .Text "old" "new"}}` | Replace all |
| `repeat` | `{{repeat "-" 10}}` | Repeat string |

*Math:*
| Function | Usage | Description |
|----------|-------|-------------|
| `add` | `{{add .A .B}}` | Addition |
| `sub` | `{{sub .A .B}}` | Subtraction |
| `mul` | `{{mul .Price .Qty}}` | Multiplication |
| `div` | `{{div .Total .Count}}` | Division |
| `mod` | `{{mod .Index 2}}` | Modulo |

*Number Formatting:*
| Function | Usage | Description |
|----------|-------|-------------|
| `formatNumber` | `{{formatNumber 1234.5 2}}` | Format with commas (→ "1,234.50") |
| `formatMoney` | `{{formatMoney 1234.5 "$"}}` | Currency format (→ "$1,234.50") |
| `formatPercent` | `{{formatPercent 0.156 1}}` | Percentage (→ "15.6%") |

*Date/Time:*
| Function | Usage | Description |
|----------|-------|-------------|
| `now` | `{{now}}` | Current time |
| `formatDate` | `{{formatDate .Date "Jan 2, 2006"}}` | Format date |
| `parseDate` | `{{parseDate "2024-01-15" "2006-01-02"}}` | Parse string to date |
| `addDays` | `{{addDays .Date 7}}` | Add days to date |
| `addMonths` | `{{addMonths .Date 1}}` | Add months to date |
| `addYears` | `{{addYears .Date 1}}` | Add years to date |

*Document Structure:*
| Function | Usage | Description |
|----------|-------|-------------|
| `pageBreak` | `{{pageBreak}}` | Insert page break |
| `sectionBreak` | `{{sectionBreak}}` | Insert section break |
| `link` | `{{link "https://..." "Click"}}` | Clickable hyperlink |

*Additional Utilities:*
| Function | Usage | Description |
|----------|-------|-------------|
| `uuid` | `{{uuid}}` | Generate UUID |
| `pluralize` | `{{pluralize 5 "item" "items"}}` | Singular/plural |
| `truncate` | `{{truncate .Text 50}}` | Truncate with ellipsis |
| `wordwrap` | `{{wordwrap .Text 80}}` | Wrap at width |
| `capitalize` | `{{capitalize .name}}` | Capitalize first letter |
| `camelCase` | `{{camelCase "hello world"}}` | → "helloWorld" |
| `snakeCase` | `{{snakeCase "Hello World"}}` | → "hello_world" |
| `kebabCase` | `{{kebabCase "Hello World"}}` | → "hello-world" |

### Internal Packages
- `internal/contenttypes/` - Manages `[Content_Types].xml` for added media
- `internal/functions/` - Function name/signature validation, default FuncMap (60+ functions)
- `internal/headerfooter/` - Processable file extraction (headers, footers, footnotes, endnotes), watermark handling
- `internal/hyperlinks/` - Hyperlink registry and relationship management for clickable links
- `internal/tags/` - Tag detection (`tag_checking.go`), merging, and replacement
- `internal/templatedata/` - Struct-to-map conversion, image path detection
- `internal/xmlutils/` - XML escaping (including newline→line break), `MergeFragmentedTagsInXml`, manipulation helpers

## Template Syntax

Uses standard Go `text/template` syntax in DOCX files:
- Simple: `{{.FieldName}}`
- Nested: `{{.Person.Name}}`
- Functions: `{{.Name | upper}}`
- Conditionals: `{{if .Show}}...{{end}}`
- Loops: `{{range .Items}}...{{end}}`

## Test Templates

Example templates in `test_templates/`:
- `test_basic.docx` - Simple field replacement
- `test_basic_with_images.docx` - Image insertion
- `test_with_tables.docx` - Table iteration with `{{range}}`
- `test_with_custom_functions.docx` - Custom function usage

Generated outputs are prefixed with `generated_`.
