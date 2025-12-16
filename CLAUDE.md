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

### Document Builder API (`builder.go`)
Create documents programmatically without templates:
- `New()` / `NewWithOptions(PageSize)` - Create new empty document (A4, A3, Letter, Legal)
- `AddParagraph(text)` - Add paragraph, returns `*Paragraph` for formatting
- `AddEmptyParagraph()` - Add empty paragraph for spacing
- `AddHeading(text, level)` - Add heading (level 0-9)
- `AddPageBreak()` - Insert page break
- `AddTable(rows, cols)` - Create table, returns `*Table`
- `AddTableWithWidths(rows, colWidths)` - Table with custom column widths (twips)
- `AddTableWithBorders(rows, cols, colors)` - Table with custom border colors

### Paragraph Formatting (`formatting.go`)
Fluent API for paragraph styling:
- `Bold()`, `Italic()`, `Underline()`, `Strike()` - Text formatting
- `Color(hex)`, `Highlight(color)`, `Background(hex)` - Colors
- `Size(halfPoints)`, `SizePoints(points)`, `Font(name)` - Typography
- `Center()`, `Left()`, `Right()`, `Justified()` - Alignment
- `Justify(Justification)` - Custom alignment
- `Style(styleID)` - Apply paragraph style (Heading1, ListBullet, etc.)
- `Bullet()`, `Numbered()` - List formatting
- `AddText(text)` - Add more text, returns `*Run`
- **Spacing**: `SpacingBefore(pts)`, `LineSpacing(twips)`, `LineSpacingSingle()`, `LineSpacingOneAndHalf()`, `LineSpacingDouble()`
- **Indentation**: `IndentLeft(inches)`, `IndentRight(inches)`, `IndentFirstLine(inches)`, `IndentHanging(inches)`, `Indent(IndentOptions)`
- **Tab Stops**: `AddTabStop(pos, align, leader)`, `AddTabStops([]TabStop)`, `ClearTabStops()`
- **Hyperlinks**: `AddLink(text, url)` - returns `*Hyperlink`
- **Images**: `AddAnchorImage(data)`, `AddAnchorImageFromFile(path)`, `AddInlineImage(data)`, `AddInlineImageFromFile(path)`
- **Shapes**: `AddAnchorShape(ShapeOptions)`, `AddInlineShape(ShapeOptions)` - rectangles, ellipses, arrows, stars, etc.

### Run Formatting (`formatting.go`)
Format individual text runs within paragraphs:
- `Bold()`, `Italic()`, `Underline(style)`, `Strike()`, `DoubleStrike()` - Basic formatting
- `Color(hex)`, `Highlight(color)`, `Background(hex)` - Colors
- `Size(halfPoints)`, `SizePoints(points)`, `Font(name)` - Typography
- `Shade(pattern, color, fill)` - Background shading
- `Superscript()`, `Subscript()` - Vertical alignment
- `CharacterSpacing(twips)`, `Expand(pts)`, `Condense(pts)` - Character spacing
- `Kern(halfPoints)` - Kerning threshold
- `Then()` - Return to parent paragraph for continued building

### Table Builder (`table.go`)
Programmatic table creation and formatting:
- `Cell(row, col)` - Get cell at position, returns `*TableCell`
- `SetCell(row, col, text)` - Set cell text, returns `*TableCell`
- `Row(index)` - Get row, returns `*TableRow`
- `Rows()`, `Cols()` - Get dimensions
- `Justify(alignment)`, `Center()` - Table alignment
- `SetBorderColors(TableBorderColors)` - Apply border colors to all cells
- **TableCell**:
  - Content: `SetText()`, `AddParagraph()`
  - Formatting: `Bold()`, `Center()`, `Background(hex)`, `Shade()`
  - Borders: `Borders(color, width)`, `NoBorders()`
  - Width: `Width(twips)`, `WidthInches(in)`, `WidthCm(cm)`, `WidthPercent(pct)`
  - Vertical Align: `VAlign(align)`, `VAlignTop()`, `VAlignCenter()`, `VAlignBottom()`
  - Merging: `MergeHorizontal(count)`, `MergeVerticalStart()`, `MergeVerticalContinue()`
- **TableRow**: `Cell(col)`, `SetCell()`, `Justify()`

### Document Properties (`properties.go`)
Get and set document metadata:
- `GetProperties()` - Returns `*DocumentProperties`
- `SetProperties(props)` - Update all properties
- `SetTitle(title)`, `SetAuthor(author)`, `SetSubject(subject)` - Convenience setters
- `SetKeywords(keywords)`, `SetDescription(desc)`, `SetCategory(cat)`
- `SetContentStatus(status)` - e.g., "Draft", "Final"
- **DocumentProperties**: Title, Subject, Creator, Keywords, Description, LastModifiedBy, Revision, Created, Modified, Category, ContentStatus

### Rendering Pipeline
1. **Tag Merging** (`internal/tags/tag_merging.go`): Scans paragraphs and tables for incomplete tags (unmatched `{{` or `}}`), accumulates text across runs until complete, then writes merged text back. Supports whitespace-trimming syntax (`{{- ... -}}`)
2. **Data Processing** (`internal/templatedata/`): Converts structs to maps, handles pointers, nested structs, slices of any type (`[]string`, `[]int`, `[]*Struct`), maps with various key types, XML-escapes string values, detects image file paths
3. **Tag Replacement** (`internal/tags/tag_replacement.go`): Uses Go's `text/template` to execute template against prepared XML
4. **Processable Files** (`internal/headerfooter/`): Extracts and processes headers (`word/header*.xml`), footers (`word/footer*.xml`), footnotes (`word/footnotes.xml`), endnotes (`word/endnotes.xml`), and document properties (`docProps/core.xml`, `docProps/app.xml`). Also handles watermark extraction/replacement via VML `<v:textpath string="...">` elements. Watermarks support Go template syntax (e.g., `{{.Status}}`)
5. **XML Utilities** (`internal/xmlutils/`): Pre-processes XML for template compatibility, `MergeFragmentedTagsInXml` handles fragmented tags in raw XML strings, converts newlines to `<w:br/>` line breaks, fixes issues post-replacement (including replacing `<no value>` and `<nil>` template outputs with empty strings)

### Supported Data Types
The library supports all common Go data types in templates:
- **Primitives**: `string`, `int`, `int64`, `float64`, `bool`, etc.
- **Structs**: Regular structs and nested structs
- **Pointers**: `*string`, `*Struct` - automatically dereferenced
- **Slices**: `[]string`, `[]int`, `[]Struct`, `[]*Struct`
- **Maps**: `map[string]any`, `map[string]string`, etc.
- **Nil values**: Nil pointers and nil interface values are handled gracefully (output empty string instead of `<nil>` or `<no value>`)

### Inline Images (`image.go`)
- `CreateInlineImage(filepath)` - Load image from file path
- String values that are valid image paths are auto-converted to inline images
- `InlineImage.Resize(width, height)` - Resize before rendering
- Supports JPEG (.jpg, .jpeg) and PNG formats
- Reads EXIF data for proper DPI/sizing (defaults to 72 DPI)

### Template Functions (`functions.go`)
- Register via `doc.RegisterFunction(name, fn)` - validates function signature
- Register function libraries via `doc.RegisterFuncMap(funcMap)` - for Sprig/Sprout integration

**Built-in Function:**
| Function | Usage | Description |
|----------|-------|-------------|
| `link` | `{{link "https://..." "Click"}}` | Create clickable hyperlink |

**Go Template Built-ins** (always available):
| Function | Usage | Description |
|----------|-------|-------------|
| `and`, `or`, `not` | `{{if and .A .B}}` | Logical operations |
| `eq`, `ne`, `lt`, `le`, `gt`, `ge` | `{{if eq .Status "active"}}` | Comparisons |
| `len` | `{{len .Items}}` | Length of slice/map/string |
| `index` | `{{index .Items 0}}` | Element at index |
| `slice` | `{{slice .Items 1 3}}` | Sub-slice |
| `print`, `printf`, `println` | `{{printf "%.2f" .Price}}` | Formatted output |

**Community Function Libraries:**
- [Sprig](https://github.com/Masterminds/sprig) - 100+ functions: `doc.RegisterFuncMap(sprig.FuncMap())`
- [Sprout](https://github.com/go-sprout/sprout) - Modern alternative: `doc.RegisterFuncMap(sprout.New().Build())`

Functions like `upper`, `lower`, `title`, `default`, `ternary`, `join`, `add`, `mul`, `date`, etc. require registering Sprig or Sprout.

### Document Operations (`operations.go`)
Batch processing and document manipulation:
- **Cloning**: `Clone()` - Create independent copy of document
- **Merging**: `MergeDocuments(docs...)`, `AppendDocument(doc)` - Combine documents
- **Mail Merge**: `MailMerge(records)` - Generate multiple documents from data
- **Batch Processing**: `BatchRender(records)`, `ProcessDirectory(input, output, data)`
- **Text Operations**: `ReplaceText(old, new)`, `ReplaceTextInHeaders/Footers()`
- **Lists**: `CreateBulletList(items)`, `CreateNumberedList(items)`, `AddListItem()`

### Document Analysis (`analysis.go`)
Advanced document inspection and manipulation:
- **Comparison**: `Compare(other)` - Diff two documents, returns `DocumentDiff`
- **Bookmarks**: `GetBookmarks()`, `AddBookmark()`, `GetInternalLinks()`, `GetTableOfContents()`
- **Comments**: `GetComments()`, `AddComment()`, `DeleteComment()`, `ReplyToComment()`
- **Tracked Changes**: `GetTrackedChanges()`, `AcceptChange()`, `RejectChange()`, `AcceptAllChanges()`
- **Protection**: `GetProtectionInfo()`, `Protect()`, `Unprotect()`, `IsProtected()`
- **Export**: `ToJSON()`, `ToMarkdown()`, `ToHTML()`, `ToStructuredDocument()`
- **XML Access**: `UnpackToDirectory()`, `PackFromDirectory()`, `GetXMLContent()`, `SetXMLContent()`

### Validation (`validation.go`)
Template validation and error handling:
- **Validation**: `Validate()` - Check template syntax, returns `[]ValidationError`
- **Data Validation**: `ValidateData(data)` - Check data against template requirements
- **Field Discovery**: `GetRequiredFields()` - Extract required field names
- **Error Types**: `TemplateError`, `ErrorSummary`, `ValidationError`
- **Error Constructors**: `ErrSyntax()`, `ErrUnclosedTag()`, `ErrUndefinedField()`, etc.

### Internal Packages
- `internal/contenttypes/` - Manages `[Content_Types].xml` for added media
- `internal/functions/` - Function name/signature validation, empty DefaultFuncMap (users register functions via Sprig/Sprout)
- `internal/headerfooter/` - Processable file extraction (headers, footers, footnotes, endnotes), watermark handling
- `internal/hyperlinks/` - Hyperlink registry and relationship management for clickable links
- `internal/tags/` - Tag detection (`tag_checking.go`), merging, and replacement
- `internal/templatedata/` - Struct-to-map conversion, image path detection
- `internal/xmlutils/` - XML escaping (including newline→line break), `MergeFragmentedTagsInXml`, manipulation helpers, post-processing to replace `<no value>` and `<nil>` with empty strings

## Template Syntax

Uses standard Go `text/template` syntax in DOCX files:
- Simple: `{{.FieldName}}`
- Nested: `{{.Person.Name}}`
- Functions: `{{.Name | upper}}` (requires Sprig/Sprout registration)
- Built-in functions: `{{printf "%.2f" .Price}}`, `{{len .Items}}`, `{{if eq .Status "active"}}`
- Conditionals: `{{if .Show}}...{{end}}`
- Loops: `{{range .Items}}...{{end}}`

## Test Templates

Example templates in `test/testdata/templates/`:
- `test_basic.docx` - Simple field replacement
- `test_basic_with_images.docx` - Image insertion
- `test_with_tables.docx` - Table iteration with `{{range}}`
- `test_with_custom_functions.docx` - Custom function usage

Generated outputs are prefixed with `generated_`.
