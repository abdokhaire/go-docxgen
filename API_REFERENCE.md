# go-docxgen API Reference

Complete API reference for the go-docxgen library.

## Table of Contents

- [Document Parsing](#document-parsing)
- [Document Rendering](#document-rendering)
- [Document Saving](#document-saving)
- [Document Creation](#document-creation)
- [Document Operations](#document-operations)
- [Paragraph API](#paragraph-api)
- [Run API](#run-api)
- [Table API](#table-api)
- [Document Properties](#document-properties)
- [Inline Images](#inline-images)
- [Custom Functions](#custom-functions)
- [Template Functions](#template-functions)

---

## Document Parsing

### Parse
```go
func Parse(reader io.ReaderAt, size int64) (*DocxTmpl, error)
```
Parse a DOCX document from a reader.

**Example:**
```go
file, _ := os.Open("template.docx")
info, _ := file.Stat()
doc, err := docxtpl.Parse(file, info.Size())
```

### ParseFromFilename
```go
func ParseFromFilename(filename string) (*DocxTmpl, error)
```
Parse a DOCX document from a file path.

**Example:**
```go
doc, err := docxtpl.ParseFromFilename("template.docx")
```

### ParseFromBytes
```go
func ParseFromBytes(data []byte) (*DocxTmpl, error)
```
Parse a DOCX document from a byte slice.

**Example:**
```go
data, _ := os.ReadFile("template.docx")
doc, err := docxtpl.ParseFromBytes(data)
```

---

## Document Rendering

### Render
```go
func (d *DocxTmpl) Render(data any) error
```
Replace template placeholders with provided data. Accepts structs or maps.

**Example:**
```go
data := map[string]any{
    "Name": "John Doe",
    "Date": time.Now(),
}
err := doc.Render(data)
```

### GetPlaceholders
```go
func (d *DocxTmpl) GetPlaceholders() ([]string, error)
```
Returns all unique placeholders found in the document.

**Example:**
```go
placeholders, err := doc.GetPlaceholders()
// Returns: []string{"{{.Name}}", "{{.Date}}", "{{range .Items}}"}
```

### GetWatermarks
```go
func (d *DocxTmpl) GetWatermarks() []string
```
Returns all watermark texts from document headers.

### ReplaceWatermark
```go
func (d *DocxTmpl) ReplaceWatermark(oldText, newText string)
```
Replace watermark text before rendering.

**Example:**
```go
doc.ReplaceWatermark("DRAFT", "FINAL")
```

---

## Document Saving

### Save
```go
func (d *DocxTmpl) Save(writer io.Writer) error
```
Save the document to a writer.

**Example:**
```go
file, _ := os.Create("output.docx")
err := doc.Save(file)
file.Close()
```

### SaveToFile
```go
func (d *DocxTmpl) SaveToFile(filename string) error
```
Save the document directly to a file.

**Example:**
```go
err := doc.SaveToFile("output.docx")
```

---

## Document Creation

### New
```go
func New() *DocxTmpl
```
Create a new empty document.

**Example:**
```go
doc := docxtpl.New()
doc.AddHeading("Hello World", 1)
doc.SaveToFile("output.docx")
```

### NewWithOptions
```go
func NewWithOptions(pageSize ...PageSize) *DocxTmpl
```
Create a new document with optional page size.

**Page Sizes:**
- `PageSizeA4` - A4 paper (default)
- `PageSizeA3` - A3 paper
- `PageSizeLetter` - US Letter
- `PageSizeLegal` - US Legal

**Example:**
```go
doc := docxtpl.NewWithOptions(docxtpl.PageSizeA3)
```

### AddParagraph
```go
func (d *DocxTmpl) AddParagraph(text string) *Paragraph
```
Add a paragraph with text. Returns a Paragraph for formatting.

**Example:**
```go
para := doc.AddParagraph("Hello, World!")
para.Bold().Color("FF0000")
```

### AddEmptyParagraph
```go
func (d *DocxTmpl) AddEmptyParagraph() *Paragraph
```
Add an empty paragraph for spacing.

### AddHeading
```go
func (d *DocxTmpl) AddHeading(text string, level int) *Paragraph
```
Add a heading (level 0-9). Level 0 is title, 1-9 are headings.

**Example:**
```go
doc.AddHeading("Document Title", 0)
doc.AddHeading("Chapter 1", 1)
doc.AddHeading("Section 1.1", 2)
```

### AddPageBreak
```go
func (d *DocxTmpl) AddPageBreak()
```
Insert a page break.

### AddTable
```go
func (d *DocxTmpl) AddTable(rows, cols int) *Table
```
Create a table with specified dimensions.

**Example:**
```go
table := doc.AddTable(3, 4) // 3 rows, 4 columns
table.SetCell(0, 0, "Header")
```

### AddTableWithWidths
```go
func (d *DocxTmpl) AddTableWithWidths(rows int, colWidths []int) *Table
```
Create a table with custom column widths in twips (1 inch = 1440 twips).

**Example:**
```go
table := doc.AddTableWithWidths(3, []int{2880, 1440, 1440}) // 2", 1", 1"
```

### AddTableWithBorders
```go
func (d *DocxTmpl) AddTableWithBorders(rows, cols int, colors TableBorderColors) *Table
```
Create a table with custom border colors.

**Example:**
```go
table := doc.AddTableWithBorders(3, 3, docxtpl.TableBorderColors{
    Top: "FF0000", Bottom: "FF0000",
    Left: "0000FF", Right: "0000FF",
    InsideH: "00FF00", InsideV: "00FF00",
})
```

---

## Document Operations

### Document Merging

#### AppendDocument
```go
func (d *DocxTmpl) AppendDocument(other *DocxTmpl)
```
Append all contents from another document.

**Example:**
```go
doc1, _ := docxtpl.ParseFromFilename("chapter1.docx")
doc2, _ := docxtpl.ParseFromFilename("chapter2.docx")
doc1.AppendDocument(doc2)
doc1.SaveToFile("combined.docx")
```

### Document Splitting

#### SplitAt
```go
func (d *DocxTmpl) SplitAt(rule SplitRule) []*DocxTmpl
```
Split document based on a custom rule.

**Example:**
```go
docs := doc.SplitAt(func(text string, isHeading bool, level int) bool {
    return isHeading && level == 1
})
```

#### SplitAtHeading
```go
func (d *DocxTmpl) SplitAtHeading(level int) []*DocxTmpl
```
Split at each heading of specified level.

**Example:**
```go
chapters := doc.SplitAtHeading(1) // Split at Heading 1
```

#### SplitAtText
```go
func (d *DocxTmpl) SplitAtText(text string) []*DocxTmpl
```
Split at paragraphs containing specified text.

#### SplitAtRegex
```go
func (d *DocxTmpl) SplitAtRegex(pattern string) []*DocxTmpl
```
Split at paragraphs matching regex pattern.

### Text Extraction

#### GetText
```go
func (d *DocxTmpl) GetText() string
```
Extract all plain text from document.

**Example:**
```go
text := doc.GetText()
fmt.Println(text)
```

#### GetParagraphTexts
```go
func (d *DocxTmpl) GetParagraphTexts() []string
```
Get text of each paragraph as a slice.

### Text Search

#### HasText
```go
func (d *DocxTmpl) HasText(text string) bool
```
Check if document contains text.

#### HasTextMatch
```go
func (d *DocxTmpl) HasTextMatch(pattern string) bool
```
Check if document contains text matching regex.

#### FindText
```go
func (d *DocxTmpl) FindText(text string) []string
```
Find all paragraphs containing text.

#### FindTextMatch
```go
func (d *DocxTmpl) FindTextMatch(pattern string) []string
```
Find all paragraphs matching regex.

### Text Replacement

#### ReplaceText
```go
func (d *DocxTmpl) ReplaceText(oldText, newText string)
```
Replace all occurrences of text.

**Example:**
```go
doc.ReplaceText("COMPANY", "Acme Corp")
```

#### ReplaceTextRegex
```go
func (d *DocxTmpl) ReplaceTextRegex(pattern, replacement string)
```
Replace text matching regex pattern.

**Example:**
```go
doc.ReplaceTextRegex(`\d{3}-\d{2}-\d{4}`, "[REDACTED]")
```

### Media Access

#### GetMedia
```go
func (d *DocxTmpl) GetMedia(name string) *Media
```
Get embedded media file by name.

**Example:**
```go
media := doc.GetMedia("image1.png")
if media != nil {
    os.WriteFile("extracted.png", media.Data, 0644)
}
```

#### GetAllMedia
```go
func (d *DocxTmpl) GetAllMedia() []*Media
```
Get all embedded media files.

### Document Counting

| Method | Description |
|--------|-------------|
| `CountParagraphs() int` | Count paragraphs |
| `CountTables() int` | Count tables |

### Element Cleanup

| Method | Description |
|--------|-------------|
| `DropAllDrawings()` | Remove all shapes, canvases, groups |
| `DropShapes()` | Remove all shapes |
| `DropCanvas()` | Remove all canvas elements |
| `DropGroups()` | Remove all group elements |
| `DropEmptyPictures()` | Remove nil picture references |
| `KeepBodyElements(names...)` | Keep only specified element types |

### Document Optimization

#### MergeAllRuns
```go
func (d *DocxTmpl) MergeAllRuns()
```
Merge runs with same formatting in all paragraphs.

#### CleanDocument
```go
func (d *DocxTmpl) CleanDocument()
```
Perform common cleanup (merge runs, remove empty pictures).

---

## Paragraph API

### Text Formatting

| Method | Description |
|--------|-------------|
| `Bold()` | Apply bold formatting |
| `Italic()` | Apply italic formatting |
| `Underline()` | Apply underline formatting |
| `Strike()` | Apply strikethrough |
| `Color(hex string)` | Set text color (e.g., "FF0000") |
| `Highlight(color string)` | Apply highlight (yellow, green, cyan, etc.) |
| `Background(hex string)` | Set background color |
| `Size(halfPoints int)` | Set font size in half-points (24 = 12pt) |
| `SizePoints(points int)` | Set font size in points |
| `Font(name string)` | Set font family |
| `Shade(pattern, color, fill string)` | Apply shading pattern |

### Alignment

| Method | Description |
|--------|-------------|
| `Left()` | Left align |
| `Center()` | Center align |
| `Right()` | Right align |
| `Justified()` | Justify text |
| `Justify(j Justification)` | Custom alignment |

**Justification Constants:**
- `JustifyLeft`
- `JustifyCenter`
- `JustifyRight`
- `JustifyBoth`
- `JustifyDistribute`

### Spacing

| Method | Description |
|--------|-------------|
| `Spacing(opts SpacingOptions)` | Set all spacing options |
| `SpacingBefore(points int)` | Space before paragraph in points |
| `SpacingAfter(points int)` | Space after paragraph in points |
| `LineSpacing(twips int)` | Line spacing in twips |
| `LineSpacingSingle()` | Single line spacing |
| `LineSpacingOneAndHalf()` | 1.5 line spacing |
| `LineSpacingDouble()` | Double line spacing |

**SpacingOptions:**
```go
type SpacingOptions struct {
    Before      int    // Space before (twips)
    After       int    // Space after (twips)
    Line        int    // Line spacing (twips)
    LineRule    string // "auto", "exact", "atLeast"
    BeforeLines int    // Space before in lines
}
```

### Indentation

| Method | Description |
|--------|-------------|
| `Indent(opts IndentOptions)` | Set all indent options |
| `IndentLeft(inches float64)` | Left indent in inches |
| `IndentRight(inches float64)` | Right indent in inches |
| `IndentFirstLine(inches float64)` | First line indent |
| `IndentHanging(inches float64)` | Hanging indent |

**IndentOptions:**
```go
type IndentOptions struct {
    Left           int // Left indent (twips)
    Right          int // Right indent (twips)
    FirstLine      int // First line indent (twips)
    Hanging        int // Hanging indent (twips)
    LeftChars      int // Left indent (character units)
    FirstLineChars int // First line (character units)
}
```

### Tab Stops

| Method | Description |
|--------|-------------|
| `AddTabStop(position int, align, leader string)` | Add a tab stop |
| `AddTabStops(stops []TabStop)` | Add multiple tab stops |
| `ClearTabStops()` | Remove all tab stops |

**TabStop:**
```go
type TabStop struct {
    Position int    // Position in twips (1440 = 1 inch)
    Align    string // "left", "center", "right", "decimal"
    Leader   string // "none", "dot", "hyphen", "underscore"
}
```

### Lists

| Method | Description |
|--------|-------------|
| `Bullet()` | Make bullet point |
| `Numbered()` | Make numbered list item |
| `Style(styleID string)` | Apply paragraph style |

### Adding Content

| Method | Description |
|--------|-------------|
| `AddText(text string) *Run` | Add text, returns Run for formatting |
| `AddTab()` | Add tab character |
| `AddBreak()` | Add line break |
| `AddPageBreak()` | Add page break |
| `AddLink(text, url string) *Hyperlink` | Add hyperlink |

### Images

| Method | Description |
|--------|-------------|
| `AddAnchorImage(data []byte) (*Run, error)` | Add floating image from bytes |
| `AddAnchorImageFromFile(path string) (*Run, error)` | Add floating image from file |
| `AddInlineImage(data []byte) (*Run, error)` | Add inline image from bytes |
| `AddInlineImageFromFile(path string) (*Run, error)` | Add inline image from file |

### Shapes

| Method | Description |
|--------|-------------|
| `AddAnchorShape(opts ShapeOptions) *Run` | Add floating shape |
| `AddInlineShape(opts ShapeOptions) *Run` | Add inline shape |

**ShapeOptions:**
```go
type ShapeOptions struct {
    Width     int64       // Width in EMUs (914400 = 1 inch)
    Height    int64       // Height in EMUs
    Preset    ShapePreset // Shape type
    Name      string      // Shape name
    LineColor string      // Outline color (hex)
    LineWidth int64       // Outline width (EMUs)
    BWMode    string      // Black/white mode
}
```

**Shape Presets:**
- `ShapeRectangle`, `ShapeRoundedRect`, `ShapeEllipse`
- `ShapeTriangle`, `ShapeDiamond`, `ShapePentagon`, `ShapeHexagon`
- `ShapeArrowRight`, `ShapeArrowLeft`, `ShapeArrowUp`, `ShapeArrowDown`
- `ShapeStar5`, `ShapeStar6`, `ShapeHeart`, `ShapeLightningBolt`
- `ShapeSun`, `ShapeMoon`, `ShapeCloud`, `ShapeLine`

### Text Extraction

| Method | Description |
|--------|-------------|
| `GetText() string` | Get plain text content |

### Element Cleanup

| Method | Description |
|--------|-------------|
| `DropShapes()` | Remove all shapes |
| `DropCanvas()` | Remove all canvas elements |
| `DropGroups()` | Remove all group elements |
| `DropAllDrawings()` | Remove all shapes, canvases, groups |
| `DropEmptyPictures()` | Remove nil picture references |
| `KeepElements(names...)` | Keep only specified element types |

### Run Merging

| Method | Description |
|--------|-------------|
| `MergeRuns()` | Merge runs with same formatting |
| `MergeAllRuns()` | Merge all runs (may lose formatting) |

### Raw Access

| Method | Description |
|--------|-------------|
| `GetRaw() *docx.Paragraph` | Get underlying go-docx paragraph |

---

## Run API

A Run represents a contiguous piece of text with the same formatting.

### Text Formatting

| Method | Description |
|--------|-------------|
| `Bold()` | Apply bold |
| `Italic()` | Apply italic |
| `Underline(style ...string)` | Apply underline (single, double, thick, dotted, dash, wave) |
| `Strike()` | Apply strikethrough |
| `DoubleStrike()` | Apply double strikethrough |
| `Color(hex string)` | Set text color |
| `Highlight(color string)` | Apply highlight |
| `Background(hex string)` | Set background color |
| `Shade(pattern, color, fill string)` | Apply shading |

### Typography

| Method | Description |
|--------|-------------|
| `Size(halfPoints int)` | Font size in half-points |
| `SizePoints(points int)` | Font size in points |
| `Font(name string)` | Font family |

### Vertical Alignment

| Method | Description |
|--------|-------------|
| `Superscript()` | Superscript text |
| `Subscript()` | Subscript text |

### Character Spacing

| Method | Description |
|--------|-------------|
| `CharacterSpacing(twips int)` | Set character spacing |
| `Expand(points float64)` | Expand spacing |
| `Condense(points float64)` | Condense spacing |
| `Kern(halfPoints int)` | Set kerning threshold |

### Text & Navigation

| Method | Description |
|--------|-------------|
| `GetText() string` | Get plain text content of run |
| `AddTab()` | Add tab after run |
| `Then() *Paragraph` | Return to parent paragraph |
| `KeepElements(names...)` | Keep only specified elements |
| `GetRaw() *docx.Run` | Get underlying go-docx run |

---

## Table API

### Table Methods

| Method | Description |
|--------|-------------|
| `Cell(row, col int) *TableCell` | Get cell at position |
| `SetCell(row, col int, text string) *TableCell` | Set cell text |
| `Row(index int) *TableRow` | Get row |
| `Rows() int` | Get row count |
| `Cols() int` | Get column count |
| `Justify(alignment string)` | Set table alignment |
| `Center()` | Center table |
| `SetBorderColors(colors TableBorderColors)` | Set all border colors |
| `GetRaw() *docx.Table` | Get underlying table |

### TableCell Methods

#### Content
| Method | Description |
|--------|-------------|
| `SetText(text string)` | Set cell text |
| `AddParagraph(text string) *Paragraph` | Add paragraph to cell |

#### Formatting
| Method | Description |
|--------|-------------|
| `Bold()` | Bold cell text |
| `Center()` | Center cell text |
| `Background(hex string)` | Set background color |
| `Shade(pattern, color, fill string)` | Apply shading |

#### Borders
| Method | Description |
|--------|-------------|
| `Borders(color string, width int)` | Set all borders |
| `NoBorders()` | Remove borders |

#### Width
| Method | Description |
|--------|-------------|
| `Width(twips int)` | Set width in twips |
| `WidthInches(inches float64)` | Set width in inches |
| `WidthCm(cm float64)` | Set width in centimeters |
| `WidthPercent(percent int)` | Set width as percentage |

#### Vertical Alignment
| Method | Description |
|--------|-------------|
| `VAlign(align VerticalAlignment)` | Set vertical alignment |
| `VAlignTop()` | Align to top |
| `VAlignCenter()` | Align to center |
| `VAlignBottom()` | Align to bottom |

#### Cell Merging
| Method | Description |
|--------|-------------|
| `MergeHorizontal(count int)` | Merge with cells to the right |
| `MergeVerticalStart()` | Start vertical merge |
| `MergeVerticalContinue()` | Continue vertical merge |

### TableRow Methods

| Method | Description |
|--------|-------------|
| `Cell(col int) *TableCell` | Get cell in row |
| `SetCell(col int, text string) *TableCell` | Set cell text |
| `Justify(alignment string)` | Set row alignment |
| `GetRaw() *docx.WTableRow` | Get underlying row |

---

## Document Properties

### Get/Set Properties

```go
func (d *DocxTmpl) GetProperties() *DocumentProperties
func (d *DocxTmpl) SetProperties(props *DocumentProperties)
```

### Convenience Setters

| Method | Description |
|--------|-------------|
| `SetTitle(title string)` | Set document title |
| `SetAuthor(author string)` | Set author/creator |
| `SetSubject(subject string)` | Set subject |
| `SetKeywords(keywords string)` | Set keywords |
| `SetDescription(desc string)` | Set description |
| `SetCategory(category string)` | Set category |
| `SetContentStatus(status string)` | Set status (Draft, Final, etc.) |

### DocumentProperties Struct

```go
type DocumentProperties struct {
    Title          string
    Subject        string
    Creator        string    // Author
    Keywords       string
    Description    string
    LastModifiedBy string
    Revision       string
    Created        time.Time
    Modified       time.Time
    Category       string
    ContentStatus  string
}
```

---

## Inline Images

### CreateInlineImage
```go
func CreateInlineImage(filepath string) (*InlineImage, error)
```
Load an image from file for template rendering.

### InlineImage Methods

| Method | Description |
|--------|-------------|
| `Resize(width, height int)` | Resize image in pixels |

**Example:**
```go
img, _ := docxtpl.CreateInlineImage("logo.png")
img.Resize(200, 100)
data := map[string]any{"Logo": img}
doc.Render(data)
```

---

## Custom Functions

### RegisterFunction
```go
func (d *DocxTmpl) RegisterFunction(name string, fn any) error
```
Register a custom template function.

**Example:**
```go
doc.RegisterFunction("currency", func(amount float64) string {
    return fmt.Sprintf("$%.2f", amount)
})
```

---

## Template Functions

### Text Formatting

| Function | Usage | Description |
|----------|-------|-------------|
| `upper` | `{{.Name \| upper}}` | Uppercase |
| `lower` | `{{.Name \| lower}}` | Lowercase |
| `title` | `{{.Name \| title}}` | Title case |
| `bold` | `{{bold .Name}}` | Bold text |
| `italic` | `{{italic .Name}}` | Italic text |
| `underline` | `{{underline .Name}}` | Underlined text |
| `strikethrough` | `{{strikethrough .Name}}` | Strikethrough |
| `doubleStrike` | `{{doubleStrike .Name}}` | Double strikethrough |
| `color` | `{{color "FF0000" .Name}}` | Colored text |
| `highlight` | `{{highlight "yellow" .Name}}` | Highlighted text |
| `bgColor` | `{{bgColor "FFFF00" .Name}}` | Background color |
| `fontSize` | `{{fontSize 28 .Name}}` | Font size (half-points) |
| `fontFamily` | `{{fontFamily "Arial" .Name}}` | Font family |
| `font` | `{{font "Arial" 24 "FF0000" .Name}}` | Combined font settings |
| `subscript` | `{{subscript .Name}}` | Subscript |
| `superscript` | `{{superscript .Name}}` | Superscript |
| `smallCaps` | `{{smallCaps .Name}}` | Small capitals |
| `allCaps` | `{{allCaps .Name}}` | All capitals |
| `shadow` | `{{shadow .Name}}` | Shadow effect |
| `outline` | `{{outline .Name}}` | Outline effect |
| `emboss` | `{{emboss .Name}}` | Emboss effect |
| `imprint` | `{{imprint .Name}}` | Imprint effect |

### Comparison

| Function | Usage | Description |
|----------|-------|-------------|
| `eq` | `{{if eq .A .B}}` | Equal |
| `ne` | `{{if ne .A .B}}` | Not equal |
| `lt` | `{{if lt .A .B}}` | Less than |
| `le` | `{{if le .A .B}}` | Less than or equal |
| `gt` | `{{if gt .A .B}}` | Greater than |
| `ge` | `{{if ge .A .B}}` | Greater than or equal |

### Logical

| Function | Usage | Description |
|----------|-------|-------------|
| `and` | `{{if and .A .B}}` | All args truthy |
| `or` | `{{if or .A .B}}` | Any arg truthy |
| `not` | `{{if not .A}}` | Negation |

### Collections

| Function | Usage | Description |
|----------|-------|-------------|
| `len` | `{{len .Items}}` | Length |
| `first` | `{{first .Items}}` | First element |
| `last` | `{{last .Items}}` | Last element |
| `index` | `{{index .Items 0}}` | Element at index |
| `slice` | `{{slice .Items 1 3}}` | Sub-slice |
| `join` | `{{join .Names ", "}}` | Join with separator |
| `contains` | `{{if contains .Roles "admin"}}` | Check membership |

### Math

| Function | Usage | Description |
|----------|-------|-------------|
| `add` | `{{add .A .B}}` | Addition |
| `sub` | `{{sub .A .B}}` | Subtraction |
| `mul` | `{{mul .A .B}}` | Multiplication |
| `div` | `{{div .A .B}}` | Division |
| `mod` | `{{mod .A .B}}` | Modulo |

### Number Formatting

| Function | Usage | Description |
|----------|-------|-------------|
| `formatNumber` | `{{formatNumber 1234.5 2}}` | Format with commas |
| `formatMoney` | `{{formatMoney 1234.5 "$"}}` | Currency format |
| `formatPercent` | `{{formatPercent 0.156 1}}` | Percentage |

### Date/Time

| Function | Usage | Description |
|----------|-------|-------------|
| `now` | `{{now}}` | Current time |
| `formatDate` | `{{formatDate .Date "Jan 2, 2006"}}` | Format date |
| `parseDate` | `{{parseDate "2024-01-15" "2006-01-02"}}` | Parse date |
| `addDays` | `{{addDays .Date 7}}` | Add days |
| `addMonths` | `{{addMonths .Date 1}}` | Add months |
| `addYears` | `{{addYears .Date 1}}` | Add years |

### String Utilities

| Function | Usage | Description |
|----------|-------|-------------|
| `default` | `{{default "N/A" .Name}}` | Default if empty |
| `coalesce` | `{{coalesce .A .B "default"}}` | First non-empty |
| `ternary` | `{{ternary "Yes" "No" .Active}}` | Conditional |
| `split` | `{{split .Text ","}}` | Split string |
| `concat` | `{{concat .A " " .B}}` | Concatenate |
| `trim` | `{{trim .Text}}` | Trim whitespace |
| `replace` | `{{replace .Text "old" "new"}}` | Replace all |
| `repeat` | `{{repeat "-" 10}}` | Repeat string |
| `truncate` | `{{truncate .Text 50}}` | Truncate with ellipsis |
| `wordwrap` | `{{wordwrap .Text 80}}` | Wrap at width |
| `capitalize` | `{{capitalize .name}}` | Capitalize first |
| `camelCase` | `{{camelCase "hello world"}}` | camelCase |
| `snakeCase` | `{{snakeCase "Hello World"}}` | snake_case |
| `kebabCase` | `{{kebabCase "Hello World"}}` | kebab-case |

### Document Structure

| Function | Usage | Description |
|----------|-------|-------------|
| `br` | `{{br}}` | Line break |
| `tab` | `{{tab}}` | Tab character |
| `pageBreak` | `{{pageBreak}}` | Page break |
| `sectionBreak` | `{{sectionBreak}}` | Section break |
| `link` | `{{link "https://..." "Click"}}` | Hyperlink |

### Miscellaneous

| Function | Usage | Description |
|----------|-------|-------------|
| `uuid` | `{{uuid}}` | Generate UUID |
| `pluralize` | `{{pluralize 5 "item" "items"}}` | Singular/plural |

---

## Units Reference

| Unit | Value | Description |
|------|-------|-------------|
| Twip | 1/20 point | 1440 twips = 1 inch |
| Half-point | 1/2 point | 24 half-points = 12pt |
| EMU | English Metric Unit | 914400 EMUs = 1 inch |
| Point | 1/72 inch | Standard typography unit |

## Color Format

Colors are specified as 6-character hex codes without the `#` prefix:
- `"FF0000"` - Red
- `"00FF00"` - Green
- `"0000FF"` - Blue
- `"000000"` - Black
- `"FFFFFF"` - White

## Highlight Colors

Valid highlight colors: `yellow`, `green`, `cyan`, `magenta`, `blue`, `red`, `darkBlue`, `darkCyan`, `darkGreen`, `darkMagenta`, `darkRed`, `darkYellow`, `darkGray`, `lightGray`, `black`
