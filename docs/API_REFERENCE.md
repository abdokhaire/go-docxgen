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
- [Document Metadata](#document-metadata)
- [List Operations](#list-operations)
- [Enhanced Table Operations](#enhanced-table-operations)
- [Export/Conversion Functions](#exportconversion-functions)
- [Section Operations](#section-operations)
- [Template Validation](#template-validation)
- [Document Comparison](#document-comparison)
- [Batch Operations](#batch-operations)
- [Tracked Changes](#tracked-changes)
- [Comments](#comments)
- [Raw XML Access](#raw-xml-access)
- [Bookmarks](#bookmarks)
- [Document Protection](#document-protection)

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

## Template Functions

### Built-in Function

The library includes one built-in function:

| Function | Usage | Description |
|----------|-------|-------------|
| `link` | `{{link "https://example.com" "Click here"}}` | Create clickable hyperlink |

### Go Template Built-ins (Always Available)

These functions are provided by Go's `text/template` package:

| Function | Usage | Description |
|----------|-------|-------------|
| `and` | `{{if and .A .B}}` | Logical AND |
| `or` | `{{if or .A .B}}` | Logical OR |
| `not` | `{{if not .A}}` | Logical NOT |
| `eq` | `{{if eq .A .B}}` | Equal |
| `ne` | `{{if ne .A .B}}` | Not equal |
| `lt` | `{{if lt .A .B}}` | Less than |
| `le` | `{{if le .A .B}}` | Less than or equal |
| `gt` | `{{if gt .A .B}}` | Greater than |
| `ge` | `{{if ge .A .B}}` | Greater than or equal |
| `len` | `{{len .Items}}` | Length of slice/map/string |
| `index` | `{{index .Items 0}}` | Element at index |
| `slice` | `{{slice .Items 1 3}}` | Sub-slice |
| `print` | `{{print .A .B}}` | Concatenate values |
| `printf` | `{{printf "%.2f" .Price}}` | Formatted output |
| `println` | `{{println .A}}` | Print with newline |

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

### RegisterFuncMap
```go
func (d *DocxTmpl) RegisterFuncMap(funcs template.FuncMap)
```
Register multiple functions at once. Useful for adding external function libraries.

**Example:**
```go
doc.RegisterFuncMap(sprig.FuncMap())
```

### GetRegisteredFunctions
```go
func (d *DocxTmpl) GetRegisteredFunctions() *template.FuncMap
```
Get a copy of all registered functions.

---

## Using Community Function Libraries

For a rich set of template functions, use community libraries like **Sprig** or **Sprout**.

### Using Sprig (100+ functions)

[Sprig](https://github.com/Masterminds/sprig) provides 100+ template functions.

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
    doc.RegisterFuncMap(sprig.FuncMap())
    doc.Render(data)
    doc.SaveToFile("output.docx")
}
```

### Using Sprout (Modern Alternative)

[Sprout](https://github.com/go-sprout/sprout) is a modern, modular alternative.

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
    handler := sprout.New()
    doc.RegisterFuncMap(handler.Build())
    doc.Render(data)
    doc.SaveToFile("output.docx")
}
```

### Common Functions (via Sprig/Sprout)

Once registered, these functions become available:

| Category | Functions |
|----------|-----------|
| **Strings** | `upper`, `lower`, `title`, `trim`, `replace`, `contains`, `hasPrefix`, `hasSuffix` |
| **Math** | `add`, `sub`, `mul`, `div`, `mod`, `max`, `min`, `ceil`, `floor`, `round` |
| **Dates** | `now`, `date`, `dateModify`, `toDate`, `dateInZone` |
| **Lists** | `list`, `first`, `last`, `append`, `prepend`, `concat`, `join` |
| **Logic** | `default`, `ternary`, `coalesce` |
| **Encoding** | `b64enc`, `b64dec`, `toJson`, `fromJson` |

See [Sprig documentation](http://masterminds.github.io/sprig/) or [Sprout documentation](https://docs.atom.codes/sprout) for the complete function reference.

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

---

## Document Metadata

### GetMetadata
```go
func (d *DocxTmpl) GetMetadata() *DocumentMetadata
```
Extract document metadata (title, author, dates, etc.).

**Example:**
```go
meta := doc.GetMetadata()
fmt.Println("Author:", meta.Creator)
fmt.Println("Created:", meta.Created)
```

### SetMetadata
```go
func (d *DocxTmpl) SetMetadata(meta *DocumentMetadata)
```
Set document metadata. Only non-empty fields are updated.

**Example:**
```go
doc.SetMetadata(&DocumentMetadata{
    Title:   "My Report",
    Creator: "John Doe",
    Keywords: "report, quarterly, finance",
})
```

### GetStats
```go
func (d *DocxTmpl) GetStats() *DocumentStats
```
Get document statistics.

**Example:**
```go
stats := doc.GetStats()
fmt.Printf("Words: %d, Paragraphs: %d\n", stats.WordCount, stats.ParagraphCount)
```

**DocumentStats:**
```go
type DocumentStats struct {
    ParagraphCount int
    TableCount     int
    WordCount      int
    CharCount      int // without spaces
    CharCountSpace int // with spaces
    LineCount      int
    ImageCount     int
    LinkCount      int
}
```

### GetOutline
```go
func (d *DocxTmpl) GetOutline() []OutlineItem
```
Extract document structure as heading outline.

**Example:**
```go
outline := doc.GetOutline()
for _, item := range outline {
    fmt.Printf("H%d: %s\n", item.Level, item.Text)
}
```

### GetAllHyperlinks
```go
func (d *DocxTmpl) GetAllHyperlinks() []HyperlinkInfo
```
Get all hyperlinks in the document.

### GetAllStyles
```go
func (d *DocxTmpl) GetAllStyles() []string
```
Get all paragraph styles used in the document.

### GetTextByStyle
```go
func (d *DocxTmpl) GetTextByStyle(style string) []string
```
Get all text from paragraphs with a specific style.

**Example:**
```go
headings := doc.GetTextByStyle("Heading1")
```

---

## List Operations

### AddBulletList
```go
func (d *DocxTmpl) AddBulletList(items []string) *List
```
Add a bullet list.

**Example:**
```go
doc.AddBulletList([]string{
    "First item",
    "Second item",
    "Third item",
})
```

### AddNumberedList
```go
func (d *DocxTmpl) AddNumberedList(items []string) *List
```
Add a numbered list.

**Example:**
```go
doc.AddNumberedList([]string{
    "Step one",
    "Step two",
    "Step three",
})
```

### AddNestedList
```go
func (d *DocxTmpl) AddNestedList(listType ListType, items []ListItem) *List
```
Add a nested list with multiple levels.

**Example:**
```go
doc.AddNestedList(ListTypeBullet, []ListItem{
    {Text: "Item 1", Children: []ListItem{
        {Text: "Sub-item 1.1"},
        {Text: "Sub-item 1.2"},
    }},
    {Text: "Item 2"},
})
```

### NewListBuilder
```go
func (d *DocxTmpl) NewListBuilder(listType ListType) *ListBuilder
```
Create a fluent list builder.

**Example:**
```go
doc.NewListBuilder(ListTypeBullet).
    Item("First").
    Item("Second").
    Indent().Item("Nested").
    Outdent().Item("Third").
    Build()
```

### AddChecklistItem
```go
func (d *DocxTmpl) AddChecklistItem(text string, checked bool) *Paragraph
```
Add a checkbox item.

**Example:**
```go
doc.AddChecklistItem("Complete task", true)
doc.AddChecklistItem("Pending task", false)
```

---

## Enhanced Table Operations

### Table Row/Column Operations

| Method | Description |
|--------|-------------|
| `AddRow() *TableRow` | Add new empty row |
| `AddRowWithData(values ...string) *TableRow` | Add row with data |
| `InsertRow(index int) *TableRow` | Insert row at position |
| `DeleteRow(index int)` | Delete row |
| `AddColumn()` | Add new column |
| `InsertColumn(index int)` | Insert column at position |
| `DeleteColumn(index int)` | Delete column |

### Table Sorting
```go
func (t *Table) SortByColumn(col int, ascending, skipHeader bool) *Table
```
Sort table rows by column.

**Example:**
```go
table.SortByColumn(0, true, true) // Sort by first column, ascending, skip header
```

### Table Export

#### ToJSON
```go
func (t *Table) ToJSON(headers bool) (string, error)
```
Export table as JSON.

**Example:**
```go
jsonStr, _ := table.ToJSON(true) // Use first row as headers
// Returns: [{"Name":"John","Age":"30"},{"Name":"Jane","Age":"25"}]
```

#### ToCSV
```go
func (t *Table) ToCSV() string
```
Export table as CSV.

#### ToSlice
```go
func (t *Table) ToSlice() [][]string
```
Export table as 2D string slice.

### Table Import

#### AddTableFromJSON
```go
func (d *DocxTmpl) AddTableFromJSON(jsonStr string, headers ...string) (*Table, error)
```
Create table from JSON.

**Example:**
```go
doc.AddTableFromJSON(`[{"Name":"John","Age":"30"}]`)
```

#### AddTableFromCSV
```go
func (d *DocxTmpl) AddTableFromCSV(csvStr string) (*Table, error)
```
Create table from CSV.

#### AddTableFromSlice
```go
func (d *DocxTmpl) AddTableFromSlice(data [][]string) *Table
```
Create table from 2D slice.

#### AddTableWithHeaders
```go
func (d *DocxTmpl) AddTableWithHeaders(headers []string, data [][]string) *Table
```
Create table with header row (automatically bolded and centered).

**Example:**
```go
doc.AddTableWithHeaders(
    []string{"Name", "Age", "City"},
    [][]string{
        {"John", "30", "NYC"},
        {"Jane", "25", "LA"},
    },
)
```

### Table Cell Operations

| Method | Description |
|--------|-------------|
| `SetColumnWidth(col, twips int)` | Set column width |
| `SetAllColumnWidths(widths []int)` | Set all column widths |
| `SetRowHeight(row, twips int)` | Set row height |
| `ClearRow(index int)` | Clear row content |
| `ClearColumn(col int)` | Clear column content |
| `FillColumn(col int, value string)` | Fill column with value |
| `FillRow(row int, value string)` | Fill row with value |
| `GetColumn(col int) []string` | Get column values |
| `GetRowData(row int) []string` | Get row values |
| `FindRow(col int, value string) int` | Find row by value |
| `FindAllRows(col int, value string) []int` | Find all matching rows |

---

## Export/Conversion Functions

### ToStructured
```go
func (d *DocxTmpl) ToStructured() *StructuredDocument
```
Convert document to structured format for AI/LLM consumption.

**Example:**
```go
structured := doc.ToStructured()
fmt.Println("Paragraphs:", len(structured.Paragraphs))
```

### ToJSON
```go
func (d *DocxTmpl) ToJSON() (string, error)
```
Export document structure as JSON.

**Example:**
```go
jsonStr, _ := doc.ToJSON()
fmt.Println(jsonStr)
```

### ToMarkdown
```go
func (d *DocxTmpl) ToMarkdown() string
```
Convert document to Markdown format.

**Example:**
```go
markdown := doc.ToMarkdown()
// Output:
// # Title
// ## Heading
// **Bold text** and *italic text*
// | Header | Header |
// | --- | --- |
// | Cell | Cell |
```

### ToHTML
```go
func (d *DocxTmpl) ToHTML() string
```
Convert document to basic HTML.

**Example:**
```go
html := doc.ToHTML()
// Output: <!DOCTYPE html><html>...
```

---

## Section Operations

### AddSection
```go
func (d *DocxTmpl) AddSection() *DocxTmpl
```
Add a section break (page break).

### AddSectionBreak
```go
func (d *DocxTmpl) AddSectionBreak(breakType SectionBreakType) *DocxTmpl
```
Add a section break of specified type.

**Section Break Types:**
- `SectionBreakNextPage` - Start on next page
- `SectionBreakContinuous` - Continue on same page
- `SectionBreakEvenPage` - Start on next even page
- `SectionBreakOddPage` - Start on next odd page

### EstimatePageCount
```go
func (d *DocxTmpl) EstimatePageCount() int
```
Estimate number of pages in document.

---

## Template Validation

### ValidateData
```go
func (d *DocxTmpl) ValidateData(data any) []ValidationError
```
Validate that data contains all required template fields.

**Example:**
```go
errors := doc.ValidateData(data)
if len(errors) > 0 {
    for _, err := range errors {
        fmt.Printf("Missing: %s\n", err.Field)
    }
}
```

### GetRequiredFields
```go
func (d *DocxTmpl) GetRequiredFields() []FieldInfo
```
Get information about all required template fields.

**Example:**
```go
fields := doc.GetRequiredFields()
for _, f := range fields {
    fmt.Printf("Field: %s, Type: %s, Occurrences: %d\n",
        f.Name, f.Type, f.Occurrences)
}
```

### GetPlaceholderSchema
```go
func (d *DocxTmpl) GetPlaceholderSchema() map[string]FieldSchema
```
Get JSON schema-like representation of placeholders.

### PreviewRender
```go
func (d *DocxTmpl) PreviewRender(data any) (string, error)
```
Render template and return text content (for preview/validation).

### GenerateSampleData
```go
func (d *DocxTmpl) GenerateSampleData() map[string]interface{}
```
Generate sample data matching template placeholders.

### ValidatePlaceholderSyntax
```go
func (d *DocxTmpl) ValidatePlaceholderSyntax() []ValidationError
```
Check placeholders for syntax errors.

---

## Document Comparison

### CompareDocuments
```go
func CompareDocuments(doc1, doc2 *DocxTmpl) *DocumentDiff
```
Compare two documents and return differences.

**Example:**
```go
doc1, _ := docxtpl.ParseFromFilename("version1.docx")
doc2, _ := docxtpl.ParseFromFilename("version2.docx")
diff := docxtpl.CompareDocuments(doc1, doc2)
fmt.Println(diff.String())
```

### DiffWith
```go
func (d *DocxTmpl) DiffWith(other *DocxTmpl) *DocumentDiff
```
Compare this document with another.

**Example:**
```go
diff := doc1.DiffWith(doc2)
if diff.HasChanges() {
    fmt.Printf("Total changes: %d\n", diff.Summary.TotalChanges)
}
```

### DocumentDiff Methods

| Method | Description |
|--------|-------------|
| `HasChanges() bool` | Check if there are any differences |
| `String() string` | Get human-readable summary |
| `GetChanges() []DiffItem` | Get all changes as flat list |

### CompareStats
```go
func CompareStats(stats1, stats2 *DocumentStats) map[string]int
```
Compare document statistics.

### CompareMetadata
```go
func CompareMetadata(meta1, meta2 *DocumentMetadata) map[string][2]string
```
Compare document metadata.

---

## Batch Operations

### Clone
```go
func (d *DocxTmpl) Clone() (*DocxTmpl, error)
```
Create a deep copy of the document.

**Example:**
```go
clone, _ := doc.Clone()
clone.Render(differentData)
clone.SaveToFile("copy.docx")
```

### MergeDocuments
```go
func MergeDocuments(docs ...*DocxTmpl) (*DocxTmpl, error)
```
Combine multiple documents into one.

**Example:**
```go
merged, _ := docxtpl.MergeDocuments(doc1, doc2, doc3)
merged.SaveToFile("combined.docx")
```

### MailMerge
```go
func (d *DocxTmpl) MailMerge(records []map[string]any) ([]*DocxTmpl, error)
```
Render template with multiple data records.

**Example:**
```go
records := []map[string]any{
    {"Name": "John", "Email": "john@example.com"},
    {"Name": "Jane", "Email": "jane@example.com"},
}
docs, _ := template.MailMerge(records)
```

### MailMergeToFiles
```go
func (d *DocxTmpl) MailMergeToFiles(records []map[string]any, pattern string) error
```
Mail merge and save each result to a file.

**Example:**
```go
template.MailMergeToFiles(records, "output/letter_%d.docx")
// Creates: output/letter_1.docx, output/letter_2.docx, ...
```

### MailMergeToSingle
```go
func (d *DocxTmpl) MailMergeToSingle(records []map[string]any) (*DocxTmpl, error)
```
Mail merge into a single document with page breaks.

### SearchWithContext
```go
func (d *DocxTmpl) SearchWithContext(text string, contextLines int) []SearchResult
```
Search for text and return matches with surrounding context.

**Example:**
```go
results := doc.SearchWithContext("important", 2)
for _, r := range results {
    fmt.Printf("Found at paragraph %d:\n%s\n", r.Paragraph, r.Context)
}
```

### Batch Utilities

| Function | Description |
|----------|-------------|
| `BatchProcess(docs, fn)` | Apply function to each document |
| `BatchRender(templates, dataList)` | Render multiple templates |
| `ReplaceInAll(docs, old, new)` | Replace text in all documents |
| `LoadAllFromDirectory(dir)` | Load all .docx files from directory |
| `SaveAllToDirectory(docs, dir, names)` | Save all documents to directory |
| `ProcessDirectory(in, out, fn)` | Process all .docx files in directory |

**Example:**
```go
// Load all documents from directory
docs, _ := docxtpl.LoadAllFromDirectory("templates/")

// Replace text in all
docxtpl.ReplaceInAll(docs, "2023", "2024")

// Save all to output directory
names := []string{"doc1.docx", "doc2.docx"}
docxtpl.SaveAllToDirectory(docs, "output/", names)
```

---

## Tracked Changes

Track changes (revisions) support for reviewing document modifications.

### Types

```go
type TrackedChangeType string

const (
    ChangeTypeInsertion TrackedChangeType = "insertion"
    ChangeTypeDeletion  TrackedChangeType = "deletion"
)

type TrackedChange struct {
    ID       int               // Change ID
    Type     TrackedChangeType // insertion or deletion
    Author   string            // Author who made the change
    Date     time.Time         // When the change was made
    Text     string            // The changed text
    Location string            // Location description
}
```

### Methods

| Method | Description |
|--------|-------------|
| `GetTrackedChanges()` | Get all tracked changes |
| `HasTrackedChanges()` | Check if document has changes |
| `GetInsertions()` | Get only insertions |
| `GetDeletions()` | Get only deletions |
| `CountTrackedChanges()` | Count insertions and deletions |
| `AcceptAllChanges()` | Accept all tracked changes |
| `RejectAllChanges()` | Reject all tracked changes |
| `GetChangesByAuthor(author)` | Filter changes by author |
| `EnableTrackChanges()` | Enable track changes mode |
| `DisableTrackChanges()` | Disable track changes mode |
| `IsTrackChangesEnabled()` | Check if track changes is enabled |
| `TrackedChangesSummary()` | Get text summary of changes |

**Example:**
```go
// Get all tracked changes
changes := doc.GetTrackedChanges()
for _, change := range changes {
    fmt.Printf("%s by %s: %s\n", change.Type, change.Author, change.Text)
}

// Count changes
ins, del := doc.CountTrackedChanges()
fmt.Printf("Insertions: %d, Deletions: %d\n", ins, del)

// Accept all changes
doc.AcceptAllChanges()
```

---

## Comments

Extract and manage document comments.

### Types

```go
type Comment struct {
    ID        int       // Comment ID
    Author    string    // Comment author
    Initials  string    // Author initials
    Date      time.Time // When the comment was created
    Text      string    // Comment text content
    ParentID  int       // Parent comment ID for replies (-1 if not a reply)
    Paragraph int       // Paragraph index
}
```

### Methods

| Method | Description |
|--------|-------------|
| `GetComments()` | Get all comments |
| `HasComments()` | Check if document has comments |
| `CountComments()` | Count total comments |
| `GetCommentsByAuthor(author)` | Filter comments by author |
| `GetCommentReplies(commentID)` | Get replies to a comment |
| `GetTopLevelComments()` | Get only top-level comments |
| `GetCommentAuthors()` | Get list of comment authors |
| `GetCommentsInDateRange(start, end)` | Filter by date range |
| `DeleteAllComments()` | Remove all comments |
| `CommentsSummary()` | Get text summary |

**Example:**
```go
// Get all comments
comments := doc.GetComments()
for _, c := range comments {
    fmt.Printf("[%s] %s: %s\n", c.Date.Format("2006-01-02"), c.Author, c.Text)
}

// Get comments by author
myComments := doc.GetCommentsByAuthor("John Doe")

// Get summary
fmt.Println(doc.CommentsSummary())
```

---

## Raw XML Access

Direct access to the underlying DOCX archive structure.

### Types

```go
type XMLFile struct {
    Path    string // File path within archive
    Content string // XML content
}

type ArchiveInfo struct {
    TotalFiles   int      // Total files in archive
    XMLFiles     int      // Number of XML files
    MediaFiles   int      // Number of media files
    TotalSize    int64    // Total uncompressed size
    FileList     []string // List of all file paths
    HasComments  bool     // Has comments.xml
    HasSettings  bool     // Has settings.xml
    HasFootnotes bool     // Has footnotes.xml
    HasEndnotes  bool     // Has endnotes.xml
}
```

### Methods

| Method | Description |
|--------|-------------|
| `UnpackToDirectory(dirPath)` | Extract DOCX to directory |
| `PackFromDirectory(dirPath)` | Create DOCX from directory |
| `GetXMLFiles()` | List all XML files |
| `GetXMLContent(filePath)` | Get XML content |
| `SetXMLContent(filePath, content)` | Set XML content |
| `GetArchiveInfo()` | Get archive information |
| `GetRelationships()` | Get document relationships |
| `GetContentTypesXML()` | Get content types |
| `GetDocumentXML()` | Get main document XML |
| `GetSettingsXML()` | Get settings XML |
| `GetStylesXML()` | Get styles XML |

**Example:**
```go
// Unpack to directory for manual editing
err := doc.UnpackToDirectory("/tmp/unpacked")

// Get list of XML files
files := doc.GetXMLFiles()
for _, f := range files {
    fmt.Println(f)
}

// Get document XML
xml, _ := doc.GetDocumentXML()

// Get archive info
info := doc.GetArchiveInfo()
fmt.Printf("Total files: %d, XML: %d, Media: %d\n",
    info.TotalFiles, info.XMLFiles, info.MediaFiles)

// Pack from directory
doc, _ := docxtpl.PackFromDirectory("/tmp/unpacked")
```

---

## Bookmarks

Navigate and manage document bookmarks.

### Types

```go
type Bookmark struct {
    ID   int    // Bookmark ID
    Name string // Bookmark name
    Text string // Text at bookmark location
}

type InternalLink struct {
    Anchor string // Bookmark name
    Text   string // Display text
}

type TableOfContentsEntry struct {
    Level    int    // Entry level (1, 2, 3, etc.)
    Text     string // Entry text
    Page     string // Page number
    Bookmark string // Bookmark reference
}
```

### Methods

| Method | Description |
|--------|-------------|
| `GetBookmarks()` | Get all bookmarks |
| `HasBookmark(name)` | Check if bookmark exists |
| `GetBookmarkByName(name)` | Get bookmark by name |
| `CountBookmarks()` | Count bookmarks |
| `GetBookmarkNames()` | Get list of bookmark names |
| `GetInternalLinks()` | Get internal hyperlinks |
| `GetTableOfContents()` | Extract table of contents |
| `HasTableOfContents()` | Check if TOC exists |
| `BookmarksSummary()` | Get text summary |

**Example:**
```go
// Get all bookmarks
bookmarks := doc.GetBookmarks()
for _, b := range bookmarks {
    fmt.Printf("%s (ID: %d)\n", b.Name, b.ID)
}

// Check for specific bookmark
if doc.HasBookmark("Chapter1") {
    bookmark, _ := doc.GetBookmarkByName("Chapter1")
    fmt.Println(bookmark.Name)
}

// Get internal links
links := doc.GetInternalLinks()
for _, l := range links {
    fmt.Printf("'%s' -> #%s\n", l.Text, l.Anchor)
}

// Get table of contents
toc := doc.GetTableOfContents()
for _, entry := range toc {
    fmt.Printf("%s%s\n", strings.Repeat("  ", entry.Level-1), entry.Text)
}
```

---

## Document Protection

Manage document protection and restrictions.

### Types

```go
type ProtectionType string

const (
    ProtectionNone           ProtectionType = "none"
    ProtectionReadOnly       ProtectionType = "readOnly"
    ProtectionComments       ProtectionType = "comments"
    ProtectionTrackedChanges ProtectionType = "trackedChanges"
    ProtectionForms          ProtectionType = "forms"
)

type ProtectionInfo struct {
    IsProtected    bool           // Whether protected
    Type           ProtectionType // Protection type
    HasPassword    bool           // Password set
    EnforceMessage string         // Enforcement message
}

type RestrictionInfo struct {
    CanEdit       bool // Can edit
    CanComment    bool // Can comment
    CanTrack      bool // Can make tracked changes
    CanFillForms  bool // Can fill forms
    CanFormatText bool // Can format text
}
```

### Methods

| Method | Description |
|--------|-------------|
| `GetProtectionInfo()` | Get protection information |
| `IsProtected()` | Check if document is protected |
| `IsReadOnly()` | Check if read-only |
| `SetProtection(type)` | Set protection type |
| `RemoveProtection()` | Remove protection |
| `SetReadOnly()` | Set read-only |
| `AllowOnlyComments()` | Allow only comments |
| `AllowOnlyTrackedChanges()` | Allow only tracked changes |
| `AllowOnlyFormFilling()` | Allow only form filling |
| `GetRestrictions()` | Get detailed restrictions |
| `ProtectionSummary()` | Get text summary |

**Example:**
```go
// Check protection
info := doc.GetProtectionInfo()
if info.IsProtected {
    fmt.Printf("Protected: %s, Password: %v\n", info.Type, info.HasPassword)
}

// Set read-only protection
doc.SetReadOnly()

// Allow only comments
doc.AllowOnlyComments()

// Check restrictions
restrictions := doc.GetRestrictions()
if restrictions.CanComment {
    fmt.Println("Comments are allowed")
}

// Remove protection
doc.RemoveProtection()
```

**Note:** Password-based protection requires the Word application or a dedicated library.
