package docxtpl

import (
	"text/template"

	"github.com/fumiama/go-docx"
	"github.com/abdokhaire/go-docxgen/internal/contenttypes"
	"github.com/abdokhaire/go-docxgen/internal/headerfooter"
	"github.com/abdokhaire/go-docxgen/internal/hyperlinks"
)

// PageSize represents standard page sizes
type PageSize int

const (
	PageSizeA4 PageSize = iota
	PageSizeA3
	PageSizeLetter
	PageSizeLegal
)

// New creates a new empty document from scratch.
// Use this when you want to build a document programmatically without a template.
//
//	doc := docxtpl.New()
//	doc.AddHeading("My Document", 1)
//	doc.AddParagraph("Hello, World!")
//	doc.SaveToFile("output.docx")
func New() *DocxTmpl {
	return NewWithOptions()
}

// NewWithOptions creates a new document with optional configuration.
//
//	doc := docxtpl.NewWithOptions(docxtpl.PageSizeA4)
func NewWithOptions(pageSize ...PageSize) *DocxTmpl {
	w := docx.New().WithDefaultTheme()

	// Apply page size if specified
	if len(pageSize) > 0 {
		switch pageSize[0] {
		case PageSizeA4:
			w = w.WithA4Page()
		case PageSizeA3:
			w = w.WithA3Page()
		// Letter and Legal use default sizing
		}
	}

	funcMap := make(template.FuncMap)

	hyperlinkReg := hyperlinks.NewHyperlinkRegistry()
	contentTypes := contenttypes.NewContentTypes()

	docTmpl := &DocxTmpl{
		Docx:             w,
		funcMap:          funcMap,
		contentTypes:     contentTypes,
		processableFiles: []headerfooter.DocxFile{},
		hyperlinkReg:     hyperlinkReg,
	}

	// Override the link function to use our hyperlink registry
	docTmpl.funcMap["link"] = docTmpl.createLink

	return docTmpl
}

// AddParagraph adds a new paragraph with the given text to the document.
// Returns a Paragraph wrapper for further formatting.
//
//	para := doc.AddParagraph("Hello, World!")
//	para.Bold().Color("FF0000")
func (d *DocxTmpl) AddParagraph(text string) *Paragraph {
	p := d.Docx.AddParagraph()
	run := p.AddText(text)
	return &Paragraph{
		paragraph: p,
		lastRun:   run,
		doc:       d,
	}
}

// AddEmptyParagraph adds an empty paragraph to the document.
// Useful for adding spacing between elements.
func (d *DocxTmpl) AddEmptyParagraph() *Paragraph {
	p := d.Docx.AddParagraph()
	return &Paragraph{
		paragraph: p,
		doc:       d,
	}
}

// AddHeading adds a heading with the specified level (0-9).
// Level 0 is the document title, level 1 is Heading 1, etc.
//
//	doc.AddHeading("Chapter 1", 1)
//	doc.AddHeading("Section 1.1", 2)
func (d *DocxTmpl) AddHeading(text string, level int) *Paragraph {
	if level < 0 {
		level = 0
	}
	if level > 9 {
		level = 9
	}

	p := d.Docx.AddParagraph()
	run := p.AddText(text)

	// Apply heading style
	styleID := "Title"
	if level > 0 {
		styleID = headingStyleID(level)
	}
	p.Style(styleID)

	return &Paragraph{
		paragraph: p,
		lastRun:   run,
		doc:       d,
	}
}

// AddPageBreak inserts a page break at the current position.
func (d *DocxTmpl) AddPageBreak() {
	p := d.Docx.AddParagraph()
	p.AddPageBreaks()
}

// AddTable creates a new table with the specified number of rows and columns.
// Returns a Table wrapper for populating and formatting.
//
//	table := doc.AddTable(3, 4) // 3 rows, 4 columns
//	table.SetCell(0, 0, "Header 1")
func (d *DocxTmpl) AddTable(rows, cols int) *Table {
	// Default table width (full page width in twips, ~6.5 inches)
	tableWidth := int64(9360)
	t := d.Docx.AddTable(rows, cols, tableWidth, nil)
	return &Table{
		table: t,
		doc:   d,
		rows:  rows,
		cols:  cols,
	}
}

// AddTableWithWidths creates a table with custom column widths (in twips).
// 1 inch = 1440 twips, 1 cm = 567 twips.
//
//	doc.AddTableWithWidths(3, []int{2880, 1440, 1440}) // 3 rows, columns: 2", 1", 1"
func (d *DocxTmpl) AddTableWithWidths(rows int, colWidths []int) *Table {
	// Convert int slice to int64 slice
	widths := make([]int64, len(colWidths))
	totalWidth := int64(0)
	for i, w := range colWidths {
		widths[i] = int64(w)
		totalWidth += widths[i]
	}

	// Create row heights (default height)
	rowHeights := make([]int64, rows)
	for i := range rowHeights {
		rowHeights[i] = 400 // Default row height
	}

	t := d.Docx.AddTableTwips(rowHeights, widths, totalWidth, nil)
	return &Table{
		table: t,
		doc:   d,
		rows:  rows,
		cols:  len(colWidths),
	}
}

// AddTableWithBorders creates a table with custom border colors.
// Use TableBorderColors to specify colors for each border type.
//
//	doc.AddTableWithBorders(3, 4, TableBorderColors{
//	    Top: "FF0000", Bottom: "FF0000",
//	    Left: "0000FF", Right: "0000FF",
//	    InsideH: "00FF00", InsideV: "00FF00",
//	})
func (d *DocxTmpl) AddTableWithBorders(rows, cols int, colors TableBorderColors) *Table {
	tableWidth := int64(9360)
	apiColors := &docx.APITableBorderColors{
		Top:     colors.Top,
		Left:    colors.Left,
		Bottom:  colors.Bottom,
		Right:   colors.Right,
		InsideH: colors.InsideH,
		InsideV: colors.InsideV,
	}
	t := d.Docx.AddTable(rows, cols, tableWidth, apiColors)
	return &Table{
		table: t,
		doc:   d,
		rows:  rows,
		cols:  cols,
	}
}

// headingStyleID returns the style ID for a heading level
func headingStyleID(level int) string {
	switch level {
	case 1:
		return "Heading1"
	case 2:
		return "Heading2"
	case 3:
		return "Heading3"
	case 4:
		return "Heading4"
	case 5:
		return "Heading5"
	case 6:
		return "Heading6"
	case 7:
		return "Heading7"
	case 8:
		return "Heading8"
	case 9:
		return "Heading9"
	default:
		return "Heading1"
	}
}
