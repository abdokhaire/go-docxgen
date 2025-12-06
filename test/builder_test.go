package docxtpl_test

import (
	"os"
	"testing"

	"github.com/abdokhaire/go-docxgen"
	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

func TestNew(t *testing.T) {
	t.Run("Should create a new empty document", func(t *testing.T) {
		doc := docxtpl.New()
		assert.NotNil(t, doc)
		assert.NotNil(t, doc.Docx)
	})

	t.Run("Should create document with A4 page size", func(t *testing.T) {
		doc := docxtpl.NewWithOptions(docxtpl.PageSizeA4)
		assert.NotNil(t, doc)
	})
}

func TestAddParagraph(t *testing.T) {
	t.Run("Should add paragraph with text", func(t *testing.T) {
		doc := docxtpl.New()
		para := doc.AddParagraph("Hello, World!")
		assert.NotNil(t, para)
		assert.NotNil(t, para.GetRaw())
	})

	t.Run("Should support chaining formatting", func(t *testing.T) {
		doc := docxtpl.New()
		para := doc.AddParagraph("Formatted text").Bold().Italic().Color("FF0000")
		assert.NotNil(t, para)
	})
}

func TestAddHeading(t *testing.T) {
	t.Run("Should add heading with level", func(t *testing.T) {
		doc := docxtpl.New()
		heading := doc.AddHeading("Chapter 1", 1)
		assert.NotNil(t, heading)
	})

	t.Run("Should clamp level to valid range", func(t *testing.T) {
		doc := docxtpl.New()
		// Level -1 should be clamped to 0
		heading := doc.AddHeading("Title", -1)
		assert.NotNil(t, heading)

		// Level 10 should be clamped to 9
		heading2 := doc.AddHeading("Heading", 10)
		assert.NotNil(t, heading2)
	})
}

func TestAddTable(t *testing.T) {
	t.Run("Should create table with rows and columns", func(t *testing.T) {
		doc := docxtpl.New()
		table := doc.AddTable(3, 4)
		assert.NotNil(t, table)
		assert.Equal(t, 3, table.Rows())
		assert.Equal(t, 4, table.Cols())
	})

	t.Run("Should set cell content", func(t *testing.T) {
		doc := docxtpl.New()
		table := doc.AddTable(2, 2)
		cell := table.SetCell(0, 0, "Header")
		assert.NotNil(t, cell)
	})

	t.Run("Should return nil for out of bounds cell", func(t *testing.T) {
		doc := docxtpl.New()
		table := doc.AddTable(2, 2)
		cell := table.Cell(5, 5)
		assert.Nil(t, cell)
	})
}

func TestParagraphJustification(t *testing.T) {
	t.Run("Should support different alignments", func(t *testing.T) {
		doc := docxtpl.New()

		para1 := doc.AddParagraph("Left aligned").Left()
		assert.NotNil(t, para1)

		para2 := doc.AddParagraph("Center aligned").Center()
		assert.NotNil(t, para2)

		para3 := doc.AddParagraph("Right aligned").Right()
		assert.NotNil(t, para3)

		para4 := doc.AddParagraph("Justified").Justified()
		assert.NotNil(t, para4)
	})
}

func TestRunFormatting(t *testing.T) {
	t.Run("Should support run-level formatting", func(t *testing.T) {
		doc := docxtpl.New()
		para := doc.AddParagraph("Start ")
		run := para.AddText("formatted").Bold().Italic().Color("0000FF")
		assert.NotNil(t, run)

		// Continue with paragraph
		para2 := run.Then()
		assert.Equal(t, para, para2)
	})
}

func TestDocumentProperties(t *testing.T) {
	t.Run("Should set document properties and retrieve them", func(t *testing.T) {
		doc := docxtpl.New()

		// Set properties
		doc.SetTitle("Test Document")
		doc.SetAuthor("Test Author")
		doc.SetSubject("Test Subject")
		doc.SetKeywords("test, document, go")

		// After setting, GetProperties should return the values
		props := doc.GetProperties()
		assert.Equal(t, "Test Document", props.Title)
		assert.Equal(t, "Test Author", props.Creator)
		assert.Equal(t, "Test Subject", props.Subject)
		assert.Equal(t, "test, document, go", props.Keywords)
	})

	t.Run("Should return empty properties for new document", func(t *testing.T) {
		doc := docxtpl.New()
		props := doc.GetProperties()
		assert.Equal(t, "", props.Title)
		assert.Equal(t, "", props.Creator)
	})
}

func TestSaveNewDocument(t *testing.T) {
	t.Run("Should save a new document to file", func(t *testing.T) {
		doc := docxtpl.New()

		doc.SetTitle("Generated Document")
		doc.SetAuthor("go-docxgen")

		doc.AddHeading("Welcome", 1)
		doc.AddParagraph("This is a programmatically generated document.")
		doc.AddEmptyParagraph()

		para := doc.AddParagraph("This text is ").Bold()
		para.AddText("bold and ").Bold()
		para.AddText("this is italic.").Italic()

		doc.AddHeading("Table Example", 2)
		table := doc.AddTable(3, 3)
		table.SetCell(0, 0, "Name").Bold().Background("CCCCCC")
		table.SetCell(0, 1, "Age").Bold().Background("CCCCCC")
		table.SetCell(0, 2, "City").Bold().Background("CCCCCC")
		table.SetCell(1, 0, "Alice")
		table.SetCell(1, 1, "30")
		table.SetCell(1, 2, "New York")
		table.SetCell(2, 0, "Bob")
		table.SetCell(2, 1, "25")
		table.SetCell(2, 2, "Los Angeles")

		outputPath := "testdata/templates/generated_builder_test.docx"
		err := doc.SaveToFile(outputPath)
		require.NoError(t, err)

		// Verify file exists
		_, err = os.Stat(outputPath)
		assert.NoError(t, err)

		// Clean up
		os.Remove(outputPath)
	})
}

func TestParagraphSpacingAndIndent(t *testing.T) {
	t.Run("Should set paragraph spacing", func(t *testing.T) {
		doc := docxtpl.New()
		para := doc.AddParagraph("Test paragraph")
		para.SpacingBefore(12).LineSpacingDouble()
		assert.NotNil(t, para.GetRaw().Properties)
		assert.NotNil(t, para.GetRaw().Properties.Spacing)
	})

	t.Run("Should set paragraph indentation", func(t *testing.T) {
		doc := docxtpl.New()
		para := doc.AddParagraph("Indented paragraph")
		para.IndentLeft(0.5).IndentFirstLine(0.25)
		assert.NotNil(t, para.GetRaw().Properties)
		assert.NotNil(t, para.GetRaw().Properties.Ind)
	})

	t.Run("Should support line spacing presets", func(t *testing.T) {
		doc := docxtpl.New()
		p1 := doc.AddParagraph("Single").LineSpacingSingle()
		p2 := doc.AddParagraph("One and half").LineSpacingOneAndHalf()
		p3 := doc.AddParagraph("Double").LineSpacingDouble()
		assert.NotNil(t, p1)
		assert.NotNil(t, p2)
		assert.NotNil(t, p3)
	})
}

func TestHyperlinks(t *testing.T) {
	t.Run("Should add hyperlink to paragraph", func(t *testing.T) {
		doc := docxtpl.New()
		para := doc.AddParagraph("Visit ")
		link := para.AddLink("our website", "https://example.com")
		assert.NotNil(t, link)
		assert.NotNil(t, link.GetRaw())
	})
}

func TestTableCellMerging(t *testing.T) {
	t.Run("Should merge cells horizontally", func(t *testing.T) {
		doc := docxtpl.New()
		table := doc.AddTable(2, 4)
		cell := table.Cell(0, 0).MergeHorizontal(2) // Span 3 columns
		assert.NotNil(t, cell)
		assert.NotNil(t, cell.GetRaw().TableCellProperties)
		assert.NotNil(t, cell.GetRaw().TableCellProperties.GridSpan)
	})

	t.Run("Should merge cells vertically", func(t *testing.T) {
		doc := docxtpl.New()
		table := doc.AddTable(3, 2)
		table.Cell(0, 0).MergeVerticalStart()
		table.Cell(1, 0).MergeVerticalContinue()
		table.Cell(2, 0).MergeVerticalContinue()
		assert.NotNil(t, table.Cell(0, 0).GetRaw().TableCellProperties.VMerge)
	})
}

func TestTableCellWidth(t *testing.T) {
	t.Run("Should set cell width in twips", func(t *testing.T) {
		doc := docxtpl.New()
		table := doc.AddTable(2, 2)
		cell := table.Cell(0, 0).Width(2880)
		assert.NotNil(t, cell.GetRaw().TableCellProperties)
		assert.NotNil(t, cell.GetRaw().TableCellProperties.TableCellWidth)
	})

	t.Run("Should set cell width in inches", func(t *testing.T) {
		doc := docxtpl.New()
		table := doc.AddTable(2, 2)
		cell := table.Cell(0, 0).WidthInches(1.5)
		assert.NotNil(t, cell)
	})

	t.Run("Should set cell width as percentage", func(t *testing.T) {
		doc := docxtpl.New()
		table := doc.AddTable(2, 2)
		cell := table.Cell(0, 0).WidthPercent(50)
		assert.NotNil(t, cell)
	})
}

func TestTableCellVerticalAlignment(t *testing.T) {
	t.Run("Should set vertical alignment", func(t *testing.T) {
		doc := docxtpl.New()
		table := doc.AddTable(2, 2)
		cell := table.Cell(0, 0).VAlignCenter()
		assert.NotNil(t, cell.GetRaw().TableCellProperties)
		assert.NotNil(t, cell.GetRaw().TableCellProperties.VAlign)
	})
}

func TestTableBorders(t *testing.T) {
	t.Run("Should create table with custom borders", func(t *testing.T) {
		doc := docxtpl.New()
		table := doc.AddTableWithBorders(2, 2, docxtpl.TableBorderColors{
			Top: "FF0000", Bottom: "FF0000",
			Left: "0000FF", Right: "0000FF",
		})
		assert.NotNil(t, table)
	})

	t.Run("Should set cell borders", func(t *testing.T) {
		doc := docxtpl.New()
		table := doc.AddTable(2, 2)
		cell := table.Cell(0, 0).Borders("000000", 8)
		assert.NotNil(t, cell.GetRaw().TableCellProperties.TableBorders)
	})

	t.Run("Should remove cell borders", func(t *testing.T) {
		doc := docxtpl.New()
		table := doc.AddTable(2, 2)
		cell := table.Cell(0, 0).NoBorders()
		assert.NotNil(t, cell.GetRaw().TableCellProperties.TableBorders)
	})
}

func TestRunFormatting_Extended(t *testing.T) {
	t.Run("Should apply superscript", func(t *testing.T) {
		doc := docxtpl.New()
		para := doc.AddParagraph("x")
		run := para.AddText("2").Superscript()
		assert.NotNil(t, run.GetRaw().RunProperties)
		assert.NotNil(t, run.GetRaw().RunProperties.VertAlign)
	})

	t.Run("Should apply subscript", func(t *testing.T) {
		doc := docxtpl.New()
		para := doc.AddParagraph("H")
		run := para.AddText("2").Subscript()
		assert.NotNil(t, run.GetRaw().RunProperties.VertAlign)
	})

	t.Run("Should set character spacing", func(t *testing.T) {
		doc := docxtpl.New()
		para := doc.AddParagraph("")
		run := para.AddText("Expanded").Expand(2)
		assert.NotNil(t, run.GetRaw().RunProperties.Spacing)
	})

	t.Run("Should set kerning", func(t *testing.T) {
		doc := docxtpl.New()
		para := doc.AddParagraph("")
		run := para.AddText("Kerned").Kern(24)
		assert.NotNil(t, run.GetRaw().RunProperties.Kern)
	})
}

func TestTabStops(t *testing.T) {
	t.Run("Should add tab stops", func(t *testing.T) {
		doc := docxtpl.New()
		para := doc.AddParagraph("Column1\tColumn2\tColumn3")
		para.AddTabStop(1440, "left", "none")
		para.AddTabStop(4320, "center", "dot")
		assert.NotNil(t, para.GetRaw().Properties.Tabs)
		assert.Len(t, para.GetRaw().Properties.Tabs.Tabs, 2)
	})

	t.Run("Should add multiple tab stops at once", func(t *testing.T) {
		doc := docxtpl.New()
		para := doc.AddParagraph("Text")
		para.AddTabStops([]docxtpl.TabStop{
			{Position: 1440, Align: "left"},
			{Position: 2880, Align: "center"},
			{Position: 4320, Align: "right"},
		})
		assert.Len(t, para.GetRaw().Properties.Tabs.Tabs, 3)
	})

	t.Run("Should clear tab stops", func(t *testing.T) {
		doc := docxtpl.New()
		para := doc.AddParagraph("Text")
		para.AddTabStop(1440, "left", "none")
		para.ClearTabStops()
		assert.Nil(t, para.GetRaw().Properties.Tabs)
	})
}

func TestA3PageSize(t *testing.T) {
	t.Run("Should create document with A3 page size", func(t *testing.T) {
		doc := docxtpl.NewWithOptions(docxtpl.PageSizeA3)
		assert.NotNil(t, doc)
		doc.AddParagraph("A3 document")
	})
}

func TestShapes(t *testing.T) {
	t.Run("Should add anchor shape", func(t *testing.T) {
		doc := docxtpl.New()
		para := doc.AddParagraph("")
		run := para.AddAnchorShape(docxtpl.ShapeOptions{
			Width:     914400,
			Height:    914400,
			Preset:    docxtpl.ShapeRectangle,
			LineColor: "000000",
		})
		assert.NotNil(t, run)
	})

	t.Run("Should add inline shape", func(t *testing.T) {
		doc := docxtpl.New()
		para := doc.AddParagraph("")
		run := para.AddInlineShape(docxtpl.ShapeOptions{
			Width:  457200,
			Height: 457200,
			Preset: docxtpl.ShapeEllipse,
		})
		assert.NotNil(t, run)
	})
}

func TestDocumentMerging(t *testing.T) {
	t.Run("Should append document", func(t *testing.T) {
		doc1 := docxtpl.New()
		doc1.AddParagraph("Document 1")

		doc2 := docxtpl.New()
		doc2.AddParagraph("Document 2")

		initialCount := doc1.CountParagraphs()
		doc1.AppendDocument(doc2)

		// Should have more paragraphs after append
		assert.Greater(t, doc1.CountParagraphs(), initialCount)
	})
}

func TestTextExtraction(t *testing.T) {
	t.Run("Should extract text from document", func(t *testing.T) {
		doc := docxtpl.New()
		doc.AddParagraph("Hello World")
		doc.AddParagraph("Second paragraph")

		text := doc.GetText()
		assert.Contains(t, text, "Hello World")
		assert.Contains(t, text, "Second paragraph")
	})

	t.Run("Should get paragraph texts", func(t *testing.T) {
		doc := docxtpl.New()
		doc.AddParagraph("Para 1")
		doc.AddParagraph("Para 2")
		doc.AddParagraph("Para 3")

		texts := doc.GetParagraphTexts()
		assert.Len(t, texts, 3)
	})

	t.Run("Should get text from paragraph", func(t *testing.T) {
		doc := docxtpl.New()
		para := doc.AddParagraph("Hello")
		para.AddText(" World")

		text := para.GetText()
		assert.Contains(t, text, "Hello")
		assert.Contains(t, text, "World")
	})
}

func TestTextSearch(t *testing.T) {
	t.Run("Should find text in document", func(t *testing.T) {
		doc := docxtpl.New()
		doc.AddParagraph("The quick brown fox")
		doc.AddParagraph("jumps over the lazy dog")

		assert.True(t, doc.HasText("quick"))
		assert.True(t, doc.HasText("lazy"))
		assert.False(t, doc.HasText("cat"))
	})

	t.Run("Should find text with regex", func(t *testing.T) {
		doc := docxtpl.New()
		doc.AddParagraph("Order #12345")
		doc.AddParagraph("Customer: John")

		assert.True(t, doc.HasTextMatch(`#\d+`))
		assert.False(t, doc.HasTextMatch(`#[A-Z]+`))
	})

	t.Run("Should find matching paragraphs", func(t *testing.T) {
		doc := docxtpl.New()
		doc.AddParagraph("Error: Something went wrong")
		doc.AddParagraph("Info: All good")
		doc.AddParagraph("Error: Another problem")

		results := doc.FindText("Error")
		assert.Len(t, results, 2)
	})
}

func TestDocumentCounting(t *testing.T) {
	t.Run("Should count paragraphs", func(t *testing.T) {
		doc := docxtpl.New()
		doc.AddParagraph("Para 1")
		doc.AddParagraph("Para 2")

		assert.Equal(t, 2, doc.CountParagraphs())
	})

	t.Run("Should count tables", func(t *testing.T) {
		doc := docxtpl.New()
		doc.AddTable(2, 2)
		doc.AddTable(3, 3)

		assert.Equal(t, 2, doc.CountTables())
	})
}

func TestDocumentSplitting(t *testing.T) {
	t.Run("Should split at heading", func(t *testing.T) {
		doc := docxtpl.New()
		doc.AddHeading("Chapter 1", 1)
		doc.AddParagraph("Content 1")
		doc.AddHeading("Chapter 2", 1)
		doc.AddParagraph("Content 2")

		parts := doc.SplitAtHeading(1)
		// First split includes content before first heading
		assert.GreaterOrEqual(t, len(parts), 1)
	})

	t.Run("Should split at text", func(t *testing.T) {
		doc := docxtpl.New()
		doc.AddParagraph("Section A")
		doc.AddParagraph("---")
		doc.AddParagraph("Section B")

		parts := doc.SplitAtText("---")
		assert.GreaterOrEqual(t, len(parts), 1)
	})
}

func TestTextReplacement(t *testing.T) {
	t.Run("Should replace text", func(t *testing.T) {
		doc := docxtpl.New()
		doc.AddParagraph("Hello PLACEHOLDER World")

		doc.ReplaceText("PLACEHOLDER", "Beautiful")

		text := doc.GetText()
		assert.Contains(t, text, "Beautiful")
		assert.NotContains(t, text, "PLACEHOLDER")
	})

	t.Run("Should replace with regex", func(t *testing.T) {
		doc := docxtpl.New()
		doc.AddParagraph("Date: 2024-01-15")

		doc.ReplaceTextRegex(`\d{4}-\d{2}-\d{2}`, "REDACTED")

		text := doc.GetText()
		assert.Contains(t, text, "REDACTED")
	})
}

func TestParagraphOperations(t *testing.T) {
	t.Run("Should merge runs", func(t *testing.T) {
		doc := docxtpl.New()
		para := doc.AddParagraph("Hello")
		para.AddText(" ")
		para.AddText("World")

		para.MergeRuns()
		assert.NotNil(t, para)
	})

	t.Run("Should drop shapes from paragraph", func(t *testing.T) {
		doc := docxtpl.New()
		para := doc.AddParagraph("Text with shape")
		para.AddAnchorShape(docxtpl.ShapeOptions{
			Width: 914400, Height: 914400,
			Preset: docxtpl.ShapeRectangle,
		})

		para.DropShapes()
		assert.NotNil(t, para)
	})
}

func TestDocumentCleanup(t *testing.T) {
	t.Run("Should clean document", func(t *testing.T) {
		doc := docxtpl.New()
		doc.AddParagraph("Test content")

		// Should not panic
		doc.CleanDocument()
		assert.NotNil(t, doc)
	})

	t.Run("Should merge all runs in document", func(t *testing.T) {
		doc := docxtpl.New()
		para := doc.AddParagraph("A")
		para.AddText("B")
		para.AddText("C")

		doc.MergeAllRuns()
		assert.NotNil(t, doc)
	})
}
