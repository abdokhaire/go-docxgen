package docxtpl

import (
	"regexp"
	"strings"

	"github.com/fumiama/go-docx"
)

// AppendDocument appends all contents from another document to this document.
// This is useful for merging multiple documents into one.
//
//	doc1, _ := docxtpl.ParseFromFilename("chapter1.docx")
//	doc2, _ := docxtpl.ParseFromFilename("chapter2.docx")
//	doc1.AppendDocument(doc2)
//	doc1.SaveToFile("combined.docx")
func (d *DocxTmpl) AppendDocument(other *DocxTmpl) {
	d.Docx.AppendFile(other.Docx)
}

// SplitRule defines a function that determines where to split a document.
// Return true to split before the paragraph that matches.
type SplitRule func(text string, isHeading bool, headingLevel int) bool

// SplitAt splits the document into multiple documents based on a rule.
// The rule function receives paragraph text, whether it's a heading, and heading level.
// Returns a slice of new DocxTmpl documents.
//
//	// Split at each Heading 1
//	docs := doc.SplitAt(func(text string, isHeading bool, level int) bool {
//	    return isHeading && level == 1
//	})
func (d *DocxTmpl) SplitAt(rule SplitRule) []*DocxTmpl {
	// Create a wrapper rule that checks our conditions
	splitter := func(p *docx.Paragraph) bool {
		text := p.String()
		isHeading := false
		headingLevel := 0

		// Check if paragraph has a heading style
		if p.Properties != nil && p.Properties.Style != nil {
			style := p.Properties.Style.Val
			if strings.HasPrefix(style, "Heading") {
				isHeading = true
				// Extract level from "Heading1", "Heading2", etc.
				if len(style) > 7 {
					level := style[7:]
					switch level {
					case "1":
						headingLevel = 1
					case "2":
						headingLevel = 2
					case "3":
						headingLevel = 3
					case "4":
						headingLevel = 4
					case "5":
						headingLevel = 5
					case "6":
						headingLevel = 6
					case "7":
						headingLevel = 7
					case "8":
						headingLevel = 8
					case "9":
						headingLevel = 9
					}
				}
			} else if style == "Title" {
				isHeading = true
				headingLevel = 0
			}
		}

		return rule(text, isHeading, headingLevel)
	}

	// Use go-docx's SplitByParagraph
	rawDocs := d.Docx.SplitByParagraph(splitter)

	// Wrap each result
	result := make([]*DocxTmpl, len(rawDocs))
	for i, rawDoc := range rawDocs {
		result[i] = &DocxTmpl{
			Docx:             rawDoc,
			funcMap:          d.funcMap,
			contentTypes:     d.contentTypes,
			processableFiles: nil, // Split docs don't have processable files
			hyperlinkReg:     d.hyperlinkReg,
		}
	}

	return result
}

// SplitAtHeading splits the document at each heading of the specified level.
// Level 0 splits at Title, level 1-9 splits at Heading1-Heading9.
//
//	chapters := doc.SplitAtHeading(1) // Split at each Heading 1
func (d *DocxTmpl) SplitAtHeading(level int) []*DocxTmpl {
	return d.SplitAt(func(text string, isHeading bool, headingLevel int) bool {
		return isHeading && headingLevel == level
	})
}

// SplitAtText splits the document at paragraphs containing the specified text.
//
//	parts := doc.SplitAtText("---") // Split at paragraphs containing "---"
func (d *DocxTmpl) SplitAtText(text string) []*DocxTmpl {
	return d.SplitAt(func(paraText string, isHeading bool, level int) bool {
		return strings.Contains(paraText, text)
	})
}

// SplitAtRegex splits the document at paragraphs matching the regex pattern.
//
//	parts := doc.SplitAtRegex(`^Chapter \d+`) // Split at "Chapter 1", "Chapter 2", etc.
func (d *DocxTmpl) SplitAtRegex(pattern string) []*DocxTmpl {
	re := regexp.MustCompile(pattern)
	return d.SplitAt(func(text string, isHeading bool, level int) bool {
		return re.MatchString(text)
	})
}

// GetText extracts all plain text from the document body.
// Paragraphs are separated by newlines.
//
//	text := doc.GetText()
func (d *DocxTmpl) GetText() string {
	var texts []string

	for _, item := range d.Document.Body.Items {
		switch v := item.(type) {
		case *docx.Paragraph:
			texts = append(texts, v.String())
		case *docx.Table:
			texts = append(texts, v.String())
		}
	}

	return strings.Join(texts, "\n")
}

// GetParagraphTexts returns the text content of each paragraph as a slice.
//
//	paragraphs := doc.GetParagraphTexts()
func (d *DocxTmpl) GetParagraphTexts() []string {
	var texts []string

	for _, item := range d.Document.Body.Items {
		if p, ok := item.(*docx.Paragraph); ok {
			texts = append(texts, p.String())
		}
	}

	return texts
}

// Media represents an embedded media file in the document.
type Media struct {
	Name string // Filename (e.g., "image1.png")
	Data []byte // File contents
}

// GetMedia retrieves an embedded media file by name.
// Returns nil if the media is not found.
//
//	media := doc.GetMedia("image1.png")
//	if media != nil {
//	    os.WriteFile("extracted.png", media.Data, 0644)
//	}
func (d *DocxTmpl) GetMedia(name string) *Media {
	m := d.Docx.Media(name)
	if m == nil {
		return nil
	}
	return &Media{
		Name: m.Name,
		Data: m.Data,
	}
}

// GetAllMedia returns all embedded media files in the document.
//
//	mediaFiles := doc.GetAllMedia()
//	for _, m := range mediaFiles {
//	    os.WriteFile(m.Name, m.Data, 0644)
//	}
func (d *DocxTmpl) GetAllMedia() []*Media {
	var result []*Media

	// Iterate through relationships to find media
	d.Docx.RangeRelationships(func(rel *docx.Relationship) error {
		if strings.HasPrefix(rel.Target, "media/") {
			name := strings.TrimPrefix(rel.Target, "media/")
			m := d.Docx.Media(name)
			if m != nil {
				result = append(result, &Media{
					Name: m.Name,
					Data: m.Data,
				})
			}
		}
		return nil
	})

	return result
}

// CountParagraphs returns the number of paragraphs in the document.
func (d *DocxTmpl) CountParagraphs() int {
	count := 0
	for _, item := range d.Document.Body.Items {
		if _, ok := item.(*docx.Paragraph); ok {
			count++
		}
	}
	return count
}

// CountTables returns the number of tables in the document.
func (d *DocxTmpl) CountTables() int {
	count := 0
	for _, item := range d.Document.Body.Items {
		if _, ok := item.(*docx.Table); ok {
			count++
		}
	}
	return count
}

// HasText checks if the document contains the specified text.
//
//	if doc.HasText("confidential") {
//	    // Handle confidential document
//	}
func (d *DocxTmpl) HasText(text string) bool {
	return strings.Contains(d.GetText(), text)
}

// HasTextMatch checks if the document contains text matching the regex pattern.
//
//	if doc.HasTextMatch(`\d{3}-\d{2}-\d{4}`) {
//	    // Document contains SSN pattern
//	}
func (d *DocxTmpl) HasTextMatch(pattern string) bool {
	re := regexp.MustCompile(pattern)
	return re.MatchString(d.GetText())
}

// FindText returns all paragraphs containing the specified text.
func (d *DocxTmpl) FindText(text string) []string {
	var results []string
	for _, item := range d.Document.Body.Items {
		if p, ok := item.(*docx.Paragraph); ok {
			paraText := p.String()
			if strings.Contains(paraText, text) {
				results = append(results, paraText)
			}
		}
	}
	return results
}

// FindTextMatch returns all paragraphs matching the regex pattern.
func (d *DocxTmpl) FindTextMatch(pattern string) []string {
	re := regexp.MustCompile(pattern)
	var results []string
	for _, item := range d.Document.Body.Items {
		if p, ok := item.(*docx.Paragraph); ok {
			paraText := p.String()
			if re.MatchString(paraText) {
				results = append(results, paraText)
			}
		}
	}
	return results
}

// DropAllDrawings removes all shapes, canvases, and groups from the entire document.
// This is useful for extracting just the text content.
func (d *DocxTmpl) DropAllDrawings() {
	d.Document.Body.DropDrawingOf("ShapeAndCanvasAndGroup")
}

// DropShapes removes all shapes from the entire document.
func (d *DocxTmpl) DropShapes() {
	d.Document.Body.DropDrawingOf("Shape")
}

// DropCanvas removes all canvas elements from the entire document.
func (d *DocxTmpl) DropCanvas() {
	d.Document.Body.DropDrawingOf("Canvas")
}

// DropGroups removes all group elements from the entire document.
func (d *DocxTmpl) DropGroups() {
	d.Document.Body.DropDrawingOf("Group")
}

// DropEmptyPictures removes all nil/empty picture references from the document.
func (d *DocxTmpl) DropEmptyPictures() {
	d.Document.Body.DropDrawingOf("NilPicture")
}

// KeepBodyElements keeps only specified element types in the document body.
// Valid names: "*docx.Paragraph", "*docx.Table"
//
//	doc.KeepBodyElements("*docx.Paragraph") // Keep only paragraphs, remove tables
func (d *DocxTmpl) KeepBodyElements(names ...string) {
	d.Document.Body.KeepElements(names...)
}

// MergeAllRuns merges contiguous runs with same formatting in all paragraphs.
// This reduces document complexity and can decrease file size.
func (d *DocxTmpl) MergeAllRuns() {
	for _, item := range d.Document.Body.Items {
		if p, ok := item.(*docx.Paragraph); ok {
			*p = p.MergeText(docx.MergeSamePropRuns)
		}
	}
}

// CleanDocument performs common cleanup operations:
// - Merges runs with same formatting
// - Removes empty picture references
// This is useful before saving to reduce file size.
func (d *DocxTmpl) CleanDocument() {
	d.MergeAllRuns()
	d.DropEmptyPictures()
}

// ReplaceText replaces all occurrences of old text with new text in the document.
// This is a simple text replacement that works on the rendered document.
// For template-based replacement, use Render() instead.
//
//	doc.ReplaceText("COMPANY_NAME", "Acme Corp")
func (d *DocxTmpl) ReplaceText(oldText, newText string) {
	for _, item := range d.Document.Body.Items {
		if p, ok := item.(*docx.Paragraph); ok {
			for _, child := range p.Children {
				if r, ok := child.(*docx.Run); ok {
					for _, rc := range r.Children {
						if t, ok := rc.(*docx.Text); ok {
							t.Text = strings.ReplaceAll(t.Text, oldText, newText)
						}
					}
				}
			}
		}
	}
}

// ReplaceTextRegex replaces text matching the pattern with replacement.
// The replacement can include $1, $2, etc. for captured groups.
//
//	doc.ReplaceTextRegex(`\bfoo\b`, "bar")
func (d *DocxTmpl) ReplaceTextRegex(pattern, replacement string) {
	re := regexp.MustCompile(pattern)
	for _, item := range d.Document.Body.Items {
		if p, ok := item.(*docx.Paragraph); ok {
			for _, child := range p.Children {
				if r, ok := child.(*docx.Run); ok {
					for _, rc := range r.Children {
						if t, ok := rc.(*docx.Text); ok {
							t.Text = re.ReplaceAllString(t.Text, replacement)
						}
					}
				}
			}
		}
	}
}
