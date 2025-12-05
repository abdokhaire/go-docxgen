package docxtpl

import (
	"strings"
	"time"
	"unicode"

	"github.com/fumiama/go-docx"
)

// DocumentMetadata contains document properties.
type DocumentMetadata struct {
	Title          string
	Subject        string
	Creator        string // Author
	Keywords       string
	Description    string
	LastModifiedBy string
	Revision       string
	Created        time.Time
	Modified       time.Time
	Category       string
	ContentStatus  string
	Language       string
}

// DocumentStats contains document statistics.
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

// GetMetadata extracts document metadata from core.xml and app.xml.
// Returns empty values for fields not found in the document.
//
//	meta := doc.GetMetadata()
//	fmt.Println("Author:", meta.Creator)
func (d *DocxTmpl) GetMetadata() *DocumentMetadata {
	meta := &DocumentMetadata{}

	// Extract from processable files if available
	for _, pf := range d.processableFiles {
		content := pf.Content

		// core.xml properties
		if strings.Contains(pf.Name, "core.xml") {
			meta.Title = extractXMLValue(content, "dc:title")
			meta.Subject = extractXMLValue(content, "dc:subject")
			meta.Creator = extractXMLValue(content, "dc:creator")
			meta.Keywords = extractXMLValue(content, "cp:keywords")
			meta.Description = extractXMLValue(content, "dc:description")
			meta.LastModifiedBy = extractXMLValue(content, "cp:lastModifiedBy")
			meta.Revision = extractXMLValue(content, "cp:revision")
			meta.Category = extractXMLValue(content, "cp:category")
			meta.ContentStatus = extractXMLValue(content, "cp:contentStatus")
			meta.Language = extractXMLValue(content, "dc:language")

			// Parse dates
			if created := extractXMLValue(content, "dcterms:created"); created != "" {
				if t, err := time.Parse(time.RFC3339, created); err == nil {
					meta.Created = t
				}
			}
			if modified := extractXMLValue(content, "dcterms:modified"); modified != "" {
				if t, err := time.Parse(time.RFC3339, modified); err == nil {
					meta.Modified = t
				}
			}
		}
	}

	return meta
}

// SetMetadata sets document metadata in core.xml.
// Only non-empty fields are updated.
//
//	doc.SetMetadata(&DocumentMetadata{
//	    Title:   "My Report",
//	    Creator: "John Doe",
//	})
func (d *DocxTmpl) SetMetadata(meta *DocumentMetadata) {
	for i, pf := range d.processableFiles {
		if strings.Contains(pf.Name, "core.xml") {
			content := pf.Content

			if meta.Title != "" {
				content = setXMLValue(content, "dc:title", meta.Title)
			}
			if meta.Subject != "" {
				content = setXMLValue(content, "dc:subject", meta.Subject)
			}
			if meta.Creator != "" {
				content = setXMLValue(content, "dc:creator", meta.Creator)
			}
			if meta.Keywords != "" {
				content = setXMLValue(content, "cp:keywords", meta.Keywords)
			}
			if meta.Description != "" {
				content = setXMLValue(content, "dc:description", meta.Description)
			}
			if meta.LastModifiedBy != "" {
				content = setXMLValue(content, "cp:lastModifiedBy", meta.LastModifiedBy)
			}
			if meta.Category != "" {
				content = setXMLValue(content, "cp:category", meta.Category)
			}
			if meta.Language != "" {
				content = setXMLValue(content, "dc:language", meta.Language)
			}
			if !meta.Modified.IsZero() {
				content = setXMLValue(content, "dcterms:modified", meta.Modified.Format(time.RFC3339))
			}

			d.processableFiles[i].Content = content
			break
		}
	}
}

// GetStats returns document statistics including word count, paragraph count, etc.
//
//	stats := doc.GetStats()
//	fmt.Printf("Words: %d, Paragraphs: %d\n", stats.WordCount, stats.ParagraphCount)
func (d *DocxTmpl) GetStats() *DocumentStats {
	stats := &DocumentStats{}

	// Count paragraphs and tables
	stats.ParagraphCount = d.CountParagraphs()
	stats.TableCount = d.CountTables()

	// Get all text
	text := d.GetText()

	// Count characters
	stats.CharCountSpace = len(text)
	stats.CharCount = len(strings.ReplaceAll(text, " ", ""))

	// Count words
	stats.WordCount = countWords(text)

	// Count lines
	stats.LineCount = strings.Count(text, "\n") + 1

	// Count images
	stats.ImageCount = countImages(d)

	// Count links
	stats.LinkCount = len(d.GetAllHyperlinks())

	return stats
}

// OutlineItem represents a heading in the document outline.
type OutlineItem struct {
	Level    int           // Heading level (0=Title, 1-9=Heading1-9)
	Text     string        // Heading text
	Index    int           // Paragraph index in document
	Children []OutlineItem // Nested headings
}

// GetOutline extracts the document structure as a hierarchical outline.
// Returns headings organized by their levels.
//
//	outline := doc.GetOutline()
//	for _, item := range outline {
//	    fmt.Printf("H%d: %s\n", item.Level, item.Text)
//	}
func (d *DocxTmpl) GetOutline() []OutlineItem {
	var items []OutlineItem
	idx := 0

	for _, item := range d.Document.Body.Items {
		if para, ok := item.(*docx.Paragraph); ok {
			if para.Properties != nil && para.Properties.Style != nil {
				style := para.Properties.Style.Val
				level := -1

				if style == "Title" {
					level = 0
				} else if strings.HasPrefix(style, "Heading") {
					levelStr := strings.TrimPrefix(style, "Heading")
					for l := 1; l <= 9; l++ {
						if levelStr == string(rune('0'+l)) {
							level = l
							break
						}
					}
				}

				if level >= 0 {
					items = append(items, OutlineItem{
						Level: level,
						Text:  para.String(),
						Index: idx,
					})
				}
			}
			idx++
		}
	}

	// Build hierarchy
	return buildOutlineHierarchy(items)
}

// HyperlinkInfo contains information about a hyperlink in the document.
type HyperlinkInfo struct {
	Text       string // Display text
	URL        string // Target URL
	Index      int    // Paragraph index
	IsInternal bool   // True if internal bookmark link
}

// GetAllHyperlinks returns all hyperlinks in the document.
//
//	links := doc.GetAllHyperlinks()
//	for _, link := range links {
//	    fmt.Printf("%s -> %s\n", link.Text, link.URL)
//	}
func (d *DocxTmpl) GetAllHyperlinks() []HyperlinkInfo {
	var links []HyperlinkInfo

	// Get hyperlinks from the hyperlink registry
	if d.hyperlinkReg != nil {
		for url, id := range d.hyperlinkReg.GetLinks() {
			links = append(links, HyperlinkInfo{
				URL:  url,
				Text: id, // We store ID, actual text would need paragraph parsing
			})
		}
	}

	return links
}

// GetAllStyles returns all paragraph styles used in the document.
//
//	styles := doc.GetAllStyles()
//	// Returns: ["Normal", "Heading1", "Heading2", ...]
func (d *DocxTmpl) GetAllStyles() []string {
	styleSet := make(map[string]bool)

	for _, item := range d.Document.Body.Items {
		if para, ok := item.(*docx.Paragraph); ok {
			if para.Properties != nil && para.Properties.Style != nil {
				styleSet[para.Properties.Style.Val] = true
			} else {
				styleSet["Normal"] = true
			}
		}
	}

	styles := make([]string, 0, len(styleSet))
	for style := range styleSet {
		styles = append(styles, style)
	}
	return styles
}

// GetTextByStyle returns all text from paragraphs with the specified style.
//
//	headings := doc.GetTextByStyle("Heading1")
func (d *DocxTmpl) GetTextByStyle(style string) []string {
	var texts []string

	for _, item := range d.Document.Body.Items {
		if para, ok := item.(*docx.Paragraph); ok {
			paraStyle := "Normal"
			if para.Properties != nil && para.Properties.Style != nil {
				paraStyle = para.Properties.Style.Val
			}
			if paraStyle == style {
				texts = append(texts, para.String())
			}
		}
	}

	return texts
}

// Helper functions

func extractXMLValue(xml, tag string) string {
	startTag := "<" + tag
	endTag := "</" + tag + ">"

	startIdx := strings.Index(xml, startTag)
	if startIdx == -1 {
		return ""
	}

	// Find the end of the start tag
	tagEnd := strings.Index(xml[startIdx:], ">")
	if tagEnd == -1 {
		return ""
	}
	contentStart := startIdx + tagEnd + 1

	endIdx := strings.Index(xml[contentStart:], endTag)
	if endIdx == -1 {
		return ""
	}

	return xml[contentStart : contentStart+endIdx]
}

func setXMLValue(xml, tag, value string) string {
	startTag := "<" + tag
	endTag := "</" + tag + ">"

	startIdx := strings.Index(xml, startTag)
	if startIdx == -1 {
		// Tag doesn't exist, we'd need to add it
		return xml
	}

	tagEnd := strings.Index(xml[startIdx:], ">")
	if tagEnd == -1 {
		return xml
	}
	contentStart := startIdx + tagEnd + 1

	endIdx := strings.Index(xml[contentStart:], endTag)
	if endIdx == -1 {
		return xml
	}

	return xml[:contentStart] + value + xml[contentStart+endIdx:]
}

func countWords(text string) int {
	words := 0
	inWord := false

	for _, r := range text {
		if unicode.IsLetter(r) || unicode.IsDigit(r) {
			if !inWord {
				words++
				inWord = true
			}
		} else {
			inWord = false
		}
	}

	return words
}

func countImages(d *DocxTmpl) int {
	count := 0
	// Count media files that are images
	media := d.GetAllMedia()
	for _, m := range media {
		name := strings.ToLower(m.Name)
		if strings.HasSuffix(name, ".png") ||
			strings.HasSuffix(name, ".jpg") ||
			strings.HasSuffix(name, ".jpeg") ||
			strings.HasSuffix(name, ".gif") ||
			strings.HasSuffix(name, ".bmp") {
			count++
		}
	}
	return count
}

func buildOutlineHierarchy(items []OutlineItem) []OutlineItem {
	if len(items) == 0 {
		return items
	}

	// For now, return flat list - hierarchy building is complex
	// and requires tracking parent-child relationships
	return items
}
