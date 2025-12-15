package docxtpl

import (
	"encoding/xml"
	"strings"
	"time"
	"unicode"

	"github.com/abdokhaire/go-docxgen/internal/docx"
	"github.com/abdokhaire/go-docxgen/internal/headerfooter"
)

// =============================================================================
// Document Properties - Core Metadata
// =============================================================================

// DocumentProperties contains the core document metadata.
type DocumentProperties struct {
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
	ContentStatus  string // e.g., "Draft", "Final"
}

// coreProperties represents the XML structure of docProps/core.xml
type coreProperties struct {
	XMLName        xml.Name `xml:"cp:coreProperties"`
	XMLNScp        string   `xml:"xmlns:cp,attr"`
	XMLNSdc        string   `xml:"xmlns:dc,attr"`
	XMLNSdcterms   string   `xml:"xmlns:dcterms,attr"`
	XMLNSdcmitype  string   `xml:"xmlns:dcmitype,attr"`
	XMLNSxsi       string   `xml:"xmlns:xsi,attr"`
	Title          string   `xml:"dc:title,omitempty"`
	Subject        string   `xml:"dc:subject,omitempty"`
	Creator        string   `xml:"dc:creator,omitempty"`
	Keywords       string   `xml:"cp:keywords,omitempty"`
	Description    string   `xml:"dc:description,omitempty"`
	LastModifiedBy string   `xml:"cp:lastModifiedBy,omitempty"`
	Revision       string   `xml:"cp:revision,omitempty"`
	Created        *dcTime  `xml:"dcterms:created,omitempty"`
	Modified       *dcTime  `xml:"dcterms:modified,omitempty"`
	Category       string   `xml:"cp:category,omitempty"`
	ContentStatus  string   `xml:"cp:contentStatus,omitempty"`
}

// dcTime represents a date/time in Dublin Core format
type dcTime struct {
	Type  string `xml:"xsi:type,attr"`
	Value string `xml:",chardata"`
}

const (
	cpNamespace       = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
	dcNamespace       = "http://purl.org/dc/elements/1.1/"
	dctermsNamespace  = "http://purl.org/dc/terms/"
	dcmitypeNamespace = "http://purl.org/dc/dcmitype/"
	xsiNamespace      = "http://www.w3.org/2001/XMLSchema-instance"
	w3cDateFormat     = "2006-01-02T15:04:05Z"
)

// GetProperties returns the document properties.
// For documents created from scratch, this returns default/empty properties.
// For parsed documents, this extracts properties from docProps/core.xml.
func (d *DocxTmpl) GetProperties() *DocumentProperties {
	// If we have in-memory properties, return them
	if d.properties != nil {
		return d.properties
	}

	props := &DocumentProperties{}

	// Try to find and parse core.xml from processable files
	for _, pf := range d.processableFiles {
		if strings.HasSuffix(pf.Name, "core.xml") {
			var core coreProperties
			if err := xml.Unmarshal([]byte(pf.Content), &core); err == nil {
				props.Title = core.Title
				props.Subject = core.Subject
				props.Creator = core.Creator
				props.Keywords = core.Keywords
				props.Description = core.Description
				props.LastModifiedBy = core.LastModifiedBy
				props.Revision = core.Revision
				props.Category = core.Category
				props.ContentStatus = core.ContentStatus

				if core.Created != nil {
					if t, err := time.Parse(w3cDateFormat, core.Created.Value); err == nil {
						props.Created = t
					}
				}
				if core.Modified != nil {
					if t, err := time.Parse(w3cDateFormat, core.Modified.Value); err == nil {
						props.Modified = t
					}
				}
			}
			d.properties = props
			return props
		}
	}

	return props
}

// SetProperties updates the document properties.
// The properties will be written when the document is saved.
func (d *DocxTmpl) SetProperties(props *DocumentProperties) {
	// Store in memory
	d.properties = props

	// Also serialize to processable files for saving
	core := &coreProperties{
		XMLNScp:        cpNamespace,
		XMLNSdc:        dcNamespace,
		XMLNSdcterms:   dctermsNamespace,
		XMLNSdcmitype:  dcmitypeNamespace,
		XMLNSxsi:       xsiNamespace,
		Title:          props.Title,
		Subject:        props.Subject,
		Creator:        props.Creator,
		Keywords:       props.Keywords,
		Description:    props.Description,
		LastModifiedBy: props.LastModifiedBy,
		Revision:       props.Revision,
		Category:       props.Category,
		ContentStatus:  props.ContentStatus,
	}

	if !props.Created.IsZero() {
		core.Created = &dcTime{
			Type:  "dcterms:W3CDTF",
			Value: props.Created.Format(w3cDateFormat),
		}
	}
	if !props.Modified.IsZero() {
		core.Modified = &dcTime{
			Type:  "dcterms:W3CDTF",
			Value: props.Modified.Format(w3cDateFormat),
		}
	}

	// Marshal to XML
	output, err := xml.MarshalIndent(core, "", "  ")
	if err != nil {
		return
	}
	content := xml.Header + string(output)

	// Update or add to processable files
	found := false
	for i, pf := range d.processableFiles {
		if strings.HasSuffix(pf.Name, "core.xml") {
			d.processableFiles[i].Content = content
			found = true
			break
		}
	}

	if !found {
		d.processableFiles = append(d.processableFiles, headerfooter.DocxFile{
			Name:    "docProps/core.xml",
			Content: content,
		})
	}
}

// SetTitle sets the document title.
func (d *DocxTmpl) SetTitle(title string) *DocxTmpl {
	props := d.GetProperties()
	props.Title = title
	d.SetProperties(props)
	return d
}

// SetAuthor sets the document author (creator).
func (d *DocxTmpl) SetAuthor(author string) *DocxTmpl {
	props := d.GetProperties()
	props.Creator = author
	d.SetProperties(props)
	return d
}

// SetSubject sets the document subject.
func (d *DocxTmpl) SetSubject(subject string) *DocxTmpl {
	props := d.GetProperties()
	props.Subject = subject
	d.SetProperties(props)
	return d
}

// SetKeywords sets the document keywords.
func (d *DocxTmpl) SetKeywords(keywords string) *DocxTmpl {
	props := d.GetProperties()
	props.Keywords = keywords
	d.SetProperties(props)
	return d
}

// SetDescription sets the document description/comments.
func (d *DocxTmpl) SetDescription(description string) *DocxTmpl {
	props := d.GetProperties()
	props.Description = description
	d.SetProperties(props)
	return d
}

// SetCategory sets the document category.
func (d *DocxTmpl) SetCategory(category string) *DocxTmpl {
	props := d.GetProperties()
	props.Category = category
	d.SetProperties(props)
	return d
}

// SetContentStatus sets the document status (e.g., "Draft", "Final", "Approved").
func (d *DocxTmpl) SetContentStatus(status string) *DocxTmpl {
	props := d.GetProperties()
	props.ContentStatus = status
	d.SetProperties(props)
	return d
}

// =============================================================================
// Extended Metadata and Statistics
// =============================================================================

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

// =============================================================================
// Document Outline and Structure
// =============================================================================

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

// =============================================================================
// Hyperlinks and Styles
// =============================================================================

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

// =============================================================================
// Sections and Page Layout
// =============================================================================

// Orientation represents page orientation
type Orientation string

const (
	OrientationPortrait  Orientation = "portrait"
	OrientationLandscape Orientation = "landscape"
)

// SectionBreakType represents the type of section break
type SectionBreakType string

const (
	SectionBreakNextPage   SectionBreakType = "nextPage"
	SectionBreakContinuous SectionBreakType = "continuous"
	SectionBreakEvenPage   SectionBreakType = "evenPage"
	SectionBreakOddPage    SectionBreakType = "oddPage"
)

// Margins represents page margins in inches
type Margins struct {
	Top    float64
	Right  float64
	Bottom float64
	Left   float64
}

// DefaultMargins returns the default Word margins (1 inch all around)
func DefaultMargins() Margins {
	return Margins{
		Top:    1.0,
		Right:  1.0,
		Bottom: 1.0,
		Left:   1.0,
	}
}

// NarrowMargins returns narrow margins (0.5 inch all around)
func NarrowMargins() Margins {
	return Margins{
		Top:    0.5,
		Right:  0.5,
		Bottom: 0.5,
		Left:   0.5,
	}
}

// WideMargins returns wide margins (1 inch top/bottom, 2 inch left/right)
func WideMargins() Margins {
	return Margins{
		Top:    1.0,
		Right:  2.0,
		Bottom: 1.0,
		Left:   2.0,
	}
}

// AddSectionBreak adds a visual section break to the document.
// For SectionBreakNextPage, this creates a page break.
// For other types, it creates appropriate spacing.
//
//	doc.AddSectionBreak(SectionBreakNextPage)
func (d *DocxTmpl) AddSectionBreak(breakType SectionBreakType) *DocxTmpl {
	switch breakType {
	case SectionBreakNextPage:
		d.AddPageBreak()
	case SectionBreakContinuous:
		// Continuous sections don't need a visual break
		d.AddEmptyParagraph()
	case SectionBreakEvenPage, SectionBreakOddPage:
		// These are approximated with page breaks
		d.AddPageBreak()
	}
	return d
}

// AddSection adds a new section with a page break.
// This is a convenience method for common use cases.
//
//	doc.AddSection()
func (d *DocxTmpl) AddSection() *DocxTmpl {
	return d.AddSectionBreak(SectionBreakNextPage)
}

// Custom page size constants (in inches)
const (
	PageWidthA4      = 8.27
	PageHeightA4     = 11.69
	PageWidthA3      = 11.69
	PageHeightA3     = 16.54
	PageWidthLetter  = 8.5
	PageHeightLetter = 11.0
	PageWidthLegal   = 8.5
	PageHeightLegal  = 14.0
)

// EstimatePageCount estimates the number of pages in the document.
// This is a rough estimate based on paragraph count and assumes:
// - Single spacing, 12pt font
// - About 50 lines per page for Letter size
// - Tables count as 3 paragraphs
//
//	pages := doc.EstimatePageCount()
func (d *DocxTmpl) EstimatePageCount() int {
	stats := d.GetStats()

	// Rough estimate: 50 lines per page, average 2 lines per paragraph
	linesPerPage := 50.0
	avgLinesPerPara := 2.0
	linesPerTable := 6.0 // Tables take more space

	totalLines := float64(stats.ParagraphCount)*avgLinesPerPara +
		float64(stats.TableCount)*linesPerTable

	pages := int(totalLines / linesPerPage)
	if pages < 1 {
		pages = 1
	}
	return pages
}

// =============================================================================
// Helper Functions
// =============================================================================

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
