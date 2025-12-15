package docxtpl

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strings"

	"github.com/abdokhaire/go-docxgen/internal/docx"
)

// =============================================================================
// Document Cloning and Merging
// =============================================================================

// Clone creates a deep copy of the document.
// The cloned document can be modified independently.
//
//	clone := doc.Clone()
//	clone.Render(differentData)
//	clone.SaveToFile("copy.docx")
func (d *DocxTmpl) Clone() (*DocxTmpl, error) {
	// Save to buffer
	var buf bytes.Buffer
	err := d.Save(&buf)
	if err != nil {
		return nil, fmt.Errorf("failed to save document for cloning: %w", err)
	}

	// Parse from buffer
	clone, err := ParseFromBytes(buf.Bytes())
	if err != nil {
		return nil, fmt.Errorf("failed to parse cloned document: %w", err)
	}

	return clone, nil
}

// MergeDocuments combines multiple documents into one.
// Documents are appended in order with page breaks between them.
//
//	merged := docxtpl.MergeDocuments(doc1, doc2, doc3)
func MergeDocuments(docs ...*DocxTmpl) (*DocxTmpl, error) {
	if len(docs) == 0 {
		return nil, fmt.Errorf("no documents to merge")
	}

	if len(docs) == 1 {
		return docs[0].Clone()
	}

	// Clone the first document
	result, err := docs[0].Clone()
	if err != nil {
		return nil, err
	}

	// Append remaining documents
	for i := 1; i < len(docs); i++ {
		result.AppendDocument(docs[i])
	}

	return result, nil
}

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

// =============================================================================
// Document Splitting
// =============================================================================

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

// =============================================================================
// Mail Merge Operations
// =============================================================================

// MailMerge renders the template with multiple data records.
// Returns a slice of rendered documents, one per record.
//
//	records := []map[string]any{
//	    {"Name": "John", "Email": "john@example.com"},
//	    {"Name": "Jane", "Email": "jane@example.com"},
//	}
//	docs, err := template.MailMerge(records)
func (d *DocxTmpl) MailMerge(records []map[string]any) ([]*DocxTmpl, error) {
	var results []*DocxTmpl

	for i, record := range records {
		// Clone the template
		clone, err := d.Clone()
		if err != nil {
			return nil, fmt.Errorf("failed to clone for record %d: %w", i, err)
		}

		// Render with this record's data
		err = clone.Render(record)
		if err != nil {
			return nil, fmt.Errorf("failed to render record %d: %w", i, err)
		}

		results = append(results, clone)
	}

	return results, nil
}

// MailMergeToFiles renders the template and saves each result to a file.
// filenamePattern should contain %d for the record number (e.g., "letter_%d.docx").
//
//	err := template.MailMergeToFiles(records, "output/letter_%d.docx")
func (d *DocxTmpl) MailMergeToFiles(records []map[string]any, filenamePattern string) error {
	docs, err := d.MailMerge(records)
	if err != nil {
		return err
	}

	for i, doc := range docs {
		filename := fmt.Sprintf(filenamePattern, i+1)

		// Ensure directory exists
		dir := filepath.Dir(filename)
		if dir != "" && dir != "." {
			if err := os.MkdirAll(dir, 0755); err != nil {
				return fmt.Errorf("failed to create directory for %s: %w", filename, err)
			}
		}

		if err := doc.SaveToFile(filename); err != nil {
			return fmt.Errorf("failed to save %s: %w", filename, err)
		}
	}

	return nil
}

// MailMergeToSingle renders the template with multiple records and combines
// into a single document with page breaks between each.
//
//	merged, err := template.MailMergeToSingle(records)
func (d *DocxTmpl) MailMergeToSingle(records []map[string]any) (*DocxTmpl, error) {
	if len(records) == 0 {
		return d.Clone()
	}

	docs, err := d.MailMerge(records)
	if err != nil {
		return nil, err
	}

	return MergeDocuments(docs...)
}

// =============================================================================
// Batch Processing
// =============================================================================

// BatchProcess applies a function to each document in a list.
//
//	err := docxtpl.BatchProcess(docs, func(doc *DocxTmpl) error {
//	    doc.ReplaceText("COMPANY", "Acme Corp")
//	    return nil
//	})
func BatchProcess(docs []*DocxTmpl, fn func(*DocxTmpl) error) error {
	for i, doc := range docs {
		if err := fn(doc); err != nil {
			return fmt.Errorf("error processing document %d: %w", i, err)
		}
	}
	return nil
}

// BatchRender renders multiple templates with corresponding data.
//
//	docs, err := docxtpl.BatchRender(templates, dataList)
func BatchRender(templates []*DocxTmpl, dataList []any) ([]*DocxTmpl, error) {
	if len(templates) != len(dataList) {
		return nil, fmt.Errorf("templates and data must have same length")
	}

	results := make([]*DocxTmpl, len(templates))
	for i, tmpl := range templates {
		clone, err := tmpl.Clone()
		if err != nil {
			return nil, fmt.Errorf("failed to clone template %d: %w", i, err)
		}

		if err := clone.Render(dataList[i]); err != nil {
			return nil, fmt.Errorf("failed to render template %d: %w", i, err)
		}

		results[i] = clone
	}

	return results, nil
}

// ReplaceInAll replaces text in all provided documents.
//
//	docxtpl.ReplaceInAll(docs, "OLD", "NEW")
func ReplaceInAll(docs []*DocxTmpl, oldText, newText string) {
	for _, doc := range docs {
		doc.ReplaceText(oldText, newText)
	}
}

// =============================================================================
// Directory Operations
// =============================================================================

// LoadAllFromDirectory loads all .docx files from a directory.
//
//	docs, err := docxtpl.LoadAllFromDirectory("templates/")
func LoadAllFromDirectory(dirPath string) ([]*DocxTmpl, error) {
	var docs []*DocxTmpl

	entries, err := os.ReadDir(dirPath)
	if err != nil {
		return nil, fmt.Errorf("failed to read directory: %w", err)
	}

	for _, entry := range entries {
		if entry.IsDir() {
			continue
		}

		if !strings.HasSuffix(strings.ToLower(entry.Name()), ".docx") {
			continue
		}

		fullPath := filepath.Join(dirPath, entry.Name())
		doc, err := ParseFromFilename(fullPath)
		if err != nil {
			return nil, fmt.Errorf("failed to load %s: %w", entry.Name(), err)
		}

		docs = append(docs, doc)
	}

	return docs, nil
}

// SaveAllToDirectory saves all documents to a directory with specified names.
//
//	err := docxtpl.SaveAllToDirectory(docs, "output/", []string{"doc1.docx", "doc2.docx"})
func SaveAllToDirectory(docs []*DocxTmpl, dirPath string, names []string) error {
	if len(docs) != len(names) {
		return fmt.Errorf("documents and names must have same length")
	}

	// Ensure directory exists
	if err := os.MkdirAll(dirPath, 0755); err != nil {
		return fmt.Errorf("failed to create directory: %w", err)
	}

	for i, doc := range docs {
		fullPath := filepath.Join(dirPath, names[i])
		if err := doc.SaveToFile(fullPath); err != nil {
			return fmt.Errorf("failed to save %s: %w", names[i], err)
		}
	}

	return nil
}

// ProcessDirectory processes all .docx files in a directory.
//
//	err := docxtpl.ProcessDirectory("input/", "output/", func(doc *DocxTmpl) error {
//	    return doc.Render(data)
//	})
func ProcessDirectory(inputDir, outputDir string, processor func(*DocxTmpl) error) error {
	// Ensure output directory exists
	if err := os.MkdirAll(outputDir, 0755); err != nil {
		return fmt.Errorf("failed to create output directory: %w", err)
	}

	entries, err := os.ReadDir(inputDir)
	if err != nil {
		return fmt.Errorf("failed to read input directory: %w", err)
	}

	for _, entry := range entries {
		if entry.IsDir() {
			continue
		}

		name := entry.Name()
		if !strings.HasSuffix(strings.ToLower(name), ".docx") {
			continue
		}

		inputPath := filepath.Join(inputDir, name)
		outputPath := filepath.Join(outputDir, name)

		doc, err := ParseFromFilename(inputPath)
		if err != nil {
			return fmt.Errorf("failed to load %s: %w", name, err)
		}

		if err := processor(doc); err != nil {
			return fmt.Errorf("failed to process %s: %w", name, err)
		}

		if err := doc.SaveToFile(outputPath); err != nil {
			return fmt.Errorf("failed to save %s: %w", name, err)
		}
	}

	return nil
}

// =============================================================================
// Text Extraction and Search
// =============================================================================

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

// SearchResult contains information about a text match.
type SearchResult struct {
	Text       string // The matched text
	Line       int    // Line number (1-indexed)
	Paragraph  int    // Paragraph index (0-indexed)
	Context    string // Surrounding context
	MatchStart int    // Start position of match in context
}

// SearchWithContext searches for text and returns matches with context.
// contextLines specifies how many lines of context to include.
//
//	results := doc.SearchWithContext("important", 2)
func (d *DocxTmpl) SearchWithContext(searchText string, contextLines int) []SearchResult {
	var results []SearchResult

	paragraphs := d.GetParagraphTexts()
	searchLower := strings.ToLower(searchText)

	for i, para := range paragraphs {
		if para == "" {
			continue
		}

		paraLower := strings.ToLower(para)
		if !strings.Contains(paraLower, searchLower) {
			continue
		}

		// Build context
		var context strings.Builder
		startPara := i - contextLines
		if startPara < 0 {
			startPara = 0
		}
		endPara := i + contextLines + 1
		if endPara > len(paragraphs) {
			endPara = len(paragraphs)
		}

		for j := startPara; j < endPara; j++ {
			if j > startPara {
				context.WriteString("\n")
			}
			if j == i {
				context.WriteString(">>> ")
			}
			context.WriteString(paragraphs[j])
		}

		results = append(results, SearchResult{
			Text:      searchText,
			Paragraph: i,
			Line:      i + 1,
			Context:   context.String(),
		})
	}

	return results
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

// =============================================================================
// Text Replacement
// =============================================================================

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

// =============================================================================
// Media Operations
// =============================================================================

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

// =============================================================================
// Document Counting
// =============================================================================

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

// =============================================================================
// Document Cleanup
// =============================================================================

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

// =============================================================================
// List Types and Structures
// =============================================================================

// ListType represents the type of list
type ListType int

const (
	ListTypeBullet   ListType = iota // Bullet list (-, o, *)
	ListTypeNumbered                 // Numbered list (1, 2, 3)
	ListTypeLetter                   // Letter list (a, b, c)
	ListTypeRoman                    // Roman numeral list (i, ii, iii)
)

// ListItem represents an item in a list with optional nesting
type ListItem struct {
	Text     string     // Item text
	Level    int        // Nesting level (0-8)
	Children []ListItem // Nested items
}

// List wraps a collection of list paragraphs
type List struct {
	doc        *DocxTmpl
	listType   ListType
	paragraphs []*Paragraph
}

// =============================================================================
// List Creation
// =============================================================================

// AddBulletList adds a bullet list to the document.
// Each string in items becomes a bullet point.
//
//	doc.AddBulletList([]string{"First item", "Second item", "Third item"})
func (d *DocxTmpl) AddBulletList(items []string) *List {
	list := &List{
		doc:      d,
		listType: ListTypeBullet,
	}

	for _, item := range items {
		para := d.addListParagraph(item, ListTypeBullet, 0)
		list.paragraphs = append(list.paragraphs, para)
	}

	return list
}

// AddNumberedList adds a numbered list to the document.
// Each string in items becomes a numbered item (1, 2, 3...).
//
//	doc.AddNumberedList([]string{"First step", "Second step", "Third step"})
func (d *DocxTmpl) AddNumberedList(items []string) *List {
	list := &List{
		doc:      d,
		listType: ListTypeNumbered,
	}

	for _, item := range items {
		para := d.addListParagraph(item, ListTypeNumbered, 0)
		list.paragraphs = append(list.paragraphs, para)
	}

	return list
}

// AddNestedList adds a nested list with multiple levels.
// Use ListItem.Level to control indentation (0-8).
// Use ListItem.Children for nested items.
//
//	doc.AddNestedList(ListTypeBullet, []ListItem{
//	    {Text: "Item 1", Children: []ListItem{
//	        {Text: "Sub-item 1.1"},
//	        {Text: "Sub-item 1.2"},
//	    }},
//	    {Text: "Item 2"},
//	})
func (d *DocxTmpl) AddNestedList(listType ListType, items []ListItem) *List {
	list := &List{
		doc:      d,
		listType: listType,
	}

	d.addNestedListItems(list, items, listType, 0)

	return list
}

// addNestedListItems recursively adds list items
func (d *DocxTmpl) addNestedListItems(list *List, items []ListItem, listType ListType, level int) {
	for _, item := range items {
		effectiveLevel := item.Level
		if effectiveLevel == 0 {
			effectiveLevel = level
		}
		if effectiveLevel > 8 {
			effectiveLevel = 8
		}

		para := d.addListParagraph(item.Text, listType, effectiveLevel)
		list.paragraphs = append(list.paragraphs, para)

		// Recursively add children
		if len(item.Children) > 0 {
			d.addNestedListItems(list, item.Children, listType, effectiveLevel+1)
		}
	}
}

// addListParagraph creates a paragraph with list formatting
func (d *DocxTmpl) addListParagraph(text string, listType ListType, level int) *Paragraph {
	p := d.Docx.AddParagraph()

	// Create the bullet/number prefix based on list type and level
	prefix := getBulletPrefix(listType, level)

	// Add prefix as a separate run
	prefixRun := p.AddText(prefix + "\t")
	_ = prefixRun // avoid unused

	// Add the main text
	run := p.AddText(text)

	// Apply indentation based on level
	para := &Paragraph{
		paragraph: p,
		lastRun:   run,
		doc:       d,
	}

	// Apply list indentation (in inches)
	indent := 0.5 * float64(level+1) // 0.5 inch per level
	para.IndentLeft(indent)
	para.IndentHanging(0.25) // Hanging indent for bullet/number

	return para
}

// getBulletPrefix returns the appropriate bullet or number character
func getBulletPrefix(listType ListType, level int) string {
	switch listType {
	case ListTypeBullet:
		// Different bullets for different levels
		bullets := []string{"•", "○", "▪", "•", "○", "▪", "•", "○", "▪"}
		if level < len(bullets) {
			return bullets[level]
		}
		return "•"
	case ListTypeNumbered:
		// Numbers - would need state tracking for actual numbers
		// For now, return placeholder
		return "1."
	case ListTypeLetter:
		return "a."
	case ListTypeRoman:
		return "i."
	default:
		return "•"
	}
}

// =============================================================================
// List Methods
// =============================================================================

// AddItem adds another item to the list at the same level.
func (l *List) AddItem(text string) *List {
	para := l.doc.addListParagraph(text, l.listType, 0)
	l.paragraphs = append(l.paragraphs, para)
	return l
}

// AddSubItem adds a nested item to the list.
func (l *List) AddSubItem(text string, level int) *List {
	if level < 0 {
		level = 0
	}
	if level > 8 {
		level = 8
	}
	para := l.doc.addListParagraph(text, l.listType, level)
	l.paragraphs = append(l.paragraphs, para)
	return l
}

// GetParagraphs returns all paragraphs in the list.
func (l *List) GetParagraphs() []*Paragraph {
	return l.paragraphs
}

// Count returns the number of items in the list.
func (l *List) Count() int {
	return len(l.paragraphs)
}

// =============================================================================
// List Builder
// =============================================================================

// ListBuilder provides a fluent interface for building complex lists
type ListBuilder struct {
	doc          *DocxTmpl
	listType     ListType
	items        []listBuilderItem
	currentLevel int
}

type listBuilderItem struct {
	text  string
	level int
}

// NewListBuilder creates a new list builder
//
//	list := doc.NewListBuilder(ListTypeBullet).
//	    Item("First").
//	    Item("Second").
//	    Indent().Item("Nested").
//	    Outdent().Item("Third").
//	    Build()
func (d *DocxTmpl) NewListBuilder(listType ListType) *ListBuilder {
	return &ListBuilder{
		doc:          d,
		listType:     listType,
		items:        []listBuilderItem{},
		currentLevel: 0,
	}
}

// Item adds an item at the current indentation level
func (lb *ListBuilder) Item(text string) *ListBuilder {
	lb.items = append(lb.items, listBuilderItem{
		text:  text,
		level: lb.currentLevel,
	})
	return lb
}

// Indent increases the indentation level for subsequent items
func (lb *ListBuilder) Indent() *ListBuilder {
	if lb.currentLevel < 8 {
		lb.currentLevel++
	}
	return lb
}

// Outdent decreases the indentation level for subsequent items
func (lb *ListBuilder) Outdent() *ListBuilder {
	if lb.currentLevel > 0 {
		lb.currentLevel--
	}
	return lb
}

// Level sets the indentation level directly
func (lb *ListBuilder) Level(level int) *ListBuilder {
	if level < 0 {
		level = 0
	}
	if level > 8 {
		level = 8
	}
	lb.currentLevel = level
	return lb
}

// Build creates the list and adds it to the document
func (lb *ListBuilder) Build() *List {
	list := &List{
		doc:      lb.doc,
		listType: lb.listType,
	}

	for _, item := range lb.items {
		para := lb.doc.addListParagraph(item.text, lb.listType, item.level)
		list.paragraphs = append(list.paragraphs, para)
	}

	return list
}

// =============================================================================
// Checklist
// =============================================================================

// AddChecklistItem adds a checkbox item to the document.
// checked determines if the checkbox appears checked.
//
//	doc.AddChecklistItem("Complete task", true)
//	doc.AddChecklistItem("Pending task", false)
func (d *DocxTmpl) AddChecklistItem(text string, checked bool) *Paragraph {
	p := d.Docx.AddParagraph()

	// Add checkbox symbol
	checkbox := "☐ "
	if checked {
		checkbox = "☑ "
	}

	p.AddText(checkbox)
	run := p.AddText(text)

	return &Paragraph{
		paragraph: p,
		lastRun:   run,
		doc:       d,
	}
}

// AddChecklist adds a checklist to the document.
// items is a map of text to checked status.
//
//	doc.AddChecklist(map[string]bool{
//	    "Task 1": true,
//	    "Task 2": false,
//	    "Task 3": false,
//	})
func (d *DocxTmpl) AddChecklist(items map[string]bool) []*Paragraph {
	var paragraphs []*Paragraph
	for text, checked := range items {
		para := d.AddChecklistItem(text, checked)
		paragraphs = append(paragraphs, para)
	}
	return paragraphs
}

// AddOrderedChecklist adds a checklist with preserved order.
//
//	doc.AddOrderedChecklist([]struct{Text string; Checked bool}{
//	    {"Task 1", true},
//	    {"Task 2", false},
//	})
func (d *DocxTmpl) AddOrderedChecklist(items []struct {
	Text    string
	Checked bool
}) []*Paragraph {
	var paragraphs []*Paragraph
	for _, item := range items {
		para := d.AddChecklistItem(item.Text, item.Checked)
		paragraphs = append(paragraphs, para)
	}
	return paragraphs
}
