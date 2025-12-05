package docxtpl

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

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

// SearchWithContext returns text matches with surrounding context.
//
//	results := doc.SearchWithContext("error", 1)
//	for _, r := range results {
//	    fmt.Printf("Found at line %d:\n%s\n", r.Line, r.Context)
//	}
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

// ReplaceInAll replaces text in all provided documents.
//
//	docxtpl.ReplaceInAll(docs, "OLD", "NEW")
func ReplaceInAll(docs []*DocxTmpl, oldText, newText string) {
	for _, doc := range docs {
		doc.ReplaceText(oldText, newText)
	}
}

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
