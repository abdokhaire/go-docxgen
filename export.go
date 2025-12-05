package docxtpl

import (
	"encoding/json"
	"fmt"
	"strings"

	"github.com/fumiama/go-docx"
)

// StructuredDocument represents the document in a structured format suitable for AI consumption
type StructuredDocument struct {
	Metadata   DocumentMetadata       `json:"metadata"`
	Stats      DocumentStats          `json:"stats"`
	Outline    []OutlineItem          `json:"outline"`
	Paragraphs []StructuredParagraph  `json:"paragraphs"`
	Tables     []StructuredTable      `json:"tables"`
	Images     []ImageInfo            `json:"images"`
	Links      []HyperlinkInfo        `json:"links"`
}

// StructuredParagraph represents a paragraph with its metadata
type StructuredParagraph struct {
	Index     int              `json:"index"`
	Text      string           `json:"text"`
	Style     string           `json:"style"`
	IsList    bool             `json:"is_list,omitempty"`
	ListLevel int              `json:"list_level,omitempty"`
	Alignment string           `json:"alignment,omitempty"`
	Runs      []StructuredRun  `json:"runs,omitempty"`
}

// StructuredRun represents a text run with formatting
type StructuredRun struct {
	Text       string `json:"text"`
	Bold       bool   `json:"bold,omitempty"`
	Italic     bool   `json:"italic,omitempty"`
	Underline  bool   `json:"underline,omitempty"`
	Strike     bool   `json:"strike,omitempty"`
	Color      string `json:"color,omitempty"`
	FontSize   int    `json:"font_size,omitempty"`
	FontFamily string `json:"font_family,omitempty"`
	Highlight  string `json:"highlight,omitempty"`
}

// StructuredTable represents a table with its data
type StructuredTable struct {
	Index   int        `json:"index"`
	Rows    int        `json:"rows"`
	Cols    int        `json:"cols"`
	Data    [][]string `json:"data"`
	Headers []string   `json:"headers,omitempty"`
}

// ImageInfo represents an embedded image
type ImageInfo struct {
	Name   string `json:"name"`
	Format string `json:"format"`
	Size   int    `json:"size"` // bytes
}

// ToStructured converts the document to a structured representation.
// This is useful for AI/LLM consumption and document analysis.
//
//	structured := doc.ToStructured()
//	jsonBytes, _ := json.Marshal(structured)
func (d *DocxTmpl) ToStructured() *StructuredDocument {
	sd := &StructuredDocument{
		Metadata:   *d.GetMetadata(),
		Stats:      *d.GetStats(),
		Outline:    d.GetOutline(),
		Links:      d.GetAllHyperlinks(),
	}

	// Extract paragraphs
	paraIdx := 0
	tableIdx := 0
	for _, item := range d.Document.Body.Items {
		switch v := item.(type) {
		case *docx.Paragraph:
			sp := extractStructuredParagraph(v, paraIdx)
			sd.Paragraphs = append(sd.Paragraphs, sp)
			paraIdx++
		case *docx.Table:
			st := extractStructuredTable(v, tableIdx)
			sd.Tables = append(sd.Tables, st)
			tableIdx++
		}
	}

	// Extract images
	for _, m := range d.GetAllMedia() {
		name := strings.ToLower(m.Name)
		if isImageFile(name) {
			sd.Images = append(sd.Images, ImageInfo{
				Name:   m.Name,
				Format: getImageFormat(name),
				Size:   len(m.Data),
			})
		}
	}

	return sd
}

// ToJSON returns the document structure as JSON.
//
//	jsonStr, err := doc.ToJSON()
func (d *DocxTmpl) ToJSON() (string, error) {
	structured := d.ToStructured()
	bytes, err := json.MarshalIndent(structured, "", "  ")
	if err != nil {
		return "", err
	}
	return string(bytes), nil
}

// ToMarkdown converts the document to Markdown format.
// Supports headings, bold, italic, lists, tables, and links.
//
//	markdown := doc.ToMarkdown()
func (d *DocxTmpl) ToMarkdown() string {
	var sb strings.Builder

	for _, item := range d.Document.Body.Items {
		switch v := item.(type) {
		case *docx.Paragraph:
			md := paragraphToMarkdown(v)
			if md != "" {
				sb.WriteString(md)
				sb.WriteString("\n\n")
			}
		case *docx.Table:
			md := tableToMarkdown(v)
			sb.WriteString(md)
			sb.WriteString("\n")
		}
	}

	return strings.TrimSpace(sb.String())
}

// ToHTML converts the document to basic HTML format.
//
//	html := doc.ToHTML()
func (d *DocxTmpl) ToHTML() string {
	var sb strings.Builder

	sb.WriteString("<!DOCTYPE html>\n<html>\n<head>\n")
	sb.WriteString("<meta charset=\"UTF-8\">\n")

	meta := d.GetMetadata()
	if meta.Title != "" {
		sb.WriteString(fmt.Sprintf("<title>%s</title>\n", escapeHTML(meta.Title)))
	}
	sb.WriteString("</head>\n<body>\n")

	for _, item := range d.Document.Body.Items {
		switch v := item.(type) {
		case *docx.Paragraph:
			html := paragraphToHTML(v)
			if html != "" {
				sb.WriteString(html)
				sb.WriteString("\n")
			}
		case *docx.Table:
			html := tableToHTML(v)
			sb.WriteString(html)
			sb.WriteString("\n")
		}
	}

	sb.WriteString("</body>\n</html>")
	return sb.String()
}

// Helper functions

func extractStructuredParagraph(p *docx.Paragraph, idx int) StructuredParagraph {
	sp := StructuredParagraph{
		Index: idx,
		Text:  p.String(),
		Style: "Normal",
	}

	if p.Properties != nil {
		if p.Properties.Style != nil {
			sp.Style = p.Properties.Style.Val
		}
		if p.Properties.Justification != nil {
			sp.Alignment = p.Properties.Justification.Val
		}
	}

	// Extract runs
	for _, child := range p.Children {
		if run, ok := child.(*docx.Run); ok {
			sr := extractStructuredRun(run)
			if sr.Text != "" {
				sp.Runs = append(sp.Runs, sr)
			}
		}
	}

	return sp
}

func extractStructuredRun(r *docx.Run) StructuredRun {
	sr := StructuredRun{}

	// Get text
	for _, child := range r.Children {
		if t, ok := child.(*docx.Text); ok {
			sr.Text += t.Text
		}
	}

	// Get formatting
	if r.RunProperties != nil {
		if r.RunProperties.Bold != nil {
			sr.Bold = true
		}
		if r.RunProperties.Italic != nil {
			sr.Italic = true
		}
		if r.RunProperties.Underline != nil {
			sr.Underline = true
		}
		if r.RunProperties.Strike != nil {
			sr.Strike = true
		}
		if r.RunProperties.Color != nil {
			sr.Color = r.RunProperties.Color.Val
		}
		if r.RunProperties.Size != nil {
			if size, err := parseSize(r.RunProperties.Size.Val); err == nil {
				sr.FontSize = size / 2 // Half-points to points
			}
		}
		if r.RunProperties.Highlight != nil {
			sr.Highlight = r.RunProperties.Highlight.Val
		}
	}

	return sr
}

func extractStructuredTable(t *docx.Table, idx int) StructuredTable {
	st := StructuredTable{
		Index: idx,
		Rows:  len(t.TableRows),
	}

	if len(t.TableRows) > 0 {
		st.Cols = len(t.TableRows[0].TableCells)
	}

	// Extract data
	for _, row := range t.TableRows {
		var rowData []string
		for _, cell := range row.TableCells {
			cellText := ""
			for _, p := range cell.Paragraphs {
				if cellText != "" {
					cellText += "\n"
				}
				cellText += p.String()
			}
			rowData = append(rowData, cellText)
		}
		st.Data = append(st.Data, rowData)
	}

	// First row as headers (optional)
	if len(st.Data) > 0 {
		st.Headers = st.Data[0]
	}

	return st
}

func paragraphToMarkdown(p *docx.Paragraph) string {
	text := p.String()
	if text == "" {
		return ""
	}

	// Check for heading style
	if p.Properties != nil && p.Properties.Style != nil {
		style := p.Properties.Style.Val
		if style == "Title" {
			return "# " + text
		}
		if strings.HasPrefix(style, "Heading") {
			level := style[len("Heading"):]
			prefix := strings.Repeat("#", len(level)+1)
			return prefix + " " + text
		}
	}

	// Process inline formatting
	var result strings.Builder
	for _, child := range p.Children {
		if run, ok := child.(*docx.Run); ok {
			runText := getRunText(run)
			if runText == "" {
				continue
			}

			// Apply formatting
			if run.RunProperties != nil {
				if run.RunProperties.Bold != nil && run.RunProperties.Italic != nil {
					runText = "***" + runText + "***"
				} else if run.RunProperties.Bold != nil {
					runText = "**" + runText + "**"
				} else if run.RunProperties.Italic != nil {
					runText = "*" + runText + "*"
				}
				if run.RunProperties.Strike != nil {
					runText = "~~" + runText + "~~"
				}
			}
			result.WriteString(runText)
		}
	}

	return result.String()
}

func tableToMarkdown(t *docx.Table) string {
	if len(t.TableRows) == 0 {
		return ""
	}

	var sb strings.Builder

	// Header row
	if len(t.TableRows) > 0 {
		row := t.TableRows[0]
		sb.WriteString("|")
		for _, cell := range row.TableCells {
			cellText := getCellText(cell)
			sb.WriteString(" " + cellText + " |")
		}
		sb.WriteString("\n|")
		for range row.TableCells {
			sb.WriteString(" --- |")
		}
		sb.WriteString("\n")
	}

	// Data rows
	for i := 1; i < len(t.TableRows); i++ {
		row := t.TableRows[i]
		sb.WriteString("|")
		for _, cell := range row.TableCells {
			cellText := getCellText(cell)
			sb.WriteString(" " + cellText + " |")
		}
		sb.WriteString("\n")
	}

	return sb.String()
}

func paragraphToHTML(p *docx.Paragraph) string {
	text := p.String()
	if text == "" {
		return ""
	}

	// Check for heading style
	tag := "p"
	if p.Properties != nil && p.Properties.Style != nil {
		style := p.Properties.Style.Val
		if style == "Title" {
			tag = "h1"
		} else if strings.HasPrefix(style, "Heading") {
			level := style[len("Heading"):]
			tag = "h" + level
		}
	}

	// Build content with formatting
	var content strings.Builder
	for _, child := range p.Children {
		if run, ok := child.(*docx.Run); ok {
			runText := getRunText(run)
			if runText == "" {
				continue
			}

			runText = escapeHTML(runText)

			// Apply formatting
			if run.RunProperties != nil {
				if run.RunProperties.Bold != nil {
					runText = "<strong>" + runText + "</strong>"
				}
				if run.RunProperties.Italic != nil {
					runText = "<em>" + runText + "</em>"
				}
				if run.RunProperties.Underline != nil {
					runText = "<u>" + runText + "</u>"
				}
				if run.RunProperties.Strike != nil {
					runText = "<s>" + runText + "</s>"
				}
				if run.RunProperties.Color != nil {
					runText = fmt.Sprintf("<span style=\"color:#%s\">%s</span>",
						run.RunProperties.Color.Val, runText)
				}
			}
			content.WriteString(runText)
		}
	}

	return fmt.Sprintf("<%s>%s</%s>", tag, content.String(), tag)
}

func tableToHTML(t *docx.Table) string {
	if len(t.TableRows) == 0 {
		return ""
	}

	var sb strings.Builder
	sb.WriteString("<table border=\"1\">\n")

	for i, row := range t.TableRows {
		sb.WriteString("<tr>")
		cellTag := "td"
		if i == 0 {
			cellTag = "th"
		}
		for _, cell := range row.TableCells {
			cellText := escapeHTML(getCellText(cell))
			sb.WriteString(fmt.Sprintf("<%s>%s</%s>", cellTag, cellText, cellTag))
		}
		sb.WriteString("</tr>\n")
	}

	sb.WriteString("</table>")
	return sb.String()
}

func getRunText(r *docx.Run) string {
	var text string
	for _, child := range r.Children {
		if t, ok := child.(*docx.Text); ok {
			text += t.Text
		}
	}
	return text
}

func getCellText(cell *docx.WTableCell) string {
	var texts []string
	for _, p := range cell.Paragraphs {
		texts = append(texts, p.String())
	}
	return strings.Join(texts, " ")
}

func escapeHTML(s string) string {
	s = strings.ReplaceAll(s, "&", "&amp;")
	s = strings.ReplaceAll(s, "<", "&lt;")
	s = strings.ReplaceAll(s, ">", "&gt;")
	s = strings.ReplaceAll(s, "\"", "&quot;")
	return s
}

func isImageFile(name string) bool {
	return strings.HasSuffix(name, ".png") ||
		strings.HasSuffix(name, ".jpg") ||
		strings.HasSuffix(name, ".jpeg") ||
		strings.HasSuffix(name, ".gif") ||
		strings.HasSuffix(name, ".bmp")
}

func getImageFormat(name string) string {
	if strings.HasSuffix(name, ".png") {
		return "PNG"
	}
	if strings.HasSuffix(name, ".jpg") || strings.HasSuffix(name, ".jpeg") {
		return "JPEG"
	}
	if strings.HasSuffix(name, ".gif") {
		return "GIF"
	}
	if strings.HasSuffix(name, ".bmp") {
		return "BMP"
	}
	return "Unknown"
}

func parseSize(s string) (int, error) {
	var size int
	_, err := fmt.Sscanf(s, "%d", &size)
	return size, err
}
