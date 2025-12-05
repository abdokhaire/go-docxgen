package docxtpl

import (
	"fmt"
	"regexp"
	"strings"
	"time"
)

// TrackedChangeType represents the type of tracked change
type TrackedChangeType string

const (
	ChangeTypeInsertion TrackedChangeType = "insertion"
	ChangeTypeDeletion  TrackedChangeType = "deletion"
)

// TrackedChange represents a tracked change in the document
type TrackedChange struct {
	ID       int               // Change ID
	Type     TrackedChangeType // insertion or deletion
	Author   string            // Author who made the change
	Date     time.Time         // When the change was made
	Text     string            // The changed text
	Location string            // Location description (paragraph index, etc.)
}

// TrackedChangesConfig holds configuration for tracked changes
type TrackedChangesConfig struct {
	Author   string // Default author name
	Initials string // Default author initials
}

// DefaultTrackedChangesConfig returns default configuration
func DefaultTrackedChangesConfig() TrackedChangesConfig {
	return TrackedChangesConfig{
		Author:   "Go-DocxGen",
		Initials: "GD",
	}
}

// GetTrackedChanges extracts all tracked changes from the document.
// Returns insertions and deletions found in the document XML.
//
//	changes := doc.GetTrackedChanges()
//	for _, change := range changes {
//	    fmt.Printf("%s by %s: %s\n", change.Type, change.Author, change.Text)
//	}
func (d *DocxTmpl) GetTrackedChanges() []TrackedChange {
	var changes []TrackedChange

	// Get the document XML
	xmlContent, err := d.getDocumentXml()
	if err != nil || xmlContent == "" {
		return changes
	}

	// Find insertions - <w:ins>...</w:ins>
	insChanges := extractTrackedChanges(xmlContent, "ins", ChangeTypeInsertion)
	changes = append(changes, insChanges...)

	// Find deletions - <w:del>...</w:del>
	delChanges := extractTrackedChanges(xmlContent, "del", ChangeTypeDeletion)
	changes = append(changes, delChanges...)

	// Also check processable files (headers, footers, etc.)
	for _, pf := range d.processableFiles {
		insChanges := extractTrackedChanges(pf.Content, "ins", ChangeTypeInsertion)
		changes = append(changes, insChanges...)

		delChanges := extractTrackedChanges(pf.Content, "del", ChangeTypeDeletion)
		changes = append(changes, delChanges...)
	}

	return changes
}

// HasTrackedChanges returns true if the document contains any tracked changes.
//
//	if doc.HasTrackedChanges() {
//	    fmt.Println("Document has pending changes")
//	}
func (d *DocxTmpl) HasTrackedChanges() bool {
	return len(d.GetTrackedChanges()) > 0
}

// GetInsertions returns only insertion changes.
//
//	insertions := doc.GetInsertions()
func (d *DocxTmpl) GetInsertions() []TrackedChange {
	var insertions []TrackedChange
	for _, change := range d.GetTrackedChanges() {
		if change.Type == ChangeTypeInsertion {
			insertions = append(insertions, change)
		}
	}
	return insertions
}

// GetDeletions returns only deletion changes.
//
//	deletions := doc.GetDeletions()
func (d *DocxTmpl) GetDeletions() []TrackedChange {
	var deletions []TrackedChange
	for _, change := range d.GetTrackedChanges() {
		if change.Type == ChangeTypeDeletion {
			deletions = append(deletions, change)
		}
	}
	return deletions
}

// CountTrackedChanges returns the count of tracked changes by type.
//
//	ins, del := doc.CountTrackedChanges()
func (d *DocxTmpl) CountTrackedChanges() (insertions, deletions int) {
	for _, change := range d.GetTrackedChanges() {
		switch change.Type {
		case ChangeTypeInsertion:
			insertions++
		case ChangeTypeDeletion:
			deletions++
		}
	}
	return
}

// AcceptAllChanges accepts all tracked changes, making insertions permanent
// and removing deleted text.
// Note: This modifies the processable files (headers, footers, etc.).
// Document body changes require re-parsing after save.
//
//	doc.AcceptAllChanges()
func (d *DocxTmpl) AcceptAllChanges() {
	// Process headers, footers, footnotes, endnotes
	for i := range d.processableFiles {
		content := d.processableFiles[i].Content

		// Accept insertions: remove <w:ins> tags but keep content
		content = acceptInsertions(content)

		// Accept deletions: remove <w:del> tags and their content
		content = acceptDeletions(content)

		d.processableFiles[i].Content = content
	}
}

// RejectAllChanges rejects all tracked changes, removing insertions
// and restoring deleted text.
// Note: This modifies the processable files (headers, footers, etc.).
// Document body changes require re-parsing after save.
//
//	doc.RejectAllChanges()
func (d *DocxTmpl) RejectAllChanges() {
	// Process headers, footers, footnotes, endnotes
	for i := range d.processableFiles {
		content := d.processableFiles[i].Content

		// Reject insertions: remove <w:ins> tags and their content
		content = rejectInsertions(content)

		// Reject deletions: remove <w:del> tags but convert delText to regular text
		content = rejectDeletions(content)

		d.processableFiles[i].Content = content
	}
}

// GetChangesByAuthor returns tracked changes filtered by author.
//
//	claudeChanges := doc.GetChangesByAuthor("Claude")
func (d *DocxTmpl) GetChangesByAuthor(author string) []TrackedChange {
	var filtered []TrackedChange
	authorLower := strings.ToLower(author)

	for _, change := range d.GetTrackedChanges() {
		if strings.ToLower(change.Author) == authorLower {
			filtered = append(filtered, change)
		}
	}
	return filtered
}

// TrackedChangesSummary returns a text summary of all tracked changes.
//
//	summary := doc.TrackedChangesSummary()
//	fmt.Println(summary)
func (d *DocxTmpl) TrackedChangesSummary() string {
	changes := d.GetTrackedChanges()
	if len(changes) == 0 {
		return "No tracked changes found."
	}

	var sb strings.Builder
	sb.WriteString(fmt.Sprintf("Tracked Changes Summary (%d total)\n", len(changes)))
	sb.WriteString(strings.Repeat("-", 40) + "\n")

	ins, del := d.CountTrackedChanges()
	sb.WriteString(fmt.Sprintf("Insertions: %d\n", ins))
	sb.WriteString(fmt.Sprintf("Deletions: %d\n", del))
	sb.WriteString("\nChanges:\n")

	for i, change := range changes {
		typeStr := "+"
		if change.Type == ChangeTypeDeletion {
			typeStr = "-"
		}
		text := change.Text
		if len(text) > 50 {
			text = text[:50] + "..."
		}
		sb.WriteString(fmt.Sprintf("%d. [%s] %s (by %s)\n", i+1, typeStr, text, change.Author))
	}

	return sb.String()
}

// Helper functions

func extractTrackedChanges(xml string, tagName string, changeType TrackedChangeType) []TrackedChange {
	var changes []TrackedChange

	// Pattern to match tracked change elements
	pattern := fmt.Sprintf(`<w:%s[^>]*w:author="([^"]*)"[^>]*w:date="([^"]*)"[^>]*>(.*?)</w:%s>`, tagName, tagName)
	re := regexp.MustCompile(pattern)

	matches := re.FindAllStringSubmatch(xml, -1)
	for i, match := range matches {
		if len(match) >= 4 {
			author := match[1]
			dateStr := match[2]
			content := match[3]

			// Extract text from the content
			text := extractTextFromXML(content, changeType)

			// Parse date
			date, _ := time.Parse(time.RFC3339, dateStr)

			changes = append(changes, TrackedChange{
				ID:     i + 1,
				Type:   changeType,
				Author: author,
				Date:   date,
				Text:   text,
			})
		}
	}

	return changes
}

func extractTextFromXML(xml string, changeType TrackedChangeType) string {
	var text string

	// For deletions, look for <w:delText>
	if changeType == ChangeTypeDeletion {
		delTextRe := regexp.MustCompile(`<w:delText[^>]*>([^<]*)</w:delText>`)
		matches := delTextRe.FindAllStringSubmatch(xml, -1)
		for _, match := range matches {
			if len(match) >= 2 {
				text += match[1]
			}
		}
	} else {
		// For insertions, look for <w:t>
		textRe := regexp.MustCompile(`<w:t[^>]*>([^<]*)</w:t>`)
		matches := textRe.FindAllStringSubmatch(xml, -1)
		for _, match := range matches {
			if len(match) >= 2 {
				text += match[1]
			}
		}
	}

	return strings.TrimSpace(text)
}

func acceptInsertions(xml string) string {
	// Remove <w:ins> wrapper but keep content
	insStartRe := regexp.MustCompile(`<w:ins[^>]*>`)
	insEndRe := regexp.MustCompile(`</w:ins>`)

	xml = insStartRe.ReplaceAllString(xml, "")
	xml = insEndRe.ReplaceAllString(xml, "")

	return xml
}

func acceptDeletions(xml string) string {
	// Remove entire <w:del>...</w:del> blocks including content
	delRe := regexp.MustCompile(`<w:del[^>]*>.*?</w:del>`)
	return delRe.ReplaceAllString(xml, "")
}

func rejectInsertions(xml string) string {
	// Remove entire <w:ins>...</w:ins> blocks including content
	insRe := regexp.MustCompile(`<w:ins[^>]*>.*?</w:ins>`)
	return insRe.ReplaceAllString(xml, "")
}

func rejectDeletions(xml string) string {
	// Remove <w:del> wrapper and convert delText to regular text
	delStartRe := regexp.MustCompile(`<w:del[^>]*>`)
	delEndRe := regexp.MustCompile(`</w:del>`)
	delTextRe := regexp.MustCompile(`<w:delText([^>]*)>`)
	delTextEndRe := regexp.MustCompile(`</w:delText>`)

	xml = delStartRe.ReplaceAllString(xml, "")
	xml = delEndRe.ReplaceAllString(xml, "")
	xml = delTextRe.ReplaceAllString(xml, "<w:t$1>")
	xml = delTextEndRe.ReplaceAllString(xml, "</w:t>")

	return xml
}

// EnableTrackChanges enables track changes mode in document settings.
// Note: This modifies settings.xml to enable revision tracking.
// Requires settings.xml to be in processable files.
//
//	doc.EnableTrackChanges()
func (d *DocxTmpl) EnableTrackChanges() error {
	return d.setTrackChangesEnabled(true)
}

// DisableTrackChanges disables track changes mode in document settings.
//
//	doc.DisableTrackChanges()
func (d *DocxTmpl) DisableTrackChanges() error {
	return d.setTrackChangesEnabled(false)
}

func (d *DocxTmpl) setTrackChangesEnabled(enabled bool) error {
	// Look in processable files for settings.xml
	for i := range d.processableFiles {
		if strings.HasSuffix(d.processableFiles[i].Name, "settings.xml") {
			content := d.processableFiles[i].Content

			if enabled {
				// Add trackRevisions if not present
				if !strings.Contains(content, "<w:trackRevisions") {
					// Insert before </w:settings>
					content = strings.Replace(content, "</w:settings>", "<w:trackRevisions/></w:settings>", 1)
				}
			} else {
				// Remove trackRevisions
				trackRevRe := regexp.MustCompile(`<w:trackRevisions[^>]*/>`)
				content = trackRevRe.ReplaceAllString(content, "")
			}

			d.processableFiles[i].Content = content
			return nil
		}
	}
	return fmt.Errorf("settings.xml not found in processable files")
}

// IsTrackChangesEnabled checks if track changes mode is enabled.
//
//	if doc.IsTrackChangesEnabled() {
//	    fmt.Println("Track changes is enabled")
//	}
func (d *DocxTmpl) IsTrackChangesEnabled() bool {
	for _, pf := range d.processableFiles {
		if strings.HasSuffix(pf.Name, "settings.xml") {
			return strings.Contains(pf.Content, "<w:trackRevisions")
		}
	}
	return false
}
