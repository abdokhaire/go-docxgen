package docxtpl

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strings"
	"time"

	"github.com/abdokhaire/go-docxgen/internal/docx"
)

// =============================================================================
// Document Comparison
// =============================================================================

// DiffType represents the type of difference
type DiffType string

const (
	DiffTypeAdded    DiffType = "added"
	DiffTypeRemoved  DiffType = "removed"
	DiffTypeModified DiffType = "modified"
)

// DiffItem represents a single difference between two documents
type DiffItem struct {
	Type      DiffType // Type of change
	Location  string   // Where the change occurred (paragraph index, table location, etc.)
	OldValue  string   // Original value (empty for additions)
	NewValue  string   // New value (empty for removals)
}

// DocumentDiff represents the differences between two documents
type DocumentDiff struct {
	Added    []DiffItem // Content added in the new document
	Removed  []DiffItem // Content removed from the original
	Modified []DiffItem // Content that was modified
	Summary  DiffSummary
}

// DiffSummary contains summary statistics of differences
type DiffSummary struct {
	TotalChanges       int
	AddedParagraphs    int
	RemovedParagraphs  int
	ModifiedParagraphs int
	AddedTables        int
	RemovedTables      int
}

// CompareDocuments compares two documents and returns their differences.
//
//	doc1, _ := docxtpl.ParseFromFilename("version1.docx")
//	doc2, _ := docxtpl.ParseFromFilename("version2.docx")
//	diff := docxtpl.CompareDocuments(doc1, doc2)
func CompareDocuments(doc1, doc2 *DocxTmpl) *DocumentDiff {
	diff := &DocumentDiff{
		Added:    []DiffItem{},
		Removed:  []DiffItem{},
		Modified: []DiffItem{},
	}

	// Get text from both documents
	paras1 := doc1.GetParagraphTexts()
	paras2 := doc2.GetParagraphTexts()

	// Create maps for quick lookup
	paraMap1 := make(map[string]bool)
	paraMap2 := make(map[string]bool)

	for _, p := range paras1 {
		paraMap1[p] = true
	}
	for _, p := range paras2 {
		paraMap2[p] = true
	}

	// Find removed paragraphs (in doc1 but not in doc2)
	for i, p := range paras1 {
		if p == "" {
			continue
		}
		if !paraMap2[p] {
			diff.Removed = append(diff.Removed, DiffItem{
				Type:     DiffTypeRemoved,
				Location: formatDiffLocation("paragraph", i),
				OldValue: p,
			})
			diff.Summary.RemovedParagraphs++
		}
	}

	// Find added paragraphs (in doc2 but not in doc1)
	for i, p := range paras2 {
		if p == "" {
			continue
		}
		if !paraMap1[p] {
			diff.Added = append(diff.Added, DiffItem{
				Type:     DiffTypeAdded,
				Location: formatDiffLocation("paragraph", i),
				NewValue: p,
			})
			diff.Summary.AddedParagraphs++
		}
	}

	// Find modified paragraphs using Levenshtein distance
	diff.Modified = findModifiedParagraphs(paras1, paras2, paraMap1, paraMap2)
	diff.Summary.ModifiedParagraphs = len(diff.Modified)

	// Compare tables
	tables1 := doc1.CountTables()
	tables2 := doc2.CountTables()

	if tables2 > tables1 {
		diff.Summary.AddedTables = tables2 - tables1
	} else if tables1 > tables2 {
		diff.Summary.RemovedTables = tables1 - tables2
	}

	// Update summary
	diff.Summary.TotalChanges = len(diff.Added) + len(diff.Removed) + len(diff.Modified)

	return diff
}

// DiffWith compares this document with another and returns differences.
//
//	diff := doc1.DiffWith(doc2)
func (d *DocxTmpl) DiffWith(other *DocxTmpl) *DocumentDiff {
	return CompareDocuments(d, other)
}

// HasChanges returns true if there are any differences
func (d *DocumentDiff) HasChanges() bool {
	return d.Summary.TotalChanges > 0
}

// String returns a human-readable summary of changes
func (d *DocumentDiff) String() string {
	var sb strings.Builder

	sb.WriteString("Document Comparison Summary:\n")
	sb.WriteString(strings.Repeat("-", 40) + "\n")
	sb.WriteString("Total changes: " + formatDiffInt(d.Summary.TotalChanges) + "\n")
	sb.WriteString("  Added paragraphs: " + formatDiffInt(d.Summary.AddedParagraphs) + "\n")
	sb.WriteString("  Removed paragraphs: " + formatDiffInt(d.Summary.RemovedParagraphs) + "\n")
	sb.WriteString("  Modified paragraphs: " + formatDiffInt(d.Summary.ModifiedParagraphs) + "\n")

	if d.Summary.AddedTables > 0 {
		sb.WriteString("  Added tables: " + formatDiffInt(d.Summary.AddedTables) + "\n")
	}
	if d.Summary.RemovedTables > 0 {
		sb.WriteString("  Removed tables: " + formatDiffInt(d.Summary.RemovedTables) + "\n")
	}

	return sb.String()
}

// GetChanges returns all changes as a flat list
func (d *DocumentDiff) GetChanges() []DiffItem {
	all := make([]DiffItem, 0, len(d.Added)+len(d.Removed)+len(d.Modified))
	all = append(all, d.Added...)
	all = append(all, d.Removed...)
	all = append(all, d.Modified...)
	return all
}

// CompareStats compares document statistics
//
//	statsDiff := docxtpl.CompareStats(doc1.GetStats(), doc2.GetStats())
func CompareStats(stats1, stats2 *DocumentStats) map[string]int {
	diff := make(map[string]int)

	diff["paragraphs"] = stats2.ParagraphCount - stats1.ParagraphCount
	diff["tables"] = stats2.TableCount - stats1.TableCount
	diff["words"] = stats2.WordCount - stats1.WordCount
	diff["characters"] = stats2.CharCount - stats1.CharCount
	diff["images"] = stats2.ImageCount - stats1.ImageCount
	diff["links"] = stats2.LinkCount - stats1.LinkCount

	return diff
}

// CompareMetadata compares document metadata
func CompareMetadata(meta1, meta2 *DocumentMetadata) map[string][2]string {
	diff := make(map[string][2]string)

	if meta1.Title != meta2.Title {
		diff["title"] = [2]string{meta1.Title, meta2.Title}
	}
	if meta1.Subject != meta2.Subject {
		diff["subject"] = [2]string{meta1.Subject, meta2.Subject}
	}
	if meta1.Creator != meta2.Creator {
		diff["creator"] = [2]string{meta1.Creator, meta2.Creator}
	}
	if meta1.Keywords != meta2.Keywords {
		diff["keywords"] = [2]string{meta1.Keywords, meta2.Keywords}
	}

	return diff
}

// =============================================================================
// Bookmarks
// =============================================================================

// Bookmark represents a bookmark in the document
type Bookmark struct {
	ID   int    // Bookmark ID
	Name string // Bookmark name
	Text string // Text content at the bookmark location
}

// GetBookmarks extracts all bookmarks from the document.
// Bookmarks are named locations in the document that can be linked to.
//
//	bookmarks := doc.GetBookmarks()
//	for _, b := range bookmarks {
//	    fmt.Printf("%s (ID: %d)\n", b.Name, b.ID)
//	}
func (d *DocxTmpl) GetBookmarks() []Bookmark {
	var bookmarks []Bookmark

	// Get document XML
	xmlContent, err := d.getDocumentXml()
	if err != nil || xmlContent == "" {
		return bookmarks
	}

	// Pattern to match bookmark start elements
	// <w:bookmarkStart w:id="0" w:name="MyBookmark"/>
	pattern := regexp.MustCompile(`<w:bookmarkStart[^>]*w:id="(\d+)"[^>]*w:name="([^"]*)"[^>]*/>`)

	matches := pattern.FindAllStringSubmatch(xmlContent, -1)
	for _, match := range matches {
		if len(match) >= 3 {
			var id int
			fmt.Sscanf(match[1], "%d", &id)
			name := match[2]

			// Skip internal bookmarks (like _GoBack)
			if strings.HasPrefix(name, "_") {
				continue
			}

			bookmarks = append(bookmarks, Bookmark{
				ID:   id,
				Name: name,
			})
		}
	}

	// Also check processable files
	for _, pf := range d.processableFiles {
		matches := pattern.FindAllStringSubmatch(pf.Content, -1)
		for _, match := range matches {
			if len(match) >= 3 {
				var id int
				fmt.Sscanf(match[1], "%d", &id)
				name := match[2]

				if !strings.HasPrefix(name, "_") {
					bookmarks = append(bookmarks, Bookmark{
						ID:   id,
						Name: name,
					})
				}
			}
		}
	}

	return bookmarks
}

// HasBookmark checks if a bookmark with the given name exists.
//
//	if doc.HasBookmark("Chapter1") {
//	    fmt.Println("Found Chapter1 bookmark")
//	}
func (d *DocxTmpl) HasBookmark(name string) bool {
	for _, b := range d.GetBookmarks() {
		if b.Name == name {
			return true
		}
	}
	return false
}

// GetBookmarkByName returns a bookmark by its name.
//
//	bookmark, found := doc.GetBookmarkByName("Chapter1")
func (d *DocxTmpl) GetBookmarkByName(name string) (Bookmark, bool) {
	for _, b := range d.GetBookmarks() {
		if b.Name == name {
			return b, true
		}
	}
	return Bookmark{}, false
}

// CountBookmarks returns the total number of bookmarks in the document.
//
//	count := doc.CountBookmarks()
func (d *DocxTmpl) CountBookmarks() int {
	return len(d.GetBookmarks())
}

// GetBookmarkNames returns just the names of all bookmarks.
//
//	names := doc.GetBookmarkNames()
func (d *DocxTmpl) GetBookmarkNames() []string {
	bookmarks := d.GetBookmarks()
	names := make([]string, len(bookmarks))
	for i, b := range bookmarks {
		names[i] = b.Name
	}
	return names
}

// InternalLink represents an internal hyperlink (link to bookmark)
type InternalLink struct {
	Anchor string // Bookmark name this link points to
	Text   string // Display text
}

// GetInternalLinks returns all internal hyperlinks (links to bookmarks).
//
//	links := doc.GetInternalLinks()
func (d *DocxTmpl) GetInternalLinks() []InternalLink {
	var links []InternalLink

	// Get document XML
	xmlContent, err := d.getDocumentXml()
	if err != nil || xmlContent == "" {
		return links
	}

	// Pattern to match internal hyperlinks
	// <w:hyperlink w:anchor="BookmarkName">
	pattern := regexp.MustCompile(`<w:hyperlink[^>]*w:anchor="([^"]*)"[^>]*>(.*?)</w:hyperlink>`)

	matches := pattern.FindAllStringSubmatch(xmlContent, -1)
	for _, match := range matches {
		if len(match) >= 3 {
			anchor := match[1]
			content := match[2]

			// Extract text from content
			text := extractLinkText(content)

			links = append(links, InternalLink{
				Anchor: anchor,
				Text:   text,
			})
		}
	}

	return links
}

// BookmarksSummary returns a text summary of all bookmarks.
//
//	summary := doc.BookmarksSummary()
func (d *DocxTmpl) BookmarksSummary() string {
	bookmarks := d.GetBookmarks()
	if len(bookmarks) == 0 {
		return "No bookmarks found."
	}

	var sb strings.Builder
	sb.WriteString(fmt.Sprintf("Bookmarks (%d total):\n", len(bookmarks)))
	sb.WriteString(strings.Repeat("-", 40) + "\n")

	for _, b := range bookmarks {
		sb.WriteString(fmt.Sprintf("  - %s (ID: %d)\n", b.Name, b.ID))
	}

	// Show internal links
	links := d.GetInternalLinks()
	if len(links) > 0 {
		sb.WriteString(fmt.Sprintf("\nInternal Links (%d total):\n", len(links)))
		for _, l := range links {
			text := l.Text
			if len(text) > 30 {
				text = text[:30] + "..."
			}
			sb.WriteString(fmt.Sprintf("  - '%s' -> #%s\n", text, l.Anchor))
		}
	}

	return sb.String()
}

// TableOfContentsEntry represents an entry in the table of contents
type TableOfContentsEntry struct {
	Level    int    // Entry level (1, 2, 3, etc.)
	Text     string // Entry text
	Page     string // Page number (may be empty if not updated)
	Bookmark string // Bookmark reference
}

// GetTableOfContents extracts the table of contents if present.
// Returns nil if no TOC is found.
//
//	toc := doc.GetTableOfContents()
func (d *DocxTmpl) GetTableOfContents() []TableOfContentsEntry {
	var entries []TableOfContentsEntry

	// Get document XML
	xmlContent, err := d.getDocumentXml()
	if err != nil || xmlContent == "" {
		return entries
	}

	// TOC entries are typically in SDT (Structured Document Tag) blocks
	// with w:hyperlink pointing to bookmarks like _Toc123456
	tocPattern := regexp.MustCompile(`<w:hyperlink[^>]*w:anchor="(_Toc\d+)"[^>]*>(.*?)</w:hyperlink>`)

	matches := tocPattern.FindAllStringSubmatch(xmlContent, -1)
	for _, match := range matches {
		if len(match) >= 3 {
			bookmark := match[1]
			content := match[2]

			// Extract text
			text := extractLinkText(content)

			// Determine level from style if available
			level := 1 // Default level

			entries = append(entries, TableOfContentsEntry{
				Level:    level,
				Text:     text,
				Bookmark: bookmark,
			})
		}
	}

	return entries
}

// HasTableOfContents checks if the document has a table of contents.
//
//	if doc.HasTableOfContents() {
//	    fmt.Println("Document has TOC")
//	}
func (d *DocxTmpl) HasTableOfContents() bool {
	return len(d.GetTableOfContents()) > 0
}

// =============================================================================
// Comments
// =============================================================================

// Comment represents a comment in the document
type Comment struct {
	ID        int       // Comment ID
	Author    string    // Comment author
	Initials  string    // Author initials
	Date      time.Time // When the comment was created
	Text      string    // Comment text content
	ParentID  int       // Parent comment ID for replies (-1 if not a reply)
	Paragraph int       // Paragraph index where comment is attached
}

// GetComments extracts all comments from the document.
// Returns both top-level comments and replies.
//
//	comments := doc.GetComments()
//	for _, c := range comments {
//	    fmt.Printf("[%s] %s: %s\n", c.Date.Format("2006-01-02"), c.Author, c.Text)
//	}
func (d *DocxTmpl) GetComments() []Comment {
	var comments []Comment

	// Look for comments.xml in processable files
	for _, pf := range d.processableFiles {
		if strings.HasSuffix(pf.Name, "comments.xml") {
			comments = extractComments(pf.Content)
			break
		}
	}

	return comments
}

// HasComments returns true if the document contains any comments.
//
//	if doc.HasComments() {
//	    fmt.Println("Document has comments")
//	}
func (d *DocxTmpl) HasComments() bool {
	return len(d.GetComments()) > 0
}

// CountComments returns the total number of comments in the document.
//
//	count := doc.CountComments()
func (d *DocxTmpl) CountComments() int {
	return len(d.GetComments())
}

// GetCommentsByAuthor returns comments filtered by author name.
//
//	myComments := doc.GetCommentsByAuthor("John Doe")
func (d *DocxTmpl) GetCommentsByAuthor(author string) []Comment {
	var filtered []Comment
	authorLower := strings.ToLower(author)

	for _, c := range d.GetComments() {
		if strings.ToLower(c.Author) == authorLower {
			filtered = append(filtered, c)
		}
	}
	return filtered
}

// GetCommentReplies returns all replies to a specific comment.
//
//	replies := doc.GetCommentReplies(0) // Get replies to comment with ID 0
func (d *DocxTmpl) GetCommentReplies(commentID int) []Comment {
	var replies []Comment

	for _, c := range d.GetComments() {
		if c.ParentID == commentID {
			replies = append(replies, c)
		}
	}
	return replies
}

// GetTopLevelComments returns only top-level comments (not replies).
//
//	topComments := doc.GetTopLevelComments()
func (d *DocxTmpl) GetTopLevelComments() []Comment {
	var topLevel []Comment

	for _, c := range d.GetComments() {
		if c.ParentID == -1 {
			topLevel = append(topLevel, c)
		}
	}
	return topLevel
}

// CommentsSummary returns a text summary of all comments.
//
//	summary := doc.CommentsSummary()
//	fmt.Println(summary)
func (d *DocxTmpl) CommentsSummary() string {
	comments := d.GetComments()
	if len(comments) == 0 {
		return "No comments found."
	}

	var sb strings.Builder
	sb.WriteString(fmt.Sprintf("Comments Summary (%d total)\n", len(comments)))
	sb.WriteString(strings.Repeat("-", 40) + "\n")

	// Group by author
	byAuthor := make(map[string]int)
	for _, c := range comments {
		byAuthor[c.Author]++
	}

	sb.WriteString("By author:\n")
	for author, count := range byAuthor {
		sb.WriteString(fmt.Sprintf("  %s: %d comment(s)\n", author, count))
	}

	sb.WriteString("\nComments:\n")
	for _, c := range comments {
		text := c.Text
		if len(text) > 50 {
			text = text[:50] + "..."
		}
		prefix := ""
		if c.ParentID >= 0 {
			prefix = "  [Reply] "
		}
		sb.WriteString(fmt.Sprintf("%s%s (%s): %s\n", prefix, c.Author, c.Date.Format("2006-01-02"), text))
	}

	return sb.String()
}

// DeleteAllComments removes all comments from the document.
// This modifies the comments.xml file.
//
//	doc.DeleteAllComments()
func (d *DocxTmpl) DeleteAllComments() {
	for i := range d.processableFiles {
		if strings.HasSuffix(d.processableFiles[i].Name, "comments.xml") {
			// Create an empty comments XML
			d.processableFiles[i].Content = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
</w:comments>`
			break
		}
	}
}

// GetCommentAuthors returns a list of unique comment authors.
//
//	authors := doc.GetCommentAuthors()
func (d *DocxTmpl) GetCommentAuthors() []string {
	authorSet := make(map[string]bool)

	for _, c := range d.GetComments() {
		if c.Author != "" {
			authorSet[c.Author] = true
		}
	}

	authors := make([]string, 0, len(authorSet))
	for author := range authorSet {
		authors = append(authors, author)
	}
	return authors
}

// GetCommentsInDateRange returns comments within a specific date range.
//
//	start := time.Date(2024, 1, 1, 0, 0, 0, 0, time.UTC)
//	end := time.Date(2024, 12, 31, 23, 59, 59, 0, time.UTC)
//	comments := doc.GetCommentsInDateRange(start, end)
func (d *DocxTmpl) GetCommentsInDateRange(start, end time.Time) []Comment {
	var filtered []Comment

	for _, c := range d.GetComments() {
		if !c.Date.IsZero() && !c.Date.Before(start) && !c.Date.After(end) {
			filtered = append(filtered, c)
		}
	}
	return filtered
}

// =============================================================================
// Tracked Changes
// =============================================================================

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

// =============================================================================
// Document Protection
// =============================================================================

// ProtectionType represents the type of document protection
type ProtectionType string

const (
	ProtectionNone           ProtectionType = "none"
	ProtectionReadOnly       ProtectionType = "readOnly"
	ProtectionComments       ProtectionType = "comments"       // Allow only comments
	ProtectionTrackedChanges ProtectionType = "trackedChanges" // Allow only tracked changes
	ProtectionForms          ProtectionType = "forms"          // Allow only form filling
)

// ProtectionInfo contains information about document protection
type ProtectionInfo struct {
	IsProtected    bool           // Whether the document is protected
	Type           ProtectionType // Type of protection
	HasPassword    bool           // Whether a password is set
	EnforceMessage string         // Enforcement message if any
}

// GetProtectionInfo returns information about document protection.
//
//	info := doc.GetProtectionInfo()
//	if info.IsProtected {
//	    fmt.Printf("Document is protected: %s\n", info.Type)
//	}
func (d *DocxTmpl) GetProtectionInfo() *ProtectionInfo {
	info := &ProtectionInfo{
		Type: ProtectionNone,
	}

	// Look for settings.xml
	var settingsContent string
	for _, pf := range d.processableFiles {
		if strings.HasSuffix(pf.Name, "settings.xml") {
			settingsContent = pf.Content
			break
		}
	}

	if settingsContent == "" {
		return info
	}

	// Check for document protection element
	// <w:documentProtection w:edit="readOnly" w:enforcement="1" w:cryptProviderType="..." />
	protPattern := regexp.MustCompile(`<w:documentProtection([^>]*)/>`)
	match := protPattern.FindStringSubmatch(settingsContent)

	if len(match) >= 2 {
		attrs := match[1]
		info.IsProtected = strings.Contains(attrs, `w:enforcement="1"`)

		// Check protection type
		if strings.Contains(attrs, `w:edit="readOnly"`) {
			info.Type = ProtectionReadOnly
		} else if strings.Contains(attrs, `w:edit="comments"`) {
			info.Type = ProtectionComments
		} else if strings.Contains(attrs, `w:edit="trackedChanges"`) {
			info.Type = ProtectionTrackedChanges
		} else if strings.Contains(attrs, `w:edit="forms"`) {
			info.Type = ProtectionForms
		}

		// Check if password is set
		info.HasPassword = strings.Contains(attrs, "w:cryptProviderType") ||
			strings.Contains(attrs, "w:hash") ||
			strings.Contains(attrs, "w:salt")
	}

	return info
}

// IsProtected returns true if the document has any protection enabled.
//
//	if doc.IsProtected() {
//	    fmt.Println("Document is protected")
//	}
func (d *DocxTmpl) IsProtected() bool {
	return d.GetProtectionInfo().IsProtected
}

// IsReadOnly returns true if the document is read-only protected.
//
//	if doc.IsReadOnly() {
//	    fmt.Println("Document is read-only")
//	}
func (d *DocxTmpl) IsReadOnly() bool {
	info := d.GetProtectionInfo()
	return info.IsProtected && info.Type == ProtectionReadOnly
}

// SetProtection sets document protection.
// Note: This sets protection without a password. For password protection,
// use the Word application or a dedicated library.
//
//	err := doc.SetProtection(docxtpl.ProtectionReadOnly)
func (d *DocxTmpl) SetProtection(protType ProtectionType) error {
	// Find settings.xml
	for i := range d.processableFiles {
		if strings.HasSuffix(d.processableFiles[i].Name, "settings.xml") {
			content := d.processableFiles[i].Content

			// Remove existing protection
			content = removeProtection(content)

			if protType != ProtectionNone {
				// Add new protection before </w:settings>
				protXML := fmt.Sprintf(`<w:documentProtection w:edit="%s" w:enforcement="1"/>`, protType)
				content = strings.Replace(content, "</w:settings>", protXML+"</w:settings>", 1)
			}

			d.processableFiles[i].Content = content
			return nil
		}
	}

	return fmt.Errorf("settings.xml not found")
}

// RemoveProtection removes document protection (without password).
// Note: This only works for documents without password protection.
//
//	err := doc.RemoveProtection()
func (d *DocxTmpl) RemoveProtection() error {
	return d.SetProtection(ProtectionNone)
}

// SetReadOnly sets the document to read-only mode.
//
//	err := doc.SetReadOnly()
func (d *DocxTmpl) SetReadOnly() error {
	return d.SetProtection(ProtectionReadOnly)
}

// AllowOnlyComments sets protection to allow only comments.
//
//	err := doc.AllowOnlyComments()
func (d *DocxTmpl) AllowOnlyComments() error {
	return d.SetProtection(ProtectionComments)
}

// AllowOnlyTrackedChanges sets protection to allow only tracked changes.
//
//	err := doc.AllowOnlyTrackedChanges()
func (d *DocxTmpl) AllowOnlyTrackedChanges() error {
	return d.SetProtection(ProtectionTrackedChanges)
}

// AllowOnlyFormFilling sets protection to allow only form filling.
//
//	err := doc.AllowOnlyFormFilling()
func (d *DocxTmpl) AllowOnlyFormFilling() error {
	return d.SetProtection(ProtectionForms)
}

// ProtectionSummary returns a text summary of protection status.
//
//	summary := doc.ProtectionSummary()
func (d *DocxTmpl) ProtectionSummary() string {
	info := d.GetProtectionInfo()

	if !info.IsProtected {
		return "Document is not protected."
	}

	var sb strings.Builder
	sb.WriteString("Document Protection:\n")
	sb.WriteString(strings.Repeat("-", 30) + "\n")
	sb.WriteString("Protected: Yes\n")
	sb.WriteString(fmt.Sprintf("Type: %s\n", info.Type))
	sb.WriteString(fmt.Sprintf("Password: %v\n", info.HasPassword))

	return sb.String()
}

// RestrictionInfo provides detailed restriction information
type RestrictionInfo struct {
	CanEdit       bool // Can edit the document
	CanComment    bool // Can add comments
	CanTrack      bool // Can make tracked changes
	CanFillForms  bool // Can fill forms
	CanFormatText bool // Can format text
}

// GetRestrictions returns detailed information about what actions are allowed.
//
//	restrictions := doc.GetRestrictions()
//	if restrictions.CanComment {
//	    fmt.Println("Comments are allowed")
//	}
func (d *DocxTmpl) GetRestrictions() *RestrictionInfo {
	info := d.GetProtectionInfo()

	if !info.IsProtected {
		return &RestrictionInfo{
			CanEdit:       true,
			CanComment:    true,
			CanTrack:      true,
			CanFillForms:  true,
			CanFormatText: true,
		}
	}

	switch info.Type {
	case ProtectionReadOnly:
		return &RestrictionInfo{
			CanEdit:       false,
			CanComment:    false,
			CanTrack:      false,
			CanFillForms:  false,
			CanFormatText: false,
		}
	case ProtectionComments:
		return &RestrictionInfo{
			CanEdit:       false,
			CanComment:    true,
			CanTrack:      false,
			CanFillForms:  false,
			CanFormatText: false,
		}
	case ProtectionTrackedChanges:
		return &RestrictionInfo{
			CanEdit:       false,
			CanComment:    true,
			CanTrack:      true,
			CanFillForms:  false,
			CanFormatText: false,
		}
	case ProtectionForms:
		return &RestrictionInfo{
			CanEdit:       false,
			CanComment:    false,
			CanTrack:      false,
			CanFillForms:  true,
			CanFormatText: false,
		}
	default:
		return &RestrictionInfo{
			CanEdit:       true,
			CanComment:    true,
			CanTrack:      true,
			CanFillForms:  true,
			CanFormatText: true,
		}
	}
}

// =============================================================================
// Document Export
// =============================================================================

// StructuredDocument represents the document in a structured format suitable for AI consumption
type StructuredDocument struct {
	Metadata   DocumentMetadata      `json:"metadata"`
	Stats      DocumentStats         `json:"stats"`
	Outline    []OutlineItem         `json:"outline"`
	Paragraphs []StructuredParagraph `json:"paragraphs"`
	Tables     []StructuredTable     `json:"tables"`
	Images     []ImageInfo           `json:"images"`
	Links      []HyperlinkInfo       `json:"links"`
}

// StructuredParagraph represents a paragraph with its metadata
type StructuredParagraph struct {
	Index     int             `json:"index"`
	Text      string          `json:"text"`
	Style     string          `json:"style"`
	IsList    bool            `json:"is_list,omitempty"`
	ListLevel int             `json:"list_level,omitempty"`
	Alignment string          `json:"alignment,omitempty"`
	Runs      []StructuredRun `json:"runs,omitempty"`
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
		Metadata: *d.GetMetadata(),
		Stats:    *d.GetStats(),
		Outline:  d.GetOutline(),
		Links:    d.GetAllHyperlinks(),
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

// =============================================================================
// XML Access
// =============================================================================

// XMLFile represents an XML file from the DOCX archive
type XMLFile struct {
	Path    string // File path within the archive (e.g., "word/document.xml")
	Content string // XML content
}

// UnpackedDocument represents an unpacked DOCX document
type UnpackedDocument struct {
	Files map[string][]byte // All files in the archive
}

// UnpackToDirectory extracts the DOCX file to a directory.
// This allows direct access to all XML files for advanced manipulation.
//
//	err := doc.UnpackToDirectory("/tmp/unpacked")
func (d *DocxTmpl) UnpackToDirectory(dirPath string) error {
	// Create the directory
	if err := os.MkdirAll(dirPath, 0755); err != nil {
		return fmt.Errorf("failed to create directory: %w", err)
	}

	// Get the document as bytes
	var buf bytes.Buffer
	if err := d.Save(&buf); err != nil {
		return fmt.Errorf("failed to save document: %w", err)
	}

	// Open as zip
	reader := bytes.NewReader(buf.Bytes())
	zipReader, err := zip.NewReader(reader, int64(buf.Len()))
	if err != nil {
		return fmt.Errorf("failed to read zip: %w", err)
	}

	// Extract all files
	for _, f := range zipReader.File {
		targetPath := filepath.Join(dirPath, f.Name)

		// Create directory if needed
		if f.FileInfo().IsDir() {
			if err := os.MkdirAll(targetPath, 0755); err != nil {
				return err
			}
			continue
		}

		// Ensure parent directory exists
		if err := os.MkdirAll(filepath.Dir(targetPath), 0755); err != nil {
			return err
		}

		// Extract file
		rc, err := f.Open()
		if err != nil {
			return err
		}

		outFile, err := os.Create(targetPath)
		if err != nil {
			rc.Close()
			return err
		}

		_, err = io.Copy(outFile, rc)
		outFile.Close()
		rc.Close()

		if err != nil {
			return err
		}
	}

	return nil
}

// PackFromDirectory creates a DOCX file from an unpacked directory.
// This is the reverse of UnpackToDirectory.
//
//	doc, err := docxtpl.PackFromDirectory("/tmp/unpacked")
func PackFromDirectory(dirPath string) (*DocxTmpl, error) {
	// Create a buffer to write the zip
	var buf bytes.Buffer
	zipWriter := zip.NewWriter(&buf)

	// Walk the directory
	err := filepath.Walk(dirPath, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}

		// Skip directories
		if info.IsDir() {
			return nil
		}

		// Get relative path
		relPath, err := filepath.Rel(dirPath, path)
		if err != nil {
			return err
		}

		// Use forward slashes for zip paths
		relPath = strings.ReplaceAll(relPath, string(os.PathSeparator), "/")

		// Create file in zip
		writer, err := zipWriter.Create(relPath)
		if err != nil {
			return err
		}

		// Read and write content
		content, err := os.ReadFile(path)
		if err != nil {
			return err
		}

		_, err = writer.Write(content)
		return err
	})

	if err != nil {
		return nil, fmt.Errorf("failed to walk directory: %w", err)
	}

	if err := zipWriter.Close(); err != nil {
		return nil, fmt.Errorf("failed to close zip: %w", err)
	}

	// Parse the packed document
	return ParseFromBytes(buf.Bytes())
}

// GetXMLFiles returns a list of all XML files in the document.
//
//	files := doc.GetXMLFiles()
//	for _, f := range files {
//	    fmt.Println(f)
//	}
func (d *DocxTmpl) GetXMLFiles() []string {
	var files []string

	var buf bytes.Buffer
	if err := d.Save(&buf); err != nil {
		return files
	}

	reader := bytes.NewReader(buf.Bytes())
	zipReader, err := zip.NewReader(reader, int64(buf.Len()))
	if err != nil {
		return files
	}

	for _, f := range zipReader.File {
		if strings.HasSuffix(f.Name, ".xml") {
			files = append(files, f.Name)
		}
	}

	sort.Strings(files)
	return files
}

// GetXMLContent returns the content of a specific XML file within the document.
//
//	content, err := doc.GetXMLContent("word/document.xml")
func (d *DocxTmpl) GetXMLContent(filePath string) (string, error) {
	// Check processable files first
	for _, pf := range d.processableFiles {
		if pf.Name == filePath {
			return pf.Content, nil
		}
	}

	// Otherwise, extract from the document
	var buf bytes.Buffer
	if err := d.Save(&buf); err != nil {
		return "", fmt.Errorf("failed to save document: %w", err)
	}

	reader := bytes.NewReader(buf.Bytes())
	zipReader, err := zip.NewReader(reader, int64(buf.Len()))
	if err != nil {
		return "", fmt.Errorf("failed to read zip: %w", err)
	}

	for _, f := range zipReader.File {
		if f.Name == filePath {
			rc, err := f.Open()
			if err != nil {
				return "", err
			}
			defer rc.Close()

			content, err := io.ReadAll(rc)
			if err != nil {
				return "", err
			}
			return string(content), nil
		}
	}

	return "", fmt.Errorf("file not found: %s", filePath)
}

// SetXMLContent sets the content of a processable XML file.
// Only works for files in the processable files list (headers, footers, settings, etc.)
//
//	err := doc.SetXMLContent("word/settings.xml", newContent)
func (d *DocxTmpl) SetXMLContent(filePath string, content string) error {
	for i := range d.processableFiles {
		if d.processableFiles[i].Name == filePath {
			d.processableFiles[i].Content = content
			return nil
		}
	}
	return fmt.Errorf("file not in processable files: %s", filePath)
}

// GetRelationships returns the document relationships.
//
//	rels, err := doc.GetRelationships()
func (d *DocxTmpl) GetRelationships() (string, error) {
	return d.GetXMLContent("word/_rels/document.xml.rels")
}

// GetContentTypesXML returns the content types XML.
//
//	types, err := doc.GetContentTypesXML()
func (d *DocxTmpl) GetContentTypesXML() (string, error) {
	return d.GetXMLContent("[Content_Types].xml")
}

// GetDocumentXML returns the main document XML.
//
//	xml, err := doc.GetDocumentXML()
func (d *DocxTmpl) GetDocumentXML() (string, error) {
	return d.getDocumentXml()
}

// GetSettingsXML returns the settings XML.
//
//	settings, err := doc.GetSettingsXML()
func (d *DocxTmpl) GetSettingsXML() (string, error) {
	for _, pf := range d.processableFiles {
		if strings.HasSuffix(pf.Name, "settings.xml") {
			return pf.Content, nil
		}
	}
	return d.GetXMLContent("word/settings.xml")
}

// GetStylesXML returns the styles XML.
//
//	styles, err := doc.GetStylesXML()
func (d *DocxTmpl) GetStylesXML() (string, error) {
	return d.GetXMLContent("word/styles.xml")
}

// ArchiveInfo returns information about the DOCX archive.
type ArchiveInfo struct {
	TotalFiles   int      // Total number of files in archive
	XMLFiles     int      // Number of XML files
	MediaFiles   int      // Number of media files
	TotalSize    int64    // Total uncompressed size
	FileList     []string // List of all file paths
	HasComments  bool     // Whether the document has comments.xml
	HasSettings  bool     // Whether the document has settings.xml
	HasFootnotes bool     // Whether the document has footnotes.xml
	HasEndnotes  bool     // Whether the document has endnotes.xml
}

// GetArchiveInfo returns information about the DOCX archive structure.
//
//	info := doc.GetArchiveInfo()
//	fmt.Printf("Total files: %d\n", info.TotalFiles)
func (d *DocxTmpl) GetArchiveInfo() *ArchiveInfo {
	info := &ArchiveInfo{}

	var buf bytes.Buffer
	if err := d.Save(&buf); err != nil {
		return info
	}

	reader := bytes.NewReader(buf.Bytes())
	zipReader, err := zip.NewReader(reader, int64(buf.Len()))
	if err != nil {
		return info
	}

	for _, f := range zipReader.File {
		info.TotalFiles++
		info.TotalSize += int64(f.UncompressedSize64)
		info.FileList = append(info.FileList, f.Name)

		if strings.HasSuffix(f.Name, ".xml") {
			info.XMLFiles++
		}
		if strings.HasPrefix(f.Name, "word/media/") {
			info.MediaFiles++
		}
		if f.Name == "word/comments.xml" {
			info.HasComments = true
		}
		if f.Name == "word/settings.xml" {
			info.HasSettings = true
		}
		if f.Name == "word/footnotes.xml" {
			info.HasFootnotes = true
		}
		if f.Name == "word/endnotes.xml" {
			info.HasEndnotes = true
		}
	}

	sort.Strings(info.FileList)
	return info
}

// =============================================================================
// Helper Functions
// =============================================================================

func formatDiffLocation(itemType string, index int) string {
	return itemType + " " + formatDiffInt(index+1)
}

func formatDiffInt(n int) string {
	return fmt.Sprintf("%d", n)
}

func findModifiedParagraphs(paras1, paras2 []string, map1, map2 map[string]bool) []DiffItem {
	var modified []DiffItem

	// For each removed paragraph, find if there's a similar one in added
	removed := []string{}
	added := []string{}

	for _, p := range paras1 {
		if p != "" && !map2[p] {
			removed = append(removed, p)
		}
	}
	for _, p := range paras2 {
		if p != "" && !map1[p] {
			added = append(added, p)
		}
	}

	// Find pairs with high similarity
	usedAdded := make(map[int]bool)
	for i, old := range removed {
		bestMatch := -1
		bestScore := 0.5 // Minimum similarity threshold

		for j, new := range added {
			if usedAdded[j] {
				continue
			}

			score := similarity(old, new)
			if score > bestScore {
				bestScore = score
				bestMatch = j
			}
		}

		if bestMatch >= 0 {
			modified = append(modified, DiffItem{
				Type:     DiffTypeModified,
				Location: formatDiffLocation("paragraph", i),
				OldValue: old,
				NewValue: added[bestMatch],
			})
			usedAdded[bestMatch] = true
		}
	}

	return modified
}

// similarity calculates a simple similarity score between two strings
func similarity(s1, s2 string) float64 {
	if s1 == s2 {
		return 1.0
	}
	if len(s1) == 0 || len(s2) == 0 {
		return 0.0
	}

	// Use word overlap for similarity
	words1 := strings.Fields(strings.ToLower(s1))
	words2 := strings.Fields(strings.ToLower(s2))

	if len(words1) == 0 || len(words2) == 0 {
		return 0.0
	}

	// Count common words
	wordSet := make(map[string]bool)
	for _, w := range words1 {
		wordSet[w] = true
	}

	common := 0
	for _, w := range words2 {
		if wordSet[w] {
			common++
		}
	}

	// Jaccard similarity
	total := len(words1) + len(words2) - common
	if total == 0 {
		return 0.0
	}

	return float64(common) / float64(total)
}

func extractLinkText(content string) string {
	var text strings.Builder
	textPattern := regexp.MustCompile(`<w:t[^>]*>([^<]*)</w:t>`)
	matches := textPattern.FindAllStringSubmatch(content, -1)
	for _, match := range matches {
		if len(match) >= 2 {
			text.WriteString(match[1])
		}
	}
	return text.String()
}

func extractComments(xml string) []Comment {
	var comments []Comment

	// Pattern to match comment elements
	// <w:comment w:id="0" w:author="John" w:initials="J" w:date="2024-01-15T10:30:00Z">
	commentPattern := regexp.MustCompile(`<w:comment[^>]*w:id="(\d+)"[^>]*w:author="([^"]*)"[^>]*(?:w:initials="([^"]*)")?[^>]*(?:w:date="([^"]*)")?[^>]*>(.*?)</w:comment>`)

	matches := commentPattern.FindAllStringSubmatch(xml, -1)
	for _, match := range matches {
		if len(match) >= 6 {
			var id int
			fmt.Sscanf(match[1], "%d", &id)

			author := match[2]
			initials := match[3]
			dateStr := match[4]
			content := match[5]

			// Parse date
			var date time.Time
			if dateStr != "" {
				date, _ = time.Parse(time.RFC3339, dateStr)
			}

			// Extract text from content
			text := extractCommentText(content)

			comments = append(comments, Comment{
				ID:       id,
				Author:   author,
				Initials: initials,
				Date:     date,
				Text:     text,
				ParentID: -1, // Will be set later if it's a reply
			})
		}
	}

	return comments
}

func extractCommentText(content string) string {
	var text strings.Builder

	// Extract text from <w:t> elements
	textPattern := regexp.MustCompile(`<w:t[^>]*>([^<]*)</w:t>`)
	matches := textPattern.FindAllStringSubmatch(content, -1)

	for _, match := range matches {
		if len(match) >= 2 {
			text.WriteString(match[1])
		}
	}

	return strings.TrimSpace(text.String())
}

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
			text := extractTextFromTrackedXML(content, changeType)

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

func extractTextFromTrackedXML(xml string, changeType TrackedChangeType) string {
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

func removeProtection(content string) string {
	protPattern := regexp.MustCompile(`<w:documentProtection[^>]*/>`)
	return protPattern.ReplaceAllString(content, "")
}

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
			if size, err := parseExportSize(r.RunProperties.Size.Val); err == nil {
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
			cellText := getExportCellText(cell)
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
			cellText := getExportCellText(cell)
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
			cellText := escapeHTML(getExportCellText(cell))
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

func getExportCellText(cell *docx.WTableCell) string {
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

func parseExportSize(s string) (int, error) {
	var size int
	_, err := fmt.Sscanf(s, "%d", &size)
	return size, err
}
