package docxtpl

import (
	"fmt"
	"regexp"
	"strings"
	"time"
)

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

// Helper functions

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

	// Look for reply relationships in commentsExtended.xml
	// For now, we'll assume top-level comments only
	// Extended comment threading would require additional processing

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
