package docxtpl

import (
	"fmt"
	"regexp"
	"strings"
)

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

// Helper function
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
