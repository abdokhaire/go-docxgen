package hyperlinks

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"regexp"
	"strings"
	"sync"
)

// HyperlinkRegistry tracks hyperlinks and their relationship IDs
type HyperlinkRegistry struct {
	mu        sync.RWMutex
	links     map[string]string // URL -> rId
	nextID    int
	existingIDs map[string]bool // Track existing rIds to avoid conflicts
}

// NewHyperlinkRegistry creates a new hyperlink registry
func NewHyperlinkRegistry() *HyperlinkRegistry {
	return &HyperlinkRegistry{
		links:       make(map[string]string),
		nextID:      100, // Start high to avoid conflicts with existing IDs
		existingIDs: make(map[string]bool),
	}
}

// RegisterLink registers a URL and returns its relationship ID
func (r *HyperlinkRegistry) RegisterLink(url string) string {
	r.mu.Lock()
	defer r.mu.Unlock()

	// Check if URL already registered
	if rId, exists := r.links[url]; exists {
		return rId
	}

	// Generate new ID
	rId := fmt.Sprintf("rIdLink%d", r.nextID)
	r.nextID++
	r.links[url] = rId
	return rId
}

// GetLinks returns all registered links (URL -> rId)
func (r *HyperlinkRegistry) GetLinks() map[string]string {
	r.mu.RLock()
	defer r.mu.RUnlock()

	result := make(map[string]string)
	for url, rId := range r.links {
		result[url] = rId
	}
	return result
}

// HasLinks returns true if any hyperlinks have been registered
func (r *HyperlinkRegistry) HasLinks() bool {
	r.mu.RLock()
	defer r.mu.RUnlock()
	return len(r.links) > 0
}

// Relationships represents the XML structure of a .rels file
type Relationships struct {
	XMLName       xml.Name       `xml:"Relationships"`
	Xmlns         string         `xml:"xmlns,attr"`
	Relationships []Relationship `xml:"Relationship"`
}

// Relationship represents a single relationship entry
type Relationship struct {
	ID         string `xml:"Id,attr"`
	Type       string `xml:"Type,attr"`
	Target     string `xml:"Target,attr"`
	TargetMode string `xml:"TargetMode,attr,omitempty"`
}

const (
	RelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships"
	HyperlinkType          = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
)

// GetDocumentRels reads the document.xml.rels file from a zip
func GetDocumentRels(reader io.ReaderAt, size int64) (*Relationships, error) {
	zipReader, err := zip.NewReader(reader, size)
	if err != nil {
		return nil, err
	}

	for _, f := range zipReader.File {
		if f.Name == "word/_rels/document.xml.rels" {
			zf, err := f.Open()
			if err != nil {
				return nil, err
			}
			defer zf.Close()

			content, err := io.ReadAll(zf)
			if err != nil {
				return nil, err
			}

			var rels Relationships
			if err := xml.Unmarshal(content, &rels); err != nil {
				return nil, err
			}
			return &rels, nil
		}
	}

	// If no rels file exists, return empty structure
	return &Relationships{
		Xmlns:         RelationshipsNamespace,
		Relationships: []Relationship{},
	}, nil
}

// AddHyperlinks adds hyperlink relationships to the relationships structure
func (r *Relationships) AddHyperlinks(links map[string]string) {
	for url, rId := range links {
		r.Relationships = append(r.Relationships, Relationship{
			ID:         rId,
			Type:       HyperlinkType,
			Target:     url,
			TargetMode: "External",
		})
	}
}

// ToXML returns the XML representation of the relationships as a string
func (r *Relationships) ToXML() (string, error) {
	output, err := xml.MarshalIndent(r, "", "  ")
	if err != nil {
		return "", err
	}
	return xml.Header + string(output), nil
}

// HyperlinkXML generates the Word XML for a hyperlink
// rId is the relationship ID that maps to the URL in document.xml.rels
func HyperlinkXML(rId, text string) string {
	return fmt.Sprintf(
		`</w:t></w:r><w:hyperlink r:id="%s" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr><w:t>%s</w:t></w:r></w:hyperlink><w:r><w:t>`,
		rId, text)
}

// rIdRegex matches relationship IDs in the XML
var rIdRegex = regexp.MustCompile(`r:id="(rId\d+)"`)

// FindExistingRIds finds all existing relationship IDs in content
func FindExistingRIds(content string) []string {
	matches := rIdRegex.FindAllStringSubmatch(content, -1)
	var ids []string
	for _, match := range matches {
		if len(match) > 1 {
			ids = append(ids, match[1])
		}
	}
	return ids
}

// ProcessRelationshipsFile updates or creates the document.xml.rels file
func ProcessRelationshipsFile(existingContent string, links map[string]string) (string, error) {
	var rels Relationships

	if existingContent != "" {
		// Parse existing content
		if err := xml.Unmarshal([]byte(existingContent), &rels); err != nil {
			// If parsing fails, create new
			rels = Relationships{
				Xmlns:         RelationshipsNamespace,
				Relationships: []Relationship{},
			}
		}
	} else {
		rels = Relationships{
			Xmlns:         RelationshipsNamespace,
			Relationships: []Relationship{},
		}
	}

	// Ensure xmlns is set
	if rels.Xmlns == "" {
		rels.Xmlns = RelationshipsNamespace
	}

	// Add hyperlink relationships
	rels.AddHyperlinks(links)

	return rels.ToXML()
}

// EnsureRelsFolder checks if _rels folder path is correct
func GetRelsPath(documentPath string) string {
	// word/document.xml -> word/_rels/document.xml.rels
	dir := ""
	file := documentPath
	if idx := strings.LastIndex(documentPath, "/"); idx != -1 {
		dir = documentPath[:idx+1]
		file = documentPath[idx+1:]
	}
	return dir + "_rels/" + file + ".rels"
}
