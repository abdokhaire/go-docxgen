package docxtpl

import (
	"strings"
)

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
	Added     []DiffItem // Content added in the new document
	Removed   []DiffItem // Content removed from the original
	Modified  []DiffItem // Content that was modified
	Summary   DiffSummary
}

// DiffSummary contains summary statistics of differences
type DiffSummary struct {
	TotalChanges      int
	AddedParagraphs   int
	RemovedParagraphs int
	ModifiedParagraphs int
	AddedTables       int
	RemovedTables     int
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
				Location: formatLocation("paragraph", i),
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
				Location: formatLocation("paragraph", i),
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
	sb.WriteString("Total changes: " + formatInt(d.Summary.TotalChanges) + "\n")
	sb.WriteString("  Added paragraphs: " + formatInt(d.Summary.AddedParagraphs) + "\n")
	sb.WriteString("  Removed paragraphs: " + formatInt(d.Summary.RemovedParagraphs) + "\n")
	sb.WriteString("  Modified paragraphs: " + formatInt(d.Summary.ModifiedParagraphs) + "\n")

	if d.Summary.AddedTables > 0 {
		sb.WriteString("  Added tables: " + formatInt(d.Summary.AddedTables) + "\n")
	}
	if d.Summary.RemovedTables > 0 {
		sb.WriteString("  Removed tables: " + formatInt(d.Summary.RemovedTables) + "\n")
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

// Helper functions

func formatLocation(itemType string, index int) string {
	return itemType + " " + formatInt(index+1)
}

func formatInt(n int) string {
	return strings.TrimSpace(strings.Replace("    "+string(rune('0'+n%10)), "    ", "", 1))
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
				Location: formatLocation("paragraph", i),
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
