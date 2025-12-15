package xmlutils

import (
	"html"
	"regexp"
	"strings"

	"github.com/dlclark/regexp2"
)

func PrepareXmlForTagReplacement(xmlString string) (string, error) {
	// First, decode XML entities within template tags
	// Word often encodes " as &#34; or &quot; which breaks template parsing
	xmlString = decodeEntitiesInTemplateTags(xmlString)

	newXmlString, err := replaceTableRangeRows(xmlString)

	return newXmlString, err
}

// decodeEntitiesInTemplateTags decodes XML/HTML entities within template tags.
// Word often encodes characters like " as &#34; or &quot; which breaks template parsing.
func decodeEntitiesInTemplateTags(xmlString string) string {
	var result strings.Builder
	i := 0

	for i < len(xmlString) {
		// Look for start of template tag
		if i+1 < len(xmlString) && xmlString[i] == '{' && xmlString[i+1] == '{' {
			// Find the end of the template tag
			start := i
			i += 2
			depth := 1

			for i < len(xmlString) && depth > 0 {
				if i+1 < len(xmlString) && xmlString[i] == '{' && xmlString[i+1] == '{' {
					depth++
					i += 2
				} else if i+1 < len(xmlString) && xmlString[i] == '}' && xmlString[i+1] == '}' {
					depth--
					i += 2
				} else {
					i++
				}
			}

			// Extract the tag and decode HTML entities
			tag := xmlString[start:i]
			decoded := html.UnescapeString(tag)
			result.WriteString(decoded)
		} else {
			result.WriteByte(xmlString[i])
			i++
		}
	}

	return result.String()
}

// MergeFragmentedTagsInXml merges template tags that are split across multiple
// <w:t> elements in the XML. This handles cases where WordprocessingML fragments
// placeholders like {{.FirstName}} into {{.First and Name}} across different runs.
func MergeFragmentedTagsInXml(xmlString string) string {
	// Regex to match <w:t> elements with their content (including attributes like xml:space)
	textTagRegex := regexp.MustCompile(`(<w:t[^>]*>)(.*?)(</w:t>)`)

	// Find all matches
	matches := textTagRegex.FindAllStringSubmatchIndex(xmlString, -1)
	if len(matches) == 0 {
		return xmlString
	}

	// Extract text contents and their positions
	type textMatch struct {
		fullStart   int
		fullEnd     int
		openTag     string
		content     string
		closeTag    string
		contentStart int
		contentEnd  int
	}

	var textMatches []textMatch
	for _, match := range matches {
		tm := textMatch{
			fullStart:    match[0],
			fullEnd:      match[1],
			openTag:      xmlString[match[2]:match[3]],
			content:      xmlString[match[4]:match[5]],
			closeTag:     xmlString[match[6]:match[7]],
			contentStart: match[4],
			contentEnd:   match[5],
		}
		textMatches = append(textMatches, tm)
	}

	// Check for incomplete tags and merge them
	result := strings.Builder{}
	lastEnd := 0
	i := 0

	for i < len(textMatches) {
		tm := textMatches[i]

		// Check if this text contains an incomplete opening tag
		if containsIncompleteOpenTag(tm.content) {
			// Accumulate text until we find the closing
			accumulated := tm.content
			startIdx := i
			j := i + 1

			for j < len(textMatches) && !containsCompleteTag(accumulated) {
				accumulated += textMatches[j].content
				j++
			}

			if containsCompleteTag(accumulated) {
				// Write everything before this match
				result.WriteString(xmlString[lastEnd:textMatches[startIdx].contentStart])

				// Write accumulated text in the first text node
				result.WriteString(accumulated)
				result.WriteString(textMatches[startIdx].closeTag)

				// Empty out the other text nodes
				for k := startIdx + 1; k < j; k++ {
					result.WriteString(xmlString[textMatches[k-1].fullEnd:textMatches[k].contentStart])
					// Write empty content
					result.WriteString(textMatches[k].closeTag)
				}

				lastEnd = textMatches[j-1].fullEnd
				i = j
				continue
			}
		}

		i++
	}

	// Write remaining content
	result.WriteString(xmlString[lastEnd:])

	return result.String()
}

// containsIncompleteOpenTag checks if text has an opening {{ without a closing }}
// Also handles whitespace-trimming syntax {{- and -}}
func containsIncompleteOpenTag(text string) bool {
	// Count standard and trimming delimiters
	openCount := strings.Count(text, "{{")
	closeCount := strings.Count(text, "}}")

	if openCount > closeCount {
		return true
	}

	// Also check for partial opening like just "{"
	lastOpen := strings.LastIndex(text, "{{")
	lastClose := strings.LastIndex(text, "}}")

	if lastOpen > lastClose {
		return true
	}

	// Check for single { at end that might be start of {{
	if strings.HasSuffix(text, "{") && !strings.HasSuffix(text, "{{") {
		return true
	}

	// Check for incomplete closing (single } at end)
	if strings.HasSuffix(text, "}") && !strings.HasSuffix(text, "}}") {
		// But make sure it's part of a tag, not just a random }
		if openCount > 0 {
			return true
		}
	}

	// Check for trimming syntax that might be split
	// e.g., "{{-" split as "{{" and "-"
	if strings.HasSuffix(text, "-") {
		// Check if there's an unclosed tag before the -
		beforeDash := text[:len(text)-1]
		if strings.Count(beforeDash, "{{") > strings.Count(beforeDash, "}}") {
			return true
		}
	}

	return false
}

// containsCompleteTag checks if text has balanced {{ and }}
// Handles multiple tags and nested structures
func containsCompleteTag(text string) bool {
	openCount := strings.Count(text, "{{")
	closeCount := strings.Count(text, "}}")

	if openCount == 0 {
		return false
	}

	// Basic balance check
	if openCount != closeCount {
		return false
	}

	// Verify tags are properly formed (no partial delimiters at boundaries)
	// Check for incomplete tag at the end
	trimmed := strings.TrimRight(text, " \t\n\r")
	if strings.HasSuffix(trimmed, "{") {
		return false
	}

	return true
}

// containsAnyTemplateContent checks if text contains any template-like content
// that might need merging (partial delimiters or complete tags)
func containsAnyTemplateContent(text string) bool {
	return strings.Contains(text, "{") || strings.Contains(text, "}")
}

var tableRangeRowRegex = regexp2.MustCompile("<w:tr>(?:(?!<w:tr>).)*?({{range .*?}}|{{ range .*? }}|{{end}}|{{ end }})(?:(?!<w:tr>).)*?</w:tr>", 0)

func replaceTableRangeRows(xmlString string) (string, error) {
	tableRangeRowRegex.MatchTimeout = 500

	newXmlString := xmlString

	m, err := tableRangeRowRegex.FindStringMatch(xmlString)
	if err != nil {
		return "", err
	}
	for m != nil {
		gps := m.Groups()
		newXmlString = strings.Replace(newXmlString, m.String(), gps[1].Captures[0].String(), 1)
		m, _ = tableRangeRowRegex.FindNextMatch(m)
	}

	return newXmlString, nil
}

func FixXmlIssuesPostTagReplacement(xmlString string) string {
	// CRITICAL: Replace "<no value>" which Go templates output for nil/missing values
	// This contains literal < and > characters that break XML parsing
	// Note: missingkey=zero only handles missing map keys, not nil field access or chained nil access
	xmlString = strings.ReplaceAll(xmlString, "<no value>", "")

	// Fix issues with drawings in text nodes
	xmlString = strings.ReplaceAll(xmlString, "<w:t><w:drawing>", "<w:drawing>")
	xmlString = strings.ReplaceAll(xmlString, "</w:drawing></w:t>", "</w:drawing>")

	// Clean up empty text elements that may result from formatting functions
	xmlString = strings.ReplaceAll(xmlString, "<w:t></w:t>", "")
	xmlString = strings.ReplaceAll(xmlString, `<w:t xml:space="preserve"></w:t>`, "")

	// Clean up empty runs that may result from formatting
	xmlString = strings.ReplaceAll(xmlString, "<w:r></w:r>", "")
	xmlString = strings.ReplaceAll(xmlString, "<w:r><w:rPr></w:rPr></w:r>", "")

	return xmlString
}
