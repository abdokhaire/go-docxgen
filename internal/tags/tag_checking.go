package tags

import (
	"regexp"
	"strings"
)

// tagRegex matches complete template tags including whitespace-trimming syntax
// Matches: {{...}}, {{- ... }}, {{ ... -}}, {{- ... -}}
var tagRegex = regexp.MustCompile(`\{\{-?\s*.*?\s*-?\}\}`)

// incompleteTagRegex matches incomplete tags that need merging
var incompleteTagRegex = regexp.MustCompile(`\{\{[^}]*$|^[^{]*\}\}|\{$`)

func textContainsTags(text string) bool {
	return tagRegex.MatchString(text)
}

func textContainsIncompleteTags(text string) bool {
	// Check for obvious incomplete patterns
	if incompleteTagRegex.MatchString(text) {
		return true
	}

	// Check for unbalanced braces
	openCount := strings.Count(text, "{{")
	closeCount := strings.Count(text, "}}")

	if openCount != closeCount {
		return true
	}

	// Check for single { at end (start of {{)
	if strings.HasSuffix(text, "{") && !strings.HasSuffix(text, "{{") {
		return true
	}

	// Check for single } at end (might be part of }})
	if strings.HasSuffix(text, "}") && !strings.HasSuffix(text, "}}") {
		// Only if there are unclosed tags
		if openCount > 0 {
			return true
		}
	}

	return false
}

// FindAllTags returns all template tags found in the given text.
// Tags are in the format {{...}}, including whitespace-trimming variants.
func FindAllTags(text string) []string {
	return tagRegex.FindAllString(text, -1)
}

// ExtractTagContent extracts the content inside a tag (without delimiters)
// e.g., "{{.Name}}" -> ".Name", "{{- .Name -}}" -> ".Name"
func ExtractTagContent(tag string) string {
	// Remove outer {{ and }}
	content := strings.TrimPrefix(tag, "{{")
	content = strings.TrimSuffix(content, "}}")

	// Remove trimming markers and whitespace
	content = strings.TrimPrefix(content, "-")
	content = strings.TrimSuffix(content, "-")
	content = strings.TrimSpace(content)

	return content
}

// IsBlockTag returns true if the tag is a block control tag (if, range, with, define, block, template)
func IsBlockTag(tag string) bool {
	content := ExtractTagContent(tag)
	blockKeywords := []string{"if", "else", "end", "range", "with", "define", "template", "block"}

	for _, keyword := range blockKeywords {
		if strings.HasPrefix(content, keyword+" ") || content == keyword || strings.HasPrefix(content, keyword+"\t") {
			return true
		}
	}
	return false
}
