package tags

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestTextContainsTags(t *testing.T) {
	tests := []struct {
		name           string
		text           string
		expectedResult bool
	}{
		{
			name:           "Single tag",
			text:           "This text contains a tag: {{.Tag}}.",
			expectedResult: true,
		},
		{
			name:           "Multiple tags",
			text:           "This text contains multiple tags: {{.Tag1}} and {{.Tag2}}.",
			expectedResult: true,
		},
		{
			name:           "No tag",
			text:           "This text contains no tags",
			expectedResult: false,
		},
		{
			name:           "Incomplete tag",
			text:           "This text contains an incomplete tag: {{.Tag",
			expectedResult: false,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := textContainsTags(tt.text)
			assert.Equal(t, result, tt.expectedResult)
		})
	}
}

func TestTextContainsIncompleteTags(t *testing.T) {
	tests := []struct {
		name           string
		text           string
		expectedResult bool
	}{
		{
			name:           "Single incomplete tag",
			text:           "This text contains a tag: {{.Tag",
			expectedResult: true,
		},
		{
			name:           "Multiple incomplete tags",
			text:           "This text contains multiple tags: {{.Tag1 and {{.Tag2.",
			expectedResult: true,
		},
		{
			name:           "No tag",
			text:           "This text contains no tags",
			expectedResult: false,
		},
		{
			name:           "Complete tag",
			text:           "This text contains an complete tag: {{.Tag}}",
			expectedResult: false,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := textContainsIncompleteTags(tt.text)
			assert.Equal(t, result, tt.expectedResult)
		})
	}
}

func TestFindAllTags(t *testing.T) {
	tests := []struct {
		name     string
		text     string
		expected []string
	}{
		{
			name:     "Single tag",
			text:     "Hello {{.Name}}!",
			expected: []string{"{{.Name}}"},
		},
		{
			name:     "Multiple tags",
			text:     "{{.FirstName}} {{.LastName}} is {{.Age}} years old",
			expected: []string{"{{.FirstName}}", "{{.LastName}}", "{{.Age}}"},
		},
		{
			name:     "Range and end tags",
			text:     "{{range .Items}}{{.Name}}{{end}}",
			expected: []string{"{{range .Items}}", "{{.Name}}", "{{end}}"},
		},
		{
			name:     "No tags",
			text:     "Plain text without any tags",
			expected: nil,
		},
		{
			name:     "Tags with functions",
			text:     "{{.Name | upper}} and {{.Title | lower}}",
			expected: []string{"{{.Name | upper}}", "{{.Title | lower}}"},
		},
		{
			name:     "Conditional tags",
			text:     "{{if .Show}}Visible{{else}}Hidden{{end}}",
			expected: []string{"{{if .Show}}", "{{else}}", "{{end}}"},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := FindAllTags(tt.text)
			assert.Equal(t, tt.expected, result)
		})
	}
}
