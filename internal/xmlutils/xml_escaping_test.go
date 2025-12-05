package xmlutils

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestEscapeXmlString(t *testing.T) {
	tests := []struct {
		name     string
		input    string
		expected string
	}{
		{
			name:     "Escape XML tags",
			input:    "<tag>Text</tag>",
			expected: "&lt;tag&gt;Text&lt;/tag&gt;",
		},
		{
			name:     "Escape ampersand",
			input:    "Text & more text",
			expected: "Text &amp; more text",
		},
		{
			name:     "Escape double quotes",
			input:    "\"Quoted text\"",
			expected: "&#34;Quoted text&#34;",
		},
		{
			name:     "Escape single quotes",
			input:    "'Single quoted text'",
			expected: "&#39;Single quoted text&#39;",
		},
		{
			name:     "Convert newlines to Word line breaks",
			input:    "Line 1\nLine 2\nLine 3",
			expected: "Line 1</w:t><w:br/><w:t>Line 2</w:t><w:br/><w:t>Line 3",
		},
		{
			name:     "Handle text without newlines",
			input:    "Single line text",
			expected: "Single line text",
		},
	}

	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			result, err := EscapeXmlString(test.input)
			assert.NoError(t, err)
			assert.Equal(t, test.expected, result)
		})
	}
}
