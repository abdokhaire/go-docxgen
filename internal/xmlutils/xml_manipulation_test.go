package xmlutils

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestReplaceTableRangeRows(t *testing.T) {
	tests := []struct {
		name              string
		inputXml          string
		expectedOutputXml string
	}{
		{
			name:              "Basic range tag",
			inputXml:          "<w:tr>{{range . }}</w:tr>",
			expectedOutputXml: "{{range . }}",
		},
		{
			name:              "Basic end tag",
			inputXml:          "<w:tr>{{end}}</w:tr>",
			expectedOutputXml: "{{end}}",
		},
		{
			name:              "Range tag in other text",
			inputXml:          "<w:tbl><w:tr>{{range . }}</w:tr></w:tbl>",
			expectedOutputXml: "<w:tbl>{{range . }}</w:tbl>",
		},
		{
			name:              "End tag in other text",
			inputXml:          "<w:tbl><w:tr>{{end}}</w:tr></w:tbl>",
			expectedOutputXml: "<w:tbl>{{end}}</w:tbl>",
		},
		{
			name:              "Multiple range tags",
			inputXml:          "<w:tbl><w:tr>{{range . }}</w:tr></w:tbl><w:tbl><w:tr>{{range . }}</w:tr></w:tbl>",
			expectedOutputXml: "<w:tbl>{{range . }}</w:tbl><w:tbl>{{range . }}</w:tbl>",
		},
		{
			name:              "Multiple end tags",
			inputXml:          "<w:tbl><w:tr>{{end}}</w:tr></w:tbl><w:tbl><w:tr>{{end}}</w:tr></w:tbl>",
			expectedOutputXml: "<w:tbl>{{end}}</w:tbl><w:tbl>{{end}}</w:tbl>",
		},
		{
			name:              "Full table",
			inputXml:          "<w:tbl><w:tr>{{range . }}</w:tr><w:tr></w:tr><w:tr>{{end}}</w:tr></w:tbl>",
			expectedOutputXml: "<w:tbl>{{range . }}<w:tr></w:tr>{{end}}</w:tbl>",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			assert := assert.New(t)

			outputXml, err := replaceTableRangeRows(tt.inputXml)
			assert.Nil(err)
			assert.Equal(outputXml, tt.expectedOutputXml)
		})
	}
}

func TestFixXmlIssuesPostTagReplacement(t *testing.T) {
	tests := []struct {
		name              string
		inputXml          string
		expectedOutputXml string
	}{
		{
			name:              "Drawing tags",
			inputXml:          "<w:t><w:drawing>...</w:drawing></w:t>",
			expectedOutputXml: "<w:drawing>...</w:drawing>",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			outputXml := FixXmlIssuesPostTagReplacement(tt.inputXml)
			assert.Equal(t, outputXml, tt.expectedOutputXml)
		})
	}
}

func TestMergeFragmentedTagsInXml(t *testing.T) {
	tests := []struct {
		name     string
		input    string
		expected string
	}{
		{
			name:     "No fragmentation - complete tag",
			input:    `<w:r><w:t>{{.Name}}</w:t></w:r>`,
			expected: `<w:r><w:t>{{.Name}}</w:t></w:r>`,
		},
		{
			name:     "Tag split across two text nodes",
			input:    `<w:r><w:t>{{.First</w:t></w:r><w:r><w:t>Name}}</w:t></w:r>`,
			expected: `<w:r><w:t>{{.FirstName}}</w:t></w:r><w:r><w:t></w:t></w:r>`,
		},
		{
			name:     "Tag split across three text nodes",
			input:    `<w:r><w:t>{{.</w:t></w:r><w:r><w:t>First</w:t></w:r><w:r><w:t>Name}}</w:t></w:r>`,
			expected: `<w:r><w:t>{{.FirstName}}</w:t></w:r><w:r><w:t></w:t></w:r><w:r><w:t></w:t></w:r>`,
		},
		{
			name:     "Multiple complete tags - no change needed",
			input:    `<w:r><w:t>{{.First}}</w:t></w:r><w:r><w:t>{{.Last}}</w:t></w:r>`,
			expected: `<w:r><w:t>{{.First}}</w:t></w:r><w:r><w:t>{{.Last}}</w:t></w:r>`,
		},
		{
			name:     "Text with xml:space attribute",
			input:    `<w:r><w:t xml:space="preserve">{{.First</w:t></w:r><w:r><w:t>Name}}</w:t></w:r>`,
			expected: `<w:r><w:t xml:space="preserve">{{.FirstName}}</w:t></w:r><w:r><w:t></w:t></w:r>`,
		},
		{
			name:     "Plain text without tags",
			input:    `<w:r><w:t>Hello World</w:t></w:r>`,
			expected: `<w:r><w:t>Hello World</w:t></w:r>`,
		},
		{
			name:     "Mixed content with fragmented tag",
			input:    `<w:r><w:t>Hello {{.Na</w:t></w:r><w:r><w:t>me}} World</w:t></w:r>`,
			expected: `<w:r><w:t>Hello {{.Name}} World</w:t></w:r><w:r><w:t></w:t></w:r>`,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := MergeFragmentedTagsInXml(tt.input)
			assert.Equal(t, tt.expected, result)
		})
	}
}

func TestContainsIncompleteOpenTag(t *testing.T) {
	tests := []struct {
		name     string
		text     string
		expected bool
	}{
		{
			name:     "Complete tag",
			text:     "{{.Name}}",
			expected: false,
		},
		{
			name:     "Incomplete opening",
			text:     "{{.Name",
			expected: true,
		},
		{
			name:     "Just opening braces",
			text:     "{{",
			expected: true,
		},
		{
			name:     "Single brace at end",
			text:     "Hello {",
			expected: true,
		},
		{
			name:     "Plain text",
			text:     "Hello World",
			expected: false,
		},
		{
			name:     "Multiple complete tags",
			text:     "{{.A}} {{.B}}",
			expected: false,
		},
		{
			name:     "Complete tag followed by incomplete",
			text:     "{{.A}} {{.B",
			expected: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := containsIncompleteOpenTag(tt.text)
			assert.Equal(t, tt.expected, result)
		})
	}
}
