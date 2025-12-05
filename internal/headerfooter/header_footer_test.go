package headerfooter

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestIsHeaderOrFooter(t *testing.T) {
	tests := []struct {
		name     string
		filename string
		expected bool
	}{
		{
			name:     "header1.xml is a header",
			filename: "word/header1.xml",
			expected: true,
		},
		{
			name:     "header2.xml is a header",
			filename: "word/header2.xml",
			expected: true,
		},
		{
			name:     "footer1.xml is a footer",
			filename: "word/footer1.xml",
			expected: true,
		},
		{
			name:     "footer10.xml is a footer",
			filename: "word/footer10.xml",
			expected: true,
		},
		{
			name:     "document.xml is not a header/footer",
			filename: "word/document.xml",
			expected: false,
		},
		{
			name:     "styles.xml is not a header/footer",
			filename: "word/styles.xml",
			expected: false,
		},
		{
			name:     "header without number is not matched",
			filename: "word/header.xml",
			expected: false,
		},
	}

	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			result := IsHeaderOrFooter(test.filename)
			assert.Equal(t, test.expected, result)
		})
	}
}

func TestIsProcessableFile(t *testing.T) {
	tests := []struct {
		name     string
		filename string
		expected bool
	}{
		{
			name:     "header1.xml is processable",
			filename: "word/header1.xml",
			expected: true,
		},
		{
			name:     "footer1.xml is processable",
			filename: "word/footer1.xml",
			expected: true,
		},
		{
			name:     "footnotes.xml is processable",
			filename: "word/footnotes.xml",
			expected: true,
		},
		{
			name:     "endnotes.xml is processable",
			filename: "word/endnotes.xml",
			expected: true,
		},
		{
			name:     "docProps/core.xml is processable",
			filename: "docProps/core.xml",
			expected: true,
		},
		{
			name:     "docProps/app.xml is processable",
			filename: "docProps/app.xml",
			expected: true,
		},
		{
			name:     "document.xml is not processable",
			filename: "word/document.xml",
			expected: false,
		},
		{
			name:     "styles.xml is not processable",
			filename: "word/styles.xml",
			expected: false,
		},
	}

	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			result := isProcessableFile(test.filename)
			assert.Equal(t, test.expected, result)
		})
	}
}

func TestIsDocProps(t *testing.T) {
	tests := []struct {
		name     string
		filename string
		expected bool
	}{
		{
			name:     "docProps/core.xml is doc props",
			filename: "docProps/core.xml",
			expected: true,
		},
		{
			name:     "docProps/app.xml is doc props",
			filename: "docProps/app.xml",
			expected: true,
		},
		{
			name:     "word/document.xml is not doc props",
			filename: "word/document.xml",
			expected: false,
		},
		{
			name:     "word/header1.xml is not doc props",
			filename: "word/header1.xml",
			expected: false,
		},
	}

	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			result := IsDocProps(test.filename)
			assert.Equal(t, test.expected, result)
		})
	}
}

func TestExtractWatermarkText(t *testing.T) {
	tests := []struct {
		name     string
		content  string
		expected []string
	}{
		{
			name:     "Simple watermark",
			content:  `<v:textpath style="font-family:'Calibri'" string="DRAFT"/>`,
			expected: []string{"DRAFT"},
		},
		{
			name:     "Multiple watermarks",
			content:  `<v:textpath string="DRAFT"/><v:textpath string="CONFIDENTIAL"/>`,
			expected: []string{"DRAFT", "CONFIDENTIAL"},
		},
		{
			name:     "No watermark",
			content:  `<w:t>Regular text</w:t>`,
			expected: nil,
		},
		{
			name:     "Watermark with attributes before string",
			content:  `<v:textpath style="font-family:'Arial';font-size:1pt" string="SECRET">`,
			expected: []string{"SECRET"},
		},
	}

	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			result := ExtractWatermarkText(test.content)
			assert.Equal(t, test.expected, result)
		})
	}
}

func TestReplaceWatermarkText(t *testing.T) {
	tests := []struct {
		name     string
		content  string
		oldText  string
		newText  string
		expected string
	}{
		{
			name:     "Replace simple watermark",
			content:  `<v:textpath style="font-family:'Calibri'" string="DRAFT"/>`,
			oldText:  "DRAFT",
			newText:  "FINAL",
			expected: `<v:textpath style="font-family:'Calibri'" string="FINAL"/>`,
		},
		{
			name:     "Replace only matching watermark",
			content:  `<v:textpath string="DRAFT"/><v:textpath string="OTHER"/>`,
			oldText:  "DRAFT",
			newText:  "APPROVED",
			expected: `<v:textpath string="APPROVED"/><v:textpath string="OTHER"/>`,
		},
		{
			name:     "No match - no change",
			content:  `<v:textpath string="DRAFT"/>`,
			oldText:  "FINAL",
			newText:  "APPROVED",
			expected: `<v:textpath string="DRAFT"/>`,
		},
	}

	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			result := ReplaceWatermarkText(test.content, test.oldText, test.newText)
			assert.Equal(t, test.expected, result)
		})
	}
}

func TestProcessWatermarkTemplates(t *testing.T) {
	tests := []struct {
		name        string
		content     string
		processFunc func(string) (string, error)
		expected    string
		expectErr   bool
	}{
		{
			name:    "Process template in watermark",
			content: `<v:textpath string="{{.Status}}"/>`,
			processFunc: func(text string) (string, error) {
				if text == "{{.Status}}" {
					return "APPROVED", nil
				}
				return text, nil
			},
			expected:  `<v:textpath string="APPROVED"/>`,
			expectErr: false,
		},
		{
			name:    "Multiple watermarks with templates",
			content: `<v:textpath string="{{.Status}}"/><v:textpath string="{{.Company}}"/>`,
			processFunc: func(text string) (string, error) {
				switch text {
				case "{{.Status}}":
					return "DRAFT", nil
				case "{{.Company}}":
					return "ACME Corp", nil
				}
				return text, nil
			},
			expected:  `<v:textpath string="DRAFT"/><v:textpath string="ACME Corp"/>`,
			expectErr: false,
		},
		{
			name:    "No template - unchanged",
			content: `<v:textpath string="STATIC TEXT"/>`,
			processFunc: func(text string) (string, error) {
				return text, nil
			},
			expected:  `<v:textpath string="STATIC TEXT"/>`,
			expectErr: false,
		},
	}

	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			result, err := ProcessWatermarkTemplates(test.content, test.processFunc)
			if test.expectErr {
				assert.Error(t, err)
			} else {
				assert.NoError(t, err)
				assert.Equal(t, test.expected, result)
			}
		})
	}
}
