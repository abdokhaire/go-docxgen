package xmlutils

import (
	"bytes"
	"encoding/xml"
	"strings"
)

// Escape special XML characters in a string and convert newlines to Word line breaks.
// Newline characters (\n) are converted to </w:t><w:br/><w:t> for proper display in Word.
func EscapeXmlString(xmlString string) (string, error) {
	var buf bytes.Buffer
	err := xml.EscapeText(&buf, []byte(xmlString))
	if err != nil {
		return "", err
	}

	// Convert newlines to Word line breaks
	// The pattern </w:t><w:br/><w:t> closes the current text run, inserts a break, and opens a new text run
	result := strings.ReplaceAll(buf.String(), "&#xA;", "</w:t><w:br/><w:t>")

	return result, nil
}
