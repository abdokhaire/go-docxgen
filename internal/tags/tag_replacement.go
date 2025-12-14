package tags

import (
	"bytes"
	"fmt"
	"log"
	"regexp"
	"text/template"

	"github.com/abdokhaire/go-docxgen/internal/xmlutils"
)

// Data should already be processed and have been XML escaped (aside from embedded objects like images) before being passed into this function
func ReplaceTagsInXml(xmlString string, data map[string]any, funcMap template.FuncMap) (string, error) {
	// Prepare the XML for tag replacement
	preparedXmlString, err := xmlutils.PrepareXmlForTagReplacement(xmlString)
	if err != nil {
		return "", err
	}

	tmpl, err := template.New("").Funcs(funcMap).Parse(preparedXmlString)
	if err != nil {
		// Debug: extract and log any template blocks that contain & character
		templateBlockRe := regexp.MustCompile(`\{\{[^}]*&[^}]*\}\}`)
		problemBlocks := templateBlockRe.FindAllString(preparedXmlString, -1)
		if len(problemBlocks) > 0 {
			log.Printf("DEBUG go-docxgen: Found template blocks with '&' character: %v", problemBlocks)
		}
		return "", fmt.Errorf("error parsing template: %v", err)
	}

	buf := &bytes.Buffer{}
	err = tmpl.Execute(buf, data)
	if err != nil {
		return "", err
	}

	// Fix any issues in the XML
	outputXmlString := xmlutils.FixXmlIssuesPostTagReplacement(buf.String())

	return outputXmlString, err
}

// ReplaceTagsInText processes Go template syntax in plain text (not XML).
// This is useful for watermarks and other non-XML text content.
func ReplaceTagsInText(text string, data map[string]any, funcMap template.FuncMap) (string, error) {
	// Check if text contains any template syntax
	if !textContainsTags(text) {
		return text, nil
	}

	tmpl, err := template.New("").Funcs(funcMap).Parse(text)
	if err != nil {
		return "", fmt.Errorf("error parsing template: %v", err)
	}

	buf := &bytes.Buffer{}
	err = tmpl.Execute(buf, data)
	if err != nil {
		return "", err
	}

	return buf.String(), nil
}
