package headerfooter

import (
	"archive/zip"
	"io"
	"regexp"
)

var headerRegex = regexp.MustCompile(`^word/header[0-9]+\.xml$`)
var footerRegex = regexp.MustCompile(`^word/footer[0-9]+\.xml$`)
var footnotesRegex = regexp.MustCompile(`^word/footnotes\.xml$`)
var endnotesRegex = regexp.MustCompile(`^word/endnotes\.xml$`)
var docPropsRegex = regexp.MustCompile(`^docProps/(core|app)\.xml$`)

// WatermarkRegex matches the string attribute in VML textpath elements (watermarks)
// Example: <v:textpath ... string="DRAFT" ...>
var WatermarkRegex = regexp.MustCompile(`(<v:textpath[^>]*\sstring=")([^"]*)("[^>]*>)`)

// DocxFile represents an XML file from the DOCX archive that can be processed.
type DocxFile struct {
	Name    string // File path within the archive (e.g., "word/header1.xml")
	Content string // XML content
}

// GetProcessableFiles extracts all XML files that should be processed for template replacement.
// This includes headers, footers, footnotes, and endnotes.
func GetProcessableFiles(reader io.ReaderAt, size int64) ([]DocxFile, error) {
	zipReader, err := zip.NewReader(reader, size)
	if err != nil {
		return nil, err
	}

	var files []DocxFile

	for _, f := range zipReader.File {
		if isProcessableFile(f.Name) {
			zf, err := f.Open()
			if err != nil {
				return nil, err
			}

			content, err := io.ReadAll(zf)
			zf.Close()
			if err != nil {
				return nil, err
			}

			files = append(files, DocxFile{
				Name:    f.Name,
				Content: string(content),
			})
		}
	}

	return files, nil
}

// isProcessableFile returns true if the file should be processed for template replacement.
func isProcessableFile(name string) bool {
	return headerRegex.MatchString(name) ||
		footerRegex.MatchString(name) ||
		footnotesRegex.MatchString(name) ||
		endnotesRegex.MatchString(name) ||
		docPropsRegex.MatchString(name)
}

// IsDocProps returns true if the given filename is a document properties file.
func IsDocProps(name string) bool {
	return docPropsRegex.MatchString(name)
}

// IsHeaderOrFooter returns true if the given filename is a header or footer file.
func IsHeaderOrFooter(name string) bool {
	return headerRegex.MatchString(name) || footerRegex.MatchString(name)
}

// ReplaceWatermarkText replaces watermark text in VML textpath elements.
// Watermarks are stored as <v:textpath string="WATERMARK TEXT"> in header files.
func ReplaceWatermarkText(content string, oldText, newText string) string {
	return WatermarkRegex.ReplaceAllStringFunc(content, func(match string) string {
		submatches := WatermarkRegex.FindStringSubmatch(match)
		if len(submatches) == 4 && submatches[2] == oldText {
			return submatches[1] + newText + submatches[3]
		}
		return match
	})
}

// ExtractWatermarkText extracts watermark text from VML textpath elements.
func ExtractWatermarkText(content string) []string {
	matches := WatermarkRegex.FindAllStringSubmatch(content, -1)
	var texts []string
	for _, match := range matches {
		if len(match) >= 3 {
			texts = append(texts, match[2])
		}
	}
	return texts
}

// ProcessWatermarkTemplates processes Go template syntax in watermark text.
// This allows using {{.FieldName}} in watermarks.
func ProcessWatermarkTemplates(content string, processFunc func(string) (string, error)) (string, error) {
	var lastErr error
	result := WatermarkRegex.ReplaceAllStringFunc(content, func(match string) string {
		submatches := WatermarkRegex.FindStringSubmatch(match)
		if len(submatches) == 4 {
			watermarkText := submatches[2]
			// Process the watermark text through the template function
			processed, err := processFunc(watermarkText)
			if err != nil {
				lastErr = err
				return match
			}
			return submatches[1] + processed + submatches[3]
		}
		return match
	})
	return result, lastErr
}
