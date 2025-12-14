package docxtpl

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"io"
	"os"
	"text/template"

	"github.com/fumiama/go-docx"
	"github.com/abdokhaire/go-docxgen/internal/contenttypes"
	"github.com/abdokhaire/go-docxgen/internal/headerfooter"
	"github.com/abdokhaire/go-docxgen/internal/hyperlinks"
	"github.com/abdokhaire/go-docxgen/internal/tags"
	"github.com/abdokhaire/go-docxgen/internal/templatedata"
	"github.com/abdokhaire/go-docxgen/internal/xmlutils"
)

type DocxTmpl struct {
	*docx.Docx
	funcMap          template.FuncMap
	contentTypes     *contenttypes.ContentTypes
	processableFiles []headerfooter.DocxFile // headers, footers, footnotes, endnotes
	hyperlinkReg     *hyperlinks.HyperlinkRegistry
	properties       *DocumentProperties // document metadata (stored in memory, serialized on save)
}

// Parse the document from a reader and store it in memory.
// You can it invoke from a file.
//
//	reader, err := os.Open("path_to_doc.docx")
//	if err != nil {
//		panic(err)
//	}
//	fileinfo, err := reader.Stat()
//	if err != nil {
//		panic(err)
//	}
//	size := fileinfo.Size()
//	doc, err := docxtpl.Parse(reader, int64(size))
func Parse(reader io.ReaderAt, size int64) (*DocxTmpl, error) {
	doc, err := docx.Parse(reader, size)
	if err != nil {
		return nil, err
	}

	contentTypes, err := contenttypes.GetContentTypes(reader, size)
	if err != nil {
		return nil, err
	}

	processableFiles, err := headerfooter.GetProcessableFiles(reader, size)
	if err != nil {
		return nil, err
	}

	funcMap := make(template.FuncMap)

	hyperlinkReg := hyperlinks.NewHyperlinkRegistry()

	docTmpl := &DocxTmpl{doc, funcMap, contentTypes, processableFiles, hyperlinkReg, nil}

	// Override the link function to use our hyperlink registry
	docTmpl.funcMap["link"] = docTmpl.createLink

	return docTmpl, nil
}

// createLink creates a hyperlink and registers it for relationship injection
func (d *DocxTmpl) createLink(url, text string) string {
	rId := d.hyperlinkReg.RegisterLink(url)
	return hyperlinks.HyperlinkXML(rId, text)
}

// Parse the document from a filename and store it in memory.
//
//	doc, err := docxtpl.ParseFromFilename("path_to_doc.docx")
func ParseFromFilename(filename string) (*DocxTmpl, error) {
	reader, err := os.Open(filename)
	if err != nil {
		return nil, err
	}

	fileinfo, err := reader.Stat()
	if err != nil {
		return nil, err
	}
	size := fileinfo.Size()

	doxtpl, err := Parse(reader, size)
	if err != nil {
		return nil, err
	}

	return doxtpl, nil
}

// Parse the document from a byte slice and store it in memory.
// Useful for documents loaded from HTTP responses or other in-memory sources.
//
//	data, err := io.ReadAll(resp.Body)
//	if err != nil {
//		panic(err)
//	}
//	doc, err := docxtpl.ParseFromBytes(data)
func ParseFromBytes(data []byte) (*DocxTmpl, error) {
	reader := bytes.NewReader(data)
	return Parse(reader, int64(len(data)))
}

// Replace the placeholders in the document with passed in data.
// Data can be a struct or map
//
//	data := struct {
//		FirstName     string
//		LastName      string
//		Gender        string
//	}{
//		FirstName: "Tom",
//		LastName:  "Watkins",
//		Gender:    "Male",
//	}
//
// err = doc.Render(data)
//
// # OR
//
//	data := map[string]any{
//		"ProjectNumber": "B-00001",
//		"Client":        "TW Software",
//		"Status":        "New",
//	}
//
// err = doc.Render(data)
func (d *DocxTmpl) Render(data any) error {
	// Ensure that there are no 'part tags' in the XML document
	tags.MergeTags(d.Document.Body.Items)

	// Process the template data
	processedData, err := d.processTemplateData(data)
	if err != nil {
		return err
	}

	// Get the document XML
	documentXmlString, err := d.getDocumentXml()
	if err != nil {
		return err
	}

	// Replace the tags in XML
	documentXmlString, err = tags.ReplaceTagsInXml(documentXmlString, processedData, d.funcMap)
	if err != nil {
		return err
	}

	// Unmarshal the modified XML and replace the document body with it
	decoder := xml.NewDecoder(bytes.NewBufferString(documentXmlString))
	for {
		t, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}
		if start, ok := t.(xml.StartElement); ok {
			if start.Name.Local == "Body" {
				clear(d.Document.Body.Items)
				err = d.Document.Body.UnmarshalXML(decoder, start)
				if err != nil {
					return err
				}
				break
			}
		}
	}

	// Process headers, footers, footnotes, endnotes, and document properties
	for i := range d.processableFiles {
		var processedContent string
		var err error

		if headerfooter.IsDocProps(d.processableFiles[i].Name) {
			// Document properties don't have <w:t> elements, process directly with templates
			processedContent, err = tags.ReplaceTagsInText(d.processableFiles[i].Content, processedData, d.funcMap)
			if err != nil {
				return err
			}
		} else {
			// Merge fragmented tags in the XML (handles tags split across multiple <w:t> elements)
			mergedContent := xmlutils.MergeFragmentedTagsInXml(d.processableFiles[i].Content)

			// Process regular text placeholders
			processedContent, err = tags.ReplaceTagsInXml(mergedContent, processedData, d.funcMap)
			if err != nil {
				return err
			}

			// Process watermark templates in headers (watermarks are VML shapes with textpath)
			if headerfooter.IsHeaderOrFooter(d.processableFiles[i].Name) {
				processedContent, err = headerfooter.ProcessWatermarkTemplates(processedContent, func(watermarkText string) (string, error) {
					return tags.ReplaceTagsInText(watermarkText, processedData, d.funcMap)
				})
				if err != nil {
					return err
				}
			}
		}

		d.processableFiles[i].Content = processedContent
	}

	return nil
}

// Save the document to a writer.
// This could be a new file.
//
//	f, err := os.Create(FILE_PATH)
//	if err != nil {
//		panic(err)
//	}
//	err = doc.Save(f)
//	if err != nil {
//		panic(err)
//	}
//	err = f.Close()
//	if err != nil {
//		panic(err)
//	}
func (d *DocxTmpl) Save(writer io.Writer) error {
	var buf bytes.Buffer
	_, err := d.WriteTo(&buf)
	if err != nil {
		return err
	}

	reader := bytes.NewReader(buf.Bytes())
	zipReader, err := zip.NewReader(reader, int64(buf.Len()))
	if err != nil {
		return err
	}

	generatedZip := zip.NewWriter(writer)

	// Track if we've seen the rels file
	hasRelsFile := false
	const documentRelsPath = "word/_rels/document.xml.rels"

	// Check if hyperlinks need to be added
	hyperlinkLinks := d.hyperlinkReg.GetLinks()
	hasHyperlinks := len(hyperlinkLinks) > 0

	for _, f := range zipReader.File {
		newFile, err := generatedZip.Create(f.Name)
		if err != nil {
			return err
		}

		// Override content types with our calculated types
		// Override header/footer files with our processed content
		// Inject hyperlinks into document.xml.rels
		// Copy across all other files
		if f.Name == "[Content_Types].xml" {
			contentTypesXml, err := d.contentTypes.MarshalXml()
			if err != nil {
				return err
			}

			_, err = newFile.Write([]byte(contentTypesXml))
			if err != nil {
				return err
			}
		} else if f.Name == documentRelsPath && hasHyperlinks {
			// Update document.xml.rels with hyperlinks
			hasRelsFile = true
			zf, err := f.Open()
			if err != nil {
				return err
			}
			existingContent, err := io.ReadAll(zf)
			zf.Close()
			if err != nil {
				return err
			}

			updatedRels, err := hyperlinks.ProcessRelationshipsFile(string(existingContent), hyperlinkLinks)
			if err != nil {
				return err
			}

			_, err = newFile.Write([]byte(updatedRels))
			if err != nil {
				return err
			}
		} else if content := d.getProcessableFileContent(f.Name); content != "" {
			// Write our processed content (headers, footers, footnotes, endnotes)
			_, err = newFile.Write([]byte(content))
			if err != nil {
				return err
			}
		} else {
			zf, err := f.Open()
			if err != nil {
				return err
			}
			defer zf.Close()

			if _, err := io.Copy(newFile, zf); err != nil {
				return err
			}
		}
	}

	// If we have hyperlinks but no rels file existed, create one
	if hasHyperlinks && !hasRelsFile {
		newFile, err := generatedZip.Create(documentRelsPath)
		if err != nil {
			return err
		}

		newRels, err := hyperlinks.ProcessRelationshipsFile("", hyperlinkLinks)
		if err != nil {
			return err
		}

		_, err = newFile.Write([]byte(newRels))
		if err != nil {
			return err
		}
	}

	if err := generatedZip.Close(); err != nil {
		return err
	}

	return nil
}

// Save the document directly to a file path.
// This is a convenience method that creates the file and calls Save().
//
//	err := doc.SaveToFile("output.docx")
//	if err != nil {
//		panic(err)
//	}
func (d *DocxTmpl) SaveToFile(filename string) error {
	f, err := os.Create(filename)
	if err != nil {
		return err
	}
	defer f.Close()

	return d.Save(f)
}

// GetPlaceholders returns a list of all unique placeholders found in the document.
// This is useful for validating templates or showing what data is expected.
// Placeholders are returned in the format they appear (e.g., "{{.Name}}", "{{range .Items}}").
//
//	placeholders, err := doc.GetPlaceholders()
//	// Returns: []string{"{{.FirstName}}", "{{.LastName}}", "{{range .Items}}", "{{end}}"}
func (d *DocxTmpl) GetPlaceholders() ([]string, error) {
	// Merge tags first to handle fragmented placeholders
	tags.MergeTags(d.Document.Body.Items)

	// Get the document XML
	documentXmlString, err := d.getDocumentXml()
	if err != nil {
		return nil, err
	}

	// Find all placeholders
	allTags := tags.FindAllTags(documentXmlString)

	// Remove duplicates while preserving order
	seen := make(map[string]bool)
	uniqueTags := make([]string, 0, len(allTags))
	for _, tag := range allTags {
		if !seen[tag] {
			seen[tag] = true
			uniqueTags = append(uniqueTags, tag)
		}
	}

	return uniqueTags, nil
}

func (d *DocxTmpl) getDocumentXml() (string, error) {
	out, err := xml.Marshal(d.Document.Body)
	if err != nil {
		return "", nil
	}

	return string(out), err
}

func (d *DocxTmpl) processTemplateData(data any) (map[string]any, error) {
	convertedData, err := templatedata.DataToMap(data)
	if err != nil {
		return nil, err
	}

	// Declare both functions first to allow mutual recursion
	var processValue func(value any) (any, error)
	var processMap func(data *map[string]any) error

	processMap = func(data *map[string]any) error {
		for key, value := range *data {
			processed, err := processValue(value)
			if err != nil {
				return err
			}
			(*data)[key] = processed
		}
		return nil
	}

	processValue = func(value any) (any, error) {
		switch v := value.(type) {
		case string:
			// Check for image files
			if isImage, err := templatedata.IsImageFilePath(v); err != nil {
				return nil, err
			} else if isImage {
				image, err := CreateInlineImage(v)
				if err != nil {
					return nil, err
				}
				return d.addInlineImage(image)
			}
			// XML escape regular strings
			return xmlutils.EscapeXmlString(v)

		case map[string]any:
			if err := processMap(&v); err != nil {
				return nil, err
			}
			return v, nil

		case []map[string]any:
			for i := range v {
				if err := processMap(&v[i]); err != nil {
					return nil, err
				}
			}
			return v, nil

		case []any:
			// Process primitive slices - XML escape string elements
			for i, elem := range v {
				processed, err := processValue(elem)
				if err != nil {
					return nil, err
				}
				v[i] = processed
			}
			return v, nil

		case *InlineImage:
			return d.addInlineImage(v)

		default:
			// Return other types as-is (int, float, bool, etc.)
			return value, nil
		}
	}

	err = processMap(&convertedData)
	if err != nil {
		return nil, err
	}

	return convertedData, nil
}

// getProcessableFileContent returns the processed content for a processable file,
// or an empty string if the file was not processed.
func (d *DocxTmpl) getProcessableFileContent(name string) string {
	for _, pf := range d.processableFiles {
		if pf.Name == name {
			return pf.Content
		}
	}
	return ""
}

// GetWatermarks returns all watermark texts found in the document headers.
// Watermarks are stored as VML shapes with textpath elements in header files.
func (d *DocxTmpl) GetWatermarks() []string {
	var watermarks []string
	for _, pf := range d.processableFiles {
		if headerfooter.IsHeaderOrFooter(pf.Name) {
			texts := headerfooter.ExtractWatermarkText(pf.Content)
			watermarks = append(watermarks, texts...)
		}
	}
	return watermarks
}

// ReplaceWatermark replaces a specific watermark text with a new value.
// This should be called before Render() if you want to change watermark text.
//
//	doc.ReplaceWatermark("DRAFT", "FINAL")
//	doc.Render(data)
func (d *DocxTmpl) ReplaceWatermark(oldText, newText string) {
	for i := range d.processableFiles {
		if headerfooter.IsHeaderOrFooter(d.processableFiles[i].Name) {
			d.processableFiles[i].Content = headerfooter.ReplaceWatermarkText(
				d.processableFiles[i].Content, oldText, newText)
		}
	}
}
