package docxtpl

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"sort"
	"strings"
)

// XMLFile represents an XML file from the DOCX archive
type XMLFile struct {
	Path    string // File path within the archive (e.g., "word/document.xml")
	Content string // XML content
}

// UnpackedDocument represents an unpacked DOCX document
type UnpackedDocument struct {
	Files map[string][]byte // All files in the archive
}

// UnpackToDirectory extracts the DOCX file to a directory.
// This allows direct access to all XML files for advanced manipulation.
//
//	err := doc.UnpackToDirectory("/tmp/unpacked")
func (d *DocxTmpl) UnpackToDirectory(dirPath string) error {
	// Create the directory
	if err := os.MkdirAll(dirPath, 0755); err != nil {
		return fmt.Errorf("failed to create directory: %w", err)
	}

	// Get the document as bytes
	var buf bytes.Buffer
	if err := d.Save(&buf); err != nil {
		return fmt.Errorf("failed to save document: %w", err)
	}

	// Open as zip
	reader := bytes.NewReader(buf.Bytes())
	zipReader, err := zip.NewReader(reader, int64(buf.Len()))
	if err != nil {
		return fmt.Errorf("failed to read zip: %w", err)
	}

	// Extract all files
	for _, f := range zipReader.File {
		targetPath := filepath.Join(dirPath, f.Name)

		// Create directory if needed
		if f.FileInfo().IsDir() {
			if err := os.MkdirAll(targetPath, 0755); err != nil {
				return err
			}
			continue
		}

		// Ensure parent directory exists
		if err := os.MkdirAll(filepath.Dir(targetPath), 0755); err != nil {
			return err
		}

		// Extract file
		rc, err := f.Open()
		if err != nil {
			return err
		}

		outFile, err := os.Create(targetPath)
		if err != nil {
			rc.Close()
			return err
		}

		_, err = io.Copy(outFile, rc)
		outFile.Close()
		rc.Close()

		if err != nil {
			return err
		}
	}

	return nil
}

// PackFromDirectory creates a DOCX file from an unpacked directory.
// This is the reverse of UnpackToDirectory.
//
//	doc, err := docxtpl.PackFromDirectory("/tmp/unpacked")
func PackFromDirectory(dirPath string) (*DocxTmpl, error) {
	// Create a buffer to write the zip
	var buf bytes.Buffer
	zipWriter := zip.NewWriter(&buf)

	// Walk the directory
	err := filepath.Walk(dirPath, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}

		// Skip directories
		if info.IsDir() {
			return nil
		}

		// Get relative path
		relPath, err := filepath.Rel(dirPath, path)
		if err != nil {
			return err
		}

		// Use forward slashes for zip paths
		relPath = strings.ReplaceAll(relPath, string(os.PathSeparator), "/")

		// Create file in zip
		writer, err := zipWriter.Create(relPath)
		if err != nil {
			return err
		}

		// Read and write content
		content, err := os.ReadFile(path)
		if err != nil {
			return err
		}

		_, err = writer.Write(content)
		return err
	})

	if err != nil {
		return nil, fmt.Errorf("failed to walk directory: %w", err)
	}

	if err := zipWriter.Close(); err != nil {
		return nil, fmt.Errorf("failed to close zip: %w", err)
	}

	// Parse the packed document
	return ParseFromBytes(buf.Bytes())
}

// GetXMLFiles returns a list of all XML files in the document.
//
//	files := doc.GetXMLFiles()
//	for _, f := range files {
//	    fmt.Println(f)
//	}
func (d *DocxTmpl) GetXMLFiles() []string {
	var files []string

	var buf bytes.Buffer
	if err := d.Save(&buf); err != nil {
		return files
	}

	reader := bytes.NewReader(buf.Bytes())
	zipReader, err := zip.NewReader(reader, int64(buf.Len()))
	if err != nil {
		return files
	}

	for _, f := range zipReader.File {
		if strings.HasSuffix(f.Name, ".xml") {
			files = append(files, f.Name)
		}
	}

	sort.Strings(files)
	return files
}

// GetXMLContent returns the content of a specific XML file within the document.
//
//	content, err := doc.GetXMLContent("word/document.xml")
func (d *DocxTmpl) GetXMLContent(filePath string) (string, error) {
	// Check processable files first
	for _, pf := range d.processableFiles {
		if pf.Name == filePath {
			return pf.Content, nil
		}
	}

	// Otherwise, extract from the document
	var buf bytes.Buffer
	if err := d.Save(&buf); err != nil {
		return "", fmt.Errorf("failed to save document: %w", err)
	}

	reader := bytes.NewReader(buf.Bytes())
	zipReader, err := zip.NewReader(reader, int64(buf.Len()))
	if err != nil {
		return "", fmt.Errorf("failed to read zip: %w", err)
	}

	for _, f := range zipReader.File {
		if f.Name == filePath {
			rc, err := f.Open()
			if err != nil {
				return "", err
			}
			defer rc.Close()

			content, err := io.ReadAll(rc)
			if err != nil {
				return "", err
			}
			return string(content), nil
		}
	}

	return "", fmt.Errorf("file not found: %s", filePath)
}

// SetXMLContent sets the content of a processable XML file.
// Only works for files in the processable files list (headers, footers, settings, etc.)
//
//	err := doc.SetXMLContent("word/settings.xml", newContent)
func (d *DocxTmpl) SetXMLContent(filePath string, content string) error {
	for i := range d.processableFiles {
		if d.processableFiles[i].Name == filePath {
			d.processableFiles[i].Content = content
			return nil
		}
	}
	return fmt.Errorf("file not in processable files: %s", filePath)
}

// Note: GetAllMedia is defined in document_ops.go

// GetRelationships returns the document relationships.
//
//	rels, err := doc.GetRelationships()
func (d *DocxTmpl) GetRelationships() (string, error) {
	return d.GetXMLContent("word/_rels/document.xml.rels")
}

// GetContentTypes returns the content types XML.
//
//	types, err := doc.GetContentTypes()
func (d *DocxTmpl) GetContentTypesXML() (string, error) {
	return d.GetXMLContent("[Content_Types].xml")
}

// GetDocumentXML returns the main document XML.
//
//	xml, err := doc.GetDocumentXML()
func (d *DocxTmpl) GetDocumentXML() (string, error) {
	return d.getDocumentXml()
}

// GetSettingsXML returns the settings XML.
//
//	settings, err := doc.GetSettingsXML()
func (d *DocxTmpl) GetSettingsXML() (string, error) {
	for _, pf := range d.processableFiles {
		if strings.HasSuffix(pf.Name, "settings.xml") {
			return pf.Content, nil
		}
	}
	return d.GetXMLContent("word/settings.xml")
}

// GetStylesXML returns the styles XML.
//
//	styles, err := doc.GetStylesXML()
func (d *DocxTmpl) GetStylesXML() (string, error) {
	return d.GetXMLContent("word/styles.xml")
}

// ArchiveInfo returns information about the DOCX archive.
type ArchiveInfo struct {
	TotalFiles   int      // Total number of files in archive
	XMLFiles     int      // Number of XML files
	MediaFiles   int      // Number of media files
	TotalSize    int64    // Total uncompressed size
	FileList     []string // List of all file paths
	HasComments  bool     // Whether the document has comments.xml
	HasSettings  bool     // Whether the document has settings.xml
	HasFootnotes bool     // Whether the document has footnotes.xml
	HasEndnotes  bool     // Whether the document has endnotes.xml
}

// GetArchiveInfo returns information about the DOCX archive structure.
//
//	info := doc.GetArchiveInfo()
//	fmt.Printf("Total files: %d\n", info.TotalFiles)
func (d *DocxTmpl) GetArchiveInfo() *ArchiveInfo {
	info := &ArchiveInfo{}

	var buf bytes.Buffer
	if err := d.Save(&buf); err != nil {
		return info
	}

	reader := bytes.NewReader(buf.Bytes())
	zipReader, err := zip.NewReader(reader, int64(buf.Len()))
	if err != nil {
		return info
	}

	for _, f := range zipReader.File {
		info.TotalFiles++
		info.TotalSize += int64(f.UncompressedSize64)
		info.FileList = append(info.FileList, f.Name)

		if strings.HasSuffix(f.Name, ".xml") {
			info.XMLFiles++
		}
		if strings.HasPrefix(f.Name, "word/media/") {
			info.MediaFiles++
		}
		if f.Name == "word/comments.xml" {
			info.HasComments = true
		}
		if f.Name == "word/settings.xml" {
			info.HasSettings = true
		}
		if f.Name == "word/footnotes.xml" {
			info.HasFootnotes = true
		}
		if f.Name == "word/endnotes.xml" {
			info.HasEndnotes = true
		}
	}

	sort.Strings(info.FileList)
	return info
}
