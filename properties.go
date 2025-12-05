package docxtpl

import (
	"encoding/xml"
	"strings"
	"time"

	"github.com/abdokhaire/go-docxgen/internal/headerfooter"
)

// DocumentProperties contains the core document metadata.
type DocumentProperties struct {
	Title          string
	Subject        string
	Creator        string // Author
	Keywords       string
	Description    string
	LastModifiedBy string
	Revision       string
	Created        time.Time
	Modified       time.Time
	Category       string
	ContentStatus  string // e.g., "Draft", "Final"
}

// coreProperties represents the XML structure of docProps/core.xml
type coreProperties struct {
	XMLName        xml.Name `xml:"cp:coreProperties"`
	XMLNScp        string   `xml:"xmlns:cp,attr"`
	XMLNSdc        string   `xml:"xmlns:dc,attr"`
	XMLNSdcterms   string   `xml:"xmlns:dcterms,attr"`
	XMLNSdcmitype  string   `xml:"xmlns:dcmitype,attr"`
	XMLNSxsi       string   `xml:"xmlns:xsi,attr"`
	Title          string   `xml:"dc:title,omitempty"`
	Subject        string   `xml:"dc:subject,omitempty"`
	Creator        string   `xml:"dc:creator,omitempty"`
	Keywords       string   `xml:"cp:keywords,omitempty"`
	Description    string   `xml:"dc:description,omitempty"`
	LastModifiedBy string   `xml:"cp:lastModifiedBy,omitempty"`
	Revision       string   `xml:"cp:revision,omitempty"`
	Created        *dcTime  `xml:"dcterms:created,omitempty"`
	Modified       *dcTime  `xml:"dcterms:modified,omitempty"`
	Category       string   `xml:"cp:category,omitempty"`
	ContentStatus  string   `xml:"cp:contentStatus,omitempty"`
}

// dcTime represents a date/time in Dublin Core format
type dcTime struct {
	Type  string `xml:"xsi:type,attr"`
	Value string `xml:",chardata"`
}

const (
	cpNamespace       = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
	dcNamespace       = "http://purl.org/dc/elements/1.1/"
	dctermsNamespace  = "http://purl.org/dc/terms/"
	dcmitypeNamespace = "http://purl.org/dc/dcmitype/"
	xsiNamespace      = "http://www.w3.org/2001/XMLSchema-instance"
	w3cDateFormat     = "2006-01-02T15:04:05Z"
)

// GetProperties returns the document properties.
// For documents created from scratch, this returns default/empty properties.
// For parsed documents, this extracts properties from docProps/core.xml.
func (d *DocxTmpl) GetProperties() *DocumentProperties {
	// If we have in-memory properties, return them
	if d.properties != nil {
		return d.properties
	}

	props := &DocumentProperties{}

	// Try to find and parse core.xml from processable files
	for _, pf := range d.processableFiles {
		if strings.HasSuffix(pf.Name, "core.xml") {
			var core coreProperties
			if err := xml.Unmarshal([]byte(pf.Content), &core); err == nil {
				props.Title = core.Title
				props.Subject = core.Subject
				props.Creator = core.Creator
				props.Keywords = core.Keywords
				props.Description = core.Description
				props.LastModifiedBy = core.LastModifiedBy
				props.Revision = core.Revision
				props.Category = core.Category
				props.ContentStatus = core.ContentStatus

				if core.Created != nil {
					if t, err := time.Parse(w3cDateFormat, core.Created.Value); err == nil {
						props.Created = t
					}
				}
				if core.Modified != nil {
					if t, err := time.Parse(w3cDateFormat, core.Modified.Value); err == nil {
						props.Modified = t
					}
				}
			}
			d.properties = props
			return props
		}
	}

	return props
}

// SetProperties updates the document properties.
// The properties will be written when the document is saved.
func (d *DocxTmpl) SetProperties(props *DocumentProperties) {
	// Store in memory
	d.properties = props

	// Also serialize to processable files for saving
	core := &coreProperties{
		XMLNScp:        cpNamespace,
		XMLNSdc:        dcNamespace,
		XMLNSdcterms:   dctermsNamespace,
		XMLNSdcmitype:  dcmitypeNamespace,
		XMLNSxsi:       xsiNamespace,
		Title:          props.Title,
		Subject:        props.Subject,
		Creator:        props.Creator,
		Keywords:       props.Keywords,
		Description:    props.Description,
		LastModifiedBy: props.LastModifiedBy,
		Revision:       props.Revision,
		Category:       props.Category,
		ContentStatus:  props.ContentStatus,
	}

	if !props.Created.IsZero() {
		core.Created = &dcTime{
			Type:  "dcterms:W3CDTF",
			Value: props.Created.Format(w3cDateFormat),
		}
	}
	if !props.Modified.IsZero() {
		core.Modified = &dcTime{
			Type:  "dcterms:W3CDTF",
			Value: props.Modified.Format(w3cDateFormat),
		}
	}

	// Marshal to XML
	output, err := xml.MarshalIndent(core, "", "  ")
	if err != nil {
		return
	}
	content := xml.Header + string(output)

	// Update or add to processable files
	found := false
	for i, pf := range d.processableFiles {
		if strings.HasSuffix(pf.Name, "core.xml") {
			d.processableFiles[i].Content = content
			found = true
			break
		}
	}

	if !found {
		d.processableFiles = append(d.processableFiles, headerfooter.DocxFile{
			Name:    "docProps/core.xml",
			Content: content,
		})
	}
}

// SetTitle sets the document title.
func (d *DocxTmpl) SetTitle(title string) *DocxTmpl {
	props := d.GetProperties()
	props.Title = title
	d.SetProperties(props)
	return d
}

// SetAuthor sets the document author (creator).
func (d *DocxTmpl) SetAuthor(author string) *DocxTmpl {
	props := d.GetProperties()
	props.Creator = author
	d.SetProperties(props)
	return d
}

// SetSubject sets the document subject.
func (d *DocxTmpl) SetSubject(subject string) *DocxTmpl {
	props := d.GetProperties()
	props.Subject = subject
	d.SetProperties(props)
	return d
}

// SetKeywords sets the document keywords.
func (d *DocxTmpl) SetKeywords(keywords string) *DocxTmpl {
	props := d.GetProperties()
	props.Keywords = keywords
	d.SetProperties(props)
	return d
}

// SetDescription sets the document description/comments.
func (d *DocxTmpl) SetDescription(description string) *DocxTmpl {
	props := d.GetProperties()
	props.Description = description
	d.SetProperties(props)
	return d
}

// SetCategory sets the document category.
func (d *DocxTmpl) SetCategory(category string) *DocxTmpl {
	props := d.GetProperties()
	props.Category = category
	d.SetProperties(props)
	return d
}

// SetContentStatus sets the document status (e.g., "Draft", "Final", "Approved").
func (d *DocxTmpl) SetContentStatus(status string) *DocxTmpl {
	props := d.GetProperties()
	props.ContentStatus = status
	d.SetProperties(props)
	return d
}
