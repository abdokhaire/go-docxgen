package docxtpl

import (
	"fmt"
	"regexp"
	"strings"
)

// ProtectionType represents the type of document protection
type ProtectionType string

const (
	ProtectionNone           ProtectionType = "none"
	ProtectionReadOnly       ProtectionType = "readOnly"
	ProtectionComments       ProtectionType = "comments"       // Allow only comments
	ProtectionTrackedChanges ProtectionType = "trackedChanges" // Allow only tracked changes
	ProtectionForms          ProtectionType = "forms"          // Allow only form filling
)

// ProtectionInfo contains information about document protection
type ProtectionInfo struct {
	IsProtected    bool           // Whether the document is protected
	Type           ProtectionType // Type of protection
	HasPassword    bool           // Whether a password is set
	EnforceMessage string         // Enforcement message if any
}

// GetProtectionInfo returns information about document protection.
//
//	info := doc.GetProtectionInfo()
//	if info.IsProtected {
//	    fmt.Printf("Document is protected: %s\n", info.Type)
//	}
func (d *DocxTmpl) GetProtectionInfo() *ProtectionInfo {
	info := &ProtectionInfo{
		Type: ProtectionNone,
	}

	// Look for settings.xml
	var settingsContent string
	for _, pf := range d.processableFiles {
		if strings.HasSuffix(pf.Name, "settings.xml") {
			settingsContent = pf.Content
			break
		}
	}

	if settingsContent == "" {
		return info
	}

	// Check for document protection element
	// <w:documentProtection w:edit="readOnly" w:enforcement="1" w:cryptProviderType="..." />
	protPattern := regexp.MustCompile(`<w:documentProtection([^>]*)/>`)
	match := protPattern.FindStringSubmatch(settingsContent)

	if len(match) >= 2 {
		attrs := match[1]
		info.IsProtected = strings.Contains(attrs, `w:enforcement="1"`)

		// Check protection type
		if strings.Contains(attrs, `w:edit="readOnly"`) {
			info.Type = ProtectionReadOnly
		} else if strings.Contains(attrs, `w:edit="comments"`) {
			info.Type = ProtectionComments
		} else if strings.Contains(attrs, `w:edit="trackedChanges"`) {
			info.Type = ProtectionTrackedChanges
		} else if strings.Contains(attrs, `w:edit="forms"`) {
			info.Type = ProtectionForms
		}

		// Check if password is set
		info.HasPassword = strings.Contains(attrs, "w:cryptProviderType") ||
			strings.Contains(attrs, "w:hash") ||
			strings.Contains(attrs, "w:salt")
	}

	return info
}

// IsProtected returns true if the document has any protection enabled.
//
//	if doc.IsProtected() {
//	    fmt.Println("Document is protected")
//	}
func (d *DocxTmpl) IsProtected() bool {
	return d.GetProtectionInfo().IsProtected
}

// IsReadOnly returns true if the document is read-only protected.
//
//	if doc.IsReadOnly() {
//	    fmt.Println("Document is read-only")
//	}
func (d *DocxTmpl) IsReadOnly() bool {
	info := d.GetProtectionInfo()
	return info.IsProtected && info.Type == ProtectionReadOnly
}

// SetProtection sets document protection.
// Note: This sets protection without a password. For password protection,
// use the Word application or a dedicated library.
//
//	err := doc.SetProtection(docxtpl.ProtectionReadOnly)
func (d *DocxTmpl) SetProtection(protType ProtectionType) error {
	// Find settings.xml
	for i := range d.processableFiles {
		if strings.HasSuffix(d.processableFiles[i].Name, "settings.xml") {
			content := d.processableFiles[i].Content

			// Remove existing protection
			content = removeProtection(content)

			if protType != ProtectionNone {
				// Add new protection before </w:settings>
				protXML := fmt.Sprintf(`<w:documentProtection w:edit="%s" w:enforcement="1"/>`, protType)
				content = strings.Replace(content, "</w:settings>", protXML+"</w:settings>", 1)
			}

			d.processableFiles[i].Content = content
			return nil
		}
	}

	return fmt.Errorf("settings.xml not found")
}

// RemoveProtection removes document protection (without password).
// Note: This only works for documents without password protection.
//
//	err := doc.RemoveProtection()
func (d *DocxTmpl) RemoveProtection() error {
	return d.SetProtection(ProtectionNone)
}

// SetReadOnly sets the document to read-only mode.
//
//	err := doc.SetReadOnly()
func (d *DocxTmpl) SetReadOnly() error {
	return d.SetProtection(ProtectionReadOnly)
}

// AllowOnlyComments sets protection to allow only comments.
//
//	err := doc.AllowOnlyComments()
func (d *DocxTmpl) AllowOnlyComments() error {
	return d.SetProtection(ProtectionComments)
}

// AllowOnlyTrackedChanges sets protection to allow only tracked changes.
//
//	err := doc.AllowOnlyTrackedChanges()
func (d *DocxTmpl) AllowOnlyTrackedChanges() error {
	return d.SetProtection(ProtectionTrackedChanges)
}

// AllowOnlyFormFilling sets protection to allow only form filling.
//
//	err := doc.AllowOnlyFormFilling()
func (d *DocxTmpl) AllowOnlyFormFilling() error {
	return d.SetProtection(ProtectionForms)
}

// ProtectionSummary returns a text summary of protection status.
//
//	summary := doc.ProtectionSummary()
func (d *DocxTmpl) ProtectionSummary() string {
	info := d.GetProtectionInfo()

	if !info.IsProtected {
		return "Document is not protected."
	}

	var sb strings.Builder
	sb.WriteString("Document Protection:\n")
	sb.WriteString(strings.Repeat("-", 30) + "\n")
	sb.WriteString("Protected: Yes\n")
	sb.WriteString(fmt.Sprintf("Type: %s\n", info.Type))
	sb.WriteString(fmt.Sprintf("Password: %v\n", info.HasPassword))

	return sb.String()
}

// Helper function
func removeProtection(content string) string {
	protPattern := regexp.MustCompile(`<w:documentProtection[^>]*/>`)
	return protPattern.ReplaceAllString(content, "")
}

// RestrictionInfo provides detailed restriction information
type RestrictionInfo struct {
	CanEdit       bool // Can edit the document
	CanComment    bool // Can add comments
	CanTrack      bool // Can make tracked changes
	CanFillForms  bool // Can fill forms
	CanFormatText bool // Can format text
}

// GetRestrictions returns detailed information about what actions are allowed.
//
//	restrictions := doc.GetRestrictions()
//	if restrictions.CanComment {
//	    fmt.Println("Comments are allowed")
//	}
func (d *DocxTmpl) GetRestrictions() *RestrictionInfo {
	info := d.GetProtectionInfo()

	if !info.IsProtected {
		return &RestrictionInfo{
			CanEdit:       true,
			CanComment:    true,
			CanTrack:      true,
			CanFillForms:  true,
			CanFormatText: true,
		}
	}

	switch info.Type {
	case ProtectionReadOnly:
		return &RestrictionInfo{
			CanEdit:       false,
			CanComment:    false,
			CanTrack:      false,
			CanFillForms:  false,
			CanFormatText: false,
		}
	case ProtectionComments:
		return &RestrictionInfo{
			CanEdit:       false,
			CanComment:    true,
			CanTrack:      false,
			CanFillForms:  false,
			CanFormatText: false,
		}
	case ProtectionTrackedChanges:
		return &RestrictionInfo{
			CanEdit:       false,
			CanComment:    true,
			CanTrack:      true,
			CanFillForms:  false,
			CanFormatText: false,
		}
	case ProtectionForms:
		return &RestrictionInfo{
			CanEdit:       false,
			CanComment:    false,
			CanTrack:      false,
			CanFillForms:  true,
			CanFormatText: false,
		}
	default:
		return &RestrictionInfo{
			CanEdit:       true,
			CanComment:    true,
			CanTrack:      true,
			CanFillForms:  true,
			CanFormatText: true,
		}
	}
}
