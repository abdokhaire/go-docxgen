/*
   Copyright (c) 2020 gingfrederik
   Copyright (c) 2021 Gonzalo Fernandez-Victorio
   Copyright (c) 2021 Basement Crowd Ltd (https://www.basementcrowd.com)
   Copyright (c) 2023 Fumiama Minamoto (源文雨)

   This program is free software: you can redistribute it and/or modify
   it under the terms of the GNU Affero General Public License as published
   by the Free Software Foundation, either version 3 of the License, or
   (at your option) any later version.

   This program is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU Affero General Public License for more details.

   You should have received a copy of the GNU Affero General Public License
   along with this program.  If not, see <https://www.gnu.org/licenses/>.
*/

package docx

import (
	"encoding/xml"
	"io"
	"reflect"
	"strings"
)

// Run is part of a paragraph that has its own style. It could be
// a piece of text in bold, or a link
type Run struct {
	XMLName xml.Name `xml:"w:r,omitempty"`
	Space   string   `xml:"xml:space,attr,omitempty"`
	// RsidR   string   `xml:"w:rsidR,attr,omitempty"`
	// RsidRPr string   `xml:"w:rsidRPr,attr,omitempty"`

	RunProperties *RunProperties `xml:"w:rPr,omitempty"`

	InstrText string `xml:"w:instrText,omitempty"`

	Children []interface{}

	file *Docx
}

// UnmarshalXML ...
func (r *Run) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "space":
			r.Space = attr.Value
		/*case "rsidR":
			r.RsidR = attr.Value
		case "rsidRPr":
			r.RsidRPr = attr.Value*/
		default:
			// ignore other attributes
		}
	}
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		if tt, ok := t.(xml.StartElement); ok {
			child, err := r.parse(d, tt)
			if err != nil {
				return err
			}
			if child != nil {
				r.Children = append(r.Children, child)
			}
		}
	}

	return nil
}

func (r *Run) parse(d *xml.Decoder, tt xml.StartElement) (child interface{}, err error) {
	switch tt.Name.Local {
	case "rPr":
		var value RunProperties
		err = d.DecodeElement(&value, &tt)
		if err != nil && !strings.HasPrefix(err.Error(), "expected") {
			return nil, err
		}
		r.RunProperties = &value
		return nil, nil
	case "instrText":
		var value string
		err = d.DecodeElement(&value, &tt)
		if err != nil && !strings.HasPrefix(err.Error(), "expected") {
			return nil, err
		}
		r.InstrText = value
		return nil, nil
	case "t":
		var value Text
		err = d.DecodeElement(&value, &tt)
		if err != nil && !strings.HasPrefix(err.Error(), "expected") {
			return nil, err
		}
		child = &value
	case "drawing":
		var value Drawing
		value.file = r.file
		err = d.DecodeElement(&value, &tt)
		if err != nil && !strings.HasPrefix(err.Error(), "expected") {
			return nil, err
		}
		child = &value
	case "tab":
		child = &Tab{}
	case "br":
		var value BarterRabbet
		err = d.DecodeElement(&value, &tt)
		if err != nil {
			return nil, err
		}
		child = &value
	case "fldChar":
		var value FldChar
		// Capture attributes first
		for _, attr := range tt.Attr {
			switch attr.Name.Local {
			case "fldCharType":
				value.FldCharType = attr.Value
			case "fldLock":
				value.FldLock = attr.Value
			case "dirty":
				value.Dirty = attr.Value
			}
		}
		// Parse children (ffData, fldData, etc.) if present
		for {
			tok, err1 := d.Token()
			if err1 == io.EOF {
				break
			}
			if err1 != nil {
				return nil, err1
			}
			if et, ok := tok.(xml.EndElement); ok {
				if et.Name.Local == "fldChar" {
					break
				}
			}
			if st, ok := tok.(xml.StartElement); ok {
				if st.Name.Local == "ffData" {
					var ffData FFData
					err = d.DecodeElement(&ffData, &st)
					if err != nil && !strings.HasPrefix(err.Error(), "expected") {
						return nil, err
					}
					value.FFData = &ffData
				} else {
					// Skip other child elements
					err = d.Skip()
					if err != nil {
						return nil, err
					}
				}
			}
		}
		child = &value
	case "delText":
		var value DelText
		err = d.DecodeElement(&value, &tt)
		if err != nil && !strings.HasPrefix(err.Error(), "expected") {
			return nil, err
		}
		child = &value
	case "sym":
		var value Sym
		for _, attr := range tt.Attr {
			switch attr.Name.Local {
			case "font":
				value.Font = attr.Value
			case "char":
				value.Char = attr.Value
			}
		}
		err = d.Skip()
		if err != nil {
			return nil, err
		}
		child = &value
	case "footnoteReference":
		var value FootnoteReference
		for _, attr := range tt.Attr {
			if attr.Name.Local == "id" {
				value.ID = attr.Value
				break
			}
		}
		err = d.Skip()
		if err != nil {
			return nil, err
		}
		child = &value
	case "endnoteReference":
		var value EndnoteReference
		for _, attr := range tt.Attr {
			if attr.Name.Local == "id" {
				value.ID = attr.Value
				break
			}
		}
		err = d.Skip()
		if err != nil {
			return nil, err
		}
		child = &value
	case "commentReference":
		var value CommentReference
		for _, attr := range tt.Attr {
			if attr.Name.Local == "id" {
				value.ID = attr.Value
				break
			}
		}
		err = d.Skip()
		if err != nil {
			return nil, err
		}
		child = &value
	case "lastRenderedPageBreak":
		child = &LastRenderedPageBreak{}
		err = d.Skip()
		if err != nil {
			return nil, err
		}
	case "noBreakHyphen":
		child = &NoBreakHyphen{}
		err = d.Skip()
		if err != nil {
			return nil, err
		}
	case "softHyphen":
		child = &SoftHyphen{}
		err = d.Skip()
		if err != nil {
			return nil, err
		}
	case "cr":
		child = &CR{}
		err = d.Skip()
		if err != nil {
			return nil, err
		}
	case "pgNum":
		child = &PgNum{}
		err = d.Skip()
		if err != nil {
			return nil, err
		}
	case "AlternateContent":
		/*var value AlternateContent
		value.file = r.file
		err = d.DecodeElement(&value, &tt)
		if err != nil && !strings.HasPrefix(err.Error(), "expected") {
			return nil, err
		}
		if value.Choice == nil {
			return nil, nil
		}
		if value.Choice.Requires != "wps" && value.Choice.Requires != "wpc" && value.Choice.Requires != "wpg" {
			return nil, nil
		}
		child = &value*/
	altcont:
		for {
			tok, err1 := d.Token()
			if err1 == io.EOF {
				break
			}
			if err1 != nil {
				return nil, err1
			}

			if ttt, ok := tok.(xml.StartElement); ok && ttt.Name.Local == "Choice" {
				for _, attr := range ttt.Attr {
					if attr.Name.Local == "Requires" {
						if attr.Value == "wps" || attr.Value == "wpc" || attr.Value == "wpg" {
							tok, err = d.Token() // go into choice
							if err != nil {
								return nil, err
							}
							if ttt, ok := tok.(xml.StartElement); ok {
								child, err = r.parse(d, ttt)
							}
							break altcont
						}
						break
					}
				}
			}
			if et, ok := tok.(xml.EndElement); ok {
				if et.Name.Local == "AlternateContent" {
					break
				}
			}
			err = d.Skip() // skip unsupported tags
			if err != nil {
				return nil, err
			}
		}
	default:
		err = d.Skip() // skip unsupported tags
	}
	return
}

// KeepElements keep named elems amd removes others
//
// names: *docx.Text *docx.Drawing *docx.Tab *docx.BarterRabbet
func (r *Run) KeepElements(name ...string) {
	items := make([]interface{}, 0, len(r.Children))
	namemap := make(map[string]struct{}, len(name)*2)
	for _, n := range name {
		namemap[n] = struct{}{}
	}
	for _, item := range r.Children {
		_, ok := namemap[reflect.ValueOf(item).Type().String()]
		if ok {
			items = append(items, item)
		}
	}
	r.Children = items
}

// FldChar represents a complex field character (begin, separate, end)
// This is used for Word field constructs like MERGEFIELD, IF, QUOTE, etc.
type FldChar struct {
	XMLName     xml.Name `xml:"w:fldChar,omitempty"`
	FldCharType string   `xml:"w:fldCharType,attr,omitempty"`
	FldLock     string   `xml:"w:fldLock,attr,omitempty"`
	Dirty       string   `xml:"w:dirty,attr,omitempty"`
	// FFData contains form field data (for checkboxes, text inputs, dropdowns)
	FFData *FFData `xml:"w:ffData,omitempty"`
}

// FFData represents form field data inside fldChar
type FFData struct {
	XMLName xml.Name `xml:"w:ffData,omitempty"`
	// Store raw inner XML to preserve all form field properties
	InnerXML string `xml:",innerxml"`
}

// DelText represents deleted text in tracked changes
type DelText struct {
	XMLName  xml.Name `xml:"w:delText,omitempty"`
	XMLSpace string   `xml:"xml:space,attr,omitempty"`
	Text     string   `xml:",chardata"`
}

// Sym represents a symbol character
type Sym struct {
	XMLName xml.Name `xml:"w:sym,omitempty"`
	Font    string   `xml:"w:font,attr,omitempty"`
	Char    string   `xml:"w:char,attr,omitempty"`
}

// FootnoteReference represents a reference to a footnote
type FootnoteReference struct {
	XMLName xml.Name `xml:"w:footnoteReference,omitempty"`
	ID      string   `xml:"w:id,attr,omitempty"`
}

// EndnoteReference represents a reference to an endnote
type EndnoteReference struct {
	XMLName xml.Name `xml:"w:endnoteReference,omitempty"`
	ID      string   `xml:"w:id,attr,omitempty"`
}

// CommentReference represents a reference to a comment
type CommentReference struct {
	XMLName xml.Name `xml:"w:commentReference,omitempty"`
	ID      string   `xml:"w:id,attr,omitempty"`
}

// LastRenderedPageBreak marks where the last page break was rendered
type LastRenderedPageBreak struct {
	XMLName xml.Name `xml:"w:lastRenderedPageBreak,omitempty"`
}

// NoBreakHyphen represents a non-breaking hyphen
type NoBreakHyphen struct {
	XMLName xml.Name `xml:"w:noBreakHyphen,omitempty"`
}

// SoftHyphen represents a soft hyphen
type SoftHyphen struct {
	XMLName xml.Name `xml:"w:softHyphen,omitempty"`
}

// CR represents a carriage return
type CR struct {
	XMLName xml.Name `xml:"w:cr,omitempty"`
}

// PgNum represents a page number field
type PgNum struct {
	XMLName xml.Name `xml:"w:pgNum,omitempty"`
}

// RawXML preserves unknown XML elements that we don't explicitly handle.
// This ensures we don't lose data when round-tripping documents.
type RawXML struct {
	XMLName xml.Name
	XML     string `xml:",innerxml"`
	Attrs   []xml.Attr
}

// MarshalXML outputs the raw XML element
func (r *RawXML) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = r.XMLName
	start.Attr = r.Attrs
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if r.XML != "" {
		// Write inner XML as raw characters using CharData
		if err := e.EncodeToken(xml.CharData(r.XML)); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// RunProperties encapsulates visual properties of a run
type RunProperties struct {
	XMLName   xml.Name `xml:"w:rPr,omitempty"`
	Fonts     *RunFonts
	Bold      *Bold
	ICs       *struct{} `xml:"w:iCs,omitempty"`
	Italic    *Italic
	Highlight *Highlight
	Color     *Color
	Size      *Size
	SizeCs    *SizeCs
	Spacing   *Spacing
	RunStyle  *RunStyle
	Style     *Style
	Shade     *Shade
	Kern      *Kern
	Underline *Underline
	VertAlign *VertAlign
	Strike    *Strike
}

// UnmarshalXML ...
func (r *RunProperties) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}

		if tt, ok := t.(xml.StartElement); ok {
			switch tt.Name.Local {
			case "rFonts":
				var value RunFonts
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				r.Fonts = &value
			case "b":
				r.Bold = &Bold{}
			case "iCs":
				r.ICs = &struct{}{}
			case "i":
				r.Italic = &Italic{}
			case "u":
				var value Underline
				value.Val = getAtt(tt.Attr, "val")
				r.Underline = &value
			case "highlight":
				var value Highlight
				value.Val = getAtt(tt.Attr, "val")
				r.Highlight = &value
			case "color":
				var value Color
				value.Val = getAtt(tt.Attr, "val")
				r.Color = &value
			case "sz":
				var value Size
				value.Val = getAtt(tt.Attr, "val")
				r.Size = &value
			case "spacing":
				var value Spacing
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				r.Spacing = &value
			case "szCs":
				var value SizeCs
				value.Val = getAtt(tt.Attr, "val")
				r.SizeCs = &value
			case "rStyle":
				var value RunStyle
				value.Val = getAtt(tt.Attr, "val")
				r.RunStyle = &value
			case "pStyle":
				var value Style
				value.Val = getAtt(tt.Attr, "val")
				r.Style = &value
			case "shd":
				var value Shade
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				r.Shade = &value
			case "kern":
				var value Kern
				v := getAtt(tt.Attr, "val")
				if v == "" {
					continue
				}
				value.Val, err = GetInt64(v)
				if err != nil {
					return err
				}
				r.Kern = &value
			case "vertAlign":
				var value VertAlign
				value.Val = getAtt(tt.Attr, "val")
				r.VertAlign = &value
			case "strike":
				var value Strike
				value.Val = getAtt(tt.Attr, "val")
				r.Strike = &value
			default:
				err = d.Skip() // skip unsupported tags
				if err != nil {
					return err
				}
				continue
			}
		}
	}

	return nil
}

// RunFonts specifies the fonts used in the text of a run.
type RunFonts struct {
	XMLName  xml.Name `xml:"w:rFonts,omitempty"`
	ASCII    string   `xml:"w:ascii,attr,omitempty"`
	EastAsia string   `xml:"w:eastAsia,attr,omitempty"`
	HAnsi    string   `xml:"w:hAnsi,attr,omitempty"`
	Hint     string   `xml:"w:hint,attr,omitempty"`
}

// UnmarshalXML ...
func (f *RunFonts) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "ascii":
			f.ASCII = attr.Value
		case "eastAsia":
			f.EastAsia = attr.Value
		case "hAnsi":
			f.HAnsi = attr.Value
		case "hint":
			f.Hint = attr.Value
		}
	}
	// Consume the end element
	_, err := d.Token()
	return err
}
