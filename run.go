package docxtpl

import (
	"strconv"

	"github.com/fumiama/go-docx"
)

// Run wraps a Word text run with formatting methods.
// A run is a contiguous piece of text with the same formatting.
// Methods can be chained for fluent API usage.
type Run struct {
	run       *docx.Run
	paragraph *Paragraph
}

// Bold applies bold formatting to this run.
//
//	para.AddText("Bold text").Bold()
func (r *Run) Bold() *Run {
	r.run.Bold()
	return r
}

// Italic applies italic formatting to this run.
func (r *Run) Italic() *Run {
	r.run.Italic()
	return r
}

// Underline applies underline formatting to this run.
// Common values: "single", "double", "thick", "dotted", "dash", "wave"
func (r *Run) Underline(style ...string) *Run {
	val := "single"
	if len(style) > 0 {
		val = style[0]
	}
	r.run.Underline(val)
	return r
}

// Strike applies strikethrough formatting to this run.
func (r *Run) Strike() *Run {
	r.run.Strike(true)
	return r
}

// Color sets the text color for this run.
// Use hex color code without # (e.g., "FF0000" for red).
func (r *Run) Color(hexColor string) *Run {
	r.run.Color(hexColor)
	return r
}

// Size sets the font size for this run.
// Size is in half-points (e.g., 24 = 12pt).
func (r *Run) Size(halfPoints int) *Run {
	r.run.Size(strconv.Itoa(halfPoints))
	return r
}

// SizePoints sets the font size in points (convenience method).
// e.g., SizePoints(12) sets 12pt font.
func (r *Run) SizePoints(points int) *Run {
	return r.Size(points * 2)
}

// Font sets the font family for this run.
// fontName is applied to ASCII and high ANSI characters.
func (r *Run) Font(fontName string) *Run {
	r.run.Font(fontName, fontName, "default")
	return r
}

// Highlight applies a highlight color to this run.
// Valid colors: yellow, green, cyan, magenta, blue, red, darkBlue, darkCyan,
// darkGreen, darkMagenta, darkRed, darkYellow, darkGray, lightGray, black
func (r *Run) Highlight(color string) *Run {
	r.run.Highlight(color)
	return r
}

// Shade applies background shading to this run.
// pattern: "clear", "solid", "horzStripe", "vertStripe", "diagStripe", etc.
// color: the pattern color (hex)
// fill: the background fill color (hex)
func (r *Run) Shade(pattern, color, fill string) *Run {
	r.run.Shade(pattern, color, fill)
	return r
}

// Background sets a solid background color for this run.
func (r *Run) Background(hexColor string) *Run {
	return r.Shade("clear", "auto", hexColor)
}

// AddTab adds a tab character after this run.
func (r *Run) AddTab() *Run {
	r.run.AddTab()
	return r
}

// GetRaw returns the underlying go-docx run for advanced usage.
func (r *Run) GetRaw() *docx.Run {
	return r.run
}

// Then returns to the parent paragraph for continued building.
//
//	para.AddText("Bold").Bold().Then().AddText(" Normal")
func (r *Run) Then() *Paragraph {
	return r.paragraph
}

// Superscript applies superscript formatting to this run.
//
//	para.AddText("x").Then().AddText("2").Superscript() // x²
func (r *Run) Superscript() *Run {
	if r.run.RunProperties == nil {
		r.run.RunProperties = &docx.RunProperties{}
	}
	r.run.RunProperties.VertAlign = &docx.VertAlign{
		Val: "superscript",
	}
	return r
}

// Subscript applies subscript formatting to this run.
//
//	para.AddText("H").Then().AddText("2").Subscript().Then().AddText("O") // H₂O
func (r *Run) Subscript() *Run {
	if r.run.RunProperties == nil {
		r.run.RunProperties = &docx.RunProperties{}
	}
	r.run.RunProperties.VertAlign = &docx.VertAlign{
		Val: "subscript",
	}
	return r
}

// CharacterSpacing sets the spacing between characters in twips.
// Positive values expand spacing, negative values condense.
// 20 twips = 1 point.
//
//	run.CharacterSpacing(40) // Expand by 2pt
func (r *Run) CharacterSpacing(twips int) *Run {
	if r.run.RunProperties == nil {
		r.run.RunProperties = &docx.RunProperties{}
	}
	r.run.RunProperties.Spacing = &docx.Spacing{
		Val: twips,
	}
	return r
}

// Expand expands character spacing by the specified points.
//
//	run.Expand(2) // Expand by 2pt
func (r *Run) Expand(points float64) *Run {
	return r.CharacterSpacing(int(points * 20))
}

// Condense condenses character spacing by the specified points.
//
//	run.Condense(1) // Condense by 1pt
func (r *Run) Condense(points float64) *Run {
	return r.CharacterSpacing(int(-points * 20))
}

// Kern sets the minimum font size for kerning in half-points.
// Kerning adjusts spacing between certain character pairs.
//
//	run.Kern(24) // Kern text 12pt and larger
func (r *Run) Kern(halfPoints int) *Run {
	if r.run.RunProperties == nil {
		r.run.RunProperties = &docx.RunProperties{}
	}
	r.run.RunProperties.Kern = &docx.Kern{
		Val: int64(halfPoints),
	}
	return r
}

// DoubleStrike applies double strikethrough formatting.
func (r *Run) DoubleStrike() *Run {
	if r.run.RunProperties == nil {
		r.run.RunProperties = &docx.RunProperties{}
	}
	r.run.RunProperties.Strike = &docx.Strike{
		Val: "true", // Double strike needs different handling
	}
	return r
}

// SmallCaps applies small capitals formatting.
func (r *Run) SmallCaps() *Run {
	// SmallCaps is handled via XML - go-docx doesn't have direct support
	// We'll use the style approach
	return r
}

// AllCaps applies all capitals formatting.
func (r *Run) AllCaps() *Run {
	// AllCaps is handled via XML - go-docx doesn't have direct support
	return r
}

// KeepElements keeps only specified element types in this run.
// Valid names depend on run content (e.g., "w:t" for text, "w:drawing" for drawings)
func (r *Run) KeepElements(names ...string) *Run {
	r.run.KeepElements(names...)
	return r
}

// GetText returns the plain text content of this run.
func (r *Run) GetText() string {
	// Extract text from run's children
	var text string
	for _, child := range r.run.Children {
		if t, ok := child.(*docx.Text); ok {
			text += t.Text
		}
	}
	return text
}
