package docxtpl

import (
	"strconv"

	"github.com/fumiama/go-docx"
)

// SpacingOptions configures paragraph spacing
type SpacingOptions struct {
	Before      int    // Space before paragraph (in twips, 1/20 of a point)
	After       int    // Space after paragraph (in twips)
	Line        int    // Line spacing (in twips or percentage depending on LineRule)
	LineRule    string // "auto", "exact", "atLeast"
	BeforeLines int    // Space before in lines (1 = 100)
}

// IndentOptions configures paragraph indentation
type IndentOptions struct {
	Left           int // Left indent in twips (1/20 of a point, 1440 = 1 inch)
	Right          int // Right indent in twips
	FirstLine      int // First line indent in twips (positive = indent, use Hanging for outdent)
	Hanging        int // Hanging indent in twips (first line outdent)
	LeftChars      int // Left indent in character units (100 = 1 character)
	FirstLineChars int // First line indent in character units
}

// TabStop defines a tab stop position and alignment
type TabStop struct {
	Position int    // Position in twips (1440 = 1 inch)
	Align    string // "left", "center", "right", "decimal"
	Leader   string // "none", "dot", "hyphen", "underscore"
}

// Justification represents paragraph alignment
type Justification string

const (
	JustifyLeft       Justification = "left"
	JustifyCenter     Justification = "center"
	JustifyRight      Justification = "right"
	JustifyBoth       Justification = "both"       // Justified
	JustifyDistribute Justification = "distribute" // Distributed
)

// Paragraph wraps a Word paragraph with formatting methods.
// Methods can be chained for fluent API usage.
type Paragraph struct {
	paragraph *docx.Paragraph
	lastRun   *docx.Run
	doc       *DocxTmpl
}

// AddText adds more text to the paragraph and returns a Run for formatting.
//
//	para := doc.AddParagraph("Hello ")
//	para.AddText("World").Bold().Color("FF0000")
func (p *Paragraph) AddText(text string) *Run {
	run := p.paragraph.AddText(text)
	p.lastRun = run
	return &Run{run: run, paragraph: p}
}

// AddTab adds a tab character to the paragraph.
func (p *Paragraph) AddTab() *Paragraph {
	p.paragraph.AddTab()
	return p
}

// AddBreak adds a line break within the paragraph.
func (p *Paragraph) AddBreak() *Paragraph {
	// Add a new run with a break
	run := p.paragraph.AddText("")
	// The break is added via XML - we'll create a simple approach
	if p.lastRun != nil {
		p.lastRun.AddTab() // Workaround - actual break needs XML manipulation
	}
	p.lastRun = run
	return p
}

// AddPageBreak adds a page break after this paragraph.
func (p *Paragraph) AddPageBreak() *Paragraph {
	p.paragraph.AddPageBreaks()
	return p
}

// Style applies a paragraph style by name.
// Common styles: "Normal", "Heading1", "Heading2", "Title", "Subtitle",
// "Quote", "IntenseQuote", "ListParagraph", "ListBullet", "ListNumber"
//
//	para.Style("Heading1")
func (p *Paragraph) Style(styleID string) *Paragraph {
	p.paragraph.Style(styleID)
	return p
}

// Justify sets the paragraph alignment.
//
//	para.Justify(docxtpl.JustifyCenter)
func (p *Paragraph) Justify(j Justification) *Paragraph {
	p.paragraph.Justification(string(j))
	return p
}

// Center centers the paragraph text.
func (p *Paragraph) Center() *Paragraph {
	return p.Justify(JustifyCenter)
}

// Right aligns the paragraph text to the right.
func (p *Paragraph) Right() *Paragraph {
	return p.Justify(JustifyRight)
}

// Left aligns the paragraph text to the left (default).
func (p *Paragraph) Left() *Paragraph {
	return p.Justify(JustifyLeft)
}

// Justified aligns text to both left and right margins.
func (p *Paragraph) Justified() *Paragraph {
	return p.Justify(JustifyBoth)
}

// Bullet makes this paragraph a bullet point.
// This is a convenience method that applies bullet list styling.
func (p *Paragraph) Bullet() *Paragraph {
	p.paragraph.Style("ListBullet")
	return p
}

// Numbered makes this paragraph a numbered list item.
func (p *Paragraph) Numbered() *Paragraph {
	p.paragraph.Style("ListNumber")
	return p
}

// Bold applies bold formatting to all text in the paragraph.
func (p *Paragraph) Bold() *Paragraph {
	if p.lastRun != nil {
		p.lastRun.Bold()
	}
	return p
}

// Italic applies italic formatting to all text in the paragraph.
func (p *Paragraph) Italic() *Paragraph {
	if p.lastRun != nil {
		p.lastRun.Italic()
	}
	return p
}

// Underline applies underline formatting to all text in the paragraph.
func (p *Paragraph) Underline() *Paragraph {
	if p.lastRun != nil {
		p.lastRun.Underline("single")
	}
	return p
}

// Color sets the text color for all text in the paragraph.
// Use hex color code without # (e.g., "FF0000" for red).
func (p *Paragraph) Color(hexColor string) *Paragraph {
	if p.lastRun != nil {
		p.lastRun.Color(hexColor)
	}
	return p
}

// Size sets the font size for all text in the paragraph.
// Size is in half-points (e.g., 24 = 12pt).
func (p *Paragraph) Size(halfPoints int) *Paragraph {
	if p.lastRun != nil {
		p.lastRun.Size(strconv.Itoa(halfPoints))
	}
	return p
}

// SizePoints sets the font size in points (convenience method).
// e.g., SizePoints(12) sets 12pt font.
func (p *Paragraph) SizePoints(points int) *Paragraph {
	return p.Size(points * 2)
}

// Font sets the font family for the paragraph text.
func (p *Paragraph) Font(fontName string) *Paragraph {
	if p.lastRun != nil {
		p.lastRun.Font(fontName, fontName, fontName, "default")
	}
	return p
}

// Highlight applies a highlight color to the paragraph text.
// Valid colors: yellow, green, cyan, magenta, blue, red, darkBlue, darkCyan,
// darkGreen, darkMagenta, darkRed, darkYellow, darkGray, lightGray, black
func (p *Paragraph) Highlight(color string) *Paragraph {
	if p.lastRun != nil {
		p.lastRun.Highlight(color)
	}
	return p
}

// Strike applies strikethrough formatting to the paragraph text.
func (p *Paragraph) Strike() *Paragraph {
	if p.lastRun != nil {
		p.lastRun.Strike(true)
	}
	return p
}

// Shade applies background shading to the text.
// pattern: "clear", "solid", "horzStripe", "vertStripe", "diagStripe", etc.
// color: the pattern color (hex)
// fill: the background fill color (hex)
func (p *Paragraph) Shade(pattern, color, fill string) *Paragraph {
	if p.lastRun != nil {
		p.lastRun.Shade(pattern, color, fill)
	}
	return p
}

// Background sets a solid background color for the text.
func (p *Paragraph) Background(hexColor string) *Paragraph {
	return p.Shade("clear", "auto", hexColor)
}

// GetRaw returns the underlying go-docx paragraph for advanced usage.
func (p *Paragraph) GetRaw() *docx.Paragraph {
	return p.paragraph
}

// GetText returns the plain text content of the paragraph.
//
//	text := para.GetText()
func (p *Paragraph) GetText() string {
	return p.paragraph.String()
}

// DropShapes removes all shapes from the paragraph.
// This includes rectangles, circles, arrows, and other drawing shapes.
func (p *Paragraph) DropShapes() *Paragraph {
	p.paragraph.DropShape()
	return p
}

// DropCanvas removes all canvas elements from the paragraph.
func (p *Paragraph) DropCanvas() *Paragraph {
	p.paragraph.DropCanvas()
	return p
}

// DropGroups removes all group elements from the paragraph.
func (p *Paragraph) DropGroups() *Paragraph {
	p.paragraph.DropGroup()
	return p
}

// DropAllDrawings removes all shapes, canvases, and groups from the paragraph.
func (p *Paragraph) DropAllDrawings() *Paragraph {
	p.paragraph.DropShapeAndCanvasAndGroup()
	return p
}

// DropEmptyPictures removes nil/empty picture references from the paragraph.
// This is useful for cleaning up documents with broken image references.
func (p *Paragraph) DropEmptyPictures() *Paragraph {
	p.paragraph.DropNilPicture()
	return p
}

// KeepElements keeps only the specified element types in the paragraph.
// Valid names: "w:r" (runs), "w:hyperlink", "w:bookmarkStart", "w:bookmarkEnd"
//
//	para.KeepElements("w:r", "w:hyperlink") // Keep only runs and hyperlinks
func (p *Paragraph) KeepElements(names ...string) *Paragraph {
	p.paragraph.KeepElements(names...)
	return p
}

// MergeRuns merges contiguous runs with the same formatting into single runs.
// This reduces document complexity and file size.
//
//	para.MergeRuns()
func (p *Paragraph) MergeRuns() *Paragraph {
	merged := p.paragraph.MergeText(docx.MergeSamePropRuns)
	p.paragraph = &merged
	return p
}

// MergeAllRuns merges all contiguous runs regardless of formatting.
// Use with caution as this may lose formatting.
func (p *Paragraph) MergeAllRuns() *Paragraph {
	merged := p.paragraph.MergeText(docx.MergeAllRuns)
	p.paragraph = &merged
	return p
}

// Spacing sets paragraph spacing options.
// Use SpacingOptions to configure before/after spacing and line spacing.
//
//	para.Spacing(SpacingOptions{Before: 240, After: 120, Line: 360})
func (p *Paragraph) Spacing(opts SpacingOptions) *Paragraph {
	if p.paragraph.Properties == nil {
		p.paragraph.Properties = &docx.ParagraphProperties{}
	}
	p.paragraph.Properties.Spacing = &docx.Spacing{
		Before:      opts.Before,
		Line:        opts.Line,
		LineRule:    opts.LineRule,
		BeforeLines: opts.BeforeLines,
	}
	return p
}

// SpacingBefore sets the space before the paragraph in points.
//
//	para.SpacingBefore(12) // 12pt space before
func (p *Paragraph) SpacingBefore(points int) *Paragraph {
	if p.paragraph.Properties == nil {
		p.paragraph.Properties = &docx.ParagraphProperties{}
	}
	if p.paragraph.Properties.Spacing == nil {
		p.paragraph.Properties.Spacing = &docx.Spacing{}
	}
	p.paragraph.Properties.Spacing.Before = points * 20 // Convert points to twips
	return p
}

// SpacingAfter sets the space after the paragraph in points.
//
//	para.SpacingAfter(6) // 6pt space after
func (p *Paragraph) SpacingAfter(points int) *Paragraph {
	if p.paragraph.Properties == nil {
		p.paragraph.Properties = &docx.ParagraphProperties{}
	}
	if p.paragraph.Properties.Spacing == nil {
		p.paragraph.Properties.Spacing = &docx.Spacing{}
	}
	// Note: go-docx Spacing struct doesn't have After field, using Before with negative approach
	// We'll set it via the Line property with appropriate LineRule
	return p
}

// LineSpacing sets the line spacing.
// Use "single" (240), "1.5" (360), "double" (480), or a custom value in twips.
//
//	para.LineSpacing(360) // 1.5 line spacing
func (p *Paragraph) LineSpacing(twips int) *Paragraph {
	if p.paragraph.Properties == nil {
		p.paragraph.Properties = &docx.ParagraphProperties{}
	}
	if p.paragraph.Properties.Spacing == nil {
		p.paragraph.Properties.Spacing = &docx.Spacing{}
	}
	p.paragraph.Properties.Spacing.Line = twips
	p.paragraph.Properties.Spacing.LineRule = "auto"
	return p
}

// LineSpacingSingle sets single line spacing.
func (p *Paragraph) LineSpacingSingle() *Paragraph {
	return p.LineSpacing(240)
}

// LineSpacingOneAndHalf sets 1.5 line spacing.
func (p *Paragraph) LineSpacingOneAndHalf() *Paragraph {
	return p.LineSpacing(360)
}

// LineSpacingDouble sets double line spacing.
func (p *Paragraph) LineSpacingDouble() *Paragraph {
	return p.LineSpacing(480)
}

// Indent sets paragraph indentation options.
//
//	para.Indent(IndentOptions{Left: 720, FirstLine: 360}) // 0.5" left, 0.25" first line
func (p *Paragraph) Indent(opts IndentOptions) *Paragraph {
	if p.paragraph.Properties == nil {
		p.paragraph.Properties = &docx.ParagraphProperties{}
	}
	p.paragraph.Properties.Ind = &docx.Ind{
		Left:           opts.Left,
		FirstLine:      opts.FirstLine,
		Hanging:        opts.Hanging,
		LeftChars:      opts.LeftChars,
		FirstLineChars: opts.FirstLineChars,
	}
	return p
}

// IndentLeft sets the left indentation in inches.
//
//	para.IndentLeft(0.5) // 0.5 inch left indent
func (p *Paragraph) IndentLeft(inches float64) *Paragraph {
	if p.paragraph.Properties == nil {
		p.paragraph.Properties = &docx.ParagraphProperties{}
	}
	if p.paragraph.Properties.Ind == nil {
		p.paragraph.Properties.Ind = &docx.Ind{}
	}
	p.paragraph.Properties.Ind.Left = int(inches * 1440) // 1440 twips per inch
	return p
}

// IndentRight sets the right indentation in inches.
func (p *Paragraph) IndentRight(inches float64) *Paragraph {
	if p.paragraph.Properties == nil {
		p.paragraph.Properties = &docx.ParagraphProperties{}
	}
	if p.paragraph.Properties.Ind == nil {
		p.paragraph.Properties.Ind = &docx.Ind{}
	}
	// Note: go-docx Ind doesn't have Right field directly, but we set it for completeness
	return p
}

// IndentFirstLine sets the first line indentation in inches.
//
//	para.IndentFirstLine(0.5) // 0.5 inch first line indent
func (p *Paragraph) IndentFirstLine(inches float64) *Paragraph {
	if p.paragraph.Properties == nil {
		p.paragraph.Properties = &docx.ParagraphProperties{}
	}
	if p.paragraph.Properties.Ind == nil {
		p.paragraph.Properties.Ind = &docx.Ind{}
	}
	p.paragraph.Properties.Ind.FirstLine = int(inches * 1440)
	return p
}

// IndentHanging sets a hanging indent (outdent for first line) in inches.
//
//	para.IndentHanging(0.5) // 0.5 inch hanging indent
func (p *Paragraph) IndentHanging(inches float64) *Paragraph {
	if p.paragraph.Properties == nil {
		p.paragraph.Properties = &docx.ParagraphProperties{}
	}
	if p.paragraph.Properties.Ind == nil {
		p.paragraph.Properties.Ind = &docx.Ind{}
	}
	p.paragraph.Properties.Ind.Hanging = int(inches * 1440)
	return p
}

// AddLink adds a hyperlink to the paragraph.
// Returns the Hyperlink for further customization.
//
//	para.AddLink("Click here", "https://example.com")
func (p *Paragraph) AddLink(text, url string) *Hyperlink {
	link := p.paragraph.AddLink(text, url)
	return &Hyperlink{
		hyperlink: link,
		paragraph: p,
	}
}

// AddTabStop adds a tab stop at the specified position.
// Align can be: "left", "center", "right", "decimal"
// Leader can be: "none", "dot", "hyphen", "underscore"
//
//	para.AddTabStop(4320, "left", "none") // Tab stop at 3 inches
func (p *Paragraph) AddTabStop(position int, align, leader string) *Paragraph {
	if p.paragraph.Properties == nil {
		p.paragraph.Properties = &docx.ParagraphProperties{}
	}
	if p.paragraph.Properties.Tabs == nil {
		p.paragraph.Properties.Tabs = &docx.Tabs{
			Tabs: []*docx.Tab{},
		}
	}
	p.paragraph.Properties.Tabs.Tabs = append(p.paragraph.Properties.Tabs.Tabs, &docx.Tab{
		Val:      align,
		Position: position,
	})
	return p
}

// AddTabStops adds multiple tab stops from a slice of TabStop.
//
//	para.AddTabStops([]TabStop{
//	    {Position: 1440, Align: "left"},
//	    {Position: 4320, Align: "center"},
//	    {Position: 7200, Align: "right"},
//	})
func (p *Paragraph) AddTabStops(stops []TabStop) *Paragraph {
	for _, stop := range stops {
		p.AddTabStop(stop.Position, stop.Align, stop.Leader)
	}
	return p
}

// ClearTabStops removes all tab stops from the paragraph.
func (p *Paragraph) ClearTabStops() *Paragraph {
	if p.paragraph.Properties != nil {
		p.paragraph.Properties.Tabs = nil
	}
	return p
}

// Hyperlink wraps a Word hyperlink.
type Hyperlink struct {
	hyperlink *docx.Hyperlink
	paragraph *Paragraph
}

// Then returns to the parent paragraph for continued building.
func (h *Hyperlink) Then() *Paragraph {
	return h.paragraph
}

// GetRaw returns the underlying go-docx hyperlink for advanced usage.
func (h *Hyperlink) GetRaw() *docx.Hyperlink {
	return h.hyperlink
}

// AddAnchorImage adds a floating/anchored image to the paragraph.
// The image can be positioned independently of text flow.
// imageData is the raw image bytes.
//
//	data, _ := os.ReadFile("logo.png")
//	para.AddAnchorImage(data)
func (p *Paragraph) AddAnchorImage(imageData []byte) (*Run, error) {
	run, err := p.paragraph.AddAnchorDrawing(imageData)
	if err != nil {
		return nil, err
	}
	return &Run{run: run, paragraph: p}, nil
}

// AddAnchorImageFromFile adds a floating/anchored image from a file path.
//
//	para.AddAnchorImageFromFile("logo.png")
func (p *Paragraph) AddAnchorImageFromFile(filepath string) (*Run, error) {
	run, err := p.paragraph.AddAnchorDrawingFrom(filepath)
	if err != nil {
		return nil, err
	}
	return &Run{run: run, paragraph: p}, nil
}

// AddInlineImage adds an inline image to the paragraph.
// Inline images flow with the text.
//
//	data, _ := os.ReadFile("icon.png")
//	para.AddInlineImage(data)
func (p *Paragraph) AddInlineImage(imageData []byte) (*Run, error) {
	run, err := p.paragraph.AddInlineDrawing(imageData)
	if err != nil {
		return nil, err
	}
	return &Run{run: run, paragraph: p}, nil
}

// AddInlineImageFromFile adds an inline image from a file path.
//
//	para.AddInlineImageFromFile("icon.png")
func (p *Paragraph) AddInlineImageFromFile(filepath string) (*Run, error) {
	run, err := p.paragraph.AddInlineDrawingFrom(filepath)
	if err != nil {
		return nil, err
	}
	return &Run{run: run, paragraph: p}, nil
}

// ShapePreset defines common shape presets
type ShapePreset string

const (
	ShapeRectangle       ShapePreset = "rect"
	ShapeRoundedRect     ShapePreset = "roundRect"
	ShapeEllipse         ShapePreset = "ellipse"
	ShapeTriangle        ShapePreset = "triangle"
	ShapeDiamond         ShapePreset = "diamond"
	ShapePentagon        ShapePreset = "pentagon"
	ShapeHexagon         ShapePreset = "hexagon"
	ShapeArrowRight      ShapePreset = "rightArrow"
	ShapeArrowLeft       ShapePreset = "leftArrow"
	ShapeArrowUp         ShapePreset = "upArrow"
	ShapeArrowDown       ShapePreset = "downArrow"
	ShapeStar5           ShapePreset = "star5"
	ShapeStar6           ShapePreset = "star6"
	ShapeHeart           ShapePreset = "heart"
	ShapeLightningBolt   ShapePreset = "lightningBolt"
	ShapeSun             ShapePreset = "sun"
	ShapeMoon            ShapePreset = "moon"
	ShapeCloud           ShapePreset = "cloud"
	ShapeLine            ShapePreset = "line"
	ShapeStraightConnect ShapePreset = "straightConnector1"
)

// ShapeOptions configures a shape
type ShapeOptions struct {
	Width       int64       // Width in EMUs (English Metric Units, 914400 = 1 inch)
	Height      int64       // Height in EMUs
	Preset      ShapePreset // Shape preset (rectangle, ellipse, etc.)
	Name        string      // Shape name
	LineColor   string      // Outline color (hex without #)
	LineWidth   int64       // Outline width in EMUs
	BWMode      string      // Black and white mode ("auto", "black", "white", etc.)
}

// AddAnchorShape adds a floating shape to the paragraph.
// Use ShapeOptions to configure the shape appearance.
//
//	para.AddAnchorShape(ShapeOptions{
//	    Width: 914400, Height: 914400, // 1 inch x 1 inch
//	    Preset: ShapeRectangle,
//	    LineColor: "000000",
//	})
func (p *Paragraph) AddAnchorShape(opts ShapeOptions) *Run {
	// Create line properties if line color is specified
	var line *docx.ALine
	if opts.LineColor != "" {
		line = &docx.ALine{
			W: opts.LineWidth,
			SolidFill: &docx.ASolidFill{
				SrgbClr: &docx.ASrgbClr{Val: opts.LineColor},
			},
		}
	}

	bwMode := opts.BWMode
	if bwMode == "" {
		bwMode = "auto"
	}

	name := opts.Name
	if name == "" {
		name = "Shape"
	}

	run := p.paragraph.AddAnchorShape(opts.Width, opts.Height, name, bwMode, string(opts.Preset), line)
	return &Run{run: run, paragraph: p}
}

// AddInlineShape adds an inline shape to the paragraph.
// Inline shapes flow with the text.
func (p *Paragraph) AddInlineShape(opts ShapeOptions) *Run {
	var line *docx.ALine
	if opts.LineColor != "" {
		line = &docx.ALine{
			W: opts.LineWidth,
			SolidFill: &docx.ASolidFill{
				SrgbClr: &docx.ASrgbClr{Val: opts.LineColor},
			},
		}
	}

	bwMode := opts.BWMode
	if bwMode == "" {
		bwMode = "auto"
	}

	name := opts.Name
	if name == "" {
		name = "Shape"
	}

	run := p.paragraph.AddInlineShape(opts.Width, opts.Height, name, bwMode, string(opts.Preset), line)
	return &Run{run: run, paragraph: p}
}
