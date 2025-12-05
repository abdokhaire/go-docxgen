package docxtpl

// Orientation represents page orientation
type Orientation string

const (
	OrientationPortrait  Orientation = "portrait"
	OrientationLandscape Orientation = "landscape"
)

// SectionBreakType represents the type of section break
type SectionBreakType string

const (
	SectionBreakNextPage   SectionBreakType = "nextPage"
	SectionBreakContinuous SectionBreakType = "continuous"
	SectionBreakEvenPage   SectionBreakType = "evenPage"
	SectionBreakOddPage    SectionBreakType = "oddPage"
)

// Margins represents page margins in inches
type Margins struct {
	Top    float64
	Right  float64
	Bottom float64
	Left   float64
}

// DefaultMargins returns the default Word margins (1 inch all around)
func DefaultMargins() Margins {
	return Margins{
		Top:    1.0,
		Right:  1.0,
		Bottom: 1.0,
		Left:   1.0,
	}
}

// NarrowMargins returns narrow margins (0.5 inch all around)
func NarrowMargins() Margins {
	return Margins{
		Top:    0.5,
		Right:  0.5,
		Bottom: 0.5,
		Left:   0.5,
	}
}

// WideMargins returns wide margins (1 inch top/bottom, 2 inch left/right)
func WideMargins() Margins {
	return Margins{
		Top:    1.0,
		Right:  2.0,
		Bottom: 1.0,
		Left:   2.0,
	}
}

// AddSectionBreak adds a visual section break to the document.
// For SectionBreakNextPage, this creates a page break.
// For other types, it creates appropriate spacing.
//
//	doc.AddSectionBreak(SectionBreakNextPage)
func (d *DocxTmpl) AddSectionBreak(breakType SectionBreakType) *DocxTmpl {
	switch breakType {
	case SectionBreakNextPage:
		d.AddPageBreak()
	case SectionBreakContinuous:
		// Continuous sections don't need a visual break
		d.AddEmptyParagraph()
	case SectionBreakEvenPage, SectionBreakOddPage:
		// These are approximated with page breaks
		d.AddPageBreak()
	}
	return d
}

// AddSection adds a new section with a page break.
// This is a convenience method for common use cases.
//
//	doc.AddSection()
func (d *DocxTmpl) AddSection() *DocxTmpl {
	return d.AddSectionBreak(SectionBreakNextPage)
}

// Custom page size constants (in inches)
const (
	PageWidthA4      = 8.27
	PageHeightA4     = 11.69
	PageWidthA3      = 11.69
	PageHeightA3     = 16.54
	PageWidthLetter  = 8.5
	PageHeightLetter = 11.0
	PageWidthLegal   = 8.5
	PageHeightLegal  = 14.0
)

// EstimatePageCount estimates the number of pages in the document.
// This is a rough estimate based on paragraph count and assumes:
// - Single spacing, 12pt font
// - About 50 lines per page for Letter size
// - Tables count as 3 paragraphs
//
//	pages := doc.EstimatePageCount()
func (d *DocxTmpl) EstimatePageCount() int {
	stats := d.GetStats()

	// Rough estimate: 50 lines per page, average 2 lines per paragraph
	linesPerPage := 50.0
	avgLinesPerPara := 2.0
	linesPerTable := 6.0 // Tables take more space

	totalLines := float64(stats.ParagraphCount)*avgLinesPerPara +
		float64(stats.TableCount)*linesPerTable

	pages := int(totalLines / linesPerPage)
	if pages < 1 {
		pages = 1
	}
	return pages
}
