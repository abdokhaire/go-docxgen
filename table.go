package docxtpl

import (
	"github.com/fumiama/go-docx"
)

// TableBorderColors defines custom border colors for tables.
// Use hex color codes without # (e.g., "000000" for black).
type TableBorderColors struct {
	Top     string // Top border color
	Left    string // Left border color
	Bottom  string // Bottom border color
	Right   string // Right border color
	InsideH string // Inside horizontal border color
	InsideV string // Inside vertical border color
}

// VerticalAlignment represents vertical alignment in table cells.
type VerticalAlignment string

const (
	VAlignTop    VerticalAlignment = "top"
	VAlignCenter VerticalAlignment = "center"
	VAlignBottom VerticalAlignment = "bottom"
)

// Table wraps a Word table with formatting and manipulation methods.
type Table struct {
	table *docx.Table
	doc   *DocxTmpl
	rows  int
	cols  int
}

// Cell returns the cell at the specified row and column (0-indexed).
// Returns nil if the indices are out of bounds.
func (t *Table) Cell(row, col int) *TableCell {
	if row < 0 || row >= t.rows || col < 0 || col >= t.cols {
		return nil
	}
	if row >= len(t.table.TableRows) {
		return nil
	}
	tableRow := t.table.TableRows[row]
	if col >= len(tableRow.TableCells) {
		return nil
	}
	return &TableCell{
		cell:  tableRow.TableCells[col],
		table: t,
		row:   row,
		col:   col,
	}
}

// SetCell sets the text content of a cell at the specified row and column.
// Returns the TableCell for further formatting.
//
//	table.SetCell(0, 0, "Header 1").Bold().Background("CCCCCC")
func (t *Table) SetCell(row, col int, text string) *TableCell {
	cell := t.Cell(row, col)
	if cell == nil {
		return nil
	}
	cell.SetText(text)
	return cell
}

// Row returns a TableRow wrapper for the specified row index.
func (t *Table) Row(index int) *TableRow {
	if index < 0 || index >= t.rows || index >= len(t.table.TableRows) {
		return nil
	}
	return &TableRow{
		row:   t.table.TableRows[index],
		table: t,
		index: index,
	}
}

// Rows returns the number of rows in the table.
func (t *Table) Rows() int {
	return t.rows
}

// Cols returns the number of columns in the table.
func (t *Table) Cols() int {
	return t.cols
}

// Justify sets the table alignment.
// Valid values: "left", "center", "right"
func (t *Table) Justify(alignment string) *Table {
	t.table.Justification(alignment)
	return t
}

// Center centers the table on the page.
func (t *Table) Center() *Table {
	return t.Justify("center")
}

// GetRaw returns the underlying go-docx table for advanced usage.
func (t *Table) GetRaw() *docx.Table {
	return t.table
}

// SetBorderColors sets border colors for all cells in the table.
// Note: For best results, use AddTableWithBorders when creating the table.
func (t *Table) SetBorderColors(colors TableBorderColors) *Table {
	// Apply borders to all cells
	for row := 0; row < t.rows; row++ {
		for col := 0; col < t.cols; col++ {
			cell := t.Cell(row, col)
			if cell != nil {
				if cell.cell.TableCellProperties == nil {
					cell.cell.TableCellProperties = &docx.WTableCellProperties{}
				}
				cell.cell.TableCellProperties.TableBorders = &docx.WTableBorders{
					Top:    &docx.WTableBorder{Val: "single", Color: colors.Top},
					Left:   &docx.WTableBorder{Val: "single", Color: colors.Left},
					Bottom: &docx.WTableBorder{Val: "single", Color: colors.Bottom},
					Right:  &docx.WTableBorder{Val: "single", Color: colors.Right},
				}
			}
		}
	}
	return t
}

// TableRow wraps a Word table row.
type TableRow struct {
	row   *docx.WTableRow
	table *Table
	index int
}

// Cell returns the cell at the specified column in this row.
func (r *TableRow) Cell(col int) *TableCell {
	return r.table.Cell(r.index, col)
}

// SetCell sets the text content of a cell in this row.
func (r *TableRow) SetCell(col int, text string) *TableCell {
	return r.table.SetCell(r.index, col, text)
}

// Justify sets the alignment for all cells in this row.
func (r *TableRow) Justify(alignment string) *TableRow {
	r.row.Justification(alignment)
	return r
}

// GetRaw returns the underlying go-docx table row for advanced usage.
func (r *TableRow) GetRaw() *docx.WTableRow {
	return r.row
}

// TableCell wraps a Word table cell.
type TableCell struct {
	cell  *docx.WTableCell
	table *Table
	row   int
	col   int
}

// SetText sets the text content of the cell.
// Clears any existing content and adds a new paragraph with the text.
func (c *TableCell) SetText(text string) *TableCell {
	// Get or create a paragraph in the cell
	var para *docx.Paragraph
	if len(c.cell.Paragraphs) > 0 {
		para = c.cell.Paragraphs[0]
		// Clear existing children
		para.Children = nil
	} else {
		para = c.cell.AddParagraph()
	}
	para.AddText(text)
	return c
}

// AddParagraph adds a new paragraph to the cell and returns it.
func (c *TableCell) AddParagraph(text string) *Paragraph {
	para := c.cell.AddParagraph()
	run := para.AddText(text)
	return &Paragraph{
		paragraph: para,
		lastRun:   run,
		doc:       c.table.doc,
	}
}

// Shade applies background shading to the cell.
// pattern: "clear", "solid", etc.
// color: pattern color (hex)
// fill: background fill color (hex)
func (c *TableCell) Shade(pattern, color, fill string) *TableCell {
	c.cell.Shade(pattern, color, fill)
	return c
}

// Background sets a solid background color for the cell.
func (c *TableCell) Background(hexColor string) *TableCell {
	return c.Shade("clear", "auto", hexColor)
}

// Bold applies bold formatting to the cell's text.
func (c *TableCell) Bold() *TableCell {
	if len(c.cell.Paragraphs) > 0 {
		para := c.cell.Paragraphs[0]
		for _, child := range para.Children {
			if run, ok := child.(*docx.Run); ok {
				run.Bold()
			}
		}
	}
	return c
}

// Center centers the text in the cell.
func (c *TableCell) Center() *TableCell {
	if len(c.cell.Paragraphs) > 0 {
		c.cell.Paragraphs[0].Justification("center")
	}
	return c
}

// GetRaw returns the underlying go-docx table cell for advanced usage.
func (c *TableCell) GetRaw() *docx.WTableCell {
	return c.cell
}

// VAlign sets the vertical alignment of the cell content.
//
//	cell.VAlign(docxtpl.VAlignCenter)
func (c *TableCell) VAlign(align VerticalAlignment) *TableCell {
	if c.cell.TableCellProperties == nil {
		c.cell.TableCellProperties = &docx.WTableCellProperties{}
	}
	c.cell.TableCellProperties.VAlign = &docx.WVerticalAlignment{
		Val: string(align),
	}
	return c
}

// VAlignTop aligns cell content to the top.
func (c *TableCell) VAlignTop() *TableCell {
	return c.VAlign(VAlignTop)
}

// VAlignCenter vertically centers cell content.
func (c *TableCell) VAlignCenter() *TableCell {
	return c.VAlign(VAlignCenter)
}

// VAlignBottom aligns cell content to the bottom.
func (c *TableCell) VAlignBottom() *TableCell {
	return c.VAlign(VAlignBottom)
}

// Width sets the cell width in twips (1/20 of a point).
// 1 inch = 1440 twips, 1 cm = 567 twips.
//
//	cell.Width(2880) // 2 inches
func (c *TableCell) Width(twips int) *TableCell {
	if c.cell.TableCellProperties == nil {
		c.cell.TableCellProperties = &docx.WTableCellProperties{}
	}
	c.cell.TableCellProperties.TableCellWidth = &docx.WTableCellWidth{
		W:    int64(twips),
		Type: "dxa",
	}
	return c
}

// WidthInches sets the cell width in inches.
//
//	cell.WidthInches(1.5) // 1.5 inches
func (c *TableCell) WidthInches(inches float64) *TableCell {
	return c.Width(int(inches * 1440))
}

// WidthCm sets the cell width in centimeters.
//
//	cell.WidthCm(3.5) // 3.5 cm
func (c *TableCell) WidthCm(cm float64) *TableCell {
	return c.Width(int(cm * 567))
}

// WidthPercent sets the cell width as a percentage of table width.
//
//	cell.WidthPercent(25) // 25% of table width
func (c *TableCell) WidthPercent(percent int) *TableCell {
	if c.cell.TableCellProperties == nil {
		c.cell.TableCellProperties = &docx.WTableCellProperties{}
	}
	c.cell.TableCellProperties.TableCellWidth = &docx.WTableCellWidth{
		W:    int64(percent * 50), // 50 = 1%
		Type: "pct",
	}
	return c
}

// MergeHorizontal merges this cell with the specified number of cells to the right.
// Uses GridSpan to span multiple columns.
//
//	table.Cell(0, 0).MergeHorizontal(2) // Merge with 2 cells to the right (3 total)
func (c *TableCell) MergeHorizontal(count int) *TableCell {
	if c.cell.TableCellProperties == nil {
		c.cell.TableCellProperties = &docx.WTableCellProperties{}
	}
	c.cell.TableCellProperties.GridSpan = &docx.WGridSpan{
		Val: count + 1, // +1 because GridSpan includes the current cell
	}
	return c
}

// MergeVerticalStart marks this cell as the start of a vertical merge.
// Use with MergeVerticalContinue on cells below.
//
//	table.Cell(0, 0).MergeVerticalStart() // Start of vertical merge
//	table.Cell(1, 0).MergeVerticalContinue() // Continue merge
func (c *TableCell) MergeVerticalStart() *TableCell {
	if c.cell.TableCellProperties == nil {
		c.cell.TableCellProperties = &docx.WTableCellProperties{}
	}
	c.cell.TableCellProperties.VMerge = &docx.WvMerge{
		Val: "restart",
	}
	return c
}

// MergeVerticalContinue marks this cell as a continuation of a vertical merge.
// The cell above must have MergeVerticalStart or MergeVerticalContinue.
func (c *TableCell) MergeVerticalContinue() *TableCell {
	if c.cell.TableCellProperties == nil {
		c.cell.TableCellProperties = &docx.WTableCellProperties{}
	}
	c.cell.TableCellProperties.VMerge = &docx.WvMerge{
		Val: "", // Empty Val means continue
	}
	return c
}

// Borders sets all borders for this cell.
// color is a hex color code without # (e.g., "000000" for black).
// width is the border width in eighths of a point.
//
//	cell.Borders("000000", 4) // Black border, 0.5pt width
func (c *TableCell) Borders(color string, width int) *TableCell {
	if c.cell.TableCellProperties == nil {
		c.cell.TableCellProperties = &docx.WTableCellProperties{}
	}
	border := &docx.WTableBorder{
		Val:   "single",
		Size:  width,
		Color: color,
	}
	c.cell.TableCellProperties.TableBorders = &docx.WTableBorders{
		Top:    border,
		Left:   border,
		Bottom: border,
		Right:  border,
	}
	return c
}

// NoBorders removes all borders from this cell.
func (c *TableCell) NoBorders() *TableCell {
	if c.cell.TableCellProperties == nil {
		c.cell.TableCellProperties = &docx.WTableCellProperties{}
	}
	border := &docx.WTableBorder{
		Val: "nil",
	}
	c.cell.TableCellProperties.TableBorders = &docx.WTableBorders{
		Top:    border,
		Left:   border,
		Bottom: border,
		Right:  border,
	}
	return c
}
