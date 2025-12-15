package docxtpl

import (
	"encoding/csv"
	"encoding/json"
	"sort"
	"strings"

	"github.com/abdokhaire/go-docxgen/internal/docx"
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

// =============================================================================
// Table Operations - Row and Column Management
// =============================================================================

// AddRow adds a new empty row to the table.
// Returns the new TableRow for populating.
//
//	row := table.AddRow()
//	row.SetCell(0, "New cell")
func (t *Table) AddRow() *TableRow {
	newRow := &docx.WTableRow{
		TableCells: make([]*docx.WTableCell, t.cols),
	}

	for i := 0; i < t.cols; i++ {
		newRow.TableCells[i] = &docx.WTableCell{
			Paragraphs: []*docx.Paragraph{{}},
		}
	}

	t.table.TableRows = append(t.table.TableRows, newRow)
	t.rows++

	return &TableRow{
		row:   newRow,
		table: t,
		index: t.rows - 1,
	}
}

// AddRowWithData adds a new row with the given cell values.
//
//	table.AddRowWithData("Col1", "Col2", "Col3")
func (t *Table) AddRowWithData(values ...string) *TableRow {
	row := t.AddRow()
	for i, val := range values {
		if i < t.cols {
			row.SetCell(i, val)
		}
	}
	return row
}

// InsertRow inserts a new row at the specified index.
// Existing rows are shifted down.
//
//	table.InsertRow(1)
func (t *Table) InsertRow(index int) *TableRow {
	if index < 0 {
		index = 0
	}
	if index > t.rows {
		index = t.rows
	}

	newRow := &docx.WTableRow{
		TableCells: make([]*docx.WTableCell, t.cols),
	}

	for i := 0; i < t.cols; i++ {
		newRow.TableCells[i] = &docx.WTableCell{
			Paragraphs: []*docx.Paragraph{{}},
		}
	}

	t.table.TableRows = append(t.table.TableRows[:index],
		append([]*docx.WTableRow{newRow}, t.table.TableRows[index:]...)...)
	t.rows++

	return &TableRow{
		row:   newRow,
		table: t,
		index: index,
	}
}

// DeleteRow removes the row at the specified index.
//
//	table.DeleteRow(2)
func (t *Table) DeleteRow(index int) *Table {
	if index < 0 || index >= t.rows || index >= len(t.table.TableRows) {
		return t
	}

	t.table.TableRows = append(t.table.TableRows[:index], t.table.TableRows[index+1:]...)
	t.rows--
	return t
}

// AddColumn adds a new column to the table.
//
//	table.AddColumn()
func (t *Table) AddColumn() *Table {
	for _, row := range t.table.TableRows {
		newCell := &docx.WTableCell{
			Paragraphs: []*docx.Paragraph{{}},
		}
		row.TableCells = append(row.TableCells, newCell)
	}
	t.cols++
	return t
}

// InsertColumn inserts a new column at the specified index.
//
//	table.InsertColumn(1)
func (t *Table) InsertColumn(index int) *Table {
	if index < 0 {
		index = 0
	}
	if index > t.cols {
		index = t.cols
	}

	for _, row := range t.table.TableRows {
		newCell := &docx.WTableCell{
			Paragraphs: []*docx.Paragraph{{}},
		}
		row.TableCells = append(row.TableCells[:index],
			append([]*docx.WTableCell{newCell}, row.TableCells[index:]...)...)
	}
	t.cols++
	return t
}

// DeleteColumn removes the column at the specified index.
//
//	table.DeleteColumn(2)
func (t *Table) DeleteColumn(index int) *Table {
	if index < 0 || index >= t.cols {
		return t
	}

	for _, row := range t.table.TableRows {
		if index < len(row.TableCells) {
			row.TableCells = append(row.TableCells[:index], row.TableCells[index+1:]...)
		}
	}
	t.cols--
	return t
}

// SortByColumn sorts the table rows by the specified column.
// Set ascending to false for descending order.
// skipHeader skips the first row (treats it as a header).
//
//	table.SortByColumn(0, true, true)
func (t *Table) SortByColumn(col int, ascending, skipHeader bool) *Table {
	if col < 0 || col >= t.cols || len(t.table.TableRows) < 2 {
		return t
	}

	startIdx := 0
	if skipHeader {
		startIdx = 1
	}

	if startIdx >= len(t.table.TableRows) {
		return t
	}

	rowsToSort := t.table.TableRows[startIdx:]

	sort.SliceStable(rowsToSort, func(i, j int) bool {
		textI := getCellTextFromRow(rowsToSort[i], col)
		textJ := getCellTextFromRow(rowsToSort[j], col)
		if ascending {
			return textI < textJ
		}
		return textI > textJ
	})

	return t
}

// =============================================================================
// Table Export Functions
// =============================================================================

// ToJSON returns the table data as JSON.
// If headers is true, the first row is used as keys.
//
//	jsonStr, err := table.ToJSON(true)
func (t *Table) ToJSON(headers bool) (string, error) {
	if len(t.table.TableRows) == 0 {
		return "[]", nil
	}

	var result interface{}

	if headers && len(t.table.TableRows) > 1 {
		headerRow := t.table.TableRows[0]
		var headerNames []string
		for _, cell := range headerRow.TableCells {
			headerNames = append(headerNames, getCellTextFromCell(cell))
		}

		var data []map[string]string
		for i := 1; i < len(t.table.TableRows); i++ {
			row := t.table.TableRows[i]
			rowData := make(map[string]string)
			for j, cell := range row.TableCells {
				key := ""
				if j < len(headerNames) {
					key = headerNames[j]
				} else {
					key = string(rune('A' + j))
				}
				rowData[key] = getCellTextFromCell(cell)
			}
			data = append(data, rowData)
		}
		result = data
	} else {
		var data [][]string
		for _, row := range t.table.TableRows {
			var rowData []string
			for _, cell := range row.TableCells {
				rowData = append(rowData, getCellTextFromCell(cell))
			}
			data = append(data, rowData)
		}
		result = data
	}

	bytes, err := json.MarshalIndent(result, "", "  ")
	if err != nil {
		return "", err
	}
	return string(bytes), nil
}

// ToCSV returns the table data as CSV.
//
//	csvStr := table.ToCSV()
func (t *Table) ToCSV() string {
	var sb strings.Builder
	writer := csv.NewWriter(&sb)

	for _, row := range t.table.TableRows {
		var record []string
		for _, cell := range row.TableCells {
			record = append(record, getCellTextFromCell(cell))
		}
		writer.Write(record)
	}
	writer.Flush()

	return sb.String()
}

// ToSlice returns the table data as a 2D string slice.
//
//	data := table.ToSlice()
func (t *Table) ToSlice() [][]string {
	var data [][]string
	for _, row := range t.table.TableRows {
		var rowData []string
		for _, cell := range row.TableCells {
			rowData = append(rowData, getCellTextFromCell(cell))
		}
		data = append(data, rowData)
	}
	return data
}

// =============================================================================
// Table Creation from Data Sources
// =============================================================================

// AddTableFromJSON creates a table from JSON data.
// Accepts either [][]string or []map[string]string format.
//
//	doc.AddTableFromJSON(`[["A","B"],["1","2"]]`)
//	doc.AddTableFromJSON(`[{"Name":"John","Age":"30"}]`)
func (d *DocxTmpl) AddTableFromJSON(jsonStr string, headers ...string) (*Table, error) {
	var arrData [][]string
	if err := json.Unmarshal([]byte(jsonStr), &arrData); err == nil {
		if len(arrData) == 0 {
			return d.AddTable(0, 0), nil
		}
		cols := len(arrData[0])
		table := d.AddTable(len(arrData), cols)
		for i, rowData := range arrData {
			for j, val := range rowData {
				table.SetCell(i, j, val)
			}
		}
		return table, nil
	}

	var objData []map[string]string
	if err := json.Unmarshal([]byte(jsonStr), &objData); err != nil {
		return nil, err
	}

	if len(objData) == 0 {
		return d.AddTable(0, 0), nil
	}

	var headerList []string
	if len(headers) > 0 {
		headerList = headers
	} else {
		for key := range objData[0] {
			headerList = append(headerList, key)
		}
		sort.Strings(headerList)
	}

	table := d.AddTable(len(objData)+1, len(headerList))

	for i, h := range headerList {
		table.SetCell(0, i, h).Bold()
	}

	for i, obj := range objData {
		for j, h := range headerList {
			table.SetCell(i+1, j, obj[h])
		}
	}

	return table, nil
}

// AddTableFromCSV creates a table from CSV data.
//
//	doc.AddTableFromCSV("Name,Age\nJohn,30\nJane,25")
func (d *DocxTmpl) AddTableFromCSV(csvStr string) (*Table, error) {
	reader := csv.NewReader(strings.NewReader(csvStr))
	records, err := reader.ReadAll()
	if err != nil {
		return nil, err
	}

	if len(records) == 0 {
		return d.AddTable(0, 0), nil
	}

	cols := len(records[0])
	table := d.AddTable(len(records), cols)

	for i, record := range records {
		for j, val := range record {
			table.SetCell(i, j, val)
		}
	}

	return table, nil
}

// AddTableFromSlice creates a table from a 2D string slice.
//
//	doc.AddTableFromSlice([][]string{{"A","B"},{"1","2"}})
func (d *DocxTmpl) AddTableFromSlice(data [][]string) *Table {
	if len(data) == 0 {
		return d.AddTable(0, 0)
	}

	cols := len(data[0])
	table := d.AddTable(len(data), cols)

	for i, rowData := range data {
		for j, val := range rowData {
			table.SetCell(i, j, val)
		}
	}

	return table
}

// AddTableWithHeaders creates a table with a header row.
// The first row is automatically bolded and centered.
//
//	doc.AddTableWithHeaders([]string{"Name", "Age", "City"}, [][]string{
//	    {"John", "30", "NYC"},
//	    {"Jane", "25", "LA"},
//	})
func (d *DocxTmpl) AddTableWithHeaders(headers []string, data [][]string) *Table {
	table := d.AddTable(len(data)+1, len(headers))

	for i, h := range headers {
		table.SetCell(0, i, h).Bold().Center()
	}

	for i, rowData := range data {
		for j, val := range rowData {
			if j < len(headers) {
				table.SetCell(i+1, j, val)
			}
		}
	}

	return table
}

// =============================================================================
// Table Helper Functions
// =============================================================================

func getCellTextFromRow(row *docx.WTableRow, col int) string {
	if col >= len(row.TableCells) {
		return ""
	}
	return getCellTextFromCell(row.TableCells[col])
}

func getCellTextFromCell(cell *docx.WTableCell) string {
	var texts []string
	for _, p := range cell.Paragraphs {
		texts = append(texts, p.String())
	}
	return strings.Join(texts, " ")
}

// SetColumnWidth sets the width for a specific column in twips.
//
//	table.SetColumnWidth(0, 2880)
func (t *Table) SetColumnWidth(col int, twips int) *Table {
	if col < 0 || col >= t.cols {
		return t
	}

	for _, row := range t.table.TableRows {
		if col < len(row.TableCells) {
			cell := row.TableCells[col]
			if cell.TableCellProperties == nil {
				cell.TableCellProperties = &docx.WTableCellProperties{}
			}
			cell.TableCellProperties.TableCellWidth = &docx.WTableCellWidth{
				W:    int64(twips),
				Type: "dxa",
			}
		}
	}
	return t
}

// SetAllColumnWidths sets widths for all columns.
//
//	table.SetAllColumnWidths([]int{2880, 1440, 1440})
func (t *Table) SetAllColumnWidths(widths []int) *Table {
	for i, w := range widths {
		if i < t.cols {
			t.SetColumnWidth(i, w)
		}
	}
	return t
}

// SetRowHeight sets the height for a specific row in twips.
//
//	table.SetRowHeight(0, 720)
func (t *Table) SetRowHeight(rowIdx int, twips int) *Table {
	if rowIdx < 0 || rowIdx >= len(t.table.TableRows) {
		return t
	}

	row := t.table.TableRows[rowIdx]
	if row.TableRowProperties == nil {
		row.TableRowProperties = &docx.WTableRowProperties{}
	}
	row.TableRowProperties.TableRowHeight = &docx.WTableRowHeight{
		Val: int64(twips),
	}
	return t
}

// ClearRow clears all cell content in a row.
//
//	table.ClearRow(0)
func (t *Table) ClearRow(rowIdx int) *Table {
	if rowIdx < 0 || rowIdx >= len(t.table.TableRows) {
		return t
	}

	row := t.table.TableRows[rowIdx]
	for _, cell := range row.TableCells {
		cell.Paragraphs = []*docx.Paragraph{{}}
	}
	return t
}

// ClearColumn clears all cell content in a column.
//
//	table.ClearColumn(0)
func (t *Table) ClearColumn(col int) *Table {
	if col < 0 || col >= t.cols {
		return t
	}

	for _, row := range t.table.TableRows {
		if col < len(row.TableCells) {
			row.TableCells[col].Paragraphs = []*docx.Paragraph{{}}
		}
	}
	return t
}

// FillColumn fills all cells in a column with the same value.
//
//	table.FillColumn(0, "N/A")
func (t *Table) FillColumn(col int, value string) *Table {
	if col < 0 || col >= t.cols {
		return t
	}

	for i := 0; i < t.rows; i++ {
		t.SetCell(i, col, value)
	}
	return t
}

// FillRow fills all cells in a row with the same value.
//
//	table.FillRow(0, "Header")
func (t *Table) FillRow(rowIdx int, value string) *Table {
	if rowIdx < 0 || rowIdx >= t.rows {
		return t
	}

	for i := 0; i < t.cols; i++ {
		t.SetCell(rowIdx, i, value)
	}
	return t
}

// GetColumn returns all values in a column.
//
//	values := table.GetColumn(0)
func (t *Table) GetColumn(col int) []string {
	if col < 0 || col >= t.cols {
		return nil
	}

	var values []string
	for _, row := range t.table.TableRows {
		if col < len(row.TableCells) {
			values = append(values, getCellTextFromCell(row.TableCells[col]))
		}
	}
	return values
}

// GetRowData returns all values in a row.
//
//	values := table.GetRowData(0)
func (t *Table) GetRowData(rowIdx int) []string {
	if rowIdx < 0 || rowIdx >= len(t.table.TableRows) {
		return nil
	}

	var values []string
	for _, cell := range t.table.TableRows[rowIdx].TableCells {
		values = append(values, getCellTextFromCell(cell))
	}
	return values
}

// FindRow finds the first row where the specified column contains the value.
// Returns -1 if not found.
//
//	rowIdx := table.FindRow(0, "John")
func (t *Table) FindRow(col int, value string) int {
	if col < 0 || col >= t.cols {
		return -1
	}

	for i, row := range t.table.TableRows {
		if col < len(row.TableCells) {
			if getCellTextFromCell(row.TableCells[col]) == value {
				return i
			}
		}
	}
	return -1
}

// FindAllRows finds all rows where the specified column contains the value.
//
//	rows := table.FindAllRows(1, "Active")
func (t *Table) FindAllRows(col int, value string) []int {
	if col < 0 || col >= t.cols {
		return nil
	}

	var results []int
	for i, row := range t.table.TableRows {
		if col < len(row.TableCells) {
			if getCellTextFromCell(row.TableCells[col]) == value {
				results = append(results, i)
			}
		}
	}
	return results
}
