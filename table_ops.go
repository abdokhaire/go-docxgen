package docxtpl

import (
	"encoding/csv"
	"encoding/json"
	"sort"
	"strings"

	"github.com/fumiama/go-docx"
)

// AddRow adds a new empty row to the table.
// Returns the new TableRow for populating.
//
//	row := table.AddRow()
//	row.SetCell(0, "New cell")
func (t *Table) AddRow() *TableRow {
	// Create a new row with the same number of cells as existing rows
	newRow := &docx.WTableRow{
		TableCells: make([]*docx.WTableCell, t.cols),
	}

	// Initialize cells
	for i := 0; i < t.cols; i++ {
		newRow.TableCells[i] = &docx.WTableCell{
			Paragraphs: []*docx.Paragraph{{}},
		}
	}

	// Append to table
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
//	table.InsertRow(1) // Insert after first row
func (t *Table) InsertRow(index int) *TableRow {
	if index < 0 {
		index = 0
	}
	if index > t.rows {
		index = t.rows
	}

	// Create a new row
	newRow := &docx.WTableRow{
		TableCells: make([]*docx.WTableCell, t.cols),
	}

	for i := 0; i < t.cols; i++ {
		newRow.TableCells[i] = &docx.WTableCell{
			Paragraphs: []*docx.Paragraph{{}},
		}
	}

	// Insert at position
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
//	table.DeleteRow(2) // Delete third row
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
//	table.InsertColumn(1) // Insert after first column
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
//	table.SortByColumn(0, true, true) // Sort by first column, ascending, skip header
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

	// Get slice to sort
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
		// Use first row as headers
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
		// Return as array of arrays
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

// AddTableFromJSON creates a table from JSON data.
// Accepts either [][]string or []map[string]string format.
//
//	doc.AddTableFromJSON(`[["A","B"],["1","2"]]`)
//	doc.AddTableFromJSON(`[{"Name":"John","Age":"30"}]`)
func (d *DocxTmpl) AddTableFromJSON(jsonStr string, headers ...string) (*Table, error) {
	// Try array of arrays first
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

	// Try array of objects
	var objData []map[string]string
	if err := json.Unmarshal([]byte(jsonStr), &objData); err != nil {
		return nil, err
	}

	if len(objData) == 0 {
		return d.AddTable(0, 0), nil
	}

	// Determine headers
	var headerList []string
	if len(headers) > 0 {
		headerList = headers
	} else {
		// Get keys from first object
		for key := range objData[0] {
			headerList = append(headerList, key)
		}
		sort.Strings(headerList) // Sort for consistent order
	}

	// Create table with header row
	table := d.AddTable(len(objData)+1, len(headerList))

	// Set headers
	for i, h := range headerList {
		table.SetCell(0, i, h).Bold()
	}

	// Set data
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

	// Set headers
	for i, h := range headers {
		table.SetCell(0, i, h).Bold().Center()
	}

	// Set data
	for i, rowData := range data {
		for j, val := range rowData {
			if j < len(headers) {
				table.SetCell(i+1, j, val)
			}
		}
	}

	return table
}

// Helper functions

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
//	table.SetColumnWidth(0, 2880) // Set first column to 2 inches
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
//	table.SetRowHeight(0, 720) // Set first row to 0.5 inch
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
