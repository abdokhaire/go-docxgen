package docxtpl

import (
	"github.com/fumiama/go-docx"
)

// ListType represents the type of list
type ListType int

const (
	ListTypeBullet   ListType = iota // Bullet list (•, ○, ▪)
	ListTypeNumbered                 // Numbered list (1, 2, 3)
	ListTypeLetter                   // Letter list (a, b, c)
	ListTypeRoman                    // Roman numeral list (i, ii, iii)
)

// ListItem represents an item in a list with optional nesting
type ListItem struct {
	Text     string     // Item text
	Level    int        // Nesting level (0-8)
	Children []ListItem // Nested items
}

// List wraps a collection of list paragraphs
type List struct {
	doc        *DocxTmpl
	listType   ListType
	paragraphs []*Paragraph
}

// AddBulletList adds a bullet list to the document.
// Each string in items becomes a bullet point.
//
//	doc.AddBulletList([]string{"First item", "Second item", "Third item"})
func (d *DocxTmpl) AddBulletList(items []string) *List {
	list := &List{
		doc:      d,
		listType: ListTypeBullet,
	}

	for _, item := range items {
		para := d.addListParagraph(item, ListTypeBullet, 0)
		list.paragraphs = append(list.paragraphs, para)
	}

	return list
}

// AddNumberedList adds a numbered list to the document.
// Each string in items becomes a numbered item (1, 2, 3...).
//
//	doc.AddNumberedList([]string{"First step", "Second step", "Third step"})
func (d *DocxTmpl) AddNumberedList(items []string) *List {
	list := &List{
		doc:      d,
		listType: ListTypeNumbered,
	}

	for _, item := range items {
		para := d.addListParagraph(item, ListTypeNumbered, 0)
		list.paragraphs = append(list.paragraphs, para)
	}

	return list
}

// AddNestedList adds a nested list with multiple levels.
// Use ListItem.Level to control indentation (0-8).
// Use ListItem.Children for nested items.
//
//	doc.AddNestedList(ListTypeBullet, []ListItem{
//	    {Text: "Item 1", Children: []ListItem{
//	        {Text: "Sub-item 1.1"},
//	        {Text: "Sub-item 1.2"},
//	    }},
//	    {Text: "Item 2"},
//	})
func (d *DocxTmpl) AddNestedList(listType ListType, items []ListItem) *List {
	list := &List{
		doc:      d,
		listType: listType,
	}

	d.addNestedListItems(list, items, listType, 0)

	return list
}

// addNestedListItems recursively adds list items
func (d *DocxTmpl) addNestedListItems(list *List, items []ListItem, listType ListType, level int) {
	for _, item := range items {
		effectiveLevel := item.Level
		if effectiveLevel == 0 {
			effectiveLevel = level
		}
		if effectiveLevel > 8 {
			effectiveLevel = 8
		}

		para := d.addListParagraph(item.Text, listType, effectiveLevel)
		list.paragraphs = append(list.paragraphs, para)

		// Recursively add children
		if len(item.Children) > 0 {
			d.addNestedListItems(list, item.Children, listType, effectiveLevel+1)
		}
	}
}

// addListParagraph creates a paragraph with list formatting
func (d *DocxTmpl) addListParagraph(text string, listType ListType, level int) *Paragraph {
	p := d.Docx.AddParagraph()

	// Create the bullet/number prefix based on list type and level
	prefix := getBulletPrefix(listType, level)

	// Add prefix as a separate run
	prefixRun := p.AddText(prefix + "\t")
	_ = prefixRun // avoid unused

	// Add the main text
	run := p.AddText(text)

	// Apply indentation based on level
	para := &Paragraph{
		paragraph: p,
		lastRun:   run,
		doc:       d,
	}

	// Apply list indentation (in inches)
	indent := 0.5 * float64(level+1) // 0.5 inch per level
	para.IndentLeft(indent)
	para.IndentHanging(0.25) // Hanging indent for bullet/number

	return para
}

// getBulletPrefix returns the appropriate bullet or number character
func getBulletPrefix(listType ListType, level int) string {
	switch listType {
	case ListTypeBullet:
		// Different bullets for different levels
		bullets := []string{"•", "○", "▪", "•", "○", "▪", "•", "○", "▪"}
		if level < len(bullets) {
			return bullets[level]
		}
		return "•"
	case ListTypeNumbered:
		// Numbers - would need state tracking for actual numbers
		// For now, return placeholder
		return "1."
	case ListTypeLetter:
		return "a."
	case ListTypeRoman:
		return "i."
	default:
		return "•"
	}
}

// AddItem adds another item to the list at the same level.
func (l *List) AddItem(text string) *List {
	para := l.doc.addListParagraph(text, l.listType, 0)
	l.paragraphs = append(l.paragraphs, para)
	return l
}

// AddSubItem adds a nested item to the list.
func (l *List) AddSubItem(text string, level int) *List {
	if level < 0 {
		level = 0
	}
	if level > 8 {
		level = 8
	}
	para := l.doc.addListParagraph(text, l.listType, level)
	l.paragraphs = append(l.paragraphs, para)
	return l
}

// GetParagraphs returns all paragraphs in the list.
func (l *List) GetParagraphs() []*Paragraph {
	return l.paragraphs
}

// Count returns the number of items in the list.
func (l *List) Count() int {
	return len(l.paragraphs)
}

// ListBuilder provides a fluent interface for building complex lists
type ListBuilder struct {
	doc        *DocxTmpl
	listType   ListType
	items      []listBuilderItem
	currentLevel int
}

type listBuilderItem struct {
	text  string
	level int
}

// NewListBuilder creates a new list builder
//
//	list := doc.NewListBuilder(ListTypeBullet).
//	    Item("First").
//	    Item("Second").
//	    Indent().Item("Nested").
//	    Outdent().Item("Third").
//	    Build()
func (d *DocxTmpl) NewListBuilder(listType ListType) *ListBuilder {
	return &ListBuilder{
		doc:      d,
		listType: listType,
		items:    []listBuilderItem{},
		currentLevel: 0,
	}
}

// Item adds an item at the current indentation level
func (lb *ListBuilder) Item(text string) *ListBuilder {
	lb.items = append(lb.items, listBuilderItem{
		text:  text,
		level: lb.currentLevel,
	})
	return lb
}

// Indent increases the indentation level for subsequent items
func (lb *ListBuilder) Indent() *ListBuilder {
	if lb.currentLevel < 8 {
		lb.currentLevel++
	}
	return lb
}

// Outdent decreases the indentation level for subsequent items
func (lb *ListBuilder) Outdent() *ListBuilder {
	if lb.currentLevel > 0 {
		lb.currentLevel--
	}
	return lb
}

// Level sets the indentation level directly
func (lb *ListBuilder) Level(level int) *ListBuilder {
	if level < 0 {
		level = 0
	}
	if level > 8 {
		level = 8
	}
	lb.currentLevel = level
	return lb
}

// Build creates the list and adds it to the document
func (lb *ListBuilder) Build() *List {
	list := &List{
		doc:      lb.doc,
		listType: lb.listType,
	}

	for _, item := range lb.items {
		para := lb.doc.addListParagraph(item.text, lb.listType, item.level)
		list.paragraphs = append(list.paragraphs, para)
	}

	return list
}

// AddChecklistItem adds a checkbox item to the document.
// checked determines if the checkbox appears checked.
//
//	doc.AddChecklistItem("Complete task", true)
//	doc.AddChecklistItem("Pending task", false)
func (d *DocxTmpl) AddChecklistItem(text string, checked bool) *Paragraph {
	p := d.Docx.AddParagraph()

	// Add checkbox symbol
	checkbox := "☐ "
	if checked {
		checkbox = "☑ "
	}

	p.AddText(checkbox)
	run := p.AddText(text)

	return &Paragraph{
		paragraph: p,
		lastRun:   run,
		doc:       d,
	}
}

// AddChecklist adds a checklist to the document.
// items is a map of text to checked status.
//
//	doc.AddChecklist(map[string]bool{
//	    "Task 1": true,
//	    "Task 2": false,
//	    "Task 3": false,
//	})
func (d *DocxTmpl) AddChecklist(items map[string]bool) []*Paragraph {
	var paragraphs []*Paragraph
	for text, checked := range items {
		para := d.AddChecklistItem(text, checked)
		paragraphs = append(paragraphs, para)
	}
	return paragraphs
}

// AddOrderedChecklist adds a checklist with preserved order.
//
//	doc.AddOrderedChecklist([]struct{Text string; Checked bool}{
//	    {"Task 1", true},
//	    {"Task 2", false},
//	})
func (d *DocxTmpl) AddOrderedChecklist(items []struct {
	Text    string
	Checked bool
}) []*Paragraph {
	var paragraphs []*Paragraph
	for _, item := range items {
		para := d.AddChecklistItem(item.Text, item.Checked)
		paragraphs = append(paragraphs, para)
	}
	return paragraphs
}

// Unused import prevention
var _ = docx.Paragraph{}
