// Example demonstrates basic usage of go-docxgen
package main

import (
	"fmt"
	"log"
	"time"

	docxtpl "github.com/abdokhaire/go-docxgen"
)

// Employee represents an employee record
type Employee struct {
	Name       string
	Department string
	Salary     float64
	StartDate  time.Time
}

// Report represents the data for a report
type Report struct {
	Title       string
	GeneratedAt time.Time
	Author      string
	Employees   []Employee
	TotalBudget float64
	IsFinalized bool
}

func main() {
	// Example 1: Template-based document generation
	templateExample()

	// Example 2: Programmatic document creation
	programmaticExample()
}

// templateExample shows how to use templates for document generation
func templateExample() {
	// Parse the template
	doc, err := docxtpl.ParseFromFilename("template.docx")
	if err != nil {
		log.Printf("Template example skipped (no template.docx): %v", err)
		return
	}

	// Prepare the data
	report := Report{
		Title:       "Q4 2025 Staff Report",
		GeneratedAt: time.Now(),
		Author:      "HR Department",
		Employees: []Employee{
			{Name: "Alice Johnson", Department: "Engineering", Salary: 95000, StartDate: time.Date(2020, 3, 15, 0, 0, 0, 0, time.UTC)},
			{Name: "Bob Smith", Department: "Marketing", Salary: 75000, StartDate: time.Date(2021, 7, 1, 0, 0, 0, 0, time.UTC)},
			{Name: "Carol White", Department: "Engineering", Salary: 105000, StartDate: time.Date(2019, 1, 10, 0, 0, 0, 0, time.UTC)},
		},
		TotalBudget: 275000,
		IsFinalized: true,
	}

	// Render the template with data
	if err := doc.Render(report); err != nil {
		log.Fatalf("Failed to render: %v", err)
	}

	// Save the output
	if err := doc.SaveToFile("output_template.docx"); err != nil {
		log.Fatalf("Failed to save: %v", err)
	}

	fmt.Println("Template document generated: output_template.docx")
}

// programmaticExample shows how to create documents from scratch
func programmaticExample() {
	// Create a new document
	doc := docxtpl.New()

	// Set document properties
	doc.SetTitle("Programmatic Report")
	doc.SetAuthor("Go Application")
	doc.SetSubject("Example Document")
	doc.SetKeywords("go, docx, example")

	// Add title
	doc.AddHeading("Welcome to go-docxgen", 0)
	doc.AddEmptyParagraph()

	// Introduction section
	doc.AddHeading("Introduction", 1)
	doc.AddParagraph("This document was created programmatically using the go-docxgen library.")
	doc.AddEmptyParagraph()

	// Formatted text examples
	doc.AddHeading("Text Formatting", 1)

	para := doc.AddParagraph("This text is ")
	para.AddText("bold").Bold()
	para.AddText(", ")
	para.AddText("italic").Italic()
	para.AddText(", ")
	para.AddText("underlined").Underline()
	para.AddText(", and ")
	para.AddText("red").Color("FF0000")
	para.AddText(".")

	doc.AddParagraph("This paragraph is centered.").Center()
	doc.AddParagraph("This paragraph is right-aligned.").Right()
	doc.AddParagraph("This text has a yellow highlight.").Highlight("yellow")

	doc.AddEmptyParagraph()

	// Table section
	doc.AddHeading("Data Table", 1)

	table := doc.AddTable(4, 3)

	// Header row
	table.SetCell(0, 0, "Name").Bold().Background("4472C4").Center()
	table.SetCell(0, 1, "Department").Bold().Background("4472C4").Center()
	table.SetCell(0, 2, "Status").Bold().Background("4472C4").Center()

	// Data rows
	table.SetCell(1, 0, "Alice Johnson")
	table.SetCell(1, 1, "Engineering")
	table.SetCell(1, 2, "Active").Background("C6EFCE")

	table.SetCell(2, 0, "Bob Smith")
	table.SetCell(2, 1, "Marketing")
	table.SetCell(2, 2, "Active").Background("C6EFCE")

	table.SetCell(3, 0, "Carol White")
	table.SetCell(3, 1, "Finance")
	table.SetCell(3, 2, "On Leave").Background("FFEB9C")

	doc.AddEmptyParagraph()

	// Conclusion
	doc.AddHeading("Conclusion", 1)
	doc.AddParagraph("This demonstrates the programmatic document creation capabilities of go-docxgen.")

	// Save the document
	if err := doc.SaveToFile("output_programmatic.docx"); err != nil {
		log.Fatalf("Failed to save: %v", err)
	}

	fmt.Println("Programmatic document generated: output_programmatic.docx")
}
