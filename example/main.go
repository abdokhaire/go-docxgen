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
	// Parse the template
	doc, err := docxtpl.ParseFromFilename("template.docx")
	if err != nil {
		log.Fatalf("Failed to parse template: %v", err)
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
	if err := doc.SaveToFile("output.docx"); err != nil {
		log.Fatalf("Failed to save: %v", err)
	}

	fmt.Println("Document generated successfully: output.docx")
}
