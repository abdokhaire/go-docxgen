package main

import (
	"fmt"
	"os"

	"github.com/Masterminds/sprig/v3"
	docxtpl "github.com/abdokhaire/go-docxgen"
)

func main() {
	if len(os.Args) < 2 {
		fmt.Println("Usage: go run main.go <template.docx> [output.docx]")
		fmt.Println("\nExample:")
		fmt.Println("  go run main.go template.docx output.docx")
		os.Exit(1)
	}

	templatePath := os.Args[1]
	outputPath := "output.docx"
	if len(os.Args) >= 3 {
		outputPath = os.Args[2]
	}

	// Parse the template
	fmt.Printf("Loading template: %s\n", templatePath)
	doc, err := docxtpl.ParseFromFilename(templatePath)
	if err != nil {
		fmt.Printf("Error parsing template: %v\n", err)
		os.Exit(1)
	}

	// Register Sprig functions
	fmt.Println("Registering Sprig functions...")
	doc.RegisterFuncMap(sprig.FuncMap())

	// Get placeholders
	placeholders, err := doc.GetPlaceholders()
	if err != nil {
		fmt.Printf("Error getting placeholders: %v\n", err)
	} else {
		fmt.Printf("\nFound %d placeholders:\n", len(placeholders))
		for i, p := range placeholders {
			fmt.Printf("  %d. %s\n", i+1, p)
		}
	}

	// Validate template syntax
	fmt.Println("\nValidating template...")
	result := doc.Validate()
	if result.HasErrors() {
		fmt.Println("Validation errors found:")
		for _, e := range result.Errors {
			fmt.Printf("  - %s\n", e.Error())
		}
	} else {
		fmt.Println("Template syntax is valid!")
	}

	// Generate sample data
	fmt.Println("\nGenerating sample data...")
	sampleData := generateTestData()

	fmt.Println("\nTest data:")
	fmt.Printf("  %+v\n", sampleData)

	// Render
	fmt.Println("\nRendering template...")
	err = doc.Render(sampleData)
	if err != nil {
		fmt.Printf("Error rendering: %v\n", err)
		os.Exit(1)
	}

	// Save
	fmt.Printf("\nSaving to: %s\n", outputPath)
	err = doc.SaveToFile(outputPath)
	if err != nil {
		fmt.Printf("Error saving: %v\n", err)
		os.Exit(1)
	}

	fmt.Println("\nSuccess!")
}

// generateTestData creates sample data matching a typical template structure
func generateTestData() map[string]any {
	return map[string]any{
		"record": map[string]any{
			"name":             "Test Project Alpha",
			"email":            "test@example.com",
			"status":           "active",
			"code":             "PRJ-001",
			"is_active":        true,
			"is_deleted":       false,
			"total_units":      5,
			"total_unit_price": 1250.50,
			"units": []map[string]any{
				{
					"name":             "Unit A",
					"unit_number":      "A-101",
					"unit_price":       250.00,
					"additional_price": 2,
				},
				{
					"name":             "Unit B",
					"unit_number":      "B-202",
					"unit_price":       300.00,
					"additional_price": 3,
				},
				{
					"name":             "Unit C",
					"unit_number":      "C-303",
					"unit_price":       200.00,
					"additional_price": 1,
				},
			},
		},
	}
}
