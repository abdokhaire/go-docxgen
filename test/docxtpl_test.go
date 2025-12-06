package docxtpl_test

import (
	"encoding/json"
	"io"
	"net/http"
	"os"
	"testing"

	"github.com/abdokhaire/go-docxgen"
)

func TestParseFromURL(t *testing.T) {
	url := os.Getenv("DOCX_TEMPLATE_URL")
	if url == "" {
		t.Skip("Skipping TestParseFromURL: DOCX_TEMPLATE_URL environment variable not set")
	}

	t.Log("Started")

	// Download the template file
	resp, err := http.Get(url)
	if err != nil {
		t.Fatalf("Failed to download template file: %v", err)
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		t.Fatalf("Failed to download template file: status code %d", resp.StatusCode)
	}

	templateFile, err := os.CreateTemp("", "template*.docx")
	if err != nil {
		t.Fatalf("Failed to create temporary template file: %v", err)
	}
	defer os.Remove(templateFile.Name())

	fileBytes, err := io.ReadAll(resp.Body)
	if err != nil {
		t.Fatalf("Failed to read template file body: %v", err)
	}

	_, err = templateFile.Write(fileBytes)
	if err != nil {
		t.Fatalf("Failed to write to temporary template file: %v", err)
	}

	// Read the JSON data
	jsonData, err := os.ReadFile("testdata/test_data.json")
	if err != nil {
		t.Fatalf("Failed to read testdata/test_data.json: %v", err)
	}

	var data map[string]interface{}
	err = json.Unmarshal(jsonData, &data)
	if err != nil {
		t.Fatalf("Failed to unmarshal JSON data: %v", err)
	}

	// Create a new DocxTpl instance
	tpl, err := docxtpl.ParseFromFilename(templateFile.Name())
	if err != nil {
		t.Fatalf("ParseFromFilename failed: %v", err)
	}

	// Render the template with the JSON data
	err = tpl.Render(data)
	if err != nil {
		t.Fatalf("Render failed: %v", err)
	}
	t.Log("Render: Done")
	// Save the output file
	outputFileName := "test_output_from_url.docx"
	outputFile, err := os.Create(outputFileName)
	if err != nil {
		t.Fatalf("Failed to create output file: %v", err)
	}
	defer outputFile.Close()

	if _, err := os.Stat(outputFileName); os.IsExist(err) {
		os.Remove(outputFileName)
	}

	err = tpl.Save(outputFile)
	if err != nil {
		t.Fatalf("Save failed: %v", err)
	}
	t.Log("Save: Done")
	// Check if the output file exists
	if _, err := os.Stat(outputFileName); os.IsNotExist(err) {
		t.Errorf("Output file was not created: %s", outputFileName)
	}
}
