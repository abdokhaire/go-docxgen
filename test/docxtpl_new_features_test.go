package docxtpl_test

import (
	"os"
	"testing"

	"github.com/abdokhaire/go-docxgen"
	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

func TestParseFromBytes(t *testing.T) {
	t.Run("Should parse document from bytes", func(t *testing.T) {
		// Read the file into bytes
		data, err := os.ReadFile("testdata/templates/test_basic.docx")
		require.NoError(t, err)

		// Parse from bytes
		doc, err := docxtpl.ParseFromBytes(data)
		require.NoError(t, err)
		assert.NotNil(t, doc)

		// Verify we can render
		renderData := map[string]any{
			"ProjectNumber": "B-00001",
			"Client":        "Test Client",
			"Status":        "Active",
		}
		err = doc.Render(renderData)
		assert.NoError(t, err)
	})
}

func TestSaveToFile(t *testing.T) {
	t.Run("Should save document to file", func(t *testing.T) {
		// Parse a document
		doc, err := docxtpl.ParseFromFilename("testdata/templates/test_basic.docx")
		require.NoError(t, err)

		// Render with data
		renderData := map[string]any{
			"ProjectNumber": "B-00001",
			"Client":        "Test Client",
			"Status":        "Active",
		}
		err = doc.Render(renderData)
		require.NoError(t, err)

		// Save to file
		outputPath := "testdata/templates/generated_test_save_to_file.docx"
		err = doc.SaveToFile(outputPath)
		assert.NoError(t, err)

		// Verify file exists
		_, err = os.Stat(outputPath)
		assert.NoError(t, err)

		// Clean up
		os.Remove(outputPath)
	})
}

func TestGetPlaceholders(t *testing.T) {
	t.Run("Should return all placeholders from document", func(t *testing.T) {
		doc, err := docxtpl.ParseFromFilename("testdata/templates/test_basic.docx")
		require.NoError(t, err)

		placeholders, err := doc.GetPlaceholders()
		require.NoError(t, err)

		// The test_basic.docx should have some placeholders
		assert.NotEmpty(t, placeholders)

		// Check that common placeholders are found
		found := make(map[string]bool)
		for _, p := range placeholders {
			found[p] = true
		}

		// These are expected placeholders in test_basic.docx
		assert.True(t, found["{{.ProjectNumber}}"] || found["{{.Client}}"] || found["{{.Status}}"],
			"Expected to find at least one of the known placeholders")
	})

	t.Run("Should return unique placeholders only", func(t *testing.T) {
		doc, err := docxtpl.ParseFromFilename("testdata/templates/test_with_tables.docx")
		require.NoError(t, err)

		placeholders, err := doc.GetPlaceholders()
		require.NoError(t, err)

		// Check for duplicates
		seen := make(map[string]bool)
		for _, p := range placeholders {
			assert.False(t, seen[p], "Found duplicate placeholder: %s", p)
			seen[p] = true
		}
	})
}

func TestNewlineConversion(t *testing.T) {
	t.Run("Should convert newlines to line breaks in rendered document", func(t *testing.T) {
		doc, err := docxtpl.ParseFromFilename("testdata/templates/test_basic.docx")
		require.NoError(t, err)

		// Render with data containing newlines
		renderData := map[string]any{
			"ProjectNumber": "B-00001",
			"Client":        "Line 1\nLine 2\nLine 3",
			"Status":        "Active",
		}
		err = doc.Render(renderData)
		require.NoError(t, err)

		// Save and verify no error
		outputPath := "testdata/templates/generated_test_newlines.docx"
		err = doc.SaveToFile(outputPath)
		assert.NoError(t, err)

		// Clean up
		os.Remove(outputPath)
	})
}
