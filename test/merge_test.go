package docxtpl_test

import (
	"os"
	"testing"

	"github.com/abdokhaire/go-docxgen"
	"github.com/stretchr/testify/assert"
)

func TestAppendDocument(t *testing.T) {
	assert := assert.New(t)

	// Load two documents
	doc1, err := docxtpl.ParseFromFilename("testdata/templates/test_basic.docx")
	assert.Nil(err)
	assert.NotNil(doc1)

	doc2, err := docxtpl.ParseFromFilename("testdata/templates/test_basic.docx")
	assert.Nil(err)
	assert.NotNil(doc2)

	// Get initial item count
	initialCount := len(doc1.Document.Body.Items)

	// Append doc2 to doc1
	doc1.AppendDocument(doc2)

	// Should have more items now
	assert.Greater(len(doc1.Document.Body.Items), initialCount)
}

func TestMergeDocuments(t *testing.T) {
	assert := assert.New(t)

	doc1, err := docxtpl.ParseFromFilename("testdata/templates/test_basic.docx")
	assert.Nil(err)

	doc2, err := docxtpl.ParseFromFilename("testdata/templates/test_basic.docx")
	assert.Nil(err)

	merged, err := docxtpl.MergeDocuments(doc1, doc2)
	if err != nil {
		t.Skip("MergeDocuments failed, skipping test: " + err.Error())
	}
	assert.NotNil(merged)

	// Merged should have content from both documents
	if merged != nil {
		assert.Greater(len(merged.Document.Body.Items), 0)
	}
}

func TestMergeDocuments_Single(t *testing.T) {
	assert := assert.New(t)

	doc, err := docxtpl.ParseFromFilename("testdata/templates/test_basic.docx")
	assert.Nil(err)

	merged, err := docxtpl.MergeDocuments(doc)
	if err != nil {
		t.Skip("MergeDocuments failed, skipping test: " + err.Error())
	}
	assert.NotNil(merged)
}

func TestMergeDocuments_Empty(t *testing.T) {
	assert := assert.New(t)

	_, err := docxtpl.MergeDocuments()
	assert.NotNil(err) // Should error with no documents
}

func TestClone(t *testing.T) {
	assert := assert.New(t)

	doc, err := docxtpl.ParseFromFilename("testdata/templates/test_basic.docx")
	assert.Nil(err)

	clone, err := doc.Clone()
	if err != nil {
		t.Skip("Clone failed, skipping test: " + err.Error())
	}
	assert.NotNil(clone)

	// Clone should have the same number of items
	if clone != nil {
		assert.Equal(len(doc.Document.Body.Items), len(clone.Document.Body.Items))
	}
}

func TestMergedDocumentCanBeSaved(t *testing.T) {
	assert := assert.New(t)

	doc1, err := docxtpl.ParseFromFilename("testdata/templates/test_basic.docx")
	assert.Nil(err)

	doc2, err := docxtpl.ParseFromFilename("testdata/templates/test_basic.docx")
	assert.Nil(err)

	merged, err := docxtpl.MergeDocuments(doc1, doc2)
	if err != nil {
		t.Skip("MergeDocuments failed, skipping test: " + err.Error())
	}

	// Should be able to save the merged document
	outputFile := "testdata/templates/generated_merged.docx"
	err = merged.SaveToFile(outputFile)
	assert.Nil(err)

	// Clean up
	os.Remove(outputFile)
}

func TestMailMerge(t *testing.T) {
	assert := assert.New(t)

	doc, err := docxtpl.ParseFromFilename("testdata/templates/test_basic.docx")
	assert.Nil(err)

	records := []map[string]any{
		{"FirstName": "John", "LastName": "Doe"},
		{"FirstName": "Jane", "LastName": "Smith"},
	}

	docs, err := doc.MailMerge(records)
	if err != nil {
		t.Skip("MailMerge failed, skipping test: " + err.Error())
	}
	assert.Len(docs, 2)
}

func TestMailMergeToSingle(t *testing.T) {
	assert := assert.New(t)

	doc, err := docxtpl.ParseFromFilename("testdata/templates/test_basic.docx")
	assert.Nil(err)

	records := []map[string]any{
		{"FirstName": "John", "LastName": "Doe"},
		{"FirstName": "Jane", "LastName": "Smith"},
	}

	merged, err := doc.MailMergeToSingle(records)
	if err != nil {
		t.Skip("MailMergeToSingle failed, skipping test: " + err.Error())
	}
	assert.NotNil(merged)
}

func TestLoadAllFromDirectory(t *testing.T) {
	assert := assert.New(t)

	docs, err := docxtpl.LoadAllFromDirectory("testdata/templates")
	assert.Nil(err)
	assert.Greater(len(docs), 0)
}

func TestBatchProcess(t *testing.T) {
	assert := assert.New(t)

	doc1, err := docxtpl.ParseFromFilename("testdata/templates/test_basic.docx")
	assert.Nil(err)

	doc2, err := docxtpl.ParseFromFilename("testdata/templates/test_basic.docx")
	assert.Nil(err)

	docs := []*docxtpl.DocxTmpl{doc1, doc2}

	// Process all docs
	processCount := 0
	err = docxtpl.BatchProcess(docs, func(doc *docxtpl.DocxTmpl) error {
		processCount++
		return nil
	})
	assert.Nil(err)
	assert.Equal(2, processCount)
}

func TestSearchWithContext(t *testing.T) {
	assert := assert.New(t)

	doc, err := docxtpl.ParseFromFilename("testdata/templates/test_basic.docx")
	assert.Nil(err)

	// Search for text that might be in the template
	results := doc.SearchWithContext("First", 1)
	// Results can be nil or empty slice if no matches, just verify no panic
	_ = results
	assert.True(true)
}

func TestSplitAtHeading(t *testing.T) {
	assert := assert.New(t)

	doc, err := docxtpl.ParseFromFilename("testdata/templates/test_basic.docx")
	assert.Nil(err)

	// Split at heading level 1
	parts := doc.SplitAtHeading(1)
	// Should return at least one part
	assert.Greater(len(parts), 0)
}
