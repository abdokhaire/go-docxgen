package docxtpl_test

import (
	"errors"
	"testing"

	"github.com/abdokhaire/go-docxgen"
	"github.com/stretchr/testify/assert"
)

func TestTemplateError_Error(t *testing.T) {
	assert := assert.New(t)

	err := docxtpl.NewTemplateError(docxtpl.ErrCodeSyntaxError, "invalid syntax")
	assert.Equal("invalid syntax", err.Error())

	err.Location = "header"
	assert.Contains(err.Error(), "[header]")

	err.Placeholder = "{{.Name}}"
	assert.Contains(err.Error(), "(at: {{.Name}})")

	err.LineNumber = 10
	assert.Contains(err.Error(), "(line 10)")
}

func TestTemplateError_String(t *testing.T) {
	assert := assert.New(t)

	err := &docxtpl.TemplateError{
		Code:        docxtpl.ErrCodeUnclosedTag,
		Message:     "unclosed tag",
		Location:    "document body",
		Placeholder: "{{.Name",
		LineNumber:  5,
		Cause:       errors.New("underlying error"),
		Suggestions: []string{"Check your tags", "Close all braces"},
	}

	str := err.String()
	assert.Contains(str, "Template Error")
	assert.Contains(str, "UNCLOSED_TAG")
	assert.Contains(str, "unclosed tag")
	assert.Contains(str, "document body")
	assert.Contains(str, "{{.Name")
	assert.Contains(str, "5")
	assert.Contains(str, "underlying error")
	assert.Contains(str, "Suggestions:")
	assert.Contains(str, "Check your tags")
}

func TestTemplateError_Unwrap(t *testing.T) {
	assert := assert.New(t)

	cause := errors.New("root cause")
	err := docxtpl.NewTemplateError(docxtpl.ErrCodeSyntaxError, "test").WithCause(cause)

	assert.Equal(cause, err.Unwrap())
}

func TestTemplateError_FluentAPI(t *testing.T) {
	assert := assert.New(t)

	err := docxtpl.NewTemplateError(docxtpl.ErrCodeSyntaxError, "test").
		WithLocation("header").
		WithPlaceholder("{{.Field}}").
		WithSuggestions("suggestion1", "suggestion2")

	assert.Equal("header", err.Location)
	assert.Equal("{{.Field}}", err.Placeholder)
	assert.Len(err.Suggestions, 2)
}

func TestErrorSummary(t *testing.T) {
	assert := assert.New(t)

	summary := &docxtpl.ErrorSummary{}
	assert.False(summary.HasErrors())
	assert.Equal("no errors", summary.Error())

	summary.Add(docxtpl.NewTemplateError(docxtpl.ErrCodeSyntaxError, "error 1"))
	assert.True(summary.HasErrors())
	assert.Contains(summary.Error(), "error 1")

	summary.Add(docxtpl.NewTemplateError(docxtpl.ErrCodeUnclosedTag, "error 2"))
	assert.Contains(summary.Error(), "2 errors occurred")
}

func TestErrorSummary_ByCode(t *testing.T) {
	assert := assert.New(t)

	summary := &docxtpl.ErrorSummary{}
	summary.Add(docxtpl.NewTemplateError(docxtpl.ErrCodeSyntaxError, "syntax 1"))
	summary.Add(docxtpl.NewTemplateError(docxtpl.ErrCodeUnclosedTag, "unclosed"))
	summary.Add(docxtpl.NewTemplateError(docxtpl.ErrCodeSyntaxError, "syntax 2"))

	syntaxErrors := summary.ByCode(docxtpl.ErrCodeSyntaxError)
	assert.Len(syntaxErrors, 2)

	unclosedErrors := summary.ByCode(docxtpl.ErrCodeUnclosedTag)
	assert.Len(unclosedErrors, 1)
}

func TestErrorSummary_ByLocation(t *testing.T) {
	assert := assert.New(t)

	summary := &docxtpl.ErrorSummary{}
	summary.Add(docxtpl.NewTemplateError(docxtpl.ErrCodeSyntaxError, "error 1").WithLocation("header"))
	summary.Add(docxtpl.NewTemplateError(docxtpl.ErrCodeSyntaxError, "error 2").WithLocation("footer"))
	summary.Add(docxtpl.NewTemplateError(docxtpl.ErrCodeSyntaxError, "error 3").WithLocation("header"))

	headerErrors := summary.ByLocation("header")
	assert.Len(headerErrors, 2)

	footerErrors := summary.ByLocation("footer")
	assert.Len(footerErrors, 1)
}

func TestErrorSummary_String(t *testing.T) {
	assert := assert.New(t)

	summary := &docxtpl.ErrorSummary{}
	assert.Equal("No errors", summary.String())

	summary.Add(docxtpl.NewTemplateError(docxtpl.ErrCodeSyntaxError, "error 1"))
	summary.Add(docxtpl.NewTemplateError(docxtpl.ErrCodeSyntaxError, "error 2"))

	str := summary.String()
	assert.Contains(str, "Found 2 error(s)")
	assert.Contains(str, "Error 1")
	assert.Contains(str, "Error 2")
}

func TestConvenienceErrorCreators(t *testing.T) {
	assert := assert.New(t)

	// ErrSyntax
	err := docxtpl.ErrSyntax("invalid")
	assert.Equal(docxtpl.ErrCodeSyntaxError, err.Code)
	assert.True(len(err.Suggestions) > 0)

	// ErrUnclosedTag
	err = docxtpl.ErrUnclosedTag("header")
	assert.Equal(docxtpl.ErrCodeUnclosedTag, err.Code)
	assert.Equal("header", err.Location)

	// ErrUnmatchedEnd
	err = docxtpl.ErrUnmatchedEnd("{{end}}", "body")
	assert.Equal(docxtpl.ErrCodeUnmatchedEnd, err.Code)
	assert.Equal("{{end}}", err.Placeholder)

	// ErrUndefinedField
	err = docxtpl.ErrUndefinedField("Name", "body")
	assert.Equal(docxtpl.ErrCodeUndefinedField, err.Code)
	assert.Contains(err.Message, "Name")

	// ErrInvalidFunc
	err = docxtpl.ErrInvalidFunc("myFunc")
	assert.Equal(docxtpl.ErrCodeInvalidFunc, err.Code)
	assert.Contains(err.Message, "myFunc")

	// ErrImageLoad
	err = docxtpl.ErrImageLoad("/path/to/image.png", errors.New("not found"))
	assert.Equal(docxtpl.ErrCodeImageError, err.Code)
	assert.NotNil(err.Cause)

	// ErrFileParse
	err = docxtpl.ErrFileParse("doc.docx", errors.New("corrupted"))
	assert.Equal(docxtpl.ErrCodeCorruptedDocx, err.Code)
}

func TestIsTemplateError(t *testing.T) {
	assert := assert.New(t)

	// Test with TemplateError
	templateErr := docxtpl.NewTemplateError(docxtpl.ErrCodeSyntaxError, "test")
	te, ok := docxtpl.IsTemplateError(templateErr)
	assert.True(ok)
	assert.Equal(templateErr, te)

	// Test with regular error
	regularErr := errors.New("regular error")
	te, ok = docxtpl.IsTemplateError(regularErr)
	assert.False(ok)
	assert.Nil(te)
}
