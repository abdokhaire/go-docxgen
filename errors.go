package docxtpl

import (
	"fmt"
	"strings"
)

// ErrorCode represents a specific error type for programmatic handling.
type ErrorCode string

const (
	// Parse errors
	ErrCodeInvalidFile    ErrorCode = "INVALID_FILE"
	ErrCodeCorruptedDocx  ErrorCode = "CORRUPTED_DOCX"
	ErrCodeFileNotFound   ErrorCode = "FILE_NOT_FOUND"
	ErrCodeReadError      ErrorCode = "READ_ERROR"

	// Template errors
	ErrCodeSyntaxError    ErrorCode = "SYNTAX_ERROR"
	ErrCodeUnclosedTag    ErrorCode = "UNCLOSED_TAG"
	ErrCodeUnmatchedEnd   ErrorCode = "UNMATCHED_END"
	ErrCodeUndefinedField ErrorCode = "UNDEFINED_FIELD"
	ErrCodeInvalidFunc    ErrorCode = "INVALID_FUNCTION"
	ErrCodeExecutionError ErrorCode = "EXECUTION_ERROR"

	// Render errors
	ErrCodeDataConversion ErrorCode = "DATA_CONVERSION"
	ErrCodeImageError     ErrorCode = "IMAGE_ERROR"
	ErrCodeMarshalError   ErrorCode = "MARSHAL_ERROR"

	// Save errors
	ErrCodeWriteError ErrorCode = "WRITE_ERROR"
	ErrCodeZipError   ErrorCode = "ZIP_ERROR"

	// Merge errors
	ErrCodeMergeError ErrorCode = "MERGE_ERROR"
)

// TemplateError is an enhanced error type that provides detailed context
// about where and why an error occurred during template processing.
type TemplateError struct {
	// Code is a programmatic error code for categorization.
	Code ErrorCode

	// Message is a human-readable error description.
	Message string

	// Location describes where the error occurred (e.g., "document body", "header1").
	Location string

	// Placeholder is the template tag that caused the error, if applicable.
	Placeholder string

	// LineNumber is the approximate line in the template, if available.
	LineNumber int

	// Cause is the underlying error that caused this error.
	Cause error

	// Suggestions provides possible fixes for the error.
	Suggestions []string
}

// Error implements the error interface.
func (e *TemplateError) Error() string {
	var parts []string

	// Build the main error message
	if e.Location != "" {
		parts = append(parts, fmt.Sprintf("[%s]", e.Location))
	}

	parts = append(parts, e.Message)

	if e.Placeholder != "" {
		parts = append(parts, fmt.Sprintf("(at: %s)", e.Placeholder))
	}

	if e.LineNumber > 0 {
		parts = append(parts, fmt.Sprintf("(line %d)", e.LineNumber))
	}

	return strings.Join(parts, " ")
}

// Unwrap returns the underlying error for errors.Is and errors.As support.
func (e *TemplateError) Unwrap() error {
	return e.Cause
}

// String returns a detailed multi-line error description.
func (e *TemplateError) String() string {
	var sb strings.Builder

	sb.WriteString("Template Error\n")
	sb.WriteString("==============\n")
	sb.WriteString(fmt.Sprintf("Code:     %s\n", e.Code))
	sb.WriteString(fmt.Sprintf("Message:  %s\n", e.Message))

	if e.Location != "" {
		sb.WriteString(fmt.Sprintf("Location: %s\n", e.Location))
	}

	if e.Placeholder != "" {
		sb.WriteString(fmt.Sprintf("Tag:      %s\n", e.Placeholder))
	}

	if e.LineNumber > 0 {
		sb.WriteString(fmt.Sprintf("Line:     %d\n", e.LineNumber))
	}

	if e.Cause != nil {
		sb.WriteString(fmt.Sprintf("Cause:    %s\n", e.Cause.Error()))
	}

	if len(e.Suggestions) > 0 {
		sb.WriteString("\nSuggestions:\n")
		for i, s := range e.Suggestions {
			sb.WriteString(fmt.Sprintf("  %d. %s\n", i+1, s))
		}
	}

	return sb.String()
}

// WithLocation returns a copy of the error with the location set.
func (e *TemplateError) WithLocation(location string) *TemplateError {
	e.Location = location
	return e
}

// WithPlaceholder returns a copy of the error with the placeholder set.
func (e *TemplateError) WithPlaceholder(placeholder string) *TemplateError {
	e.Placeholder = placeholder
	return e
}

// WithCause returns a copy of the error with the cause set.
func (e *TemplateError) WithCause(cause error) *TemplateError {
	e.Cause = cause
	return e
}

// WithSuggestions returns a copy of the error with suggestions added.
func (e *TemplateError) WithSuggestions(suggestions ...string) *TemplateError {
	e.Suggestions = append(e.Suggestions, suggestions...)
	return e
}

// NewTemplateError creates a new TemplateError.
func NewTemplateError(code ErrorCode, message string) *TemplateError {
	return &TemplateError{
		Code:    code,
		Message: message,
	}
}

// ErrorSummary provides a summary of multiple errors.
type ErrorSummary struct {
	Errors []*TemplateError
}

// Error implements the error interface.
func (s *ErrorSummary) Error() string {
	if len(s.Errors) == 0 {
		return "no errors"
	}
	if len(s.Errors) == 1 {
		return s.Errors[0].Error()
	}
	return fmt.Sprintf("%d errors occurred (first: %s)", len(s.Errors), s.Errors[0].Error())
}

// Add adds an error to the summary.
func (s *ErrorSummary) Add(err *TemplateError) {
	s.Errors = append(s.Errors, err)
}

// HasErrors returns true if there are any errors.
func (s *ErrorSummary) HasErrors() bool {
	return len(s.Errors) > 0
}

// String returns a detailed multi-line description of all errors.
func (s *ErrorSummary) String() string {
	if len(s.Errors) == 0 {
		return "No errors"
	}

	var sb strings.Builder
	sb.WriteString(fmt.Sprintf("Found %d error(s):\n\n", len(s.Errors)))

	for i, err := range s.Errors {
		sb.WriteString(fmt.Sprintf("--- Error %d ---\n", i+1))
		sb.WriteString(err.String())
		sb.WriteString("\n")
	}

	return sb.String()
}

// ByCode returns all errors with a specific code.
func (s *ErrorSummary) ByCode(code ErrorCode) []*TemplateError {
	var result []*TemplateError
	for _, err := range s.Errors {
		if err.Code == code {
			result = append(result, err)
		}
	}
	return result
}

// ByLocation returns all errors from a specific location.
func (s *ErrorSummary) ByLocation(location string) []*TemplateError {
	var result []*TemplateError
	for _, err := range s.Errors {
		if err.Location == location {
			result = append(result, err)
		}
	}
	return result
}

// Common error constructors for convenience

// ErrSyntax creates a syntax error.
func ErrSyntax(message string) *TemplateError {
	return &TemplateError{
		Code:    ErrCodeSyntaxError,
		Message: message,
		Suggestions: []string{
			"Check that all template tags use {{ and }} delimiters",
			"Verify function names and arguments are correct",
			"Ensure all strings are properly quoted",
		},
	}
}

// ErrUnclosedTag creates an unclosed tag error.
func ErrUnclosedTag(location string) *TemplateError {
	return &TemplateError{
		Code:     ErrCodeUnclosedTag,
		Message:  "found {{ without matching }}",
		Location: location,
		Suggestions: []string{
			"Check for missing }} in template tags",
			"Ensure tag delimiters are not split across formatting",
		},
	}
}

// ErrUnmatchedEnd creates an unmatched end error.
func ErrUnmatchedEnd(tag string, location string) *TemplateError {
	return &TemplateError{
		Code:        ErrCodeUnmatchedEnd,
		Message:     "{{end}} without matching block start",
		Placeholder: tag,
		Location:    location,
		Suggestions: []string{
			"Ensure every {{end}} has a matching {{if}}, {{range}}, {{with}}, or {{define}}",
			"Check for extra {{end}} tags",
		},
	}
}

// ErrUndefinedField creates an undefined field error.
func ErrUndefinedField(field string, location string) *TemplateError {
	return &TemplateError{
		Code:        ErrCodeUndefinedField,
		Message:     fmt.Sprintf("field %q not found in template data", field),
		Placeholder: "{{." + field + "}}",
		Location:    location,
		Suggestions: []string{
			fmt.Sprintf("Add a %q field to your data", field),
			"Check for typos in the field name",
			"Ensure nested fields use proper dot notation (e.g., .Parent.Child)",
		},
	}
}

// ErrInvalidFunc creates an invalid function error.
func ErrInvalidFunc(funcName string) *TemplateError {
	return &TemplateError{
		Code:    ErrCodeInvalidFunc,
		Message: fmt.Sprintf("function %q is not defined", funcName),
		Suggestions: []string{
			fmt.Sprintf("Register the function using doc.RegisterFunction(%q, fn)", funcName),
			"Use doc.RegisterFuncMap(sprig.FuncMap()) to add Sprig functions",
			"Check for typos in the function name",
		},
	}
}

// ErrImageLoad creates an image loading error.
func ErrImageLoad(path string, cause error) *TemplateError {
	return &TemplateError{
		Code:    ErrCodeImageError,
		Message: fmt.Sprintf("failed to load image: %s", path),
		Cause:   cause,
		Suggestions: []string{
			"Verify the file path is correct",
			"Ensure the file exists and is readable",
			"Check that the image is a valid JPEG or PNG file",
		},
	}
}

// ErrFileParse creates a file parsing error.
func ErrFileParse(filename string, cause error) *TemplateError {
	return &TemplateError{
		Code:    ErrCodeCorruptedDocx,
		Message: fmt.Sprintf("failed to parse DOCX file: %s", filename),
		Cause:   cause,
		Suggestions: []string{
			"Verify the file is a valid DOCX document",
			"Try opening and re-saving the file in Word",
			"Check if the file is corrupted or incomplete",
		},
	}
}

// IsTemplateError checks if an error is a TemplateError and returns it.
func IsTemplateError(err error) (*TemplateError, bool) {
	if te, ok := err.(*TemplateError); ok {
		return te, true
	}
	return nil, false
}
