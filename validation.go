package docxtpl

import (
	"fmt"
	"reflect"
	"regexp"
	"strings"
	"text/template"

	"github.com/abdokhaire/go-docxgen/internal/tags"
	"github.com/abdokhaire/go-docxgen/internal/xmlutils"
)

// =============================================================================
// Error Codes and Types
// =============================================================================

// ErrorCode represents a specific error type for programmatic handling.
type ErrorCode string

const (
	// Parse errors
	ErrCodeInvalidFile   ErrorCode = "INVALID_FILE"
	ErrCodeCorruptedDocx ErrorCode = "CORRUPTED_DOCX"
	ErrCodeFileNotFound  ErrorCode = "FILE_NOT_FOUND"
	ErrCodeReadError     ErrorCode = "READ_ERROR"

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

// =============================================================================
// Common Error Constructors
// =============================================================================

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

// =============================================================================
// Validation Error Types
// =============================================================================

// ValidationErrorType represents the type of validation error.
type ValidationErrorType string

const (
	ErrorUnclosedTag     ValidationErrorType = "UNCLOSED_TAG"
	ErrorUnmatchedEnd    ValidationErrorType = "UNMATCHED_END"
	ErrorSyntaxError     ValidationErrorType = "SYNTAX_ERROR"
	ErrorUndefinedField  ValidationErrorType = "UNDEFINED_FIELD"
	ErrorInvalidFunction ValidationErrorType = "INVALID_FUNCTION"
)

// ValidationResult contains the results of template validation.
type ValidationResult struct {
	Valid  bool
	Errors []ValidationError
}

// HasErrors returns true if there are any validation errors.
func (r *ValidationResult) HasErrors() bool {
	return len(r.Errors) > 0
}

// Error returns a combined error message for all validation errors.
func (r *ValidationResult) Error() string {
	if !r.HasErrors() {
		return ""
	}
	var messages []string
	for _, err := range r.Errors {
		messages = append(messages, err.Error())
	}
	return strings.Join(messages, "; ")
}

// Validate checks the template for common errors without rendering.
// It returns a ValidationResult containing any errors found.
//
//	result := doc.Validate()
//	if result.HasErrors() {
//	    for _, err := range result.Errors {
//	        fmt.Println(err)
//	    }
//	}
func (d *DocxTmpl) Validate() *ValidationResult {
	result := &ValidationResult{Valid: true}

	// Merge tags in document body
	tags.MergeTags(d.Document.Body.Items)

	// Get document XML
	documentXml, err := d.getDocumentXml()
	if err != nil {
		result.Valid = false
		result.Errors = append(result.Errors, ValidationError{
			Field:   "document",
			Message: "failed to get document XML: " + err.Error(),
		})
		return result
	}

	// Validate document body
	bodyErrors := d.validateXmlContent(documentXml, "document body")
	result.Errors = append(result.Errors, bodyErrors...)

	// Validate headers, footers, etc.
	for _, pf := range d.processableFiles {
		location := getLocationName(pf.Name)
		mergedContent := xmlutils.MergeFragmentedTagsInXml(pf.Content)
		pfErrors := d.validateXmlContent(mergedContent, location)
		result.Errors = append(result.Errors, pfErrors...)
	}

	result.Valid = len(result.Errors) == 0
	return result
}

// validateXmlContent validates template syntax in XML content.
func (d *DocxTmpl) validateXmlContent(content string, location string) []ValidationError {
	var errors []ValidationError

	// Check for unclosed tags
	errors = append(errors, checkUnclosedTags(content, location)...)

	// Check for unmatched end tags
	errors = append(errors, checkUnmatchedEnds(content, location)...)

	// Try to parse the template to catch syntax errors
	errors = append(errors, d.checkTemplateSyntax(content, location)...)

	return errors
}

// checkUnclosedTags checks for unclosed {{ without matching }}.
func checkUnclosedTags(content string, location string) []ValidationError {
	var errors []ValidationError

	openCount := strings.Count(content, "{{")
	closeCount := strings.Count(content, "}}")

	if openCount > closeCount {
		errors = append(errors, ValidationError{
			Field:       location,
			Message:     fmt.Sprintf("unclosed template tag detected (%d {{ vs %d }})", openCount, closeCount),
			Placeholder: "{{...}}",
		})
	}

	return errors
}

// checkUnmatchedEnds checks for {{end}} without matching {{if}}, {{range}}, etc.
func checkUnmatchedEnds(content string, location string) []ValidationError {
	var errors []ValidationError

	tagRe := regexp.MustCompile(`\{\{[^}]+\}\}`)
	allTags := tagRe.FindAllString(content, -1)

	blockStack := 0
	for _, tag := range allTags {
		tagLower := strings.ToLower(tag)
		if strings.Contains(tagLower, "{{if") ||
			strings.Contains(tagLower, "{{range") ||
			strings.Contains(tagLower, "{{with") ||
			strings.Contains(tagLower, "{{block") ||
			strings.Contains(tagLower, "{{define") {
			blockStack++
		} else if strings.Contains(tagLower, "{{end") {
			blockStack--
			if blockStack < 0 {
				errors = append(errors, ValidationError{
					Field:       location,
					Message:     "{{end}} without matching opening tag",
					Placeholder: tag,
				})
				blockStack = 0
			}
		}
	}

	if blockStack > 0 {
		errors = append(errors, ValidationError{
			Field:       location,
			Message:     fmt.Sprintf("missing %d {{end}} tag(s)", blockStack),
			Placeholder: "",
		})
	}

	return errors
}

// checkTemplateSyntax attempts to parse the template to catch syntax errors.
func (d *DocxTmpl) checkTemplateSyntax(content string, location string) []ValidationError {
	var errors []ValidationError

	preparedContent, err := xmlutils.PrepareXmlForTagReplacement(content)
	if err != nil {
		return errors
	}

	_, err = template.New("validate").Funcs(d.funcMap).Parse(preparedContent)
	if err != nil {
		errMsg := err.Error()

		// Extract the problematic tag if possible
		tagRe := regexp.MustCompile(`"([^"]*)"`)
		matches := tagRe.FindStringSubmatch(errMsg)
		problematicTag := ""
		if len(matches) >= 2 {
			problematicTag = matches[1]
		}

		errors = append(errors, ValidationError{
			Field:       location,
			Message:     errMsg,
			Placeholder: problematicTag,
		})
	}

	return errors
}

// getLocationName converts a file path to a human-readable location name.
func getLocationName(filePath string) string {
	if strings.Contains(filePath, "header") {
		return "header"
	}
	if strings.Contains(filePath, "footer") {
		return "footer"
	}
	if strings.Contains(filePath, "footnotes") {
		return "footnotes"
	}
	if strings.Contains(filePath, "endnotes") {
		return "endnotes"
	}
	if strings.Contains(filePath, "core.xml") {
		return "document properties"
	}
	return filePath
}

// ValidationError represents a template validation error
type ValidationError struct {
	Field       string // Field name that caused the error
	Message     string // Error description
	Placeholder string // Original placeholder text
}

func (e ValidationError) Error() string {
	return fmt.Sprintf("%s: %s", e.Field, e.Message)
}

// FieldInfo represents information about a template field
type FieldInfo struct {
	Name        string   // Field name (e.g., "Person.Name")
	Required    bool     // Whether the field is required
	Type        string   // Expected type (string, slice, etc.)
	Example     string   // Example value
	Occurrences int      // Number of times field appears
	Locations   []string // Where field appears (body, header, footer)
}

// FieldSchema represents a JSON schema-like field definition
type FieldSchema struct {
	Type       string                 `json:"type"`                 // "string", "number", "boolean", "array", "object"
	Required   bool                   `json:"required"`
	Properties map[string]FieldSchema `json:"properties,omitempty"` // For nested objects
	Items      *FieldSchema           `json:"items,omitempty"`      // For arrays
}

// ValidateData validates that the provided data contains all required template fields.
// Returns a slice of validation errors, or empty slice if valid.
//
//	errors := doc.ValidateData(data)
//	if len(errors) > 0 {
//	    for _, err := range errors {
//	        fmt.Println(err)
//	    }
//	}
func (d *DocxTmpl) ValidateData(data any) []ValidationError {
	var errors []ValidationError

	// Get all placeholders from template
	placeholders, _ := d.GetPlaceholders()

	// Convert data to map for checking
	dataMap := toMap(data)

	// Check each placeholder
	for _, ph := range placeholders {
		// Extract field name from placeholder (remove {{ }}, trim whitespace)
		fieldName := extractFieldName(ph)
		if fieldName == "" {
			continue
		}

		// Skip control flow statements
		if isControlFlow(fieldName) {
			continue
		}

		// Check if field exists in data
		if !fieldExists(dataMap, fieldName) {
			errors = append(errors, ValidationError{
				Field:       fieldName,
				Message:     "field not found in data",
				Placeholder: ph,
			})
		}
	}

	return errors
}

// GetRequiredFields returns information about all fields required by the template.
//
//	fields := doc.GetRequiredFields()
//	for _, f := range fields {
//	    fmt.Printf("Field: %s, Occurrences: %d\n", f.Name, f.Occurrences)
//	}
func (d *DocxTmpl) GetRequiredFields() []FieldInfo {
	fieldMap := make(map[string]*FieldInfo)

	// Get placeholders from document body
	placeholders, _ := d.GetPlaceholders()

	for _, ph := range placeholders {
		fieldName := extractFieldName(ph)
		if fieldName == "" || isControlFlow(fieldName) {
			continue
		}

		if info, exists := fieldMap[fieldName]; exists {
			info.Occurrences++
		} else {
			fieldMap[fieldName] = &FieldInfo{
				Name:        fieldName,
				Required:    true,
				Type:        inferType(fieldName),
				Occurrences: 1,
				Locations:   []string{"body"},
			}
		}
	}

	// Convert to slice
	fields := make([]FieldInfo, 0, len(fieldMap))
	for _, info := range fieldMap {
		fields = append(fields, *info)
	}

	return fields
}

// GetPlaceholderSchema returns a schema-like map of all placeholders.
// Useful for generating data templates or documentation.
//
//	schema := doc.GetPlaceholderSchema()
func (d *DocxTmpl) GetPlaceholderSchema() map[string]FieldSchema {
	schema := make(map[string]FieldSchema)

	fields := d.GetRequiredFields()
	for _, f := range fields {
		schema[f.Name] = FieldSchema{
			Type:     f.Type,
			Required: f.Required,
		}
	}

	return schema
}

// PreviewRender renders the template with data but returns the text content
// instead of saving to a file. Useful for validation and previewing.
//
//	preview, err := doc.PreviewRender(data)
func (d *DocxTmpl) PreviewRender(data any) (string, error) {
	// Create a copy to avoid modifying the original
	// For preview, we'll just render and get text
	err := d.Render(data)
	if err != nil {
		return "", err
	}

	return d.GetText(), nil
}

// GenerateSampleData generates sample data matching the template placeholders.
// This is useful for testing and documentation.
//
//	sample := doc.GenerateSampleData()
func (d *DocxTmpl) GenerateSampleData() map[string]interface{} {
	sample := make(map[string]interface{})

	fields := d.GetRequiredFields()
	for _, f := range fields {
		// Handle nested fields
		parts := strings.Split(f.Name, ".")
		current := sample

		for i, part := range parts {
			if i == len(parts)-1 {
				// Last part - set value
				current[part] = generateSampleValue(part, f.Type)
			} else {
				// Intermediate part - create nested map
				if _, exists := current[part]; !exists {
					current[part] = make(map[string]interface{})
				}
				if nested, ok := current[part].(map[string]interface{}); ok {
					current = nested
				}
			}
		}
	}

	return sample
}

// Helper functions

func extractFieldName(placeholder string) string {
	// Remove {{ }} and trim
	s := strings.TrimPrefix(placeholder, "{{")
	s = strings.TrimSuffix(s, "}}")
	s = strings.TrimSpace(s)

	// Handle whitespace trimming syntax
	s = strings.TrimPrefix(s, "-")
	s = strings.TrimSuffix(s, "-")
	s = strings.TrimSpace(s)

	// Handle pipe functions (take only the first part)
	if idx := strings.Index(s, "|"); idx > 0 {
		s = strings.TrimSpace(s[:idx])
	}

	// Handle function calls like "upper .Name" -> ".Name"
	if strings.Contains(s, " ") {
		parts := strings.Fields(s)
		for _, p := range parts {
			if strings.HasPrefix(p, ".") {
				s = p
				break
			}
		}
	}

	// Remove leading dot
	s = strings.TrimPrefix(s, ".")

	return s
}

func isControlFlow(field string) bool {
	// Check if it's a control flow statement
	lower := strings.ToLower(field)
	controlKeywords := []string{"if", "else", "end", "range", "with", "define", "template", "block"}
	for _, kw := range controlKeywords {
		if strings.HasPrefix(lower, kw) {
			return true
		}
	}
	return false
}

func fieldExists(data map[string]interface{}, fieldPath string) bool {
	parts := strings.Split(fieldPath, ".")

	current := data
	for i, part := range parts {
		val, exists := current[part]
		if !exists {
			return false
		}

		if i == len(parts)-1 {
			return true // Found the final field
		}

		// Navigate deeper
		if nested, ok := val.(map[string]interface{}); ok {
			current = nested
		} else {
			return false
		}
	}

	return true
}

func toMap(data any) map[string]interface{} {
	if data == nil {
		return nil
	}

	// If already a map
	if m, ok := data.(map[string]interface{}); ok {
		return m
	}
	if m, ok := data.(map[string]string); ok {
		result := make(map[string]interface{})
		for k, v := range m {
			result[k] = v
		}
		return result
	}

	// Convert struct to map using reflection
	v := reflect.ValueOf(data)
	if v.Kind() == reflect.Ptr {
		v = v.Elem()
	}

	if v.Kind() != reflect.Struct {
		return nil
	}

	result := make(map[string]interface{})
	t := v.Type()
	for i := 0; i < v.NumField(); i++ {
		field := t.Field(i)
		if field.PkgPath != "" {
			continue // Skip unexported fields
		}
		result[field.Name] = v.Field(i).Interface()
	}

	return result
}

func inferType(fieldName string) string {
	lower := strings.ToLower(fieldName)

	// Check for common patterns
	if strings.Contains(lower, "date") || strings.Contains(lower, "time") {
		return "date"
	}
	if strings.Contains(lower, "price") || strings.Contains(lower, "amount") ||
		strings.Contains(lower, "total") || strings.Contains(lower, "count") {
		return "number"
	}
	if strings.Contains(lower, "items") || strings.Contains(lower, "list") ||
		strings.HasSuffix(lower, "s") {
		return "array"
	}
	if strings.Contains(lower, "is") || strings.Contains(lower, "has") ||
		strings.Contains(lower, "enabled") || strings.Contains(lower, "active") {
		return "boolean"
	}
	if strings.Contains(lower, "image") || strings.Contains(lower, "photo") ||
		strings.Contains(lower, "logo") {
		return "image"
	}

	return "string"
}

func generateSampleValue(fieldName, fieldType string) interface{} {
	switch fieldType {
	case "number":
		return 0
	case "boolean":
		return false
	case "date":
		return "2024-01-01"
	case "array":
		return []interface{}{}
	case "image":
		return "/path/to/image.png"
	default:
		return fmt.Sprintf("[%s]", fieldName)
	}
}

// ValidatePlaceholderSyntax checks all placeholders for correct syntax.
// Returns errors for malformed placeholders.
//
//	errors := doc.ValidatePlaceholderSyntax()
func (d *DocxTmpl) ValidatePlaceholderSyntax() []ValidationError {
	var errors []ValidationError

	text := d.GetText()

	// Find all {{ }} patterns
	re := regexp.MustCompile(`\{\{[^}]*\}\}`)
	matches := re.FindAllString(text, -1)

	for _, match := range matches {
		// Check for common issues
		content := strings.Trim(match, "{}")
		content = strings.TrimSpace(content)

		if content == "" {
			errors = append(errors, ValidationError{
				Field:       "",
				Message:     "empty placeholder",
				Placeholder: match,
			})
			continue
		}

		// Check for unbalanced braces
		if strings.Count(match, "{") != strings.Count(match, "}") {
			errors = append(errors, ValidationError{
				Field:       content,
				Message:     "unbalanced braces",
				Placeholder: match,
			})
		}
	}

	// Check for unclosed placeholders
	openCount := strings.Count(text, "{{")
	closeCount := strings.Count(text, "}}")
	if openCount != closeCount {
		errors = append(errors, ValidationError{
			Field:   "",
			Message: fmt.Sprintf("mismatched placeholder delimiters: %d {{ vs %d }}", openCount, closeCount),
		})
	}

	return errors
}
