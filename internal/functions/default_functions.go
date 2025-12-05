package functions

import (
	"fmt"
	"reflect"
	"strings"
	"text/template"
	"time"
	"unicode"

	"github.com/google/uuid"
	"golang.org/x/text/cases"
	"golang.org/x/text/language"
	"golang.org/x/text/message"
)

var DefaultFuncMap = template.FuncMap{
	// Text case functions
	"upper": strings.ToUpper,
	"lower": strings.ToLower,
	"title": title,

	// Rich text formatting functions
	"bold":            bold,
	"italic":          italic,
	"underline":       underline,
	"strikethrough":   strikethrough,
	"doubleStrike":    doubleStrike,
	"color":           color,
	"highlight":       highlight,
	"fontSize":        fontSize,
	"fontFamily":      fontFamily,
	"font":            font,
	"subscript":       subscript,
	"superscript":     superscript,
	"smallCaps":       smallCaps,
	"allCaps":         allCaps,
	"shadow":          shadow,
	"outline":         outline,
	"emboss":          emboss,
	"imprint":         imprint,
	"bgColor":         bgColor,

	// Hyperlink function
	"link": link,

	// Special characters
	"br":  lineBreak,
	"tab": tab,

	// Comparison functions
	"eq": eq,
	"ne": ne,
	"lt": lt,
	"le": le,
	"gt": gt,
	"ge": ge,

	// Logical functions
	"and": and,
	"or":  or,
	"not": not,

	// Collection functions
	"len":     length,
	"first":   first,
	"last":    last,
	"index":   index,
	"slice":   sliceFn,
	"join":    join,
	"contains": contains,

	// Utility functions
	"default": defaultVal,
	"coalesce": coalesce,
	"ternary":  ternary,
	"repeat":   strings.Repeat,
	"replace":  strings.ReplaceAll,
	"trim":     strings.TrimSpace,
	"trimPrefix": strings.TrimPrefix,
	"trimSuffix": strings.TrimSuffix,
	"split":    split,
	"concat":   concat,

	// Math functions
	"add":  add,
	"sub":  sub,
	"mul":  mul,
	"div":  div,
	"mod":  mod,

	// Number formatting
	"formatNumber": formatNumber,
	"formatMoney":  formatMoney,
	"formatPercent": formatPercent,

	// Date/Time functions
	"now":        now,
	"formatDate": formatDate,
	"parseDate":  parseDate,
	"addDays":    addDays,
	"addMonths":  addMonths,
	"addYears":   addYears,

	// Document structure
	"pageBreak":    pageBreak,
	"sectionBreak": sectionBreak,

	// Additional utilities
	"uuid":      generateUUID,
	"pluralize": pluralize,
	"truncate":  truncate,
	"wordwrap":  wordwrap,
	"capitalize": capitalize,
	"camelCase":  camelCase,
	"snakeCase":  snakeCase,
	"kebabCase":  kebabCase,
}

func title(text string) string {
	caser := cases.Title(language.English)
	return caser.String(text)
}

// bold returns text wrapped in bold formatting.
// Usage: {{bold .Text}} or {{.Text | bold}}
func bold(text string) string {
	return `</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>` + text + `</w:t></w:r><w:r><w:t>`
}

// italic returns text wrapped in italic formatting.
// Usage: {{italic .Text}} or {{.Text | italic}}
func italic(text string) string {
	return `</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>` + text + `</w:t></w:r><w:r><w:t>`
}

// underline returns text wrapped in underline formatting.
// Usage: {{underline .Text}} or {{.Text | underline}}
func underline(text string) string {
	return `</w:t></w:r><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>` + text + `</w:t></w:r><w:r><w:t>`
}

// strikethrough returns text wrapped in strikethrough formatting.
// Usage: {{strikethrough .Text}} or {{.Text | strikethrough}}
func strikethrough(text string) string {
	return `</w:t></w:r><w:r><w:rPr><w:strike/></w:rPr><w:t>` + text + `</w:t></w:r><w:r><w:t>`
}

// doubleStrike returns text wrapped in double strikethrough formatting.
// Usage: {{doubleStrike .Text}}
func doubleStrike(text string) string {
	return `</w:t></w:r><w:r><w:rPr><w:dstrike/></w:rPr><w:t>` + text + `</w:t></w:r><w:r><w:t>`
}

// color returns text with the specified color (hex code without #).
// Usage: {{color "FF0000" .Text}} for red text
func color(hexColor string, text string) string {
	return fmt.Sprintf(`</w:t></w:r><w:r><w:rPr><w:color w:val="%s"/></w:rPr><w:t>%s</w:t></w:r><w:r><w:t>`, hexColor, text)
}

// fontSize returns text with the specified font size.
// Size is in half-points (e.g., 24 = 12pt, 28 = 14pt).
// Usage: {{fontSize 28 .Text}} for 14pt text
func fontSize(halfPoints int, text string) string {
	return fmt.Sprintf(`</w:t></w:r><w:r><w:rPr><w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr><w:t>%s</w:t></w:r><w:r><w:t>`, halfPoints, halfPoints, text)
}

// fontFamily returns text with the specified font family.
// Usage: {{fontFamily "Arial" .Text}}
func fontFamily(fontName string, text string) string {
	return fmt.Sprintf(`</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s"/></w:rPr><w:t>%s</w:t></w:r><w:r><w:t>`, fontName, fontName, text)
}

// font returns text with combined font settings.
// Usage: {{font "Arial" 24 "FF0000" .Text}} for Arial, 12pt, red
func font(fontName string, halfPoints int, hexColor string, text string) string {
	return fmt.Sprintf(`</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s"/><w:sz w:val="%d"/><w:szCs w:val="%d"/><w:color w:val="%s"/></w:rPr><w:t>%s</w:t></w:r><w:r><w:t>`,
		fontName, fontName, halfPoints, halfPoints, hexColor, text)
}

// subscript returns text formatted as subscript.
// Usage: {{subscript .Text}}
func subscript(text string) string {
	return `</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t>` + text + `</w:t></w:r><w:r><w:t>`
}

// superscript returns text formatted as superscript.
// Usage: {{superscript .Text}}
func superscript(text string) string {
	return `</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>` + text + `</w:t></w:r><w:r><w:t>`
}

// smallCaps returns text in small capitals.
// Usage: {{smallCaps .Text}}
func smallCaps(text string) string {
	return `</w:t></w:r><w:r><w:rPr><w:smallCaps/></w:rPr><w:t>` + text + `</w:t></w:r><w:r><w:t>`
}

// allCaps returns text displayed in all capitals.
// Usage: {{allCaps .Text}}
func allCaps(text string) string {
	return `</w:t></w:r><w:r><w:rPr><w:caps/></w:rPr><w:t>` + text + `</w:t></w:r><w:r><w:t>`
}

// shadow returns text with shadow effect.
// Usage: {{shadow .Text}}
func shadow(text string) string {
	return `</w:t></w:r><w:r><w:rPr><w:shadow/></w:rPr><w:t>` + text + `</w:t></w:r><w:r><w:t>`
}

// outline returns text with outline effect.
// Usage: {{outline .Text}}
func outline(text string) string {
	return `</w:t></w:r><w:r><w:rPr><w:outline/></w:rPr><w:t>` + text + `</w:t></w:r><w:r><w:t>`
}

// emboss returns text with emboss effect.
// Usage: {{emboss .Text}}
func emboss(text string) string {
	return `</w:t></w:r><w:r><w:rPr><w:emboss/></w:rPr><w:t>` + text + `</w:t></w:r><w:r><w:t>`
}

// imprint returns text with imprint (engrave) effect.
// Usage: {{imprint .Text}}
func imprint(text string) string {
	return `</w:t></w:r><w:r><w:rPr><w:imprint/></w:rPr><w:t>` + text + `</w:t></w:r><w:r><w:t>`
}

// bgColor returns text with a background/shading color.
// Usage: {{bgColor "FFFF00" .Text}} for yellow background
func bgColor(hexColor string, text string) string {
	return fmt.Sprintf(`</w:t></w:r><w:r><w:rPr><w:shd w:val="clear" w:color="auto" w:fill="%s"/></w:rPr><w:t>%s</w:t></w:r><w:r><w:t>`, hexColor, text)
}

// highlight returns text with the specified highlight color.
// Valid colors: yellow, green, cyan, magenta, blue, red, darkBlue, darkCyan, darkGreen, darkMagenta, darkRed, darkYellow, darkGray, lightGray, black
// Usage: {{highlight "yellow" .Text}}
func highlight(highlightColor string, text string) string {
	return fmt.Sprintf(`</w:t></w:r><w:r><w:rPr><w:highlight w:val="%s"/></w:rPr><w:t>%s</w:t></w:r><w:r><w:t>`, highlightColor, text)
}

// link creates a hyperlink with the given URL and display text.
// Usage: {{link "https://example.com" "Click here"}}
func link(url string, text string) string {
	// Note: Full hyperlink support requires relationship IDs and modifying document.xml.rels
	// This simplified version uses a field code approach
	return fmt.Sprintf(`</w:t></w:r><w:hyperlink r:id="" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr><w:t>%s</w:t></w:r></w:hyperlink><w:r><w:t>`, text)
}

// lineBreak inserts a line break (soft return).
// Usage: {{br}}
func lineBreak() string {
	return `</w:t><w:br/><w:t>`
}

// tab inserts a tab character.
// Usage: {{tab}}
func tab() string {
	return `</w:t><w:tab/><w:t>`
}

// ============================================================================
// Comparison Functions
// ============================================================================

// eq returns true if a == b
// Usage: {{if eq .Status "active"}}...{{end}}
func eq(a, b any) bool {
	return reflect.DeepEqual(a, b)
}

// ne returns true if a != b
// Usage: {{if ne .Status "inactive"}}...{{end}}
func ne(a, b any) bool {
	return !reflect.DeepEqual(a, b)
}

// lt returns true if a < b (for comparable types)
// Usage: {{if lt .Count 10}}...{{end}}
func lt(a, b any) bool {
	return compare(a, b) < 0
}

// le returns true if a <= b
// Usage: {{if le .Count 10}}...{{end}}
func le(a, b any) bool {
	return compare(a, b) <= 0
}

// gt returns true if a > b
// Usage: {{if gt .Count 10}}...{{end}}
func gt(a, b any) bool {
	return compare(a, b) > 0
}

// ge returns true if a >= b
// Usage: {{if ge .Count 10}}...{{end}}
func ge(a, b any) bool {
	return compare(a, b) >= 0
}

// compare compares two values and returns -1, 0, or 1
func compare(a, b any) int {
	av := reflect.ValueOf(a)
	bv := reflect.ValueOf(b)

	// Handle numeric comparisons
	if isNumeric(av) && isNumeric(bv) {
		af := toFloat64(av)
		bf := toFloat64(bv)
		if af < bf {
			return -1
		} else if af > bf {
			return 1
		}
		return 0
	}

	// Handle string comparisons
	if av.Kind() == reflect.String && bv.Kind() == reflect.String {
		as := av.String()
		bs := bv.String()
		if as < bs {
			return -1
		} else if as > bs {
			return 1
		}
		return 0
	}

	return 0
}

func isNumeric(v reflect.Value) bool {
	switch v.Kind() {
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64,
		reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64,
		reflect.Float32, reflect.Float64:
		return true
	}
	return false
}

func toFloat64(v reflect.Value) float64 {
	switch v.Kind() {
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		return float64(v.Int())
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		return float64(v.Uint())
	case reflect.Float32, reflect.Float64:
		return v.Float()
	}
	return 0
}

// ============================================================================
// Logical Functions
// ============================================================================

// and returns true if all arguments are truthy
// Usage: {{if and .IsActive .IsVerified}}...{{end}}
func and(args ...any) bool {
	for _, arg := range args {
		if !isTruthy(arg) {
			return false
		}
	}
	return true
}

// or returns true if any argument is truthy
// Usage: {{if or .IsAdmin .IsModerator}}...{{end}}
func or(args ...any) bool {
	for _, arg := range args {
		if isTruthy(arg) {
			return true
		}
	}
	return false
}

// not returns the boolean negation
// Usage: {{if not .IsDeleted}}...{{end}}
func not(arg any) bool {
	return !isTruthy(arg)
}

// isTruthy returns whether a value is considered "truthy"
func isTruthy(val any) bool {
	if val == nil {
		return false
	}

	v := reflect.ValueOf(val)
	switch v.Kind() {
	case reflect.Bool:
		return v.Bool()
	case reflect.String:
		return v.String() != ""
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		return v.Int() != 0
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		return v.Uint() != 0
	case reflect.Float32, reflect.Float64:
		return v.Float() != 0
	case reflect.Slice, reflect.Map, reflect.Array:
		return v.Len() > 0
	case reflect.Ptr, reflect.Interface:
		return !v.IsNil()
	}
	return true
}

// ============================================================================
// Collection Functions
// ============================================================================

// length returns the length of a slice, map, string, or array
// Usage: {{len .Items}}
func length(v any) int {
	if v == nil {
		return 0
	}
	rv := reflect.ValueOf(v)
	switch rv.Kind() {
	case reflect.Slice, reflect.Map, reflect.Array, reflect.String, reflect.Chan:
		return rv.Len()
	}
	return 0
}

// first returns the first element of a slice or array
// Usage: {{first .Items}}
func first(v any) any {
	if v == nil {
		return nil
	}
	rv := reflect.ValueOf(v)
	if rv.Kind() == reflect.Slice || rv.Kind() == reflect.Array {
		if rv.Len() > 0 {
			return rv.Index(0).Interface()
		}
	}
	return nil
}

// last returns the last element of a slice or array
// Usage: {{last .Items}}
func last(v any) any {
	if v == nil {
		return nil
	}
	rv := reflect.ValueOf(v)
	if rv.Kind() == reflect.Slice || rv.Kind() == reflect.Array {
		if rv.Len() > 0 {
			return rv.Index(rv.Len() - 1).Interface()
		}
	}
	return nil
}

// index returns the element at index i from a slice, array, or map
// Usage: {{index .Items 0}} or {{index .Map "key"}}
func index(v any, idx any) any {
	if v == nil {
		return nil
	}
	rv := reflect.ValueOf(v)
	switch rv.Kind() {
	case reflect.Slice, reflect.Array:
		i := toInt(idx)
		if i >= 0 && i < rv.Len() {
			return rv.Index(i).Interface()
		}
	case reflect.Map:
		key := reflect.ValueOf(idx)
		if key.Type().ConvertibleTo(rv.Type().Key()) {
			result := rv.MapIndex(key.Convert(rv.Type().Key()))
			if result.IsValid() {
				return result.Interface()
			}
		}
	case reflect.String:
		i := toInt(idx)
		s := rv.String()
		if i >= 0 && i < len(s) {
			return string(s[i])
		}
	}
	return nil
}

// sliceFn returns a slice of the input from start to end
// Usage: {{slice .Items 1 3}}
func sliceFn(v any, args ...int) any {
	if v == nil {
		return nil
	}
	rv := reflect.ValueOf(v)
	if rv.Kind() != reflect.Slice && rv.Kind() != reflect.Array && rv.Kind() != reflect.String {
		return nil
	}

	length := rv.Len()
	start := 0
	end := length

	if len(args) >= 1 {
		start = args[0]
		if start < 0 {
			start = 0
		}
	}
	if len(args) >= 2 {
		end = args[1]
		if end > length {
			end = length
		}
	}

	if start > end || start >= length {
		return reflect.MakeSlice(rv.Type(), 0, 0).Interface()
	}

	return rv.Slice(start, end).Interface()
}

// join concatenates slice elements with a separator
// Usage: {{join .Names ", "}}
func join(v any, sep string) string {
	if v == nil {
		return ""
	}
	rv := reflect.ValueOf(v)
	if rv.Kind() != reflect.Slice && rv.Kind() != reflect.Array {
		return fmt.Sprintf("%v", v)
	}

	parts := make([]string, rv.Len())
	for i := range rv.Len() {
		parts[i] = fmt.Sprintf("%v", rv.Index(i).Interface())
	}
	return strings.Join(parts, sep)
}

// contains checks if a slice contains an element or a string contains a substring
// Usage: {{if contains .Roles "admin"}}...{{end}}
func contains(v any, elem any) bool {
	if v == nil {
		return false
	}
	rv := reflect.ValueOf(v)

	// String contains
	if rv.Kind() == reflect.String {
		if es, ok := elem.(string); ok {
			return strings.Contains(rv.String(), es)
		}
		return false
	}

	// Slice/Array contains
	if rv.Kind() == reflect.Slice || rv.Kind() == reflect.Array {
		for i := range rv.Len() {
			if reflect.DeepEqual(rv.Index(i).Interface(), elem) {
				return true
			}
		}
	}

	// Map contains key
	if rv.Kind() == reflect.Map {
		key := reflect.ValueOf(elem)
		if key.Type().ConvertibleTo(rv.Type().Key()) {
			return rv.MapIndex(key.Convert(rv.Type().Key())).IsValid()
		}
	}

	return false
}

// ============================================================================
// Utility Functions
// ============================================================================

// defaultVal returns the default value if the first argument is empty/zero
// Usage: {{default "N/A" .Name}}
func defaultVal(defaultValue, value any) any {
	if !isTruthy(value) {
		return defaultValue
	}
	return value
}

// coalesce returns the first non-empty value
// Usage: {{coalesce .PreferredName .Name "Anonymous"}}
func coalesce(values ...any) any {
	for _, v := range values {
		if isTruthy(v) {
			return v
		}
	}
	return nil
}

// ternary returns trueVal if condition is true, otherwise falseVal
// Usage: {{ternary "Yes" "No" .IsActive}}
func ternary(trueVal, falseVal any, condition bool) any {
	if condition {
		return trueVal
	}
	return falseVal
}

// split splits a string by separator
// Usage: {{range split .Tags ","}}...{{end}}
func split(s, sep string) []string {
	if s == "" {
		return []string{}
	}
	return strings.Split(s, sep)
}

// concat concatenates multiple values into a string
// Usage: {{concat .FirstName " " .LastName}}
func concat(values ...any) string {
	var builder strings.Builder
	for _, v := range values {
		builder.WriteString(fmt.Sprintf("%v", v))
	}
	return builder.String()
}

// ============================================================================
// Math Functions
// ============================================================================

// add returns a + b
// Usage: {{add .Count 1}}
func add(a, b any) any {
	av := reflect.ValueOf(a)
	bv := reflect.ValueOf(b)

	if isNumeric(av) && isNumeric(bv) {
		// Check if both are integers
		if isInteger(av) && isInteger(bv) {
			return toInt64(av) + toInt64(bv)
		}
		return toFloat64(av) + toFloat64(bv)
	}

	// String concatenation
	if av.Kind() == reflect.String && bv.Kind() == reflect.String {
		return av.String() + bv.String()
	}

	return 0
}

// sub returns a - b
// Usage: {{sub .Total .Discount}}
func sub(a, b any) any {
	av := reflect.ValueOf(a)
	bv := reflect.ValueOf(b)

	if isNumeric(av) && isNumeric(bv) {
		if isInteger(av) && isInteger(bv) {
			return toInt64(av) - toInt64(bv)
		}
		return toFloat64(av) - toFloat64(bv)
	}
	return 0
}

// mul returns a * b
// Usage: {{mul .Price .Quantity}}
func mul(a, b any) any {
	av := reflect.ValueOf(a)
	bv := reflect.ValueOf(b)

	if isNumeric(av) && isNumeric(bv) {
		if isInteger(av) && isInteger(bv) {
			return toInt64(av) * toInt64(bv)
		}
		return toFloat64(av) * toFloat64(bv)
	}
	return 0
}

// div returns a / b
// Usage: {{div .Total .Count}}
func div(a, b any) any {
	av := reflect.ValueOf(a)
	bv := reflect.ValueOf(b)

	if isNumeric(av) && isNumeric(bv) {
		bf := toFloat64(bv)
		if bf == 0 {
			return 0 // Avoid division by zero
		}
		if isInteger(av) && isInteger(bv) {
			return toInt64(av) / toInt64(bv)
		}
		return toFloat64(av) / bf
	}
	return 0
}

// mod returns a % b
// Usage: {{mod .Index 2}}
func mod(a, b any) int64 {
	av := reflect.ValueOf(a)
	bv := reflect.ValueOf(b)

	if isNumeric(av) && isNumeric(bv) {
		bi := toInt64(bv)
		if bi == 0 {
			return 0
		}
		return toInt64(av) % bi
	}
	return 0
}

func isInteger(v reflect.Value) bool {
	switch v.Kind() {
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64,
		reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		return true
	}
	return false
}

func toInt64(v reflect.Value) int64 {
	switch v.Kind() {
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		return v.Int()
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		return int64(v.Uint())
	case reflect.Float32, reflect.Float64:
		return int64(v.Float())
	}
	return 0
}

func toInt(v any) int {
	rv := reflect.ValueOf(v)
	switch rv.Kind() {
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		return int(rv.Int())
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		return int(rv.Uint())
	case reflect.Float32, reflect.Float64:
		return int(rv.Float())
	}
	return 0
}

// ============================================================================
// Number Formatting Functions
// ============================================================================

// formatNumber formats a number with thousand separators and decimal places
// Usage: {{formatNumber 1234567.89 2}} -> "1,234,567.89"
func formatNumber(n any, decimals ...int) string {
	dec := 2
	if len(decimals) > 0 {
		dec = decimals[0]
	}

	f := toFloat64(reflect.ValueOf(n))
	p := message.NewPrinter(language.English)
	format := fmt.Sprintf("%%.%df", dec)
	return p.Sprintf(format, f)
}

// formatMoney formats a number as currency
// Usage: {{formatMoney 1234.5 "$"}} -> "$1,234.50"
// Usage: {{formatMoney 1234.5 "€" 2}} -> "€1,234.50"
func formatMoney(n any, symbol string, decimals ...int) string {
	dec := 2
	if len(decimals) > 0 {
		dec = decimals[0]
	}
	return symbol + formatNumber(n, dec)
}

// formatPercent formats a number as percentage
// Usage: {{formatPercent 0.156 1}} -> "15.6%"
func formatPercent(n any, decimals ...int) string {
	dec := 1
	if len(decimals) > 0 {
		dec = decimals[0]
	}
	f := toFloat64(reflect.ValueOf(n)) * 100
	format := fmt.Sprintf("%%.%df%%%%", dec)
	return fmt.Sprintf(format, f)
}

// ============================================================================
// Date/Time Functions
// ============================================================================

// now returns the current time
// Usage: {{now}} or {{formatDate now "2006-01-02"}}
func now() time.Time {
	return time.Now()
}

// formatDate formats a time value using Go's time format layout
// Usage: {{formatDate .Date "January 2, 2006"}}
// Common layouts:
//   - "2006-01-02" -> 2024-01-15
//   - "01/02/2006" -> 01/15/2024
//   - "January 2, 2006" -> January 15, 2024
//   - "Mon, 02 Jan 2006" -> Mon, 15 Jan 2024
//   - "3:04 PM" -> 3:04 PM
//   - "15:04:05" -> 15:04:05
func formatDate(t any, layout string) string {
	switch v := t.(type) {
	case time.Time:
		return v.Format(layout)
	case *time.Time:
		if v == nil {
			return ""
		}
		return v.Format(layout)
	case string:
		// Try to parse common formats
		parsed, err := parseTimeString(v)
		if err != nil {
			return v
		}
		return parsed.Format(layout)
	default:
		return fmt.Sprintf("%v", t)
	}
}

// parseDate parses a date string into a time.Time
// Usage: {{parseDate "2024-01-15" "2006-01-02"}}
func parseDate(dateStr, layout string) time.Time {
	t, err := time.Parse(layout, dateStr)
	if err != nil {
		return time.Time{}
	}
	return t
}

// parseTimeString tries to parse a time string using common formats
func parseTimeString(s string) (time.Time, error) {
	formats := []string{
		time.RFC3339,
		"2006-01-02T15:04:05",
		"2006-01-02 15:04:05",
		"2006-01-02",
		"01/02/2006",
		"02/01/2006",
		"January 2, 2006",
		"Jan 2, 2006",
	}
	for _, format := range formats {
		if t, err := time.Parse(format, s); err == nil {
			return t, nil
		}
	}
	return time.Time{}, fmt.Errorf("unable to parse time: %s", s)
}

// addDays adds days to a time
// Usage: {{addDays .Date 7}}
func addDays(t any, days int) time.Time {
	tm := toTime(t)
	return tm.AddDate(0, 0, days)
}

// addMonths adds months to a time
// Usage: {{addMonths .Date 1}}
func addMonths(t any, months int) time.Time {
	tm := toTime(t)
	return tm.AddDate(0, months, 0)
}

// addYears adds years to a time
// Usage: {{addYears .Date 1}}
func addYears(t any, years int) time.Time {
	tm := toTime(t)
	return tm.AddDate(years, 0, 0)
}

func toTime(t any) time.Time {
	switch v := t.(type) {
	case time.Time:
		return v
	case *time.Time:
		if v == nil {
			return time.Time{}
		}
		return *v
	case string:
		parsed, _ := parseTimeString(v)
		return parsed
	default:
		return time.Time{}
	}
}

// ============================================================================
// Document Structure Functions
// ============================================================================

// pageBreak inserts a page break
// Usage: {{pageBreak}}
func pageBreak() string {
	return `</w:t></w:r></w:p><w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p><w:r><w:t>`
}

// sectionBreak inserts a section break (continuous)
// Usage: {{sectionBreak}}
func sectionBreak() string {
	return `</w:t></w:r></w:p><w:p><w:pPr><w:sectPr><w:type w:val="continuous"/></w:sectPr></w:pPr></w:p><w:p><w:r><w:t>`
}

// ============================================================================
// Additional Utility Functions
// ============================================================================

// generateUUID generates a new UUID
// Usage: {{uuid}}
func generateUUID() string {
	return uuid.New().String()
}

// pluralize returns singular or plural form based on count
// Usage: {{pluralize 1 "item" "items"}} -> "item"
// Usage: {{pluralize 5 "item" "items"}} -> "items"
func pluralize(count any, singular, plural string) string {
	n := toInt(count)
	if n == 1 || n == -1 {
		return singular
	}
	return plural
}

// truncate truncates text to a maximum length, adding ellipsis if truncated
// Usage: {{truncate .Description 50}}
// Usage: {{truncate .Description 50 "..."}}
func truncate(text string, length int, suffix ...string) string {
	ellipsis := "..."
	if len(suffix) > 0 {
		ellipsis = suffix[0]
	}

	runes := []rune(text)
	if len(runes) <= length {
		return text
	}

	return string(runes[:length]) + ellipsis
}

// wordwrap wraps text at word boundaries to fit within a maximum width
// Usage: {{wordwrap .Text 80}}
func wordwrap(text string, width int) string {
	if width <= 0 {
		return text
	}

	var result strings.Builder
	var lineLen int

	words := strings.Fields(text)
	for i, word := range words {
		wordLen := len([]rune(word))

		if lineLen+wordLen > width && lineLen > 0 {
			result.WriteString("\n")
			lineLen = 0
		} else if i > 0 && lineLen > 0 {
			result.WriteString(" ")
			lineLen++
		}

		result.WriteString(word)
		lineLen += wordLen
	}

	return result.String()
}

// capitalize capitalizes the first letter of a string
// Usage: {{capitalize .name}} -> "John"
func capitalize(s string) string {
	if s == "" {
		return s
	}
	runes := []rune(s)
	runes[0] = unicode.ToUpper(runes[0])
	return string(runes)
}

// camelCase converts a string to camelCase
// Usage: {{camelCase "hello world"}} -> "helloWorld"
func camelCase(s string) string {
	words := strings.Fields(s)
	if len(words) == 0 {
		return ""
	}

	var result strings.Builder
	for i, word := range words {
		if i == 0 {
			result.WriteString(strings.ToLower(word))
		} else {
			result.WriteString(strings.Title(strings.ToLower(word)))
		}
	}
	return result.String()
}

// snakeCase converts a string to snake_case
// Usage: {{snakeCase "Hello World"}} -> "hello_world"
func snakeCase(s string) string {
	words := strings.Fields(s)
	for i, word := range words {
		words[i] = strings.ToLower(word)
	}
	return strings.Join(words, "_")
}

// kebabCase converts a string to kebab-case
// Usage: {{kebabCase "Hello World"}} -> "hello-world"
func kebabCase(s string) string {
	words := strings.Fields(s)
	for i, word := range words {
		words[i] = strings.ToLower(word)
	}
	return strings.Join(words, "-")
}
