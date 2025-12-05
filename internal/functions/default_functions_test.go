package functions

import (
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
)

func TestTitle(t *testing.T) {
	titleCase := title("tom watkins")
	if titleCase != "Tom Watkins" {
		t.Fatalf("Should return in title case but returned: %v", titleCase)
	}
}

func TestComparisonFunctions(t *testing.T) {
	t.Run("eq - equal values", func(t *testing.T) {
		assert.True(t, eq(5, 5))
		assert.True(t, eq("hello", "hello"))
		assert.False(t, eq(5, 10))
		assert.False(t, eq("hello", "world"))
	})

	t.Run("ne - not equal values", func(t *testing.T) {
		assert.True(t, ne(5, 10))
		assert.True(t, ne("hello", "world"))
		assert.False(t, ne(5, 5))
	})

	t.Run("lt - less than", func(t *testing.T) {
		assert.True(t, lt(5, 10))
		assert.False(t, lt(10, 5))
		assert.False(t, lt(5, 5))
		assert.True(t, lt("a", "b"))
	})

	t.Run("le - less than or equal", func(t *testing.T) {
		assert.True(t, le(5, 10))
		assert.True(t, le(5, 5))
		assert.False(t, le(10, 5))
	})

	t.Run("gt - greater than", func(t *testing.T) {
		assert.True(t, gt(10, 5))
		assert.False(t, gt(5, 10))
		assert.False(t, gt(5, 5))
	})

	t.Run("ge - greater than or equal", func(t *testing.T) {
		assert.True(t, ge(10, 5))
		assert.True(t, ge(5, 5))
		assert.False(t, ge(5, 10))
	})
}

func TestLogicalFunctions(t *testing.T) {
	t.Run("and - all truthy", func(t *testing.T) {
		assert.True(t, and(true, true, true))
		assert.False(t, and(true, false, true))
		assert.True(t, and(1, "hello", true))
		assert.False(t, and(0, "hello"))
	})

	t.Run("or - any truthy", func(t *testing.T) {
		assert.True(t, or(false, true, false))
		assert.True(t, or(0, "", 1))
		assert.False(t, or(false, 0, ""))
	})

	t.Run("not - negation", func(t *testing.T) {
		assert.True(t, not(false))
		assert.True(t, not(0))
		assert.True(t, not(""))
		assert.False(t, not(true))
		assert.False(t, not(1))
	})
}

func TestCollectionFunctions(t *testing.T) {
	t.Run("length", func(t *testing.T) {
		assert.Equal(t, 3, length([]int{1, 2, 3}))
		assert.Equal(t, 5, length("hello"))
		assert.Equal(t, 2, length(map[string]int{"a": 1, "b": 2}))
		assert.Equal(t, 0, length(nil))
	})

	t.Run("first", func(t *testing.T) {
		assert.Equal(t, 1, first([]int{1, 2, 3}))
		assert.Equal(t, "a", first([]string{"a", "b", "c"}))
		assert.Nil(t, first([]int{}))
		assert.Nil(t, first(nil))
	})

	t.Run("last", func(t *testing.T) {
		assert.Equal(t, 3, last([]int{1, 2, 3}))
		assert.Equal(t, "c", last([]string{"a", "b", "c"}))
		assert.Nil(t, last([]int{}))
	})

	t.Run("index", func(t *testing.T) {
		assert.Equal(t, 2, index([]int{1, 2, 3}, 1))
		assert.Equal(t, "b", index([]string{"a", "b", "c"}, 1))
		assert.Nil(t, index([]int{1, 2, 3}, 10))
		assert.Equal(t, "e", index("hello", 1))
	})

	t.Run("slice", func(t *testing.T) {
		arr := []int{1, 2, 3, 4, 5}
		assert.Equal(t, []int{2, 3}, sliceFn(arr, 1, 3))
		assert.Equal(t, []int{3, 4, 5}, sliceFn(arr, 2))
	})

	t.Run("join", func(t *testing.T) {
		assert.Equal(t, "a, b, c", join([]string{"a", "b", "c"}, ", "))
		assert.Equal(t, "1-2-3", join([]int{1, 2, 3}, "-"))
	})

	t.Run("contains", func(t *testing.T) {
		assert.True(t, contains([]int{1, 2, 3}, 2))
		assert.False(t, contains([]int{1, 2, 3}, 5))
		assert.True(t, contains("hello world", "world"))
		assert.False(t, contains("hello", "world"))
		assert.True(t, contains(map[string]int{"a": 1}, "a"))
	})
}

func TestUtilityFunctions(t *testing.T) {
	t.Run("default", func(t *testing.T) {
		assert.Equal(t, "default", defaultVal("default", ""))
		assert.Equal(t, "value", defaultVal("default", "value"))
		assert.Equal(t, "default", defaultVal("default", nil))
	})

	t.Run("coalesce", func(t *testing.T) {
		assert.Equal(t, "first", coalesce("first", "second"))
		assert.Equal(t, "second", coalesce("", "second", "third"))
		assert.Equal(t, 1, coalesce(nil, 0, 1))
	})

	t.Run("ternary", func(t *testing.T) {
		assert.Equal(t, "yes", ternary("yes", "no", true))
		assert.Equal(t, "no", ternary("yes", "no", false))
	})

	t.Run("split", func(t *testing.T) {
		assert.Equal(t, []string{"a", "b", "c"}, split("a,b,c", ","))
		assert.Equal(t, []string{}, split("", ","))
	})

	t.Run("concat", func(t *testing.T) {
		assert.Equal(t, "hello world", concat("hello", " ", "world"))
		assert.Equal(t, "123", concat(1, 2, 3))
	})
}

func TestMathFunctions(t *testing.T) {
	t.Run("add", func(t *testing.T) {
		assert.Equal(t, int64(5), add(2, 3))
		assert.Equal(t, 5.5, add(2.5, 3.0))
		assert.Equal(t, "ab", add("a", "b"))
	})

	t.Run("sub", func(t *testing.T) {
		assert.Equal(t, int64(2), sub(5, 3))
		assert.Equal(t, 2.5, sub(5.5, 3.0))
	})

	t.Run("mul", func(t *testing.T) {
		assert.Equal(t, int64(6), mul(2, 3))
		assert.Equal(t, 7.5, mul(2.5, 3.0))
	})

	t.Run("div", func(t *testing.T) {
		assert.Equal(t, int64(2), div(6, 3))
		assert.Equal(t, 2.5, div(5.0, 2.0))
		assert.Equal(t, 0, div(5, 0)) // Division by zero returns 0
	})

	t.Run("mod", func(t *testing.T) {
		assert.Equal(t, int64(1), mod(5, 2))
		assert.Equal(t, int64(0), mod(6, 3))
		assert.Equal(t, int64(0), mod(5, 0)) // Mod by zero returns 0
	})
}

func TestRichTextFunctions(t *testing.T) {
	t.Run("bold", func(t *testing.T) {
		result := bold("test")
		assert.Contains(t, result, "<w:b/>")
		assert.Contains(t, result, "test")
	})

	t.Run("italic", func(t *testing.T) {
		result := italic("test")
		assert.Contains(t, result, "<w:i/>")
		assert.Contains(t, result, "test")
	})

	t.Run("underline", func(t *testing.T) {
		result := underline("test")
		assert.Contains(t, result, `<w:u w:val="single"/>`)
		assert.Contains(t, result, "test")
	})

	t.Run("color", func(t *testing.T) {
		result := color("FF0000", "test")
		assert.Contains(t, result, `w:val="FF0000"`)
		assert.Contains(t, result, "test")
	})

	t.Run("highlight", func(t *testing.T) {
		result := highlight("yellow", "test")
		assert.Contains(t, result, `w:val="yellow"`)
		assert.Contains(t, result, "test")
	})
}

func TestNumberFormatting(t *testing.T) {
	t.Run("formatNumber", func(t *testing.T) {
		assert.Equal(t, "1,234.57", formatNumber(1234.567, 2))
		assert.Equal(t, "1,234.6", formatNumber(1234.567, 1))
		assert.Equal(t, "1,000,000.00", formatNumber(1000000, 2))
	})

	t.Run("formatMoney", func(t *testing.T) {
		assert.Equal(t, "$1,234.50", formatMoney(1234.5, "$"))
		assert.Equal(t, "€1,234.50", formatMoney(1234.5, "€"))
		assert.Equal(t, "$1,234.5", formatMoney(1234.5, "$", 1))
	})

	t.Run("formatPercent", func(t *testing.T) {
		assert.Equal(t, "15.6%", formatPercent(0.156, 1))
		assert.Equal(t, "50.00%", formatPercent(0.5, 2))
		assert.Equal(t, "100.0%", formatPercent(1.0, 1))
	})
}

func TestDateFunctions(t *testing.T) {
	t.Run("now", func(t *testing.T) {
		result := now()
		assert.False(t, result.IsZero())
	})

	t.Run("formatDate with time.Time", func(t *testing.T) {
		testDate := time.Date(2024, 1, 15, 14, 30, 0, 0, time.UTC)
		assert.Equal(t, "2024-01-15", formatDate(testDate, "2006-01-02"))
		assert.Equal(t, "January 15, 2024", formatDate(testDate, "January 2, 2006"))
		assert.Equal(t, "01/15/2024", formatDate(testDate, "01/02/2006"))
	})

	t.Run("formatDate with string", func(t *testing.T) {
		assert.Equal(t, "January 15, 2024", formatDate("2024-01-15", "January 2, 2006"))
	})

	t.Run("parseDate", func(t *testing.T) {
		result := parseDate("2024-01-15", "2006-01-02")
		assert.Equal(t, 2024, result.Year())
		assert.Equal(t, time.January, result.Month())
		assert.Equal(t, 15, result.Day())
	})

	t.Run("addDays", func(t *testing.T) {
		testDate := time.Date(2024, 1, 15, 0, 0, 0, 0, time.UTC)
		result := addDays(testDate, 7)
		assert.Equal(t, 22, result.Day())
	})

	t.Run("addMonths", func(t *testing.T) {
		testDate := time.Date(2024, 1, 15, 0, 0, 0, 0, time.UTC)
		result := addMonths(testDate, 2)
		assert.Equal(t, time.March, result.Month())
	})

	t.Run("addYears", func(t *testing.T) {
		testDate := time.Date(2024, 1, 15, 0, 0, 0, 0, time.UTC)
		result := addYears(testDate, 1)
		assert.Equal(t, 2025, result.Year())
	})
}

func TestDocumentStructureFunctions(t *testing.T) {
	t.Run("pageBreak", func(t *testing.T) {
		result := pageBreak()
		assert.Contains(t, result, `<w:br w:type="page"/>`)
	})

	t.Run("sectionBreak", func(t *testing.T) {
		result := sectionBreak()
		assert.Contains(t, result, `<w:sectPr>`)
	})
}

func TestAdditionalUtilityFunctions(t *testing.T) {
	t.Run("uuid", func(t *testing.T) {
		result := generateUUID()
		assert.NotEmpty(t, result)
		assert.Len(t, result, 36) // UUID format: 8-4-4-4-12
	})

	t.Run("pluralize", func(t *testing.T) {
		assert.Equal(t, "item", pluralize(1, "item", "items"))
		assert.Equal(t, "items", pluralize(0, "item", "items"))
		assert.Equal(t, "items", pluralize(5, "item", "items"))
		assert.Equal(t, "item", pluralize(-1, "item", "items"))
	})

	t.Run("truncate", func(t *testing.T) {
		assert.Equal(t, "Hello...", truncate("Hello World", 5))
		assert.Equal(t, "Hello World", truncate("Hello World", 20))
		assert.Equal(t, "Hello--", truncate("Hello World", 5, "--"))
	})

	t.Run("wordwrap", func(t *testing.T) {
		result := wordwrap("This is a long sentence that needs wrapping", 10)
		assert.Contains(t, result, "\n")
	})

	t.Run("capitalize", func(t *testing.T) {
		assert.Equal(t, "Hello", capitalize("hello"))
		assert.Equal(t, "Hello", capitalize("Hello"))
		assert.Equal(t, "", capitalize(""))
	})

	t.Run("camelCase", func(t *testing.T) {
		assert.Equal(t, "helloWorld", camelCase("hello world"))
		assert.Equal(t, "helloWorldTest", camelCase("Hello World Test"))
	})

	t.Run("snakeCase", func(t *testing.T) {
		assert.Equal(t, "hello_world", snakeCase("Hello World"))
	})

	t.Run("kebabCase", func(t *testing.T) {
		assert.Equal(t, "hello-world", kebabCase("Hello World"))
	})
}
