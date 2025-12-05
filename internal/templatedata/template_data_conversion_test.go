package templatedata

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestDataToMap(t *testing.T) {
	t.Run("Passing in nil should return an error", func(t *testing.T) {
		assert := assert.New(t)

		outputMap, err := DataToMap(nil)
		assert.Nil(outputMap)
		assert.NotNil(err)
	})

	t.Run("Passing in a map should return a copy of the map", func(t *testing.T) {
		assert := assert.New(t)

		inputMap := map[string]any{
			"test": 1,
		}
		outputMap, err := DataToMap(inputMap)
		assert.Equal(inputMap, outputMap)
		assert.Nil(err)
	})

	t.Run("Basic struct", func(t *testing.T) {
		assert := assert.New(t)

		data := struct {
			ProjectNumber string
			Client        string
			Status        string
		}{
			ProjectNumber: "B-00001",
			Client:        "TW Software",
			Status:        "New",
		}
		outputMap, err := DataToMap(data)
		assert.Equal(map[string]any{
			"ProjectNumber": "B-00001",
			"Client":        "TW Software",
			"Status":        "New",
		}, outputMap)
		assert.Nil(err)
	})

	t.Run("Struct with nested structs and slices", func(t *testing.T) {
		assert := assert.New(t)

		data := struct {
			ProjectNumber string
			Client        string
			Status        string
			ExtraFields   struct {
				Field1 string
				Field2 string
			}
			People []struct {
				Name   string
				Gender string
				Age    uint8
			}
		}{
			ProjectNumber: "B-00001",
			Client:        "TW Software",
			Status:        "New",
			ExtraFields: struct {
				Field1 string
				Field2 string
			}{
				Field1: "Value 1",
				Field2: "Value 2",
			},
			People: []struct {
				Name   string
				Gender string
				Age    uint8
			}{
				{
					Name:   "Tom Watkins",
					Gender: "Male",
					Age:    30,
				},
				{
					Name:   "Evie Argyle",
					Gender: "Female",
					Age:    29,
				},
			},
		}
		outputMap, err := DataToMap(data)
		assert.Equal(map[string]any{
			"ProjectNumber": "B-00001",
			"Client":        "TW Software",
			"Status":        "New",
			"ExtraFields": map[string]any{
				"Field1": "Value 1",
				"Field2": "Value 2",
			},
			"People": []map[string]any{
				{
					"Name":   "Tom Watkins",
					"Gender": "Male",
					"Age":    uint64(30), // uint8 converted to uint64
				},
				{
					"Name":   "Evie Argyle",
					"Gender": "Female",
					"Age":    uint64(29),
				},
			},
		}, outputMap)
		assert.Nil(err)
	})

	t.Run("Pointer to a struct", func(t *testing.T) {
		assert := assert.New(t)

		data := struct {
			ProjectNumber string
			Client        string
			Status        string
		}{
			ProjectNumber: "B-00001",
			Client:        "TW Software",
			Status:        "New",
		}
		outputMap, err := DataToMap(&data)
		assert.Equal(map[string]any{
			"ProjectNumber": "B-00001",
			"Client":        "TW Software",
			"Status":        "New",
		}, outputMap)
		assert.Nil(err)
	})

	t.Run("Passing in a non struct value should return error", func(t *testing.T) {
		assert := assert.New(t)

		outputMap, err := DataToMap("string")
		assert.Nil(outputMap)
		assert.NotNil(t, err)
	})

	t.Run("Slice of strings", func(t *testing.T) {
		assert := assert.New(t)

		data := struct {
			Tags []string
		}{
			Tags: []string{"go", "docx", "template"},
		}
		outputMap, err := DataToMap(data)
		assert.Nil(err)
		assert.Equal([]any{"go", "docx", "template"}, outputMap["Tags"])
	})

	t.Run("Slice of integers", func(t *testing.T) {
		assert := assert.New(t)

		data := struct {
			Numbers []int
		}{
			Numbers: []int{1, 2, 3},
		}
		outputMap, err := DataToMap(data)
		assert.Nil(err)
		assert.Equal([]any{int64(1), int64(2), int64(3)}, outputMap["Numbers"])
	})

	t.Run("Map field", func(t *testing.T) {
		assert := assert.New(t)

		data := struct {
			Metadata map[string]string
		}{
			Metadata: map[string]string{
				"author": "John",
				"date":   "2024-01-01",
			},
		}
		outputMap, err := DataToMap(data)
		assert.Nil(err)
		metadata := outputMap["Metadata"].(map[string]any)
		assert.Equal("John", metadata["author"])
		assert.Equal("2024-01-01", metadata["date"])
	})

	t.Run("Pointer fields", func(t *testing.T) {
		assert := assert.New(t)

		name := "Tom"
		data := struct {
			Name *string
		}{
			Name: &name,
		}
		outputMap, err := DataToMap(data)
		assert.Nil(err)
		assert.Equal("Tom", outputMap["Name"])
	})

	t.Run("Nil pointer field", func(t *testing.T) {
		assert := assert.New(t)

		data := struct {
			Name *string
		}{
			Name: nil,
		}
		outputMap, err := DataToMap(data)
		assert.Nil(err)
		assert.Nil(outputMap["Name"])
	})

	t.Run("Slice of pointers to structs", func(t *testing.T) {
		assert := assert.New(t)

		type Person struct {
			Name string
		}
		data := struct {
			People []*Person
		}{
			People: []*Person{
				{Name: "Alice"},
				{Name: "Bob"},
			},
		}
		outputMap, err := DataToMap(data)
		assert.Nil(err)
		people := outputMap["People"].([]map[string]any)
		assert.Len(people, 2)
		assert.Equal("Alice", people[0]["Name"])
		assert.Equal("Bob", people[1]["Name"])
	})
}
