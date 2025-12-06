package templatedata

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestIsFilePath(t *testing.T) {
	tests := []struct {
		name           string
		filepath       string
		expectedResult bool
	}{
		{
			name:           "Existing file",
			filepath:       "../../test/testdata/templates/test_image.png",
			expectedResult: true,
		},
		{
			name:           "Non existent file",
			filepath:       "../../test/testdata/templates/not_exists.docx",
			expectedResult: false,
		},
		{
			name:           "Exists but is a folder",
			filepath:       "../../test/testdata/templates",
			expectedResult: false,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			assert := assert.New(t)

			result, err := IsFilePath(tt.filepath)
			assert.Nil(err)
			assert.Equal(result, tt.expectedResult)
		})
	}
}

func TestIsImageFilePath(t *testing.T) {
	tests := []struct {
		name           string
		filepath       string
		expectedResult bool
	}{
		{
			name:           "Existing image",
			filepath:       "../../test/testdata/templates/test_image.png",
			expectedResult: true,
		},
		{
			name:           "File exists but isn't ab image",
			filepath:       "../../test/testdata/templates/test_basic.docx",
			expectedResult: false,
		},
		{
			name:           "Missing file extension",
			filepath:       "testdata/templates/test_image",
			expectedResult: false,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			assert := assert.New(t)

			result, err := IsImageFilePath(tt.filepath)
			assert.Nil(err)
			assert.Equal(result, tt.expectedResult)
		})
	}
}
