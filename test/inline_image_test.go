package docxtpl_test

import (
	"testing"

	"github.com/abdokhaire/go-docxgen"
	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

func TestCreateInlineImage(t *testing.T) {
	t.Run("Should return an image for a valid filepath", func(t *testing.T) {
		assert := assert.New(t)

		img, err := docxtpl.CreateInlineImage("testdata/templates/test_image.jpg")
		assert.Nil(err)
		assert.NotNil(img)
		assert.Equal(img.Ext, ".jpg")
	})

	t.Run("Should return error if not a valid image filename", func(t *testing.T) {
		assert := assert.New(t)

		image, err := docxtpl.CreateInlineImage("test_image.txt")
		assert.Nil(image)
		assert.NotNil(err)
	})

	t.Run("Should return error if image doesn't exist", func(t *testing.T) {
		assert := assert.New(t)

		image, err := docxtpl.CreateInlineImage("image_not_exists.png")
		assert.Nil(image)
		assert.NotNil(err)
	})
}

func TestGetExifData(t *testing.T) {
	assert := assert.New(t)
	require := require.New(t)

	img, err := docxtpl.CreateInlineImage("testdata/templates/test_image.jpg")
	require.Nil(err)

	exifData, err := img.GetExifData()
	assert.Nil(err)
	assert.Greater(len(exifData), 0)
}

func TestResize(t *testing.T) {
	require := require.New(t)
	assert := assert.New(t)

	inlineImage, err := docxtpl.CreateInlineImage("testdata/templates/test_image.jpg")
	require.Nil(err)

	originalWEmu, originalHEmu, err := inlineImage.GetSize()
	require.Nil(err)

	wDpi, hDpi := inlineImage.GetResolution()
	newWidthPx := int(originalWEmu/docxtpl.EMUS_PER_INCH) * wDpi * 2
	newHeightPx := int(originalHEmu/docxtpl.EMUS_PER_INCH) * hDpi * 2

	err = inlineImage.Resize(newWidthPx, newHeightPx)
	assert.Nil(err)

	w, h, err := inlineImage.GetSize()
	assert.Nil(err)
	assert.Equal(w, originalWEmu*2)
	assert.Equal(h, originalHEmu*2)
}

func TestGetSize(t *testing.T) {
	require := require.New(t)
	assert := assert.New(t)

	inlineImage, err := docxtpl.CreateInlineImage("testdata/templates/test_image.jpg")
	require.Nil(err)

	w, h, err := inlineImage.GetSize()
	assert.Nil(err)
	assert.Greater(w, int64(0))
	assert.Greater(h, int64(0))
}

func TestGetResolution(t *testing.T) {
	require := require.New(t)
	assert := assert.New(t)

	inlineImage, err := docxtpl.CreateInlineImage("testdata/templates/test_image.jpg")
	require.Nil(err)

	wDpi, hDpi := inlineImage.GetResolution()
	assert.Greater(wDpi, 0)
	assert.Greater(hDpi, 0)
}
