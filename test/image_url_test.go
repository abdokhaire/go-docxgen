package docxtpl_test

import (
	"testing"
	"time"

	"github.com/abdokhaire/go-docxgen"
	"github.com/stretchr/testify/assert"
)

func TestCreateInlineImageFromURL_InvalidExtension(t *testing.T) {
	assert := assert.New(t)

	// Test URL without image extension
	_, err := docxtpl.CreateInlineImageFromURL("https://example.com/file.txt")
	assert.NotNil(err)
	assert.Contains(err.Error(), "supported image format")
}

func TestCreateInlineImageFromURL_InvalidURL(t *testing.T) {
	assert := assert.New(t)

	// Test with invalid/unreachable URL
	_, err := docxtpl.CreateInlineImageFromURLWithTimeout("https://invalid-url-that-does-not-exist-12345.com/image.png", 2*time.Second)
	assert.NotNil(err)
	assert.Contains(err.Error(), "failed to download")
}

func TestCreateInlineImageFromBytes_ValidPNG(t *testing.T) {
	assert := assert.New(t)

	// Minimal valid PNG header
	pngData := []byte{0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x00}

	img, err := docxtpl.CreateInlineImageFromBytes(pngData, ".png")
	assert.Nil(err)
	assert.NotNil(img)
}

func TestCreateInlineImageFromBytes_ValidJPG(t *testing.T) {
	assert := assert.New(t)

	// Minimal valid JPEG header
	jpgData := []byte{0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 0x4A, 0x46, 0x49, 0x46}

	img, err := docxtpl.CreateInlineImageFromBytes(jpgData, ".jpg")
	assert.Nil(err)
	assert.NotNil(img)
}

func TestCreateInlineImageFromBytes_InvalidData(t *testing.T) {
	assert := assert.New(t)

	// Invalid image data
	invalidData := []byte{0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07}

	_, err := docxtpl.CreateInlineImageFromBytes(invalidData, ".png")
	assert.NotNil(err)
	assert.Contains(err.Error(), "not a valid image")
}

func TestCreateInlineImageFromBytes_InvalidExtension(t *testing.T) {
	assert := assert.New(t)

	// Valid PNG data but wrong extension
	pngData := []byte{0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x00}

	_, err := docxtpl.CreateInlineImageFromBytes(pngData, ".gif")
	assert.NotNil(err)
	assert.Contains(err.Error(), "unsupported image format")
}
