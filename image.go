package docxtpl

import (
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"image"
	"image/jpeg"
	"image/png"
	"io"
	"net/http"
	"os"
	"path"
	"regexp"
	"strconv"
	"strings"
	"time"

	"github.com/bep/imagemeta"
	"github.com/abdokhaire/go-docxgen/internal/docx"
	"github.com/fumiama/imgsz"
	"github.com/abdokhaire/go-docxgen/internal/contenttypes"
	"github.com/abdokhaire/go-docxgen/internal/templatedata"
	"golang.org/x/image/draw"
)

const (
	EMUS_PER_INCH = 914400
	DEFAULT_DPI   = 72
)

type InlineImage struct {
	data *[]byte
	Ext  string
}

type InlineImageError struct {
	Message string
}

func (e *InlineImageError) Error() string {
	return fmt.Sprintf("Image error: %v", e.Message)
}

// Take a filenane for an image and return a pointer to an InlineImage struct.
// Images can be Jpegs (.jpg or .jpeg) or PNGs
//
//	img, err := CreateInlineImage("example_img.png")
func CreateInlineImage(filepath string) (*InlineImage, error) {
	if isImage, err := templatedata.IsImageFilePath(filepath); err != nil {
		return nil, err
	} else {
		if !isImage {
			return nil, &InlineImageError{"File is not a valid image"}
		}
	}

	file, err := os.ReadFile(filepath)
	if err != nil {
		return nil, err
	}

	ext := path.Ext(filepath)

	return &InlineImage{&file, ext}, nil
}

// CreateInlineImageFromURL downloads an image from a URL and returns an InlineImage.
// Images can be Jpegs (.jpg or .jpeg) or PNGs.
// Timeout defaults to 30 seconds.
//
//	img, err := CreateInlineImageFromURL("https://example.com/image.png")
func CreateInlineImageFromURL(url string) (*InlineImage, error) {
	return CreateInlineImageFromURLWithTimeout(url, 30*time.Second)
}

// CreateInlineImageFromURLWithTimeout downloads an image from a URL with a custom timeout.
//
//	img, err := CreateInlineImageFromURLWithTimeout("https://example.com/image.png", 10*time.Second)
func CreateInlineImageFromURLWithTimeout(url string, timeout time.Duration) (*InlineImage, error) {
	// Validate URL has image extension
	ext := getExtensionFromURL(url)
	if ext == "" {
		return nil, &InlineImageError{"URL does not point to a supported image format (jpg, jpeg, png)"}
	}

	// Create HTTP client with timeout
	client := &http.Client{
		Timeout: timeout,
	}

	// Download the image
	resp, err := client.Get(url)
	if err != nil {
		return nil, &InlineImageError{fmt.Sprintf("failed to download image: %v", err)}
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		return nil, &InlineImageError{fmt.Sprintf("failed to download image: HTTP %d", resp.StatusCode)}
	}

	// Read the image data
	data, err := io.ReadAll(resp.Body)
	if err != nil {
		return nil, &InlineImageError{fmt.Sprintf("failed to read image data: %v", err)}
	}

	// Validate it's actually an image by checking the header
	contentType := resp.Header.Get("Content-Type")
	if !isValidImageContentType(contentType) && !isValidImageData(data) {
		return nil, &InlineImageError{"downloaded content is not a valid image"}
	}

	return &InlineImage{&data, ext}, nil
}

// CreateInlineImageFromBytes creates an InlineImage from raw bytes.
// You must specify the extension (.jpg, .jpeg, or .png).
//
//	img, err := CreateInlineImageFromBytes(imageData, ".png")
func CreateInlineImageFromBytes(data []byte, ext string) (*InlineImage, error) {
	ext = strings.ToLower(ext)
	if ext != ".jpg" && ext != ".jpeg" && ext != ".png" {
		return nil, &InlineImageError{"unsupported image format: must be .jpg, .jpeg, or .png"}
	}

	if !isValidImageData(data) {
		return nil, &InlineImageError{"data is not a valid image"}
	}

	return &InlineImage{&data, ext}, nil
}

// getExtensionFromURL extracts the image extension from a URL.
func getExtensionFromURL(url string) string {
	// Remove query string and fragment
	url = strings.Split(url, "?")[0]
	url = strings.Split(url, "#")[0]

	ext := strings.ToLower(path.Ext(url))
	switch ext {
	case ".jpg", ".jpeg", ".png":
		return ext
	default:
		return ""
	}
}

// isValidImageContentType checks if the content type is a supported image type.
func isValidImageContentType(contentType string) bool {
	contentType = strings.ToLower(contentType)
	return strings.Contains(contentType, "image/jpeg") ||
		strings.Contains(contentType, "image/jpg") ||
		strings.Contains(contentType, "image/png")
}

// isValidImageData checks if the data starts with valid image magic bytes.
func isValidImageData(data []byte) bool {
	if len(data) < 8 {
		return false
	}

	// Check for JPEG magic bytes (FFD8FF)
	if data[0] == 0xFF && data[1] == 0xD8 && data[2] == 0xFF {
		return true
	}

	// Check for PNG magic bytes (89504E47)
	if data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47 {
		return true
	}

	return false
}

func (i *InlineImage) getImageFormat() (imagemeta.ImageFormat, error) {
	switch i.Ext {
	case ".jpg", ".jpeg":
		return imagemeta.JPEG, nil
	case ".png":
		return imagemeta.PNG, nil
	default:
		return 0, errors.New("Unknown image format: " + i.Ext)
	}
}

// Return a map of EXIF data from the image.
func (i *InlineImage) GetExifData() (map[string]imagemeta.TagInfo, error) {
	var tags imagemeta.Tags
	handleTag := func(ti imagemeta.TagInfo) error {
		tags.Add(ti)
		return nil
	}

	imageFormat, err := i.getImageFormat()
	if err != nil {
		return nil, err
	}

	shouldHandle := func(ti imagemeta.TagInfo) bool {
		return true
	}

	knownWarnings := []*regexp.Regexp{}

	warnf := func(format string, args ...any) {
		s := fmt.Sprintf(format, args...)
		for _, re := range knownWarnings {
			if re.MatchString(s) {
				return
			}
		}
		panic(errors.New(s))
	}

	sources := imagemeta.EXIF

	err = imagemeta.Decode(imagemeta.Options{R: bytes.NewReader(*i.data), ImageFormat: imageFormat, ShouldHandleTag: shouldHandle, HandleTag: handleTag, Warnf: warnf, Sources: sources})
	if err != nil {
		return nil, err
	}

	return tags.EXIF(), nil
}

// Resize the image. Width and height should be pixel values.
func (i *InlineImage) Resize(width int, height int) error {
	src, err := i.getImage()
	if err != nil {
		return err
	}

	// Resize
	rgba := image.NewRGBA(image.Rect(0, 0, width, height))
	draw.NearestNeighbor.Scale(rgba, rgba.Rect, *src, (*src).Bounds(), draw.Over, nil)
	var resizedImage image.Image = rgba

	err = i.replaceImage(&resizedImage)
	if err != nil {
		return err
	}

	return nil
}

func (i *InlineImage) getImage() (*image.Image, error) {
	format, err := i.getImageFormat()
	if err != nil {
		return nil, err
	}

	var img image.Image
	imgReader := bytes.NewReader(*i.data)

	switch format {
	case imagemeta.JPEG:
		img, err = jpeg.Decode(imgReader)
	case imagemeta.PNG:
		img, err = png.Decode(imgReader)
	}

	return &img, err
}

func (i *InlineImage) replaceImage(rgba *image.Image) error {
	format, err := i.getImageFormat()
	if err != nil {
		return err
	}

	var buf bytes.Buffer
	switch format {
	case imagemeta.JPEG:
		err = jpeg.Encode(&buf, *rgba, &jpeg.Options{Quality: 100})
	case imagemeta.PNG:
		err = png.Encode(&buf, *rgba)
	}
	if err != nil {
		return err
	}

	newImageData := buf.Bytes()
	i.data = &newImageData

	return nil
}

// Get the size of the image in EMUs.
func (i *InlineImage) GetSize() (w int64, h int64, err error) {
	sz, _, err := imgsz.DecodeSize(bytes.NewReader(*i.data))
	if err != nil {
		return 0, 0, nil
	}

	wDpi, hDpi := i.GetResolution()

	w = int64(sz.Width/wDpi) * int64(EMUS_PER_INCH)
	h = int64(sz.Height/hDpi) * int64(EMUS_PER_INCH)

	return w, h, nil
}

// Get the resolution (DPI) of the image.
// It gets this from EXIF data and defaults to 72 if not found.
func (i *InlineImage) GetResolution() (wDpi int, hDpi int) {
	exif, err := i.GetExifData()
	if err != nil {
		return DEFAULT_DPI, DEFAULT_DPI
	}

	getResolution := func(tagName string) int {
		resolutionTag, exists := exif[tagName]
		if exists {
			if value, ok := resolutionTag.Value.(string); ok {
				resolution, err := getResolutionFromString(value)
				if err != nil || resolution == 0 {
					return DEFAULT_DPI
				}
				return resolution
			}
		}
		return DEFAULT_DPI
	}

	wDpi, hDpi = getResolution("XResolution"), getResolution("YResolution")

	return wDpi, hDpi
}

func getResolutionFromString(resolution string) (int, error) {
	// Split the string by the slash
	parts := strings.Split(resolution, "/")
	if len(parts) != 2 {
		return 0, errors.New("more than one slash found in image resolution string")
	}

	numerator, err := strconv.Atoi(parts[0])
	if err != nil {
		return 0, err
	}
	denominator, err := strconv.Atoi(parts[1])
	if err != nil {
		return 0, err
	}

	result := numerator / denominator

	return result, nil
}

func (i *InlineImage) getContentTypes() ([]*contenttypes.ContentType, error) {
	format, err := i.getImageFormat()
	if err != nil {
		return nil, err
	}

	switch format {
	case imagemeta.JPEG:
		return []*contenttypes.ContentType{&contenttypes.JPG_CONTENT_TYPE, &contenttypes.JPEG_CONTENT_TYPE}, nil
	case imagemeta.PNG:
		return []*contenttypes.ContentType{&contenttypes.PNG_CONTENT_TYPE}, nil
	}

	return []*contenttypes.ContentType{}, nil
}

func (d *DocxTmpl) addInlineImage(i *InlineImage) (xmlString string, err error) {
	// Add the image to the document (use underlying docx method)
	paragraph := d.Docx.AddParagraph()
	run, err := paragraph.AddInlineDrawing(*i.data)
	if err != nil {
		return "", err
	}

	// Append the content types
	contentTypes, err := i.getContentTypes()
	if err != nil {
		return "", err
	}
	for _, contentType := range contentTypes {
		d.contentTypes.AddContentType(contentType)
	}

	// Correctly size the image
	w, h, err := i.GetSize()
	if err != nil {
		return "", err
	}
	for _, child := range run.Children {
		if drawing, ok := child.(*docx.Drawing); ok {
			drawing.Inline.Extent.CX = w
			drawing.Inline.Extent.CY = h
			break
		}
	}

	// Get the image XML
	out, err := xml.Marshal(run)
	if err != nil {
		return "", err
	}

	// Remove run tags as the tag should be in a run already
	xmlString = string(out)
	xmlString = strings.Replace(xmlString, "<w:r>", "", 1)
	xmlString = strings.Replace(xmlString, "<w:rPr></w:rPr>", "", 1)
	lastIndex := strings.LastIndex(xmlString, "</w:r")
	if lastIndex > -1 {
		xmlString = xmlString[:lastIndex]
	}

	// Remove the paragraph from the word doc so we don't get the image twice
	var newItems []interface{}
	for _, item := range d.Document.Body.Items {
		switch o := item.(type) {
		case *docx.Paragraph:
			if o == paragraph {
				continue
			}
		}
		newItems = append(newItems, item)
	}
	d.Document.Body.Items = newItems

	return xmlString, nil
}
