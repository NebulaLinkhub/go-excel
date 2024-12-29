package go_excel

import (
	"bytes"
	"errors"
	"fmt"
	"io"
	"mime"
	"net/http"
	"net/url"
	"os"
	"path/filepath"
	"strings"
)

// 支持的图片格式
var supportedImageTypes = map[string]string{
	".bmp": ".bmp", ".emf": ".emf", ".emz": ".emz", ".gif": ".gif",
	".jpeg": ".jpeg", ".jpg": ".jpeg", ".png": ".png", ".svg": ".svg",
	".tif": ".tiff", ".tiff": ".tiff", ".wmf": ".wmf", ".wmz": ".wmz",
}

// 图片格式魔数映射
var magicNumbers = map[string][]byte{
	".png":  {0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A},
	".jpg":  {0xFF, 0xD8, 0xFF},
	".jpeg": {0xFF, 0xD8, 0xFF},
	".gif":  {0x47, 0x49, 0x46, 0x38},
	".tiff": {0x49, 0x49, 0x2A, 0x00},
	".svg":  []byte("<?xml"),
}

type ImageData struct {
	Data      []byte
	Extension string
}

// 从文件扩展名判断MIME类型
func getMimeTypeFromExtension(path string) string {
	ext := strings.ToLower(filepath.Ext(path))
	return mime.TypeByExtension(ext)
}

// 从魔数判断文件类型
func detectFileType(data []byte) string {
	for ext, magic := range magicNumbers {
		if len(data) >= len(magic) && bytes.Equal(data[:len(magic)], magic) {
			return ext
		}
	}
	return ""
}

func ReadFile(path string) (*ImageData, error) {
	if isValidURL(path) {
		return readImageFromURL(path)
	}
	return readImageFromFile(path)
}

// 从网络URL读取图片数据
func readImageFromURL(imageURL string) (*ImageData, error) {
	resp, err := http.Get(imageURL)
	if err != nil {
		return nil, fmt.Errorf("failed to fetch image from URL: %v", err)
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		return nil, fmt.Errorf("failed to fetch image, status code: %d", resp.StatusCode)
	}

	// 读取响应内容
	data, err := io.ReadAll(resp.Body)
	if err != nil {
		return nil, err
	}
	// 如果 Content-Type 不可靠，尝试通过魔数检测
	ext := detectFileType(data)
	if _, ok := supportedImageTypes[ext]; !ok {
		return nil, errors.New("unsupported image format")
	}
	return &ImageData{
		Data:      data,
		Extension: ext,
	}, nil
}

// 从本地文件读取图片数据
func readImageFromFile(imagePath string) (*ImageData, error) {
	data, err := os.ReadFile(imagePath)
	if err != nil {
		return nil, fmt.Errorf("failed to read image from file: %v", err)
	}

	ext := detectFileType(data)
	if _, ok := supportedImageTypes[ext]; !ok {
		return nil, errors.New("unsupported image format")
	}

	return &ImageData{
		Data:      data,
		Extension: filepath.Ext(imagePath),
	}, nil
}

func isValidURL(path string) bool {
	_, err := url.ParseRequestURI(path)
	return err == nil
}
