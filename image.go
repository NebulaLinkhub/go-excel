package go_excel

import (
	"fmt"
	"io"
	"net/http"
	"net/url"
	"os"
)

func ReadFile(path string) ([]byte, error) {
	// 如果路径是URL，尝试从网络读取
	if isValidURL(path) {
		return readImageFromURL(path)
	}

	// 否则，尝试读取本地文件
	return readImageFromFile(path)
}

// 判断路径是否是有效的URL
func isValidURL(path string) bool {
	_, err := url.ParseRequestURI(path)
	return err == nil
}

// 从网络URL读取图片数据
func readImageFromURL(imageURL string) ([]byte, error) {
	// 发送GET请求获取图片
	resp, err := http.Get(imageURL)
	if err != nil {
		return nil, fmt.Errorf("failed to fetch image from URL: %v", err)
	}
	defer resp.Body.Close()

	// 如果状态码不是200，则返回错误
	if resp.StatusCode != http.StatusOK {
		return nil, fmt.Errorf("failed to fetch image, status code: %d", resp.StatusCode)
	}

	// 读取响应的图片内容
	return io.ReadAll(resp.Body)
}

// 从本地文件读取图片数据
func readImageFromFile(imagePath string) ([]byte, error) {
	// 使用 ioutil.ReadFile 读取本地文件
	data, err := os.ReadFile(imagePath)
	if err != nil {
		return nil, fmt.Errorf("failed to read image from file: %v", err)
	}
	return data, nil
}
