package utils

import (
	"crypto/md5"
	"encoding/hex"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
)

// FileInfo 文件信息
type FileInfo struct {
	Path         string `json:"path"`
	Name         string `json:"name"`
	Extension    string `json:"extension"`
	Size         int64  `json:"size"`
	MD5Hash      string `json:"md5_hash"`
	LastModified string `json:"last_modified"`
}

// GetFileInfo 获取文件信息
func GetFileInfo(filePath string) (*FileInfo, error) {
	file, err := os.Open(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open file: %w", err)
	}
	defer file.Close()

	stat, err := file.Stat()
	if err != nil {
		return nil, fmt.Errorf("failed to get file stats: %w", err)
	}

	// 计算MD5哈希
	hash := md5.New()
	if _, err := io.Copy(hash, file); err != nil {
		return nil, fmt.Errorf("failed to calculate MD5 hash: %w", err)
	}

	return &FileInfo{
		Path:         filePath,
		Name:         stat.Name(),
		Extension:    strings.ToLower(filepath.Ext(filePath)),
		Size:         stat.Size(),
		MD5Hash:      hex.EncodeToString(hash.Sum(nil)),
		LastModified: stat.ModTime().Format("2006-01-02 15:04:05"),
	}, nil
}

// ValidateFile 验证文件是否存在且可读
func ValidateFile(filePath string) error {
	if filePath == "" {
		return fmt.Errorf("file path is empty")
	}

	file, err := os.Open(filePath)
	if err != nil {
		return fmt.Errorf("failed to open file: %w", err)
	}
	defer file.Close()

	stat, err := file.Stat()
	if err != nil {
		return fmt.Errorf("failed to get file stats: %w", err)
	}

	if stat.IsDir() {
		return fmt.Errorf("path is a directory, not a file")
	}

	if stat.Size() == 0 {
		return fmt.Errorf("file is empty")
	}

	return nil
}

// GetSupportedExtensions 获取支持的文档扩展名
func GetSupportedExtensions() []string {
	return []string{
		".docx", ".doc", ".rtf", ".wpd",
		".dot", ".dotx", // 模板格式
	}
}

// IsSupportedFormat 检查文件格式是否支持
func IsSupportedFormat(filePath string) bool {
	ext := strings.ToLower(filepath.Ext(filePath))
	supported := GetSupportedExtensions()

	for _, supportedExt := range supported {
		if ext == supportedExt {
			return true
		}
	}
	return false
}

// CreateOutputPath 创建输出文件路径
func CreateOutputPath(inputPath, suffix string) string {
	dir := filepath.Dir(inputPath)
	name := filepath.Base(inputPath)
	ext := filepath.Ext(name)
	baseName := strings.TrimSuffix(name, ext)

	outputName := baseName + suffix + ext
	return filepath.Join(dir, outputName)
}

// EnsureDirectoryExists 确保目录存在
func EnsureDirectoryExists(dirPath string) error {
	if err := os.MkdirAll(dirPath, 0755); err != nil {
		return fmt.Errorf("failed to create directory: %w", err)
	}
	return nil
}

// CopyFile 复制文件
func CopyFile(src, dst string) error {
	sourceFile, err := os.Open(src)
	if err != nil {
		return fmt.Errorf("failed to open source file: %w", err)
	}
	defer sourceFile.Close()

	destFile, err := os.Create(dst)
	if err != nil {
		return fmt.Errorf("failed to create destination file: %w", err)
	}
	defer destFile.Close()

	if _, err := io.Copy(destFile, sourceFile); err != nil {
		return fmt.Errorf("failed to copy file: %w", err)
	}

	return nil
}

// GetFileSize 获取文件大小
func GetFileSize(filePath string) (int64, error) {
	stat, err := os.Stat(filePath)
	if err != nil {
		return 0, fmt.Errorf("failed to get file stats: %w", err)
	}
	return stat.Size(), nil
}

// FormatFileSize 格式化文件大小
func FormatFileSize(size int64) string {
	const (
		B  = 1
		KB = 1024 * B
		MB = 1024 * KB
		GB = 1024 * MB
	)

	switch {
	case size >= GB:
		return fmt.Sprintf("%.2f GB", float64(size)/float64(GB))
	case size >= MB:
		return fmt.Sprintf("%.2f MB", float64(size)/float64(MB))
	case size >= KB:
		return fmt.Sprintf("%.2f KB", float64(size)/float64(KB))
	default:
		return fmt.Sprintf("%d B", size)
	}
}
