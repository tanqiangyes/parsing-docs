package formats

import (
	"fmt"
	"strings"

	"docs-parser/internal/core/types"
)

// WordParser 通用Word文档解析器（自动分发到具体格式解析器）
type WordParser struct {
	parsers map[string]any // 扩展名到解析器实例
}

// NewWordParser 创建通用Word解析器
func NewWordParser() *WordParser {
	return &WordParser{
		parsers: map[string]any{
			".docx": NewDocxParser(),
			".doc":  NewDocParser(),
			".rtf":  NewRtfParser(),
			".wpd":  NewWpdParser(),
			".dot":  NewLegacyParser(),
			".dotx": NewLegacyParser(),
		},
	}
}

// ParseDocument 自动识别并解析Word文档
func (wp *WordParser) ParseDocument(filePath string) (*types.Document, error) {
	// 检查文件路径是否为空
	if filePath == "" {
		return nil, fmt.Errorf("file path is empty")
	}

	ext := strings.ToLower(getFileExt(filePath))
	parser, ok := wp.parsers[ext]
	if !ok {
		return nil, fmt.Errorf("unsupported file extension: %s", ext)
	}
	switch p := parser.(type) {
	case *DocxParser:
		return p.ParseDocument(filePath)
	case *DocParser:
		return p.ParseDocument(filePath)
	case *RtfParser:
		return p.ParseDocument(filePath)
	case *WpdParser:
		return p.ParseDocument(filePath)
	case *LegacyParser:
		return p.ParseDocument(filePath)
	default:
		return nil, fmt.Errorf("unknown parser type for extension: %s", ext)
	}
}

// getFileExt 获取文件扩展名
func getFileExt(filePath string) string {
	idx := strings.LastIndex(filePath, ".")
	if idx == -1 {
		return ""
	}
	return filePath[idx:]
}
