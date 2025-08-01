package parser

import (
	"docs-parser/internal/core/types"
	"fmt"
	"os"
	"time"
)

// Parser 文档解析器接口
type Parser interface {
	// ParseDocument 解析文档文件
	ParseDocument(filePath string) (*types.Document, error)

	// ParseMetadata 解析文档元数据
	ParseMetadata(filePath string) (*types.DocumentMetadata, error)

	// ParseContent 解析文档内容
	ParseContent(filePath string) (*types.DocumentContent, error)

	// ParseStyles 解析文档样式
	ParseStyles(filePath string) (*types.DocumentStyles, error)

	// ParseFormatRules 解析格式规则
	ParseFormatRules(filePath string) (*types.FormatRules, error)

	// GetSupportedFormats 获取支持的格式
	GetSupportedFormats() []string

	// ValidateFile 验证文件格式
	ValidateFile(filePath string) error
}

// DefaultParser 默认解析器实现
type DefaultParser struct{}

// ParseDocument 解析文档
func (dp *DefaultParser) ParseDocument(filePath string) (*types.Document, error) {
	return &types.Document{
		Metadata: types.DocumentMetadata{
			Title:    "Default Document",
			Author:   "Unknown",
			Created:  time.Now(),
			Modified: time.Now(),
			Version:  "1.0",
		},
		Content: types.DocumentContent{
			Paragraphs: []types.Paragraph{
				{
					ID:    "p1",
					Text:  "Default paragraph",
					Style: types.ParagraphStyle{Name: "Normal"},
				},
			},
		},
		Styles:      types.DocumentStyles{},
		FormatRules: types.FormatRules{},
	}, nil
}

// ParseMetadata 解析元数据
func (dp *DefaultParser) ParseMetadata(filePath string) (*types.DocumentMetadata, error) {
	return &types.DocumentMetadata{
		Title:    "Default Document",
		Author:   "Unknown",
		Created:  time.Now(),
		Modified: time.Now(),
		Version:  "1.0",
	}, nil
}

// ParseContent 解析内容
func (dp *DefaultParser) ParseContent(filePath string) (*types.DocumentContent, error) {
	return &types.DocumentContent{
		Paragraphs: []types.Paragraph{
			{
				ID:    "p1",
				Text:  "Default paragraph",
				Style: types.ParagraphStyle{Name: "Normal"},
			},
		},
	}, nil
}

// ParseStyles 解析样式
func (dp *DefaultParser) ParseStyles(filePath string) (*types.DocumentStyles, error) {
	return &types.DocumentStyles{}, nil
}

// ParseFormatRules 解析格式规则
func (dp *DefaultParser) ParseFormatRules(filePath string) (*types.FormatRules, error) {
	return &types.FormatRules{}, nil
}

// GetSupportedFormats 获取支持的格式
func (dp *DefaultParser) GetSupportedFormats() []string {
	return []string{"docx", "doc", "rtf", "wpd"}
}

// ValidateFile 验证文件
func (dp *DefaultParser) ValidateFile(filePath string) error {
	if filePath == "" {
		return fmt.Errorf("file path is empty")
	}

	// 检查文件是否存在
	if _, err := os.Stat(filePath); os.IsNotExist(err) {
		return fmt.Errorf("file not found: %s", filePath)
	}

	return nil
}

// ParserFactory 解析器工厂
type ParserFactory struct {
	parsers map[string]Parser
}

// NewParserFactory 创建解析器工厂
func NewParserFactory() *ParserFactory {
	return &ParserFactory{
		parsers: make(map[string]Parser),
	}
}

// RegisterParser 注册解析器
func (pf *ParserFactory) RegisterParser(format string, parser Parser) {
	pf.parsers[format] = parser
}

// GetParser 获取解析器
func (pf *ParserFactory) GetParser(format string) (Parser, error) {
	parser, exists := pf.parsers[format]
	if !exists {
		return nil, ErrUnsupportedFormat
	}
	return parser, nil
}

// ParseDocument 解析文档
func (pf *ParserFactory) ParseDocument(filePath string) (*types.Document, error) {
	format, err := pf.detectFormat(filePath)
	if err != nil {
		return nil, err
	}

	parser, err := pf.GetParser(format)
	if err != nil {
		return nil, err
	}

	return parser.ParseDocument(filePath)
}

// ParseMetadata 解析元数据
func (pf *ParserFactory) ParseMetadata(filePath string) (*types.DocumentMetadata, error) {
	format, err := pf.detectFormat(filePath)
	if err != nil {
		return nil, err
	}

	parser, err := pf.GetParser(format)
	if err != nil {
		return nil, err
	}

	return parser.ParseMetadata(filePath)
}

// ParseContent 解析内容
func (pf *ParserFactory) ParseContent(filePath string) (*types.DocumentContent, error) {
	format, err := pf.detectFormat(filePath)
	if err != nil {
		return nil, err
	}

	parser, err := pf.GetParser(format)
	if err != nil {
		return nil, err
	}

	return parser.ParseContent(filePath)
}

// ParseStyles 解析样式
func (pf *ParserFactory) ParseStyles(filePath string) (*types.DocumentStyles, error) {
	format, err := pf.detectFormat(filePath)
	if err != nil {
		return nil, err
	}

	parser, err := pf.GetParser(format)
	if err != nil {
		return nil, err
	}

	return parser.ParseStyles(filePath)
}

// ParseFormatRules 解析格式规则
func (pf *ParserFactory) ParseFormatRules(filePath string) (*types.FormatRules, error) {
	format, err := pf.detectFormat(filePath)
	if err != nil {
		return nil, err
	}

	parser, err := pf.GetParser(format)
	if err != nil {
		return nil, err
	}

	return parser.ParseFormatRules(filePath)
}

// GetSupportedFormats 获取支持的格式
func (pf *ParserFactory) GetSupportedFormats() []string {
	return []string{"docx", "doc", "rtf", "wpd"}
}

// ValidateFile 验证文件
func (pf *ParserFactory) ValidateFile(filePath string) error {
	format, err := pf.detectFormat(filePath)
	if err != nil {
		return err
	}

	parser, err := pf.GetParser(format)
	if err != nil {
		return err
	}

	return parser.ValidateFile(filePath)
}

// detectFormat 检测文件格式
func (pf *ParserFactory) detectFormat(filePath string) (string, error) {
	// 根据文件扩展名检测格式
	// 这里可以扩展为更复杂的格式检测逻辑
	ext := getFileExtension(filePath)

	switch ext {
	case ".docx":
		return "docx", nil
	case ".doc":
		return "doc", nil
	case ".rtf":
		return "rtf", nil
	case ".wpd":
		return "wpd", nil
	default:
		return "", ErrUnsupportedFormat
	}
}

// getFileExtension 获取文件扩展名
func getFileExtension(filePath string) string {
	// 简单的扩展名提取
	// 实际实现中应该使用 path/filepath
	for i := len(filePath) - 1; i >= 0; i-- {
		if filePath[i] == '.' {
			return filePath[i:]
		}
	}
	return ""
}

// 错误定义
var (
	ErrUnsupportedFormat = fmt.Errorf("unsupported file format")
	ErrFileNotFound      = fmt.Errorf("file not found")
	ErrInvalidFile       = fmt.Errorf("invalid file")
)
