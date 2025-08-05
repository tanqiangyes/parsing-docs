package formats

import (
	"fmt"

	"docs-parser/internal/core/parser"
	"docs-parser/internal/core/types"
	"docs-parser/internal/documents"
)

// DocxParser .docx格式解析器
type DocxParser struct {
	factory *parser.ParserFactory
}

// NewDocxParser 创建.docx解析器
func NewDocxParser() *DocxParser {
	return &DocxParser{}
}

// ParseDocument 解析.docx文档
func (dp *DocxParser) ParseDocument(filePath string) (*types.Document, error) {
	fmt.Printf("开始解析DOCX文档: %s\n", filePath)

	// 验证文件
	if err := dp.ValidateFile(filePath); err != nil {
		return nil, err
	}

	// 使用新的分层架构
	wordDoc := documents.NewWordprocessingDocument(filePath)
	defer wordDoc.Close()

	// 打开文档
	if err := wordDoc.Open(); err != nil {
		return nil, fmt.Errorf("failed to open word document: %w", err)
	}

	// 解析文档
	doc, err := wordDoc.Parse()
	if err != nil {
		return nil, fmt.Errorf("failed to parse word document: %w", err)
	}

	// 打印详细的解析结果
	fmt.Printf("文档解析完成: %s\n", filePath)
	fmt.Printf("  - 段落数量: %d\n", len(doc.Content.Paragraphs))
	fmt.Printf("  - 表格数量: %d\n", len(doc.Content.Tables))
	fmt.Printf("  - 字体规则数量: %d\n", len(doc.FormatRules.FontRules))
	fmt.Printf("  - 段落规则数量: %d\n", len(doc.FormatRules.ParagraphRules))
	fmt.Printf("  - 表格规则数量: %d\n", len(doc.FormatRules.TableRules))
	fmt.Printf("  - 页面规则数量: %d\n", len(doc.FormatRules.PageRules))
	fmt.Printf("  - 段落样式数量: %d\n", len(doc.Styles.ParagraphStyles))
	fmt.Printf("  - 字符样式数量: %d\n", len(doc.Styles.CharacterStyles))
	fmt.Printf("  - 表格样式数量: %d\n", len(doc.Styles.TableStyles))

	// 打印文档元数据
	fmt.Printf("  - 标题: %s\n", doc.Metadata.Title)
	fmt.Printf("  - 作者: %s\n", doc.Metadata.Author)
	fmt.Printf("  - 页数: %d\n", doc.Metadata.PageCount)
	fmt.Printf("  - 字数: %d\n", doc.Metadata.WordCount)

	// 打印前几个段落的内容摘要
	if len(doc.Content.Paragraphs) > 0 {
		fmt.Printf("  - 段落内容摘要:\n")
		for i, para := range doc.Content.Paragraphs {
			if i >= 3 { // 只显示前3个段落
				break
			}
			text := para.Text
			if len(text) > 50 {
				text = text[:50] + "..."
			}
			fmt.Printf("    %d. [%s] %s\n", i+1, para.Style.Name, text)
		}
	}

	// 打印性能监控报告
	wordDoc.Monitor.PrintReport()

	return doc, nil
}

// ParseMetadata 解析文档元数据
func (dp *DocxParser) ParseMetadata(filePath string) (*types.DocumentMetadata, error) {
	wordDoc := documents.NewWordprocessingDocument(filePath)
	defer wordDoc.Close()

	if err := wordDoc.Open(); err != nil {
		return nil, err
	}

	doc, err := wordDoc.Parse()
	if err != nil {
		return nil, err
	}

	return &doc.Metadata, nil
}

// ParseContent 解析文档内容
func (dp *DocxParser) ParseContent(filePath string) (*types.DocumentContent, error) {
	wordDoc := documents.NewWordprocessingDocument(filePath)
	defer wordDoc.Close()

	if err := wordDoc.Open(); err != nil {
		return nil, err
	}

	doc, err := wordDoc.Parse()
	if err != nil {
		return nil, err
	}

	return &doc.Content, nil
}

// ParseStyles 解析文档样式
func (dp *DocxParser) ParseStyles(filePath string) (*types.DocumentStyles, error) {
	wordDoc := documents.NewWordprocessingDocument(filePath)
	defer wordDoc.Close()

	if err := wordDoc.Open(); err != nil {
		return nil, err
	}

	doc, err := wordDoc.Parse()
	if err != nil {
		return nil, err
	}

	return &doc.Styles, nil
}

// ParseFormatRules 解析格式规则
func (dp *DocxParser) ParseFormatRules(filePath string) (*types.FormatRules, error) {
	wordDoc := documents.NewWordprocessingDocument(filePath)
	defer wordDoc.Close()

	if err := wordDoc.Open(); err != nil {
		return nil, err
	}

	doc, err := wordDoc.Parse()
	if err != nil {
		return nil, err
	}

	return &doc.FormatRules, nil
}

// GetSupportedFormats 获取支持的格式
func (dp *DocxParser) GetSupportedFormats() []string {
	return []string{".docx"}
}

// ValidateFile 验证文件
func (dp *DocxParser) ValidateFile(filePath string) error {
	// 使用OPC容器验证
	wordDoc := documents.NewWordprocessingDocument(filePath)
	defer wordDoc.Close()

	if err := wordDoc.Open(); err != nil {
		return fmt.Errorf("invalid docx file: %w", err)
	}

	return nil
}
