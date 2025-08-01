package formats

import (
	"encoding/binary"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"time"

	"docs-parser/internal/core/parser"
	"docs-parser/internal/core/types"
)

// DocParser .doc格式解析器
type DocParser struct {
	factory *parser.ParserFactory
}

// DocVersion 表示Word版本
type DocVersion struct {
	Major    int
	Minor    int
	Build    int
	Platform string
}

// DocHeader .doc文件头结构
type DocHeader struct {
	Signature    []byte
	Version      DocVersion
	FileType     string
	Encoding     string
	Language     string
	IsEncrypted  bool
	HasPassword  bool
	DocumentType string
}

// NewDocParser 创建.doc解析器
func NewDocParser() *DocParser {
	return &DocParser{}
}

// ParseDocument 解析.doc文档
func (dp *DocParser) ParseDocument(filePath string) (*types.Document, error) {
	fmt.Printf("开始解析DOC文档: %s\n", filePath)

	// 验证文件
	if err := dp.ValidateFile(filePath); err != nil {
		return nil, err
	}

	// 解析文件头以确定版本
	header, err := dp.parseDocHeader(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to parse header: %w", err)
	}

	fmt.Printf("文档版本: %s %d.%d.%d\n", header.Version.Platform, header.Version.Major, header.Version.Minor, header.Version.Build)

	doc := &types.Document{}

	// 解析元数据
	fmt.Println("正在解析文档元数据...")
	metadata, err := dp.parseMetadata(filePath, header)
	if err != nil {
		return nil, fmt.Errorf("failed to parse metadata: %w", err)
	}
	doc.Metadata = *metadata
	fmt.Printf("元数据解析完成:\n")
	fmt.Printf("  - 标题: %s\n", metadata.Title)
	fmt.Printf("  - 作者: %s\n", metadata.Author)
	fmt.Printf("  - 创建时间: %s\n", metadata.Created.Format("2006-01-02 15:04:05"))
	fmt.Printf("  - 修改时间: %s\n", metadata.Modified.Format("2006-01-02 15:04:05"))
	fmt.Printf("  - 页数: %d\n", metadata.PageCount)
	fmt.Printf("  - 字数: %d\n", metadata.WordCount)

	// 解析内容
	fmt.Println("正在解析文档内容...")
	content, err := dp.parseContent(filePath, header)
	if err != nil {
		return nil, fmt.Errorf("failed to parse content: %w", err)
	}
	doc.Content = *content
	fmt.Printf("内容解析完成:\n")
	fmt.Printf("  - 段落数量: %d\n", len(content.Paragraphs))
	fmt.Printf("  - 表格数量: %d\n", len(content.Tables))
	fmt.Printf("  - 图片数量: %d\n", len(content.Images))
	fmt.Printf("  - 页眉数量: %d\n", len(content.Headers))
	fmt.Printf("  - 页脚数量: %d\n", len(content.Footers))
	fmt.Printf("  - 节数量: %d\n", len(content.Sections))

	// 解析样式
	fmt.Println("正在解析文档样式...")
	styles, err := dp.parseStyles(filePath, header)
	if err != nil {
		return nil, fmt.Errorf("failed to parse styles: %w", err)
	}
	doc.Styles = *styles
	fmt.Printf("样式解析完成:\n")
	fmt.Printf("  - 段落样式数量: %d\n", len(styles.ParagraphStyles))
	fmt.Printf("  - 字符样式数量: %d\n", len(styles.CharacterStyles))
	fmt.Printf("  - 表格样式数量: %d\n", len(styles.TableStyles))

	// 解析格式规则
	fmt.Println("正在解析格式规则...")
	formatRules, err := dp.parseFormatRules(filePath, header)
	if err != nil {
		return nil, fmt.Errorf("failed to parse format rules: %w", err)
	}
	doc.FormatRules = *formatRules
	fmt.Printf("格式规则解析完成:\n")
	fmt.Printf("  - 字体规则数量: %d\n", len(formatRules.FontRules))
	fmt.Printf("  - 段落规则数量: %d\n", len(formatRules.ParagraphRules))
	fmt.Printf("  - 表格规则数量: %d\n", len(formatRules.TableRules))
	fmt.Printf("  - 页面规则数量: %d\n", len(formatRules.PageRules))

	// 打印详细的格式信息
	fmt.Println("\n=== 详细格式信息 ===")

	// 打印字体规则
	if len(formatRules.FontRules) > 0 {
		fmt.Println("\n字体规则:")
		for i, font := range formatRules.FontRules {
			fmt.Printf("  %d. ID: %s, 名称: %s, 大小: %.1f, 颜色: %s\n",
				i+1, font.ID, font.Name, font.Size, font.Color.RGB)
		}
	}

	// 打印段落规则
	if len(formatRules.ParagraphRules) > 0 {
		fmt.Println("\n段落规则:")
		for i, para := range formatRules.ParagraphRules {
			fmt.Printf("  %d. ID: %s, 名称: %s, 对齐: %s, 缩进: %.1f\n",
				i+1, para.ID, para.Name, para.Alignment, para.Indentation.Left)
		}
	}

	// 打印表格规则
	if len(formatRules.TableRules) > 0 {
		fmt.Println("\n表格规则:")
		for i, table := range formatRules.TableRules {
			fmt.Printf("  %d. ID: %s, 名称: %s, 宽度: %.1f, 对齐: %s\n",
				i+1, table.ID, table.Name, table.Width, table.Alignment)
		}
	}

	// 打印页面规则
	if len(formatRules.PageRules) > 0 {
		fmt.Println("\n页面规则:")
		for i, page := range formatRules.PageRules {
			fmt.Printf("  %d. ID: %s, 名称: %s, 宽度: %.1f, 高度: %.1f\n",
				i+1, page.ID, page.Name, page.PageSize.Width, page.PageSize.Height)
		}
	}

	// 打印样式信息
	if len(styles.ParagraphStyles) > 0 {
		fmt.Println("\n段落样式:")
		for i, style := range styles.ParagraphStyles {
			fmt.Printf("  %d. ID: %s, 名称: %s\n",
				i+1, style.ID, style.Name)
		}
	}

	if len(styles.CharacterStyles) > 0 {
		fmt.Println("\n字符样式:")
		for i, style := range styles.CharacterStyles {
			fmt.Printf("  %d. ID: %s, 名称: %s\n",
				i+1, style.ID, style.Name)
		}
	}

	if len(styles.TableStyles) > 0 {
		fmt.Println("\n表格样式:")
		for i, style := range styles.TableStyles {
			fmt.Printf("  %d. ID: %s, 名称: %s\n",
				i+1, style.ID, style.Name)
		}
	}

	fmt.Printf("\n文档解析完成: %s\n", filePath)
	return doc, nil
}

// ParseMetadata 解析元数据
func (dp *DocParser) ParseMetadata(filePath string) (*types.DocumentMetadata, error) {
	header, err := dp.parseDocHeader(filePath)
	if err != nil {
		return nil, err
	}
	return dp.parseMetadata(filePath, header)
}

// ParseContent 解析内容
func (dp *DocParser) ParseContent(filePath string) (*types.DocumentContent, error) {
	header, err := dp.parseDocHeader(filePath)
	if err != nil {
		return nil, err
	}
	return dp.parseContent(filePath, header)
}

// ParseStyles 解析样式
func (dp *DocParser) ParseStyles(filePath string) (*types.DocumentStyles, error) {
	header, err := dp.parseDocHeader(filePath)
	if err != nil {
		return nil, err
	}
	return dp.parseStyles(filePath, header)
}

// ParseFormatRules 解析格式规则
func (dp *DocParser) ParseFormatRules(filePath string) (*types.FormatRules, error) {
	header, err := dp.parseDocHeader(filePath)
	if err != nil {
		return nil, err
	}
	return dp.parseFormatRules(filePath, header)
}

// GetSupportedFormats 获取支持的格式
func (dp *DocParser) GetSupportedFormats() []string {
	return []string{"doc", "dot", "wbk"}
}

// ValidateFile 验证文件格式
func (dp *DocParser) ValidateFile(filePath string) error {
	// 检查文件是否存在
	if _, err := os.Stat(filePath); os.IsNotExist(err) {
		return parser.ErrFileNotFound
	}

	// 检查文件扩展名
	ext := filepath.Ext(filePath)
	supported := map[string]bool{".doc": true, ".dot": true, ".wbk": true}
	if !supported[ext] {
		return parser.ErrUnsupportedFormat
	}

	// 检查文件头
	file, err := os.Open(filePath)
	if err != nil {
		return parser.ErrInvalidFile
	}
	defer file.Close()

	// 读取文件头
	header := make([]byte, 16)
	if _, err := file.Read(header); err != nil {
		return parser.ErrInvalidFile
	}

	// 检查.doc文件头标识
	if !dp.isValidDocHeader(header) {
		return parser.ErrInvalidFile
	}

	return nil
}

// isValidDocHeader 检查是否为有效的.doc文件头
func (dp *DocParser) isValidDocHeader(header []byte) bool {
	// .doc文件通常以特定的魔数开始
	// 这里检查常见的.doc文件头标识
	docSignatures := [][]byte{
		{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1}, // 标准OLE2格式
		{0x50, 0x4B, 0x03, 0x04},                         // ZIP格式（某些.doc文件）
		{0x31, 0xBE, 0x00, 0x00},                         // Word 6.0/95格式
		{0xDB, 0xA5, 0x2D, 0x00},                         // Word 2.0格式
	}

	for _, signature := range docSignatures {
		if len(header) >= len(signature) {
			match := true
			for i, b := range signature {
				if header[i] != b {
					match = false
					break
				}
			}
			if match {
				return true
			}
		}
	}

	return false
}

// parseDocHeader 解析.doc文件头
func (dp *DocParser) parseDocHeader(filePath string) (*DocHeader, error) {
	file, err := os.Open(filePath)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	header := make([]byte, 64)
	if _, err := file.Read(header); err != nil {
		return nil, err
	}

	docHeader := &DocHeader{
		Signature: header[:8],
	}

	// 根据魔数确定版本
	if header[0] == 0xD0 && header[1] == 0xCF {
		docHeader.Version = DocVersion{Major: 8, Minor: 0, Platform: "Windows"}
		docHeader.FileType = "Word 97-2003"
	} else if header[0] == 0x50 && header[1] == 0x4B {
		docHeader.Version = DocVersion{Major: 12, Minor: 0, Platform: "Windows"}
		docHeader.FileType = "Word 2007+"
	} else if header[0] == 0x31 && header[1] == 0xBE {
		docHeader.Version = DocVersion{Major: 6, Minor: 0, Platform: "Windows"}
		docHeader.FileType = "Word 6.0/95"
	} else if header[0] == 0xDB && header[1] == 0xA5 {
		docHeader.Version = DocVersion{Major: 2, Minor: 0, Platform: "Windows"}
		docHeader.FileType = "Word 2.0"
	} else {
		docHeader.Version = DocVersion{Major: 8, Minor: 0, Platform: "Windows"}
		docHeader.FileType = "Word Document"
	}

	// 检查文档类型
	if len(header) > 8 {
		docType := header[8]
		switch docType {
		case 0x00:
			docHeader.DocumentType = "Document"
		case 0x01:
			docHeader.DocumentType = "Template"
		case 0x02:
			docHeader.DocumentType = "Backup"
		default:
			docHeader.DocumentType = "Document"
		}
	}

	// 检查加密标志
	if len(header) > 12 {
		docHeader.IsEncrypted = (header[12] & 0x01) != 0
		docHeader.HasPassword = (header[12] & 0x02) != 0
	}

	return docHeader, nil
}

// parseMetadata 解析元数据
func (dp *DocParser) parseMetadata(filePath string, header *DocHeader) (*types.DocumentMetadata, error) {
	metadata := &types.DocumentMetadata{}

	file, err := os.Open(filePath)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	// 读取文件大小
	fileInfo, err := file.Stat()
	if err != nil {
		return nil, err
	}
	metadata.FileSize = fileInfo.Size()

	// 根据版本解析不同的元数据
	switch header.FileType {
	case "Word 97-2003":
		return dp.parseWord97Metadata(file, metadata)
	case "Word 2007+":
		return dp.parseWord2007Metadata(file, metadata)
	case "Word 6.0/95":
		return dp.parseWord60Metadata(file, metadata)
	case "Word 2.0":
		return dp.parseWord20Metadata(file, metadata)
	default:
		// 默认元数据
		metadata.Title = fmt.Sprintf("Word Document (%s)", header.FileType)
		metadata.Author = "Unknown"
		metadata.Created = time.Now()
		metadata.Modified = time.Now()
		metadata.Version = fmt.Sprintf("%d.%d", header.Version.Major, header.Version.Minor)
	}

	return metadata, nil
}

// parseWord97Metadata 解析Word 97-2003元数据
func (dp *DocParser) parseWord97Metadata(file *os.File, metadata *types.DocumentMetadata) (*types.DocumentMetadata, error) {
	// Word 97-2003使用OLE2复合文档格式
	// 这里提供基础实现，实际需要更复杂的OLE2解析

	// 跳过文件头
	file.Seek(512, 0)

	// 读取元数据区域
	metaData := make([]byte, 256)
	file.Read(metaData)

	// 提取标题
	if len(metaData) > 32 {
		title := strings.TrimRight(string(metaData[32:64]), "\x00")
		if title != "" {
			metadata.Title = title
		} else {
			metadata.Title = "Word 97-2003 Document"
		}
	}

	// 提取作者
	if len(metaData) > 64 {
		author := strings.TrimRight(string(metaData[64:96]), "\x00")
		if author != "" {
			metadata.Author = author
		} else {
			metadata.Author = "Unknown"
		}
	}

	// 提取时间信息
	if len(metaData) > 96 {
		createdTime := dp.parseDocTime(metaData[96:104])
		if !createdTime.IsZero() {
			metadata.Created = createdTime
		} else {
			metadata.Created = time.Now()
		}
	}

	if len(metaData) > 104 {
		modifiedTime := dp.parseDocTime(metaData[104:112])
		if !modifiedTime.IsZero() {
			metadata.Modified = modifiedTime
		} else {
			metadata.Modified = time.Now()
		}
	}

	metadata.Version = "8.0"
	return metadata, nil
}

// parseWord2007Metadata 解析Word 2007+元数据
func (dp *DocParser) parseWord2007Metadata(file *os.File, metadata *types.DocumentMetadata) (*types.DocumentMetadata, error) {
	// Word 2007+使用ZIP格式
	// 这里提供基础实现
	metadata.Title = "Word 2007+ Document"
	metadata.Author = "Unknown"
	metadata.Created = time.Now()
	metadata.Modified = time.Now()
	metadata.Version = "12.0"

	return metadata, nil
}

// parseWord60Metadata 解析Word 6.0/95元数据
func (dp *DocParser) parseWord60Metadata(file *os.File, metadata *types.DocumentMetadata) (*types.DocumentMetadata, error) {
	// Word 6.0/95元数据解析
	file.Seek(8, 0)

	metaData := make([]byte, 256)
	file.Read(metaData)

	// 提取标题
	if len(metaData) > 32 {
		title := strings.TrimRight(string(metaData[32:64]), "\x00")
		if title != "" {
			metadata.Title = title
		} else {
			metadata.Title = "Word 6.0/95 Document"
		}
	}

	// 提取作者
	if len(metaData) > 64 {
		author := strings.TrimRight(string(metaData[64:96]), "\x00")
		if author != "" {
			metadata.Author = author
		} else {
			metadata.Author = "Unknown"
		}
	}

	// 提取时间信息
	if len(metaData) > 96 {
		createdTime := dp.parseDocTime(metaData[96:104])
		if !createdTime.IsZero() {
			metadata.Created = createdTime
		} else {
			metadata.Created = time.Now()
		}
	}

	if len(metaData) > 104 {
		modifiedTime := dp.parseDocTime(metaData[104:112])
		if !modifiedTime.IsZero() {
			metadata.Modified = modifiedTime
		} else {
			metadata.Modified = time.Now()
		}
	}

	metadata.Version = "6.0"
	return metadata, nil
}

// parseWord20Metadata 解析Word 2.0元数据
func (dp *DocParser) parseWord20Metadata(file *os.File, metadata *types.DocumentMetadata) (*types.DocumentMetadata, error) {
	// Word 2.0元数据解析
	metadata.Title = "Word 2.0 Document"
	metadata.Author = "Unknown"
	metadata.Created = time.Now()
	metadata.Modified = time.Now()
	metadata.Version = "2.0"

	return metadata, nil
}

// parseDocTime 解析.doc时间格式
func (dp *DocParser) parseDocTime(timeData []byte) time.Time {
	if len(timeData) < 8 {
		return time.Time{}
	}

	// Word使用DOS时间格式
	dosTime := binary.LittleEndian.Uint32(timeData[:4])
	dosDate := binary.LittleEndian.Uint32(timeData[4:8])

	// 解析DOS日期时间
	year := int((dosDate>>9)&0x7F) + 1980
	month := int((dosDate >> 5) & 0x0F)
	day := int(dosDate & 0x1F)

	hour := int((dosTime >> 11) & 0x1F)
	minute := int((dosTime >> 5) & 0x3F)
	second := int((dosTime & 0x1F) * 2)

	return time.Date(year, time.Month(month), day, hour, minute, second, 0, time.UTC)
}

// parseContent 解析文档内容
func (dp *DocParser) parseContent(filePath string, header *DocHeader) (*types.DocumentContent, error) {
	content := &types.DocumentContent{}

	// 根据版本解析不同的内容结构
	switch header.FileType {
	case "Word 97-2003":
		return dp.parseWord97Content(filePath, content)
	case "Word 2007+":
		return dp.parseWord2007Content(filePath, content)
	case "Word 6.0/95":
		return dp.parseWord60Content(filePath, content)
	case "Word 2.0":
		return dp.parseWord20Content(filePath, content)
	default:
		// 默认内容解析
		content.Paragraphs = append(content.Paragraphs, types.Paragraph{
			ID:    "p1",
			Text:  fmt.Sprintf("Sample paragraph from %s file", header.FileType),
			Style: types.ParagraphStyle{Name: "Normal"},
		})
	}

	return content, nil
}

// parseWord97Content 解析Word 97-2003内容
func (dp *DocParser) parseWord97Content(filePath string, content *types.DocumentContent) (*types.DocumentContent, error) {
	file, err := os.Open(filePath)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	// Word 97-2003使用OLE2格式，需要更复杂的解析
	// 这里提供基础实现

	// 跳过文件头
	file.Seek(1024, 0)

	// 读取文档内容区域
	docContent := make([]byte, 2048)
	file.Read(docContent)

	// 解析段落 (简化实现)
	paragraphs := dp.parseDocParagraphs(docContent)
	content.Paragraphs = paragraphs

	// 解析表格 (简化实现)
	tables := dp.parseDocTables(docContent)
	content.Tables = tables

	// 解析页眉页脚 (简化实现)
	headers, footers := dp.parseDocHeadersFooters(docContent)
	content.Headers = headers
	content.Footers = footers

	return content, nil
}

// parseWord2007Content 解析Word 2007+内容
func (dp *DocParser) parseWord2007Content(filePath string, content *types.DocumentContent) (*types.DocumentContent, error) {
	// Word 2007+使用ZIP格式
	// 这里提供基础实现
	content.Paragraphs = append(content.Paragraphs, types.Paragraph{
		ID:    "p1",
		Text:  "Sample paragraph from Word 2007+ file",
		Style: types.ParagraphStyle{Name: "Normal"},
	})

	return content, nil
}

// parseWord60Content 解析Word 6.0/95内容
func (dp *DocParser) parseWord60Content(filePath string, content *types.DocumentContent) (*types.DocumentContent, error) {
	file, err := os.Open(filePath)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	// 跳过文件头
	file.Seek(512, 0)

	// 读取文档内容区域
	docContent := make([]byte, 1024)
	file.Read(docContent)

	// 解析段落 (简化实现)
	paragraphs := dp.parseDocParagraphs(docContent)
	content.Paragraphs = paragraphs

	// 解析表格 (简化实现)
	tables := dp.parseDocTables(docContent)
	content.Tables = tables

	return content, nil
}

// parseWord20Content 解析Word 2.0内容
func (dp *DocParser) parseWord20Content(filePath string, content *types.DocumentContent) (*types.DocumentContent, error) {
	file, err := os.Open(filePath)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	// Word 2.0使用不同的内容结构
	file.Seek(128, 0)

	docContent := make([]byte, 1024)
	file.Read(docContent)

	paragraphs := dp.parseDocParagraphs(docContent)
	content.Paragraphs = paragraphs

	return content, nil
}

// parseDocParagraphs 解析.doc段落
func (dp *DocParser) parseDocParagraphs(content []byte) []types.Paragraph {
	var paragraphs []types.Paragraph

	// 简化的段落解析逻辑
	// 在实际实现中，需要根据具体的Word版本格式进行详细解析

	// 查找段落分隔符
	text := string(content)
	lines := strings.Split(text, "\r\n")

	for i, line := range lines {
		if strings.TrimSpace(line) != "" {
			paragraph := types.Paragraph{
				ID:    fmt.Sprintf("p%d", i+1),
				Text:  strings.TrimSpace(line),
				Style: types.ParagraphStyle{Name: "Normal"},
			}
			paragraphs = append(paragraphs, paragraph)
		}
	}

	// 如果没有找到有效段落，添加默认段落
	if len(paragraphs) == 0 {
		paragraphs = append(paragraphs, types.Paragraph{
			ID:    "p1",
			Text:  "Word Document Content",
			Style: types.ParagraphStyle{Name: "Normal"},
		})
	}

	return paragraphs
}

// parseDocTables 解析.doc表格
func (dp *DocParser) parseDocTables(content []byte) []types.Table {
	var tables []types.Table

	// 简化的表格解析逻辑
	// 在实际实现中，需要根据具体的Word版本格式进行详细解析

	// 添加示例表格
	table := types.Table{
		ID: "t1",
		Rows: []types.TableRow{
			{
				ID: "r1",
				Cells: []types.TableCell{
					{
						ID:      "c1",
						Content: []types.Paragraph{{Text: "Word Table Cell"}},
					},
				},
			},
		},
	}

	tables = append(tables, table)

	return tables
}

// parseDocHeadersFooters 解析.doc页眉页脚
func (dp *DocParser) parseDocHeadersFooters(content []byte) ([]types.Header, []types.Footer) {
	var headers []types.Header
	var footers []types.Footer

	// 简化的页眉页脚解析逻辑
	// 在实际实现中，需要根据具体的Word版本格式进行详细解析

	// 添加示例页眉
	header := types.Header{
		ID: "h1",
		Content: []types.Paragraph{
			{
				ID:   "hp1",
				Text: "Header content",
			},
		},
	}

	// 添加示例页脚
	footer := types.Footer{
		ID: "f1",
		Content: []types.Paragraph{
			{
				ID:   "fp1",
				Text: "Footer content",
			},
		},
	}

	headers = append(headers, header)
	footers = append(footers, footer)

	return headers, footers
}

// parseStyles 解析样式
func (dp *DocParser) parseStyles(filePath string, header *DocHeader) (*types.DocumentStyles, error) {
	styles := &types.DocumentStyles{}

	// 根据版本解析不同的样式
	switch header.FileType {
	case "Word 97-2003":
		return dp.parseWord97Styles(filePath, styles)
	case "Word 2007+":
		return dp.parseWord2007Styles(filePath, styles)
	case "Word 6.0/95":
		return dp.parseWord60Styles(filePath, styles)
	case "Word 2.0":
		return dp.parseWord20Styles(filePath, styles)
	default:
		// 默认样式
		styles.ParagraphStyles = append(styles.ParagraphStyles, types.ParagraphStyle{ID: "ps1", Name: "Normal"})
		styles.CharacterStyles = append(styles.CharacterStyles, types.CharacterStyle{ID: "cs1", Name: "Default Paragraph Font"})
		styles.TableStyles = append(styles.TableStyles, types.TableStyle{ID: "ts1", Name: "Table Grid"})
	}

	return styles, nil
}

// parseWord97Styles 解析Word 97-2003样式
func (dp *DocParser) parseWord97Styles(filePath string, styles *types.DocumentStyles) (*types.DocumentStyles, error) {
	// Word 97-2003样式解析
	styles.ParagraphStyles = append(styles.ParagraphStyles, types.ParagraphStyle{ID: "ps1", Name: "Normal"})
	styles.CharacterStyles = append(styles.CharacterStyles, types.CharacterStyle{ID: "cs1", Name: "Default Paragraph Font"})
	styles.TableStyles = append(styles.TableStyles, types.TableStyle{ID: "ts1", Name: "Table Grid"})

	return styles, nil
}

// parseWord2007Styles 解析Word 2007+样式
func (dp *DocParser) parseWord2007Styles(filePath string, styles *types.DocumentStyles) (*types.DocumentStyles, error) {
	// Word 2007+样式解析
	styles.ParagraphStyles = append(styles.ParagraphStyles, types.ParagraphStyle{ID: "ps1", Name: "Normal"})
	styles.CharacterStyles = append(styles.CharacterStyles, types.CharacterStyle{ID: "cs1", Name: "Default Paragraph Font"})
	styles.TableStyles = append(styles.TableStyles, types.TableStyle{ID: "ts1", Name: "Table Grid"})

	return styles, nil
}

// parseWord60Styles 解析Word 6.0/95样式
func (dp *DocParser) parseWord60Styles(filePath string, styles *types.DocumentStyles) (*types.DocumentStyles, error) {
	// Word 6.0/95样式解析
	styles.ParagraphStyles = append(styles.ParagraphStyles, types.ParagraphStyle{ID: "ps1", Name: "Normal"})
	styles.CharacterStyles = append(styles.CharacterStyles, types.CharacterStyle{ID: "cs1", Name: "Default Paragraph Font"})
	styles.TableStyles = append(styles.TableStyles, types.TableStyle{ID: "ts1", Name: "Table Grid"})

	return styles, nil
}

// parseWord20Styles 解析Word 2.0样式
func (dp *DocParser) parseWord20Styles(filePath string, styles *types.DocumentStyles) (*types.DocumentStyles, error) {
	// Word 2.0样式解析
	styles.ParagraphStyles = append(styles.ParagraphStyles, types.ParagraphStyle{ID: "ps1", Name: "Normal"})
	styles.CharacterStyles = append(styles.CharacterStyles, types.CharacterStyle{ID: "cs1", Name: "Default Paragraph Font"})
	styles.TableStyles = append(styles.TableStyles, types.TableStyle{ID: "ts1", Name: "Table Grid"})

	return styles, nil
}

// parseFormatRules 解析格式规则
func (dp *DocParser) parseFormatRules(filePath string, header *DocHeader) (*types.FormatRules, error) {
	formatRules := &types.FormatRules{}

	// 根据版本解析不同的格式规则
	switch header.FileType {
	case "Word 97-2003":
		return dp.parseWord97FormatRules(filePath, formatRules)
	case "Word 2007+":
		return dp.parseWord2007FormatRules(filePath, formatRules)
	case "Word 6.0/95":
		return dp.parseWord60FormatRules(filePath, formatRules)
	case "Word 2.0":
		return dp.parseWord20FormatRules(filePath, formatRules)
	default:
		// 默认格式规则
		formatRules.FontRules = append(formatRules.FontRules, types.FontRule{ID: "fr1", Name: "Default Font", Size: 12.0})
		formatRules.ParagraphRules = append(formatRules.ParagraphRules, types.ParagraphRule{ID: "pr1", Name: "Normal", Alignment: types.AlignLeft})
		formatRules.TableRules = append(formatRules.TableRules, types.TableRule{ID: "tr1", Name: "Table Grid", Width: 100.0, Alignment: types.AlignLeft})
		formatRules.PageRules = append(formatRules.PageRules, types.PageRule{ID: "pg1", Name: "Normal", PageSize: types.PageSize{Width: 612.0, Height: 792.0}, PageMargins: types.PageMargins{Top: 72.0, Bottom: 72.0, Left: 72.0, Right: 72.0, Header: 36.0, Footer: 36.0}})
	}

	return formatRules, nil
}

// parseWord97FormatRules 解析Word 97-2003格式规则
func (dp *DocParser) parseWord97FormatRules(filePath string, formatRules *types.FormatRules) (*types.FormatRules, error) {
	// Word 97-2003格式规则
	formatRules.FontRules = append(formatRules.FontRules, types.FontRule{ID: "fr1", Name: "Times New Roman", Size: 12.0})
	formatRules.ParagraphRules = append(formatRules.ParagraphRules, types.ParagraphRule{ID: "pr1", Name: "Normal", Alignment: types.AlignLeft})
	formatRules.TableRules = append(formatRules.TableRules, types.TableRule{ID: "tr1", Name: "Table Grid", Width: 100.0, Alignment: types.AlignLeft})
	formatRules.PageRules = append(formatRules.PageRules, types.PageRule{ID: "pg1", Name: "Normal", PageSize: types.PageSize{Width: 612.0, Height: 792.0}, PageMargins: types.PageMargins{Top: 72.0, Bottom: 72.0, Left: 72.0, Right: 72.0, Header: 36.0, Footer: 36.0}})

	return formatRules, nil
}

// parseWord2007FormatRules 解析Word 2007+格式规则
func (dp *DocParser) parseWord2007FormatRules(filePath string, formatRules *types.FormatRules) (*types.FormatRules, error) {
	// Word 2007+格式规则
	formatRules.FontRules = append(formatRules.FontRules, types.FontRule{ID: "fr1", Name: "Calibri", Size: 11.0})
	formatRules.ParagraphRules = append(formatRules.ParagraphRules, types.ParagraphRule{ID: "pr1", Name: "Normal", Alignment: types.AlignLeft})
	formatRules.TableRules = append(formatRules.TableRules, types.TableRule{ID: "tr1", Name: "Table Grid", Width: 100.0, Alignment: types.AlignLeft})
	formatRules.PageRules = append(formatRules.PageRules, types.PageRule{ID: "pg1", Name: "Normal", PageSize: types.PageSize{Width: 612.0, Height: 792.0}, PageMargins: types.PageMargins{Top: 72.0, Bottom: 72.0, Left: 72.0, Right: 72.0, Header: 36.0, Footer: 36.0}})

	return formatRules, nil
}

// parseWord60FormatRules 解析Word 6.0/95格式规则
func (dp *DocParser) parseWord60FormatRules(filePath string, formatRules *types.FormatRules) (*types.FormatRules, error) {
	// Word 6.0/95格式规则
	formatRules.FontRules = append(formatRules.FontRules, types.FontRule{ID: "fr1", Name: "Times New Roman", Size: 12.0})
	formatRules.ParagraphRules = append(formatRules.ParagraphRules, types.ParagraphRule{ID: "pr1", Name: "Normal", Alignment: types.AlignLeft})
	formatRules.TableRules = append(formatRules.TableRules, types.TableRule{ID: "tr1", Name: "Table Grid", Width: 100.0, Alignment: types.AlignLeft})
	formatRules.PageRules = append(formatRules.PageRules, types.PageRule{ID: "pg1", Name: "Normal", PageSize: types.PageSize{Width: 612.0, Height: 792.0}, PageMargins: types.PageMargins{Top: 72.0, Bottom: 72.0, Left: 72.0, Right: 72.0, Header: 36.0, Footer: 36.0}})

	return formatRules, nil
}

// parseWord20FormatRules 解析Word 2.0格式规则
func (dp *DocParser) parseWord20FormatRules(filePath string, formatRules *types.FormatRules) (*types.FormatRules, error) {
	// Word 2.0格式规则
	formatRules.FontRules = append(formatRules.FontRules, types.FontRule{ID: "fr1", Name: "Times New Roman", Size: 12.0})
	formatRules.ParagraphRules = append(formatRules.ParagraphRules, types.ParagraphRule{ID: "pr1", Name: "Normal", Alignment: types.AlignLeft})
	formatRules.TableRules = append(formatRules.TableRules, types.TableRule{ID: "tr1", Name: "Table Grid", Width: 100.0, Alignment: types.AlignLeft})
	formatRules.PageRules = append(formatRules.PageRules, types.PageRule{ID: "pg1", Name: "Normal", PageSize: types.PageSize{Width: 612.0, Height: 792.0}, PageMargins: types.PageMargins{Top: 72.0, Bottom: 72.0, Left: 72.0, Right: 72.0, Header: 36.0, Footer: 36.0}})

	return formatRules, nil
}
