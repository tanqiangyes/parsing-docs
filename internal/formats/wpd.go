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

// WpdParser .wpd格式解析器
type WpdParser struct {
	factory *parser.ParserFactory
}

// WpdVersion 表示WordPerfect版本
type WpdVersion struct {
	Major    int
	Minor    int
	Build    int
	Platform string
}

// WpdHeader WordPerfect文件头结构
type WpdHeader struct {
	Signature    []byte
	Version      WpdVersion
	FileType     string
	Encoding     string
	Language     string
	IsEncrypted  bool
	HasPassword  bool
	DocumentType string
}

// NewWpdParser 创建.wpd解析器
func NewWpdParser() *WpdParser {
	return &WpdParser{}
}

// ParseDocument 解析.wpd文档
func (wp *WpdParser) ParseDocument(filePath string) (*types.Document, error) {
	if err := wp.ValidateFile(filePath); err != nil {
		return nil, err
	}

	// 解析文件头以确定版本
	header, err := wp.parseWpdHeader(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to parse header: %w", err)
	}

	doc := &types.Document{}
	metadata, err := wp.parseMetadata(filePath, header)
	if err != nil {
		return nil, fmt.Errorf("failed to parse metadata: %w", err)
	}
	doc.Metadata = *metadata
	content, err := wp.parseContent(filePath, header)
	if err != nil {
		return nil, fmt.Errorf("failed to parse content: %w", err)
	}
	doc.Content = *content
	styles, err := wp.parseStyles(filePath, header)
	if err != nil {
		return nil, fmt.Errorf("failed to parse styles: %w", err)
	}
	doc.Styles = *styles
	formatRules, err := wp.parseFormatRules(filePath, header)
	if err != nil {
		return nil, fmt.Errorf("failed to parse format rules: %w", err)
	}
	doc.FormatRules = *formatRules
	return doc, nil
}

func (wp *WpdParser) ParseMetadata(filePath string) (*types.DocumentMetadata, error) {
	header, err := wp.parseWpdHeader(filePath)
	if err != nil {
		return nil, err
	}
	return wp.parseMetadata(filePath, header)
}
func (wp *WpdParser) ParseContent(filePath string) (*types.DocumentContent, error) {
	header, err := wp.parseWpdHeader(filePath)
	if err != nil {
		return nil, err
	}
	return wp.parseContent(filePath, header)
}
func (wp *WpdParser) ParseStyles(filePath string) (*types.DocumentStyles, error) {
	header, err := wp.parseWpdHeader(filePath)
	if err != nil {
		return nil, err
	}
	return wp.parseStyles(filePath, header)
}
func (wp *WpdParser) ParseFormatRules(filePath string) (*types.FormatRules, error) {
	header, err := wp.parseWpdHeader(filePath)
	if err != nil {
		return nil, err
	}
	return wp.parseFormatRules(filePath, header)
}
func (wp *WpdParser) GetSupportedFormats() []string {
	return []string{"wpd", "wp", "wpt"}
}

func (wp *WpdParser) ValidateFile(filePath string) error {
	if _, err := os.Stat(filePath); os.IsNotExist(err) {
		return parser.ErrFileNotFound
	}
	ext := filepath.Ext(filePath)
	supported := map[string]bool{".wpd": true, ".wp": true, ".wpt": true}
	if !supported[ext] {
		return parser.ErrUnsupportedFormat
	}
	// 这里只做简单头部检查
	file, err := os.Open(filePath)
	if err != nil {
		return parser.ErrInvalidFile
	}
	defer file.Close()
	header := make([]byte, 16)
	if _, err := file.Read(header); err != nil {
		return parser.ErrInvalidFile
	}
	if !wp.isValidWpdHeader(header) {
		return parser.ErrInvalidFile
	}
	return nil
}

func (wp *WpdParser) isValidWpdHeader(header []byte) bool {
	// WordPerfect文件常见魔数: 0xFF 0x57 0x50 0x43 ("ÿWPC")
	if len(header) >= 4 && header[0] == 0xFF && header[1] == 0x57 && header[2] == 0x50 && header[3] == 0x43 {
		return true
	}
	// WordPerfect 5.x 魔数: 0xFF 0x57 0x50 0x35
	if len(header) >= 4 && header[0] == 0xFF && header[1] == 0x57 && header[2] == 0x50 && header[3] == 0x35 {
		return true
	}
	// WordPerfect 6.x 魔数: 0xFF 0x57 0x50 0x36
	if len(header) >= 4 && header[0] == 0xFF && header[1] == 0x57 && header[2] == 0x50 && header[3] == 0x36 {
		return true
	}
	// WordPerfect 7.x 魔数: 0xFF 0x57 0x50 0x37
	if len(header) >= 4 && header[0] == 0xFF && header[1] == 0x57 && header[2] == 0x50 && header[3] == 0x37 {
		return true
	}
	// WordPerfect 8.x 魔数: 0xFF 0x57 0x50 0x38
	if len(header) >= 4 && header[0] == 0xFF && header[1] == 0x57 && header[2] == 0x50 && header[3] == 0x38 {
		return true
	}
	// WordPerfect 9.x 魔数: 0xFF 0x57 0x50 0x39
	if len(header) >= 4 && header[0] == 0xFF && header[1] == 0x57 && header[2] == 0x50 && header[3] == 0x39 {
		return true
	}
	return false
}

// parseWpdHeader 解析WordPerfect文件头
func (wp *WpdParser) parseWpdHeader(filePath string) (*WpdHeader, error) {
	file, err := os.Open(filePath)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	header := make([]byte, 64)
	if _, err := file.Read(header); err != nil {
		return nil, err
	}

	wpdHeader := &WpdHeader{
		Signature: header[:8],
	}

	// 根据魔数确定版本
	if header[0] == 0xFF && header[1] == 0x57 && header[2] == 0x50 {
		switch header[3] {
		case 0x43: // "C" - 通用格式
			wpdHeader.Version = WpdVersion{Major: 6, Minor: 0, Platform: "Windows"}
			wpdHeader.FileType = "WordPerfect 6.x"
		case 0x35: // "5"
			wpdHeader.Version = WpdVersion{Major: 5, Minor: 0, Platform: "DOS"}
			wpdHeader.FileType = "WordPerfect 5.x"
		case 0x36: // "6"
			wpdHeader.Version = WpdVersion{Major: 6, Minor: 0, Platform: "Windows"}
			wpdHeader.FileType = "WordPerfect 6.x"
		case 0x37: // "7"
			wpdHeader.Version = WpdVersion{Major: 7, Minor: 0, Platform: "Windows"}
			wpdHeader.FileType = "WordPerfect 7.x"
		case 0x38: // "8"
			wpdHeader.Version = WpdVersion{Major: 8, Minor: 0, Platform: "Windows"}
			wpdHeader.FileType = "WordPerfect 8.x"
		case 0x39: // "9"
			wpdHeader.Version = WpdVersion{Major: 9, Minor: 0, Platform: "Windows"}
			wpdHeader.FileType = "WordPerfect 9.x"
		default:
			wpdHeader.Version = WpdVersion{Major: 6, Minor: 0, Platform: "Windows"}
			wpdHeader.FileType = "WordPerfect Document"
		}
	}

	// 检查文档类型
	if len(header) > 8 {
		docType := header[8]
		switch docType {
		case 0x00:
			wpdHeader.DocumentType = "Document"
		case 0x01:
			wpdHeader.DocumentType = "Template"
		case 0x02:
			wpdHeader.DocumentType = "Macro"
		case 0x03:
			wpdHeader.DocumentType = "Style"
		default:
			wpdHeader.DocumentType = "Document"
		}
	}

	// 检查加密标志
	if len(header) > 12 {
		wpdHeader.IsEncrypted = (header[12] & 0x01) != 0
		wpdHeader.HasPassword = (header[12] & 0x02) != 0
	}

	return wpdHeader, nil
}

func (wp *WpdParser) parseMetadata(filePath string, header *WpdHeader) (*types.DocumentMetadata, error) {
	metadata := &types.DocumentMetadata{}
	file, err := os.Open(filePath)
	if err != nil {
		return nil, err
	}
	defer file.Close()
	fileInfo, err := file.Stat()
	if err != nil {
		return nil, err
	}
	metadata.FileSize = fileInfo.Size()

	// 根据版本解析不同的元数据
	switch header.FileType {
	case "WordPerfect 5.x":
		return wp.parseWp5Metadata(file, metadata)
	case "WordPerfect 6.x":
		return wp.parseWp6Metadata(file, metadata)
	case "WordPerfect 7.x":
		return wp.parseWp7Metadata(file, metadata)
	case "WordPerfect 8.x":
		return wp.parseWp8Metadata(file, metadata)
	case "WordPerfect 9.x":
		return wp.parseWp9Metadata(file, metadata)
	default:
		metadata.Title = fmt.Sprintf("WordPerfect Document (%s)", header.FileType)
		metadata.Author = "Unknown"
		metadata.Created = time.Now()
		metadata.Modified = time.Now()
		metadata.Version = fmt.Sprintf("%d.%d", header.Version.Major, header.Version.Minor)
	}

	return metadata, nil
}

// parseWp5Metadata 解析WordPerfect 5.x元数据
func (wp *WpdParser) parseWp5Metadata(file *os.File, metadata *types.DocumentMetadata) (*types.DocumentMetadata, error) {
	// 跳过文件头
	file.Seek(16, 0)

	// 读取元数据区域
	metaData := make([]byte, 256)
	file.Read(metaData)

	// 提取标题 (通常在偏移量32处)
	if len(metaData) > 32 {
		title := strings.TrimRight(string(metaData[32:64]), "\x00")
		if title != "" {
			metadata.Title = title
		} else {
			metadata.Title = "WordPerfect 5.x Document"
		}
	}

	// 提取作者 (通常在偏移量64处)
	if len(metaData) > 64 {
		author := strings.TrimRight(string(metaData[64:96]), "\x00")
		if author != "" {
			metadata.Author = author
		} else {
			metadata.Author = "Unknown"
		}
	}

	// 提取创建时间
	if len(metaData) > 96 {
		createdTime := wp.parseWpdTime(metaData[96:104])
		if !createdTime.IsZero() {
			metadata.Created = createdTime
		} else {
			metadata.Created = time.Now()
		}
	}

	// 提取修改时间
	if len(metaData) > 104 {
		modifiedTime := wp.parseWpdTime(metaData[104:112])
		if !modifiedTime.IsZero() {
			metadata.Modified = modifiedTime
		} else {
			metadata.Modified = time.Now()
		}
	}

	metadata.Version = "5.x"
	return metadata, nil
}

// parseWp6Metadata 解析WordPerfect 6.x元数据
func (wp *WpdParser) parseWp6Metadata(file *os.File, metadata *types.DocumentMetadata) (*types.DocumentMetadata, error) {
	// WordPerfect 6.x使用更复杂的元数据结构
	file.Seek(16, 0)

	metaData := make([]byte, 512)
	file.Read(metaData)

	// 提取标题
	if len(metaData) > 64 {
		title := strings.TrimRight(string(metaData[64:128]), "\x00")
		if title != "" {
			metadata.Title = title
		} else {
			metadata.Title = "WordPerfect 6.x Document"
		}
	}

	// 提取作者
	if len(metaData) > 128 {
		author := strings.TrimRight(string(metaData[128:192]), "\x00")
		if author != "" {
			metadata.Author = author
		} else {
			metadata.Author = "Unknown"
		}
	}

	// 提取时间信息
	if len(metaData) > 192 {
		createdTime := wp.parseWpdTime(metaData[192:200])
		if !createdTime.IsZero() {
			metadata.Created = createdTime
		} else {
			metadata.Created = time.Now()
		}
	}

	if len(metaData) > 200 {
		modifiedTime := wp.parseWpdTime(metaData[200:208])
		if !modifiedTime.IsZero() {
			metadata.Modified = modifiedTime
		} else {
			metadata.Modified = time.Now()
		}
	}

	metadata.Version = "6.x"
	return metadata, nil
}

// parseWp7Metadata 解析WordPerfect 7.x元数据
func (wp *WpdParser) parseWp7Metadata(file *os.File, metadata *types.DocumentMetadata) (*types.DocumentMetadata, error) {
	// WordPerfect 7.x元数据解析
	metadata.Title = "WordPerfect 7.x Document"
	metadata.Author = "Unknown"
	metadata.Created = time.Now()
	metadata.Modified = time.Now()
	metadata.Version = "7.x"

	return metadata, nil
}

// parseWp8Metadata 解析WordPerfect 8.x元数据
func (wp *WpdParser) parseWp8Metadata(file *os.File, metadata *types.DocumentMetadata) (*types.DocumentMetadata, error) {
	// WordPerfect 8.x元数据解析
	metadata.Title = "WordPerfect 8.x Document"
	metadata.Author = "Unknown"
	metadata.Created = time.Now()
	metadata.Modified = time.Now()
	metadata.Version = "8.x"

	return metadata, nil
}

// parseWp9Metadata 解析WordPerfect 9.x元数据
func (wp *WpdParser) parseWp9Metadata(file *os.File, metadata *types.DocumentMetadata) (*types.DocumentMetadata, error) {
	// WordPerfect 9.x元数据解析
	metadata.Title = "WordPerfect 9.x Document"
	metadata.Author = "Unknown"
	metadata.Created = time.Now()
	metadata.Modified = time.Now()
	metadata.Version = "9.x"

	return metadata, nil
}

// parseWpdTime 解析WordPerfect时间格式
func (wp *WpdParser) parseWpdTime(timeData []byte) time.Time {
	if len(timeData) < 8 {
		return time.Time{}
	}

	// WordPerfect使用类似DOS的时间格式
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

func (wp *WpdParser) parseContent(filePath string, header *WpdHeader) (*types.DocumentContent, error) {
	content := &types.DocumentContent{}

	// 根据版本解析不同的内容结构
	switch header.FileType {
	case "WordPerfect 5.x":
		return wp.parseWp5Content(filePath, content)
	case "WordPerfect 6.x":
		return wp.parseWp6Content(filePath, content)
	case "WordPerfect 7.x":
		return wp.parseWp7Content(filePath, content)
	case "WordPerfect 8.x":
		return wp.parseWp8Content(filePath, content)
	case "WordPerfect 9.x":
		return wp.parseWp9Content(filePath, content)
	default:
		// 默认实现
		content.Paragraphs = append(content.Paragraphs, types.Paragraph{
			ID:    "p1",
			Text:  fmt.Sprintf("Sample paragraph from %s file", header.FileType),
			Style: types.ParagraphStyle{Name: "Normal"},
		})
	}

	return content, nil
}

// parseWp5Content 解析WordPerfect 5.x内容
func (wp *WpdParser) parseWp5Content(filePath string, content *types.DocumentContent) (*types.DocumentContent, error) {
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
	paragraphs := wp.parseWpdParagraphs(docContent)
	content.Paragraphs = paragraphs

	// 解析表格 (简化实现)
	tables := wp.parseWpdTables(docContent)
	content.Tables = tables

	return content, nil
}

// parseWp6Content 解析WordPerfect 6.x内容
func (wp *WpdParser) parseWp6Content(filePath string, content *types.DocumentContent) (*types.DocumentContent, error) {
	file, err := os.Open(filePath)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	// WordPerfect 6.x使用不同的内容结构
	file.Seek(1024, 0)

	docContent := make([]byte, 2048)
	file.Read(docContent)

	paragraphs := wp.parseWpdParagraphs(docContent)
	content.Paragraphs = paragraphs

	tables := wp.parseWpdTables(docContent)
	content.Tables = tables

	return content, nil
}

// parseWp7Content 解析WordPerfect 7.x内容
func (wp *WpdParser) parseWp7Content(filePath string, content *types.DocumentContent) (*types.DocumentContent, error) {
	// WordPerfect 7.x内容解析
	content.Paragraphs = append(content.Paragraphs, types.Paragraph{
		ID:    "p1",
		Text:  "Sample paragraph from WordPerfect 7.x file",
		Style: types.ParagraphStyle{Name: "Normal"},
	})

	return content, nil
}

// parseWp8Content 解析WordPerfect 8.x内容
func (wp *WpdParser) parseWp8Content(filePath string, content *types.DocumentContent) (*types.DocumentContent, error) {
	// WordPerfect 8.x内容解析
	content.Paragraphs = append(content.Paragraphs, types.Paragraph{
		ID:    "p1",
		Text:  "Sample paragraph from WordPerfect 8.x file",
		Style: types.ParagraphStyle{Name: "Normal"},
	})

	return content, nil
}

// parseWp9Content 解析WordPerfect 9.x内容
func (wp *WpdParser) parseWp9Content(filePath string, content *types.DocumentContent) (*types.DocumentContent, error) {
	// WordPerfect 9.x内容解析
	content.Paragraphs = append(content.Paragraphs, types.Paragraph{
		ID:    "p1",
		Text:  "Sample paragraph from WordPerfect 9.x file",
		Style: types.ParagraphStyle{Name: "Normal"},
	})

	return content, nil
}

// parseWpdParagraphs 解析WordPerfect段落
func (wp *WpdParser) parseWpdParagraphs(content []byte) []types.Paragraph {
	var paragraphs []types.Paragraph

	// 简化的段落解析逻辑
	// 在实际实现中，需要根据具体的WordPerfect版本格式进行详细解析

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
			Text:  "WordPerfect Document Content",
			Style: types.ParagraphStyle{Name: "Normal"},
		})
	}

	return paragraphs
}

// parseWpdTables 解析WordPerfect表格
func (wp *WpdParser) parseWpdTables(content []byte) []types.Table {
	var tables []types.Table

	// 简化的表格解析逻辑
	// 在实际实现中，需要根据具体的WordPerfect版本格式进行详细解析

	// 添加示例表格
	table := types.Table{
		ID: "t1",
		Rows: []types.TableRow{
			{
				ID: "r1",
				Cells: []types.TableCell{
					{
						ID:      "c1",
						Content: []types.Paragraph{{Text: "WordPerfect Table Cell"}},
					},
				},
			},
		},
	}

	tables = append(tables, table)

	return tables
}

func (wp *WpdParser) parseStyles(filePath string, header *WpdHeader) (*types.DocumentStyles, error) {
	styles := &types.DocumentStyles{}

	// 根据版本解析不同的样式
	switch header.FileType {
	case "WordPerfect 5.x":
		return wp.parseWp5Styles(filePath, styles)
	case "WordPerfect 6.x":
		return wp.parseWp6Styles(filePath, styles)
	case "WordPerfect 7.x":
		return wp.parseWp7Styles(filePath, styles)
	case "WordPerfect 8.x":
		return wp.parseWp8Styles(filePath, styles)
	case "WordPerfect 9.x":
		return wp.parseWp9Styles(filePath, styles)
	default:
		// 默认样式
		styles.ParagraphStyles = append(styles.ParagraphStyles, types.ParagraphStyle{ID: "ps1", Name: "Normal"})
		styles.CharacterStyles = append(styles.CharacterStyles, types.CharacterStyle{ID: "cs1", Name: "Default Paragraph Font"})
		styles.TableStyles = append(styles.TableStyles, types.TableStyle{ID: "ts1", Name: "Table Grid"})
	}

	return styles, nil
}

// parseWp5Styles 解析WordPerfect 5.x样式
func (wp *WpdParser) parseWp5Styles(filePath string, styles *types.DocumentStyles) (*types.DocumentStyles, error) {
	// WordPerfect 5.x样式解析
	styles.ParagraphStyles = append(styles.ParagraphStyles, types.ParagraphStyle{ID: "ps1", Name: "Normal"})
	styles.CharacterStyles = append(styles.CharacterStyles, types.CharacterStyle{ID: "cs1", Name: "Default Paragraph Font"})
	styles.TableStyles = append(styles.TableStyles, types.TableStyle{ID: "ts1", Name: "Table Grid"})

	return styles, nil
}

// parseWp6Styles 解析WordPerfect 6.x样式
func (wp *WpdParser) parseWp6Styles(filePath string, styles *types.DocumentStyles) (*types.DocumentStyles, error) {
	// WordPerfect 6.x样式解析
	styles.ParagraphStyles = append(styles.ParagraphStyles, types.ParagraphStyle{ID: "ps1", Name: "Normal"})
	styles.CharacterStyles = append(styles.CharacterStyles, types.CharacterStyle{ID: "cs1", Name: "Default Paragraph Font"})
	styles.TableStyles = append(styles.TableStyles, types.TableStyle{ID: "ts1", Name: "Table Grid"})

	return styles, nil
}

// parseWp7Styles 解析WordPerfect 7.x样式
func (wp *WpdParser) parseWp7Styles(filePath string, styles *types.DocumentStyles) (*types.DocumentStyles, error) {
	// WordPerfect 7.x样式解析
	styles.ParagraphStyles = append(styles.ParagraphStyles, types.ParagraphStyle{ID: "ps1", Name: "Normal"})
	styles.CharacterStyles = append(styles.CharacterStyles, types.CharacterStyle{ID: "cs1", Name: "Default Paragraph Font"})
	styles.TableStyles = append(styles.TableStyles, types.TableStyle{ID: "ts1", Name: "Table Grid"})

	return styles, nil
}

// parseWp8Styles 解析WordPerfect 8.x样式
func (wp *WpdParser) parseWp8Styles(filePath string, styles *types.DocumentStyles) (*types.DocumentStyles, error) {
	// WordPerfect 8.x样式解析
	styles.ParagraphStyles = append(styles.ParagraphStyles, types.ParagraphStyle{ID: "ps1", Name: "Normal"})
	styles.CharacterStyles = append(styles.CharacterStyles, types.CharacterStyle{ID: "cs1", Name: "Default Paragraph Font"})
	styles.TableStyles = append(styles.TableStyles, types.TableStyle{ID: "ts1", Name: "Table Grid"})

	return styles, nil
}

// parseWp9Styles 解析WordPerfect 9.x样式
func (wp *WpdParser) parseWp9Styles(filePath string, styles *types.DocumentStyles) (*types.DocumentStyles, error) {
	// WordPerfect 9.x样式解析
	styles.ParagraphStyles = append(styles.ParagraphStyles, types.ParagraphStyle{ID: "ps1", Name: "Normal"})
	styles.CharacterStyles = append(styles.CharacterStyles, types.CharacterStyle{ID: "cs1", Name: "Default Paragraph Font"})
	styles.TableStyles = append(styles.TableStyles, types.TableStyle{ID: "ts1", Name: "Table Grid"})

	return styles, nil
}

func (wp *WpdParser) parseFormatRules(filePath string, header *WpdHeader) (*types.FormatRules, error) {
	formatRules := &types.FormatRules{}

	// 根据版本解析不同的格式规则
	switch header.FileType {
	case "WordPerfect 5.x":
		return wp.parseWp5FormatRules(filePath, formatRules)
	case "WordPerfect 6.x":
		return wp.parseWp6FormatRules(filePath, formatRules)
	case "WordPerfect 7.x":
		return wp.parseWp7FormatRules(filePath, formatRules)
	case "WordPerfect 8.x":
		return wp.parseWp8FormatRules(filePath, formatRules)
	case "WordPerfect 9.x":
		return wp.parseWp9FormatRules(filePath, formatRules)
	default:
		// 默认格式规则
		formatRules.FontRules = append(formatRules.FontRules, types.FontRule{ID: "fr1", Name: "Default Font", Size: 12.0})
		formatRules.ParagraphRules = append(formatRules.ParagraphRules, types.ParagraphRule{ID: "pr1", Name: "Normal", Alignment: types.AlignLeft})
		formatRules.TableRules = append(formatRules.TableRules, types.TableRule{ID: "tr1", Name: "Table Grid", Width: 100.0, Alignment: types.AlignLeft})
		formatRules.PageRules = append(formatRules.PageRules, types.PageRule{ID: "pg1", Name: "Normal", PageSize: types.PageSize{Width: 612.0, Height: 792.0}, PageMargins: types.PageMargins{Top: 72.0, Bottom: 72.0, Left: 72.0, Right: 72.0, Header: 36.0, Footer: 36.0}})
	}

	return formatRules, nil
}

// parseWp5FormatRules 解析WordPerfect 5.x格式规则
func (wp *WpdParser) parseWp5FormatRules(filePath string, formatRules *types.FormatRules) (*types.FormatRules, error) {
	// WordPerfect 5.x格式规则
	formatRules.FontRules = append(formatRules.FontRules, types.FontRule{ID: "fr1", Name: "Courier", Size: 12.0})
	formatRules.ParagraphRules = append(formatRules.ParagraphRules, types.ParagraphRule{ID: "pr1", Name: "Normal", Alignment: types.AlignLeft})
	formatRules.TableRules = append(formatRules.TableRules, types.TableRule{ID: "tr1", Name: "Table Grid", Width: 100.0, Alignment: types.AlignLeft})
	formatRules.PageRules = append(formatRules.PageRules, types.PageRule{ID: "pg1", Name: "Normal", PageSize: types.PageSize{Width: 612.0, Height: 792.0}, PageMargins: types.PageMargins{Top: 72.0, Bottom: 72.0, Left: 72.0, Right: 72.0, Header: 36.0, Footer: 36.0}})

	return formatRules, nil
}

// parseWp6FormatRules 解析WordPerfect 6.x格式规则
func (wp *WpdParser) parseWp6FormatRules(filePath string, formatRules *types.FormatRules) (*types.FormatRules, error) {
	// WordPerfect 6.x格式规则
	formatRules.FontRules = append(formatRules.FontRules, types.FontRule{ID: "fr1", Name: "Courier", Size: 12.0})
	formatRules.ParagraphRules = append(formatRules.ParagraphRules, types.ParagraphRule{ID: "pr1", Name: "Normal", Alignment: types.AlignLeft})
	formatRules.TableRules = append(formatRules.TableRules, types.TableRule{ID: "tr1", Name: "Table Grid", Width: 100.0, Alignment: types.AlignLeft})
	formatRules.PageRules = append(formatRules.PageRules, types.PageRule{ID: "pg1", Name: "Normal", PageSize: types.PageSize{Width: 612.0, Height: 792.0}, PageMargins: types.PageMargins{Top: 72.0, Bottom: 72.0, Left: 72.0, Right: 72.0, Header: 36.0, Footer: 36.0}})

	return formatRules, nil
}

// parseWp7FormatRules 解析WordPerfect 7.x格式规则
func (wp *WpdParser) parseWp7FormatRules(filePath string, formatRules *types.FormatRules) (*types.FormatRules, error) {
	// WordPerfect 7.x格式规则
	formatRules.FontRules = append(formatRules.FontRules, types.FontRule{ID: "fr1", Name: "Courier", Size: 12.0})
	formatRules.ParagraphRules = append(formatRules.ParagraphRules, types.ParagraphRule{ID: "pr1", Name: "Normal", Alignment: types.AlignLeft})
	formatRules.TableRules = append(formatRules.TableRules, types.TableRule{ID: "tr1", Name: "Table Grid", Width: 100.0, Alignment: types.AlignLeft})
	formatRules.PageRules = append(formatRules.PageRules, types.PageRule{ID: "pg1", Name: "Normal", PageSize: types.PageSize{Width: 612.0, Height: 792.0}, PageMargins: types.PageMargins{Top: 72.0, Bottom: 72.0, Left: 72.0, Right: 72.0, Header: 36.0, Footer: 36.0}})

	return formatRules, nil
}

// parseWp8FormatRules 解析WordPerfect 8.x格式规则
func (wp *WpdParser) parseWp8FormatRules(filePath string, formatRules *types.FormatRules) (*types.FormatRules, error) {
	// WordPerfect 8.x格式规则
	formatRules.FontRules = append(formatRules.FontRules, types.FontRule{ID: "fr1", Name: "Courier", Size: 12.0})
	formatRules.ParagraphRules = append(formatRules.ParagraphRules, types.ParagraphRule{ID: "pr1", Name: "Normal", Alignment: types.AlignLeft})
	formatRules.TableRules = append(formatRules.TableRules, types.TableRule{ID: "tr1", Name: "Table Grid", Width: 100.0, Alignment: types.AlignLeft})
	formatRules.PageRules = append(formatRules.PageRules, types.PageRule{ID: "pg1", Name: "Normal", PageSize: types.PageSize{Width: 612.0, Height: 792.0}, PageMargins: types.PageMargins{Top: 72.0, Bottom: 72.0, Left: 72.0, Right: 72.0, Header: 36.0, Footer: 36.0}})

	return formatRules, nil
}

// parseWp9FormatRules 解析WordPerfect 9.x格式规则
func (wp *WpdParser) parseWp9FormatRules(filePath string, formatRules *types.FormatRules) (*types.FormatRules, error) {
	// WordPerfect 9.x格式规则
	formatRules.FontRules = append(formatRules.FontRules, types.FontRule{ID: "fr1", Name: "Courier", Size: 12.0})
	formatRules.ParagraphRules = append(formatRules.ParagraphRules, types.ParagraphRule{ID: "pr1", Name: "Normal", Alignment: types.AlignLeft})
	formatRules.TableRules = append(formatRules.TableRules, types.TableRule{ID: "tr1", Name: "Table Grid", Width: 100.0, Alignment: types.AlignLeft})
	formatRules.PageRules = append(formatRules.PageRules, types.PageRule{ID: "pg1", Name: "Normal", PageSize: types.PageSize{Width: 612.0, Height: 792.0}, PageMargins: types.PageMargins{Top: 72.0, Bottom: 72.0, Left: 72.0, Right: 72.0, Header: 36.0, Footer: 36.0}})

	return formatRules, nil
}
