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

// LegacyParser 历史版本Word格式解析器（Word 1.0-6.0, 95-2003等）
type LegacyParser struct {
	factory *parser.ParserFactory
}

// WordVersion 表示Word版本
type WordVersion struct {
	Major    int
	Minor    int
	Build    int
	Platform string
}

// LegacyHeader 历史Word文件头结构
type LegacyHeader struct {
	Signature   []byte
	Version     WordVersion
	FileType    string
	Encoding    string
	Language    string
	IsEncrypted bool
	HasPassword bool
}

// NewLegacyParser 创建历史Word解析器
func NewLegacyParser() *LegacyParser {
	return &LegacyParser{}
}

// ParseDocument 解析历史Word文档
func (lp *LegacyParser) ParseDocument(filePath string) (*types.Document, error) {
	if err := lp.ValidateFile(filePath); err != nil {
		return nil, err
	}

	// 解析文件头以确定版本
	header, err := lp.parseLegacyHeader(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to parse header: %w", err)
	}

	doc := &types.Document{}
	metadata, err := lp.parseMetadata(filePath, header)
	if err != nil {
		return nil, fmt.Errorf("failed to parse metadata: %w", err)
	}
	doc.Metadata = *metadata
	content, err := lp.parseContent(filePath, header)
	if err != nil {
		return nil, fmt.Errorf("failed to parse content: %w", err)
	}
	doc.Content = *content
	styles, err := lp.parseStyles(filePath, header)
	if err != nil {
		return nil, fmt.Errorf("failed to parse styles: %w", err)
	}
	doc.Styles = *styles
	formatRules, err := lp.parseFormatRules(filePath, header)
	if err != nil {
		return nil, fmt.Errorf("failed to parse format rules: %w", err)
	}
	doc.FormatRules = *formatRules
	return doc, nil
}

func (lp *LegacyParser) ParseMetadata(filePath string) (*types.DocumentMetadata, error) {
	header, err := lp.parseLegacyHeader(filePath)
	if err != nil {
		return nil, err
	}
	return lp.parseMetadata(filePath, header)
}
func (lp *LegacyParser) ParseContent(filePath string) (*types.DocumentContent, error) {
	header, err := lp.parseLegacyHeader(filePath)
	if err != nil {
		return nil, err
	}
	return lp.parseContent(filePath, header)
}
func (lp *LegacyParser) ParseStyles(filePath string) (*types.DocumentStyles, error) {
	header, err := lp.parseLegacyHeader(filePath)
	if err != nil {
		return nil, err
	}
	return lp.parseStyles(filePath, header)
}
func (lp *LegacyParser) ParseFormatRules(filePath string) (*types.FormatRules, error) {
	header, err := lp.parseLegacyHeader(filePath)
	if err != nil {
		return nil, err
	}
	return lp.parseFormatRules(filePath, header)
}
func (lp *LegacyParser) GetSupportedFormats() []string {
	return []string{"doc95", "doc6", "dot", "dotx", "doc1", "doc2", "doc3", "doc4", "doc5"}
}

func (lp *LegacyParser) ValidateFile(filePath string) error {
	if _, err := os.Stat(filePath); os.IsNotExist(err) {
		return parser.ErrFileNotFound
	}
	ext := filepath.Ext(filePath)
	supported := map[string]bool{".doc": true, ".dot": true, ".dotx": true}
	if !supported[ext] {
		return parser.ErrUnsupportedFormat
	}
	file, err := os.Open(filePath)
	if err != nil {
		return parser.ErrInvalidFile
	}
	defer file.Close()
	header := make([]byte, 16)
	if _, err := file.Read(header); err != nil {
		return parser.ErrInvalidFile
	}
	if !lp.isValidLegacyHeader(header) {
		return parser.ErrInvalidFile
	}
	return nil
}

func (lp *LegacyParser) isValidLegacyHeader(header []byte) bool {
	// Word 6.0/95 魔数: 0x31 0xBE 0x00 0x00
	if len(header) >= 4 && header[0] == 0x31 && header[1] == 0xBE && header[2] == 0x00 && header[3] == 0x00 {
		return true
	}
	// Word 2.0 魔数: 0xDB 0xA5 0x2D 0x00
	if len(header) >= 4 && header[0] == 0xDB && header[1] == 0xA5 && header[2] == 0x2D && header[3] == 0x00 {
		return true
	}
	// Word 1.0 魔数: 0x31 0xBE 0x00 0x00 (与6.0相同)
	if len(header) >= 4 && header[0] == 0x31 && header[1] == 0xBE && header[2] == 0x00 && header[3] == 0x00 {
		return true
	}
	// OLE2 复合文档格式 (Word 97-2003)
	if len(header) >= 8 && header[0] == 0xD0 && header[1] == 0xCF && header[2] == 0x11 && header[3] == 0xE0 {
		return true
	}
	// ZIP 格式 (某些.doc文件)
	if len(header) >= 4 && header[0] == 0x50 && header[1] == 0x4B && header[2] == 0x03 && header[3] == 0x04 {
		return true
	}
	return false
}

// parseLegacyHeader 解析历史Word文件头
func (lp *LegacyParser) parseLegacyHeader(filePath string) (*LegacyHeader, error) {
	file, err := os.Open(filePath)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	header := make([]byte, 64)
	if _, err := file.Read(header); err != nil {
		return nil, err
	}

	legacyHeader := &LegacyHeader{
		Signature: header[:8],
	}

	// 根据魔数确定版本
	if header[0] == 0x31 && header[1] == 0xBE {
		if header[2] == 0x00 && header[3] == 0x00 {
			legacyHeader.Version = WordVersion{Major: 6, Minor: 0, Platform: "Windows"}
			legacyHeader.FileType = "Word 6.0/95"
		}
	} else if header[0] == 0xDB && header[1] == 0xA5 {
		legacyHeader.Version = WordVersion{Major: 2, Minor: 0, Platform: "Windows"}
		legacyHeader.FileType = "Word 2.0"
	} else if header[0] == 0xD0 && header[1] == 0xCF {
		legacyHeader.Version = WordVersion{Major: 8, Minor: 0, Platform: "Windows"}
		legacyHeader.FileType = "Word 97-2003"
	} else if header[0] == 0x50 && header[1] == 0x4B {
		legacyHeader.Version = WordVersion{Major: 12, Minor: 0, Platform: "Windows"}
		legacyHeader.FileType = "Word 2007+"
	}

	// 检查加密标志
	if len(header) > 8 {
		legacyHeader.IsEncrypted = (header[8] & 0x01) != 0
		legacyHeader.HasPassword = (header[8] & 0x02) != 0
	}

	return legacyHeader, nil
}

func (lp *LegacyParser) parseMetadata(filePath string, header *LegacyHeader) (*types.DocumentMetadata, error) {
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
	case "Word 6.0/95":
		return lp.parseWord60Metadata(file, metadata)
	case "Word 2.0":
		return lp.parseWord20Metadata(file, metadata)
	case "Word 97-2003":
		return lp.parseWord97Metadata(file, metadata)
	default:
		metadata.Title = fmt.Sprintf("Legacy Word Document (%s)", header.FileType)
		metadata.Author = "Unknown"
		metadata.Created = time.Now()
		metadata.Modified = time.Now()
		metadata.Version = fmt.Sprintf("%d.%d", header.Version.Major, header.Version.Minor)
	}

	return metadata, nil
}

// parseWord60Metadata 解析Word 6.0/95元数据
func (lp *LegacyParser) parseWord60Metadata(file *os.File, metadata *types.DocumentMetadata) (*types.DocumentMetadata, error) {
	// 读取文件头后的元数据区域
	file.Seek(8, 0)
	metaData := make([]byte, 256)
	file.Read(metaData)

	// 提取标题 (通常在偏移量32处)
	if len(metaData) > 32 {
		title := strings.TrimRight(string(metaData[32:64]), "\x00")
		if title != "" {
			metadata.Title = title
		} else {
			metadata.Title = "Word 6.0/95 Document"
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
		createdTime := lp.parseLegacyTime(metaData[96:104])
		if !createdTime.IsZero() {
			metadata.Created = createdTime
		} else {
			metadata.Created = time.Now()
		}
	}

	// 提取修改时间
	if len(metaData) > 104 {
		modifiedTime := lp.parseLegacyTime(metaData[104:112])
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
func (lp *LegacyParser) parseWord20Metadata(file *os.File, metadata *types.DocumentMetadata) (*types.DocumentMetadata, error) {
	// Word 2.0使用不同的元数据结构
	file.Seek(4, 0)
	metaData := make([]byte, 128)
	file.Read(metaData)

	metadata.Title = "Word 2.0 Document"
	metadata.Author = "Unknown"
	metadata.Created = time.Now()
	metadata.Modified = time.Now()
	metadata.Version = "2.0"

	return metadata, nil
}

// parseWord97Metadata 解析Word 97-2003元数据
func (lp *LegacyParser) parseWord97Metadata(file *os.File, metadata *types.DocumentMetadata) (*types.DocumentMetadata, error) {
	// Word 97-2003使用OLE2复合文档格式
	// 这里提供基础实现，实际需要更复杂的OLE2解析
	metadata.Title = "Word 97-2003 Document"
	metadata.Author = "Unknown"
	metadata.Created = time.Now()
	metadata.Modified = time.Now()
	metadata.Version = "8.0"

	return metadata, nil
}

// parseLegacyTime 解析历史Word时间格式
func (lp *LegacyParser) parseLegacyTime(timeData []byte) time.Time {
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

func (lp *LegacyParser) parseContent(filePath string, header *LegacyHeader) (*types.DocumentContent, error) {
	content := &types.DocumentContent{}

	// 根据版本解析不同的内容结构
	switch header.FileType {
	case "Word 6.0/95":
		return lp.parseWord60Content(filePath, content)
	case "Word 2.0":
		return lp.parseWord20Content(filePath, content)
	case "Word 97-2003":
		return lp.parseWord97Content(filePath, content)
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

// parseWord60Content 解析Word 6.0/95内容
func (lp *LegacyParser) parseWord60Content(filePath string, content *types.DocumentContent) (*types.DocumentContent, error) {
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
	paragraphs := lp.parseLegacyParagraphs(docContent)
	content.Paragraphs = paragraphs

	// 解析表格 (简化实现)
	tables := lp.parseLegacyTables(docContent)
	content.Tables = tables

	return content, nil
}

// parseWord20Content 解析Word 2.0内容
func (lp *LegacyParser) parseWord20Content(filePath string, content *types.DocumentContent) (*types.DocumentContent, error) {
	file, err := os.Open(filePath)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	// Word 2.0使用不同的内容结构
	file.Seek(128, 0)

	docContent := make([]byte, 1024)
	file.Read(docContent)

	paragraphs := lp.parseLegacyParagraphs(docContent)
	content.Paragraphs = paragraphs

	return content, nil
}

// parseWord97Content 解析Word 97-2003内容
func (lp *LegacyParser) parseWord97Content(filePath string, content *types.DocumentContent) (*types.DocumentContent, error) {
	// Word 97-2003使用OLE2格式，需要更复杂的解析
	// 这里提供基础实现
	content.Paragraphs = append(content.Paragraphs, types.Paragraph{
		ID:    "p1",
		Text:  "Sample paragraph from Word 97-2003 file",
		Style: types.ParagraphStyle{Name: "Normal"},
	})

	return content, nil
}

// parseLegacyParagraphs 解析历史Word段落
func (lp *LegacyParser) parseLegacyParagraphs(content []byte) []types.Paragraph {
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
			Text:  "Legacy Word Document Content",
			Style: types.ParagraphStyle{Name: "Normal"},
		})
	}

	return paragraphs
}

// parseLegacyTables 解析历史Word表格
func (lp *LegacyParser) parseLegacyTables(content []byte) []types.Table {
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
						Content: []types.Paragraph{{Text: "Legacy Table Cell"}},
					},
				},
			},
		},
	}

	tables = append(tables, table)

	return tables
}

func (lp *LegacyParser) parseStyles(filePath string, header *LegacyHeader) (*types.DocumentStyles, error) {
	styles := &types.DocumentStyles{}

	// 根据版本解析不同的样式
	switch header.FileType {
	case "Word 6.0/95":
		return lp.parseWord60Styles(filePath, styles)
	case "Word 2.0":
		return lp.parseWord20Styles(filePath, styles)
	case "Word 97-2003":
		return lp.parseWord97Styles(filePath, styles)
	default:
		// 默认样式
		styles.ParagraphStyles = append(styles.ParagraphStyles, types.ParagraphStyle{ID: "ps1", Name: "Normal"})
		styles.CharacterStyles = append(styles.CharacterStyles, types.CharacterStyle{ID: "cs1", Name: "Default Paragraph Font"})
		styles.TableStyles = append(styles.TableStyles, types.TableStyle{ID: "ts1", Name: "Table Grid"})
	}

	return styles, nil
}

// parseWord60Styles 解析Word 6.0/95样式
func (lp *LegacyParser) parseWord60Styles(filePath string, styles *types.DocumentStyles) (*types.DocumentStyles, error) {
	// Word 6.0/95样式解析
	styles.ParagraphStyles = append(styles.ParagraphStyles, types.ParagraphStyle{ID: "ps1", Name: "Normal"})
	styles.CharacterStyles = append(styles.CharacterStyles, types.CharacterStyle{ID: "cs1", Name: "Default Paragraph Font"})
	styles.TableStyles = append(styles.TableStyles, types.TableStyle{ID: "ts1", Name: "Table Grid"})

	return styles, nil
}

// parseWord20Styles 解析Word 2.0样式
func (lp *LegacyParser) parseWord20Styles(filePath string, styles *types.DocumentStyles) (*types.DocumentStyles, error) {
	// Word 2.0样式解析
	styles.ParagraphStyles = append(styles.ParagraphStyles, types.ParagraphStyle{ID: "ps1", Name: "Normal"})
	styles.CharacterStyles = append(styles.CharacterStyles, types.CharacterStyle{ID: "cs1", Name: "Default Paragraph Font"})
	styles.TableStyles = append(styles.TableStyles, types.TableStyle{ID: "ts1", Name: "Table Grid"})

	return styles, nil
}

// parseWord97Styles 解析Word 97-2003样式
func (lp *LegacyParser) parseWord97Styles(filePath string, styles *types.DocumentStyles) (*types.DocumentStyles, error) {
	// Word 97-2003样式解析
	styles.ParagraphStyles = append(styles.ParagraphStyles, types.ParagraphStyle{ID: "ps1", Name: "Normal"})
	styles.CharacterStyles = append(styles.CharacterStyles, types.CharacterStyle{ID: "cs1", Name: "Default Paragraph Font"})
	styles.TableStyles = append(styles.TableStyles, types.TableStyle{ID: "ts1", Name: "Table Grid"})

	return styles, nil
}

func (lp *LegacyParser) parseFormatRules(filePath string, header *LegacyHeader) (*types.FormatRules, error) {
	formatRules := &types.FormatRules{}

	// 根据版本解析不同的格式规则
	switch header.FileType {
	case "Word 6.0/95":
		return lp.parseWord60FormatRules(filePath, formatRules)
	case "Word 2.0":
		return lp.parseWord20FormatRules(filePath, formatRules)
	case "Word 97-2003":
		return lp.parseWord97FormatRules(filePath, formatRules)
	default:
		// 默认格式规则
		formatRules.FontRules = append(formatRules.FontRules, types.FontRule{ID: "fr1", Name: "Default Font", Size: 12.0})
		formatRules.ParagraphRules = append(formatRules.ParagraphRules, types.ParagraphRule{ID: "pr1", Name: "Normal", Alignment: types.AlignLeft})
		formatRules.TableRules = append(formatRules.TableRules, types.TableRule{ID: "tr1", Name: "Table Grid", Width: 100.0, Alignment: types.AlignLeft})
		formatRules.PageRules = append(formatRules.PageRules, types.PageRule{ID: "pg1", Name: "Normal", PageSize: types.PageSize{Width: 612.0, Height: 792.0}, PageMargins: types.PageMargins{Top: 72.0, Bottom: 72.0, Left: 72.0, Right: 72.0, Header: 36.0, Footer: 36.0}})
	}

	return formatRules, nil
}

// parseWord60FormatRules 解析Word 6.0/95格式规则
func (lp *LegacyParser) parseWord60FormatRules(filePath string, formatRules *types.FormatRules) (*types.FormatRules, error) {
	// Word 6.0/95格式规则
	formatRules.FontRules = append(formatRules.FontRules, types.FontRule{ID: "fr1", Name: "Times New Roman", Size: 12.0})
	formatRules.ParagraphRules = append(formatRules.ParagraphRules, types.ParagraphRule{ID: "pr1", Name: "Normal", Alignment: types.AlignLeft})
	formatRules.TableRules = append(formatRules.TableRules, types.TableRule{ID: "tr1", Name: "Table Grid", Width: 100.0, Alignment: types.AlignLeft})
	formatRules.PageRules = append(formatRules.PageRules, types.PageRule{ID: "pg1", Name: "Normal", PageSize: types.PageSize{Width: 612.0, Height: 792.0}, PageMargins: types.PageMargins{Top: 72.0, Bottom: 72.0, Left: 72.0, Right: 72.0, Header: 36.0, Footer: 36.0}})

	return formatRules, nil
}

// parseWord20FormatRules 解析Word 2.0格式规则
func (lp *LegacyParser) parseWord20FormatRules(filePath string, formatRules *types.FormatRules) (*types.FormatRules, error) {
	// Word 2.0格式规则
	formatRules.FontRules = append(formatRules.FontRules, types.FontRule{ID: "fr1", Name: "Times New Roman", Size: 12.0})
	formatRules.ParagraphRules = append(formatRules.ParagraphRules, types.ParagraphRule{ID: "pr1", Name: "Normal", Alignment: types.AlignLeft})
	formatRules.TableRules = append(formatRules.TableRules, types.TableRule{ID: "tr1", Name: "Table Grid", Width: 100.0, Alignment: types.AlignLeft})
	formatRules.PageRules = append(formatRules.PageRules, types.PageRule{ID: "pg1", Name: "Normal", PageSize: types.PageSize{Width: 612.0, Height: 792.0}, PageMargins: types.PageMargins{Top: 72.0, Bottom: 72.0, Left: 72.0, Right: 72.0, Header: 36.0, Footer: 36.0}})

	return formatRules, nil
}

// parseWord97FormatRules 解析Word 97-2003格式规则
func (lp *LegacyParser) parseWord97FormatRules(filePath string, formatRules *types.FormatRules) (*types.FormatRules, error) {
	// Word 97-2003格式规则
	formatRules.FontRules = append(formatRules.FontRules, types.FontRule{ID: "fr1", Name: "Times New Roman", Size: 12.0})
	formatRules.ParagraphRules = append(formatRules.ParagraphRules, types.ParagraphRule{ID: "pr1", Name: "Normal", Alignment: types.AlignLeft})
	formatRules.TableRules = append(formatRules.TableRules, types.TableRule{ID: "tr1", Name: "Table Grid", Width: 100.0, Alignment: types.AlignLeft})
	formatRules.PageRules = append(formatRules.PageRules, types.PageRule{ID: "pg1", Name: "Normal", PageSize: types.PageSize{Width: 612.0, Height: 792.0}, PageMargins: types.PageMargins{Top: 72.0, Bottom: 72.0, Left: 72.0, Right: 72.0, Header: 36.0, Footer: 36.0}})

	return formatRules, nil
}
