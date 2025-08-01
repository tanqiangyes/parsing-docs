package formats

import (
	"bufio"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
	"time"

	"docs-parser/internal/core/parser"
	"docs-parser/internal/core/types"
)

// RtfParser .rtf格式解析器
type RtfParser struct {
	factory *parser.ParserFactory
}

// NewRtfParser 创建.rtf解析器
func NewRtfParser() *RtfParser {
	return &RtfParser{}
}

// ParseDocument 解析.rtf文档
func (rp *RtfParser) ParseDocument(filePath string) (*types.Document, error) {
	// 验证文件
	if err := rp.ValidateFile(filePath); err != nil {
		return nil, err
	}

	doc := &types.Document{}

	// 解析元数据
	metadata, err := rp.parseMetadata(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to parse metadata: %w", err)
	}
	doc.Metadata = *metadata

	// 解析内容
	content, err := rp.parseContent(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to parse content: %w", err)
	}
	doc.Content = *content

	// 解析样式
	styles, err := rp.parseStyles(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to parse styles: %w", err)
	}
	doc.Styles = *styles

	// 解析格式规则
	formatRules, err := rp.parseFormatRules(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to parse format rules: %w", err)
	}
	doc.FormatRules = *formatRules

	return doc, nil
}

// ParseMetadata 解析元数据
func (rp *RtfParser) ParseMetadata(filePath string) (*types.DocumentMetadata, error) {
	return rp.parseMetadata(filePath)
}

// ParseContent 解析内容
func (rp *RtfParser) ParseContent(filePath string) (*types.DocumentContent, error) {
	return rp.parseContent(filePath)
}

// ParseStyles 解析样式
func (rp *RtfParser) ParseStyles(filePath string) (*types.DocumentStyles, error) {
	return rp.parseStyles(filePath)
}

// ParseFormatRules 解析格式规则
func (rp *RtfParser) ParseFormatRules(filePath string) (*types.FormatRules, error) {
	return rp.parseFormatRules(filePath)
}

// GetSupportedFormats 获取支持的格式
func (rp *RtfParser) GetSupportedFormats() []string {
	return []string{"rtf"}
}

// ValidateFile 验证文件格式
func (rp *RtfParser) ValidateFile(filePath string) error {
	// 检查文件是否存在
	if _, err := os.Stat(filePath); os.IsNotExist(err) {
		return parser.ErrFileNotFound
	}

	// 检查文件扩展名
	ext := filepath.Ext(filePath)
	if ext != ".rtf" {
		return parser.ErrUnsupportedFormat
	}

	// 检查文件头
	file, err := os.Open(filePath)
	if err != nil {
		return parser.ErrInvalidFile
	}
	defer file.Close()

	// 读取文件头
	reader := bufio.NewReader(file)
	header, err := reader.ReadString('\n')
	if err != nil {
		return parser.ErrInvalidFile
	}

	// 检查RTF文件头标识
	if !rp.isValidRtfHeader(header) {
		return parser.ErrInvalidFile
	}

	return nil
}

// isValidRtfHeader 检查是否为有效的.rtf文件头
func (rp *RtfParser) isValidRtfHeader(header string) bool {
	// RTF文件通常以{\rtf1开始
	rtfPattern := regexp.MustCompile(`^\\s*\\{\\s*\\\\rtf1`)
	return rtfPattern.MatchString(header)
}

// parseMetadata 解析元数据
func (rp *RtfParser) parseMetadata(filePath string) (*types.DocumentMetadata, error) {
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

	// 解析RTF文件的基本元数据
	content, err := rp.readFileContent(file)
	if err != nil {
		return nil, err
	}

	// 提取标题
	if title := rp.extractRtfValue(content, "title"); title != "" {
		metadata.Title = title
	} else {
		metadata.Title = "Document"
	}

	// 提取作者
	if author := rp.extractRtfValue(content, "author"); author != "" {
		metadata.Author = author
	} else {
		metadata.Author = "Unknown"
	}

	// 提取主题
	if subject := rp.extractRtfValue(content, "subject"); subject != "" {
		metadata.Subject = subject
	}

	// 提取关键词
	if keywords := rp.extractRtfValue(content, "keywords"); keywords != "" {
		metadata.Keywords = strings.Split(keywords, ",")
	}

	// 提取创建时间
	if created := rp.extractRtfValue(content, "creatim"); created != "" {
		if t, err := rp.parseRtfDate(created); err == nil {
			metadata.Created = t
		}
	}

	// 提取修改时间
	if modified := rp.extractRtfValue(content, "revtim"); modified != "" {
		if t, err := rp.parseRtfDate(modified); err == nil {
			metadata.Modified = t
		}
	}

	// 设置默认值
	if metadata.Created.IsZero() {
		metadata.Created = time.Now()
	}
	if metadata.Modified.IsZero() {
		metadata.Modified = time.Now()
	}
	metadata.Version = "1.0"

	return metadata, nil
}

// parseContent 解析文档内容
func (rp *RtfParser) parseContent(filePath string) (*types.DocumentContent, error) {
	content := &types.DocumentContent{}

	file, err := os.Open(filePath)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	// 读取文件内容
	rtfContent, err := rp.readFileContent(file)
	if err != nil {
		return nil, err
	}

	// 解析段落
	if err := rp.parseParagraphs(rtfContent, content); err != nil {
		return nil, err
	}

	// 解析表格
	if err := rp.parseTables(rtfContent, content); err != nil {
		return nil, err
	}

	return content, nil
}

// parseParagraphs 解析段落
func (rp *RtfParser) parseParagraphs(rtfContent string, content *types.DocumentContent) error {
	// 分割段落
	paragraphs := rp.splitRtfParagraphs(rtfContent)

	for i, p := range paragraphs {
		paragraph := types.Paragraph{
			ID:   fmt.Sprintf("p%d", i+1),
			Text: rp.extractTextFromRtf(p),
			Style: types.ParagraphStyle{
				Name: rp.extractParagraphStyle(p),
			},
		}

		// 解析文本运行
		runs := rp.parseRtfRuns(p)
		paragraph.Runs = runs

		// 解析段落格式
		rp.parseParagraphFormat(p, &paragraph)

		content.Paragraphs = append(content.Paragraphs, paragraph)
	}

	return nil
}

// parseTables 解析表格
func (rp *RtfParser) parseTables(rtfContent string, content *types.DocumentContent) error {
	// 查找表格标记
	tablePattern := regexp.MustCompile(`\\\\trowd.*?\\\\trowd`)
	tableMatches := tablePattern.FindAllString(rtfContent, -1)

	for i, tableMatch := range tableMatches {
		table := types.Table{
			ID: fmt.Sprintf("t%d", i+1),
		}

		// 解析表格行
		rows := rp.parseRtfTableRows(tableMatch)
		table.Rows = rows

		content.Tables = append(content.Tables, table)
	}

	return nil
}

// parseRtfTableRows 解析RTF表格行
func (rp *RtfParser) parseRtfTableRows(tableContent string) []types.TableRow {
	var rows []types.TableRow

	// 分割行
	rowPattern := regexp.MustCompile(`\\\\trowd.*?\\\\row`)
	rowMatches := rowPattern.FindAllString(tableContent, -1)

	for i, rowMatch := range rowMatches {
		row := types.TableRow{
			ID: fmt.Sprintf("r%d", i+1),
		}

		// 解析单元格
		cells := rp.parseRtfTableCells(rowMatch)
		row.Cells = cells

		rows = append(rows, row)
	}

	return rows
}

// parseRtfTableCells 解析RTF表格单元格
func (rp *RtfParser) parseRtfTableCells(rowContent string) []types.TableCell {
	var cells []types.TableCell

	// 分割单元格
	cellPattern := regexp.MustCompile(`\\\\cell.*?\\\\cell`)
	cellMatches := cellPattern.FindAllString(rowContent, -1)

	for i, cellMatch := range cellMatches {
		cell := types.TableCell{
			ID: fmt.Sprintf("c%d", i+1),
			Content: []types.Paragraph{
				{
					ID:   fmt.Sprintf("cp%d", i+1),
					Text: rp.extractTextFromRtf(cellMatch),
				},
			},
		}

		cells = append(cells, cell)
	}

	return cells
}

// parseRtfRuns 解析RTF文本运行
func (rp *RtfParser) parseRtfRuns(paragraphContent string) []types.TextRun {
	var runs []types.TextRun

	// 分割文本运行
	runPattern := regexp.MustCompile(`\\\\f\\d+.*?\\\\f0`)
	runMatches := runPattern.FindAllString(paragraphContent, -1)

	for i, runMatch := range runMatches {
		run := types.TextRun{
			ID:   fmt.Sprintf("r%d", i+1),
			Text: rp.extractTextFromRtf(runMatch),
			Font: rp.parseRtfFont(runMatch),
		}

		runs = append(runs, run)
	}

	return runs
}

// parseRtfFont 解析RTF字体信息
func (rp *RtfParser) parseRtfFont(runContent string) types.Font {
	font := types.Font{
		Name:   "Times New Roman",
		Size:   12.0,
		Bold:   false,
		Italic: false,
	}

	// 提取字体名称
	if fontName := rp.extractRtfValue(runContent, "f"); fontName != "" {
		font.Name = fontName
	}

	// 提取字体大小
	if fontSize := rp.extractRtfValue(runContent, "fs"); fontSize != "" {
		if size, err := strconv.ParseFloat(fontSize, 64); err == nil {
			font.Size = size / 2.0 // RTF字体大小通常是Word字体大小的2倍
		}
	}

	// 检查粗体
	if strings.Contains(runContent, "\\b") {
		font.Bold = true
	}

	// 检查斜体
	if strings.Contains(runContent, "\\i") {
		font.Italic = true
	}

	return font
}

// parseParagraphFormat 解析段落格式
func (rp *RtfParser) parseParagraphFormat(paragraphContent string, paragraph *types.Paragraph) {
	// 解析对齐方式
	if strings.Contains(paragraphContent, "\\qc") {
		paragraph.Alignment = types.AlignCenter
	} else if strings.Contains(paragraphContent, "\\qr") {
		paragraph.Alignment = types.AlignRight
	} else if strings.Contains(paragraphContent, "\\qj") {
		paragraph.Alignment = types.AlignJustify
	} else {
		paragraph.Alignment = types.AlignLeft
	}

	// 解析缩进
	if indent := rp.extractRtfValue(paragraphContent, "li"); indent != "" {
		if value, err := strconv.ParseFloat(indent, 64); err == nil {
			paragraph.Indentation.Left = value / 20.0 // 转换为磅
		}
	}

	if indent := rp.extractRtfValue(paragraphContent, "ri"); indent != "" {
		if value, err := strconv.ParseFloat(indent, 64); err == nil {
			paragraph.Indentation.Right = value / 20.0
		}
	}

	if indent := rp.extractRtfValue(paragraphContent, "fi"); indent != "" {
		if value, err := strconv.ParseFloat(indent, 64); err == nil {
			paragraph.Indentation.First = value / 20.0
		}
	}

	// 解析间距
	if spacing := rp.extractRtfValue(paragraphContent, "sb"); spacing != "" {
		if value, err := strconv.ParseFloat(spacing, 64); err == nil {
			paragraph.Spacing.Before = value / 20.0
		}
	}

	if spacing := rp.extractRtfValue(paragraphContent, "sa"); spacing != "" {
		if value, err := strconv.ParseFloat(spacing, 64); err == nil {
			paragraph.Spacing.After = value / 20.0
		}
	}

	if spacing := rp.extractRtfValue(paragraphContent, "sl"); spacing != "" {
		if value, err := strconv.ParseFloat(spacing, 64); err == nil {
			paragraph.Spacing.Line = value / 240.0 // 转换为行间距倍数
		}
	}
}

// parseStyles 解析样式
func (rp *RtfParser) parseStyles(filePath string) (*types.DocumentStyles, error) {
	styles := &types.DocumentStyles{}

	file, err := os.Open(filePath)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	// 读取文件内容
	content, err := rp.readFileContent(file)
	if err != nil {
		return nil, err
	}

	// 解析样式表
	styleTable := rp.extractStyleTable(content)
	rp.parseRtfStyleTable(styleTable, styles)

	return styles, nil
}

// parseFormatRules 解析格式规则
func (rp *RtfParser) parseFormatRules(filePath string) (*types.FormatRules, error) {
	formatRules := &types.FormatRules{}

	// 解析字体规则
	if err := rp.parseFontRules(filePath, formatRules); err != nil {
		return nil, err
	}

	// 解析段落规则
	if err := rp.parseParagraphRules(filePath, formatRules); err != nil {
		return nil, err
	}

	// 解析表格规则
	if err := rp.parseTableRules(filePath, formatRules); err != nil {
		return nil, err
	}

	// 解析页面规则
	if err := rp.parsePageRules(filePath, formatRules); err != nil {
		return nil, err
	}

	return formatRules, nil
}

// parseFontRules 解析字体规则
func (rp *RtfParser) parseFontRules(filePath string, formatRules *types.FormatRules) error {
	// 实现字体规则解析逻辑
	// 由于RTF格式相对简单，这里提供基础实现

	fontRule := types.FontRule{
		ID:     "fr1",
		Name:   "Default Font",
		Size:   12.0,
		Color:  types.Color{},
		Bold:   false,
		Italic: false,
	}

	formatRules.FontRules = append(formatRules.FontRules, fontRule)

	return nil
}

// parseParagraphRules 解析段落规则
func (rp *RtfParser) parseParagraphRules(filePath string, formatRules *types.FormatRules) error {
	// 实现段落规则解析逻辑
	// 由于RTF格式相对简单，这里提供基础实现

	paragraphRule := types.ParagraphRule{
		ID:        "pr1",
		Name:      "Normal",
		Alignment: types.AlignLeft,
		Indentation: types.Indentation{
			Left:    0.0,
			Right:   0.0,
			First:   0.0,
			Hanging: 0.0,
		},
		Spacing: types.Spacing{
			Before: 0.0,
			After:  0.0,
			Line:   1.0,
		},
	}

	formatRules.ParagraphRules = append(formatRules.ParagraphRules, paragraphRule)

	return nil
}

// parseTableRules 解析表格规则
func (rp *RtfParser) parseTableRules(filePath string, formatRules *types.FormatRules) error {
	// 实现表格规则解析逻辑
	// 由于RTF格式相对简单，这里提供基础实现

	tableRule := types.TableRule{
		ID:        "tr1",
		Name:      "Table Grid",
		Width:     100.0,
		Alignment: types.AlignLeft,
	}

	formatRules.TableRules = append(formatRules.TableRules, tableRule)

	return nil
}

// parsePageRules 解析页面规则
func (rp *RtfParser) parsePageRules(filePath string, formatRules *types.FormatRules) error {
	// 实现页面规则解析逻辑
	// 由于RTF格式相对简单，这里提供基础实现

	pageRule := types.PageRule{
		ID:   "pg1",
		Name: "Normal",
		PageSize: types.PageSize{
			Width:  612.0, // 8.5 inches
			Height: 792.0, // 11 inches
		},
		PageMargins: types.PageMargins{
			Top:    72.0, // 1 inch
			Bottom: 72.0, // 1 inch
			Left:   72.0, // 1 inch
			Right:  72.0, // 1 inch
			Header: 36.0, // 0.5 inch
			Footer: 36.0, // 0.5 inch
		},
	}

	formatRules.PageRules = append(formatRules.PageRules, pageRule)

	return nil
}

// 辅助方法

// readFileContent 读取文件内容
func (rp *RtfParser) readFileContent(file *os.File) (string, error) {
	content, err := os.ReadFile(file.Name())
	if err != nil {
		return "", err
	}
	return string(content), nil
}

// extractRtfValue 提取RTF值
func (rp *RtfParser) extractRtfValue(content, key string) string {
	pattern := regexp.MustCompile(fmt.Sprintf(`\\\\%s\\s*([^\\\\]+)`, key))
	matches := pattern.FindStringSubmatch(content)
	if len(matches) > 1 {
		return strings.TrimSpace(matches[1])
	}
	return ""
}

// parseRtfDate 解析RTF日期
func (rp *RtfParser) parseRtfDate(dateStr string) (time.Time, error) {
	// RTF日期格式通常是 \yr2023\mo12\dy25\hr14\min30\sec45
	year := rp.extractRtfValue(dateStr, "yr")
	month := rp.extractRtfValue(dateStr, "mo")
	day := rp.extractRtfValue(dateStr, "dy")
	hour := rp.extractRtfValue(dateStr, "hr")
	minute := rp.extractRtfValue(dateStr, "min")
	second := rp.extractRtfValue(dateStr, "sec")

	if year == "" || month == "" || day == "" {
		return time.Time{}, fmt.Errorf("invalid date format")
	}

	yearInt, _ := strconv.Atoi(year)
	monthInt, _ := strconv.Atoi(month)
	dayInt, _ := strconv.Atoi(day)
	hourInt, _ := strconv.Atoi(hour)
	minuteInt, _ := strconv.Atoi(minute)
	secondInt, _ := strconv.Atoi(second)

	return time.Date(yearInt, time.Month(monthInt), dayInt, hourInt, minuteInt, secondInt, 0, time.UTC), nil
}

// splitRtfParagraphs 分割RTF段落
func (rp *RtfParser) splitRtfParagraphs(content string) []string {
	// 按段落标记分割
	paragraphPattern := regexp.MustCompile(`\\\\par`)
	paragraphs := paragraphPattern.Split(content, -1)

	// 过滤空段落
	var result []string
	for _, p := range paragraphs {
		if strings.TrimSpace(p) != "" {
			result = append(result, p)
		}
	}

	return result
}

// extractTextFromRtf 从RTF中提取纯文本
func (rp *RtfParser) extractTextFromRtf(rtfContent string) string {
	// 移除RTF控制字符
	text := rtfContent

	// 移除控制字符
	controlPattern := regexp.MustCompile(`\\\\[a-zA-Z]+\\d*`)
	text = controlPattern.ReplaceAllString(text, "")

	// 移除大括号
	text = strings.ReplaceAll(text, "{", "")
	text = strings.ReplaceAll(text, "}", "")

	return strings.TrimSpace(text)
}

// extractParagraphStyle 提取段落样式
func (rp *RtfParser) extractParagraphStyle(paragraphContent string) string {
	// 检查样式标记
	if strings.Contains(paragraphContent, "\\s1") {
		return "Heading 1"
	} else if strings.Contains(paragraphContent, "\\s2") {
		return "Heading 2"
	} else if strings.Contains(paragraphContent, "\\s3") {
		return "Heading 3"
	}
	return "Normal"
}

// extractStyleTable 提取样式表
func (rp *RtfParser) extractStyleTable(content string) string {
	// 查找样式表部分
	stylePattern := regexp.MustCompile(`\\\\stylesheet\\s*\\{([^}]+)\\}`)
	matches := stylePattern.FindStringSubmatch(content)
	if len(matches) > 1 {
		return matches[1]
	}
	return ""
}

// parseRtfStyleTable 解析RTF样式表
func (rp *RtfParser) parseRtfStyleTable(styleTable string, styles *types.DocumentStyles) {
	// 解析段落样式
	paragraphStyle := types.ParagraphStyle{
		ID:   "ps1",
		Name: "Normal",
	}
	styles.ParagraphStyles = append(styles.ParagraphStyles, paragraphStyle)

	// 解析字符样式
	characterStyle := types.CharacterStyle{
		ID:   "cs1",
		Name: "Default Paragraph Font",
	}
	styles.CharacterStyles = append(styles.CharacterStyles, characterStyle)

	// 解析表格样式
	tableStyle := types.TableStyle{
		ID:   "ts1",
		Name: "Table Grid",
	}
	styles.TableStyles = append(styles.TableStyles, tableStyle)
}
