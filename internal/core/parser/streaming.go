package parser

import (
	"bufio"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"sync"
	"time"

	"archive/zip"
	"docs-parser/internal/core/types"
)

// StreamingParser 流式解析器，用于处理大文件
type StreamingParser struct {
	bufferSize int
	workers    int
	pool       *sync.Pool
}

// StreamingResult 流式解析结果
type StreamingResult struct {
	Document *types.Document
	Error    error
	Duration time.Duration
}

// NewStreamingParser 创建新的流式解析器
func NewStreamingParser(bufferSize, workers int) *StreamingParser {
	return &StreamingParser{
		bufferSize: bufferSize,
		workers:    workers,
		pool: &sync.Pool{
			New: func() interface{} {
				return &types.Document{}
			},
		},
	}
}

// ParseStream 流式解析文档
func (sp *StreamingParser) ParseStream(filePath string) (<-chan StreamingResult, error) {
	resultChan := make(chan StreamingResult, sp.workers)

	go func() {
		defer close(resultChan)

		start := time.Now()

		// 根据文件类型选择解析策略
		ext := getFileExtensionForStreaming(filePath)
		switch ext {
		case ".docx":
			sp.parseDocxStream(filePath, resultChan)
		case ".doc":
			sp.parseDocStream(filePath, resultChan)
		case ".rtf":
			sp.parseRtfStream(filePath, resultChan)
		default:
			// 使用传统解析器作为后备
			sp.parseLegacyStream(filePath, resultChan)
		}

		duration := time.Since(start)
		resultChan <- StreamingResult{
			Document: nil,
			Error:    nil,
			Duration: duration,
		}
	}()

	return resultChan, nil
}

// parseDocxStream 流式解析DOCX文件
func (sp *StreamingParser) parseDocxStream(filePath string, resultChan chan<- StreamingResult) {
	// 从池中获取文档对象
	doc := sp.pool.Get().(*types.Document)
	defer sp.pool.Put(doc)

	// 重置文档对象
	*doc = types.Document{}

	// 打开ZIP文件
	reader, err := zip.OpenReader(filePath)
	if err != nil {
		resultChan <- StreamingResult{Error: fmt.Errorf("failed to open DOCX: %w", err)}
		return
	}
	defer reader.Close()

	// 流式解析XML文件
	for _, file := range reader.File {
		if file.Name == "word/document.xml" {
			sp.parseDocumentXML(file, doc)
		} else if file.Name == "docProps/core.xml" {
			sp.parseCoreXML(file, doc)
		} else if file.Name == "word/styles.xml" {
			sp.parseStylesXML(file, doc)
		}
	}

	resultChan <- StreamingResult{
		Document: doc,
		Error:    nil,
		Duration: time.Since(time.Now()),
	}
}

// parseDocumentXML 流式解析文档XML
func (sp *StreamingParser) parseDocumentXML(file *zip.File, doc *types.Document) {
	rc, err := file.Open()
	if err != nil {
		return
	}
	defer rc.Close()

	// 使用缓冲读取器
	reader := bufio.NewReaderSize(rc, sp.bufferSize)
	decoder := xml.NewDecoder(reader)

	// 流式解析XML
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			continue
		}

		// 处理XML标签
		switch t := token.(type) {
		case xml.StartElement:
			sp.handleXMLElement(t, doc)
		}
	}
}

// handleXMLElement 处理XML元素
func (sp *StreamingParser) handleXMLElement(element xml.StartElement, doc *types.Document) {
	switch element.Name.Local {
	case "document":
		// 开始解析文档
	case "body":
		// 开始解析正文
	case "p":
		// 解析段落
		sp.parseParagraph(element, doc)
	case "tbl":
		// 解析表格
		sp.parseTable(element, doc)
	}
}

// parseParagraph 流式解析段落
func (sp *StreamingParser) parseParagraph(element xml.StartElement, doc *types.Document) {
	paragraph := types.Paragraph{
		ID:    generateID(),
		Style: types.ParagraphStyle{Name: "Normal"},
	}

	// 解析段落属性
	for _, attr := range element.Attr {
		switch attr.Name.Local {
		case "style":
			paragraph.Style.Name = attr.Value
		}
	}

	// 解析段落内容
	// 这里简化实现，实际需要递归解析子元素
	paragraph.Runs = append(paragraph.Runs, types.TextRun{
		Text: "段落内容",
		Font: types.Font{Name: "Calibri", Size: 11.0},
	})

	doc.Content.Paragraphs = append(doc.Content.Paragraphs, paragraph)
}

// parseTable 流式解析表格
func (sp *StreamingParser) parseTable(element xml.StartElement, doc *types.Document) {
	table := types.Table{
		ID:   generateID(),
		Rows: []types.TableRow{},
	}

	// 解析表格属性
	for _, attr := range element.Attr {
		switch attr.Name.Local {
		case "style":
			table.Style.Name = attr.Value
		}
	}

	doc.Content.Tables = append(doc.Content.Tables, table)
}

// parseCoreXML 流式解析核心XML
func (sp *StreamingParser) parseCoreXML(file *zip.File, doc *types.Document) {
	rc, err := file.Open()
	if err != nil {
		return
	}
	defer rc.Close()

	reader := bufio.NewReaderSize(rc, sp.bufferSize)
	decoder := xml.NewDecoder(reader)

	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			continue
		}

		switch t := token.(type) {
		case xml.StartElement:
			sp.handleCoreElement(t, doc)
		}
	}
}

// handleCoreElement 处理核心XML元素
func (sp *StreamingParser) handleCoreElement(element xml.StartElement, doc *types.Document) {
	switch element.Name.Local {
	case "title":
		doc.Metadata.Title = sp.extractText(element)
	case "creator":
		doc.Metadata.Author = sp.extractText(element)
	case "created":
		doc.Metadata.Created = sp.parseTime(sp.extractText(element))
	case "modified":
		doc.Metadata.Modified = sp.parseTime(sp.extractText(element))
	}
}

// parseStylesXML 流式解析样式XML
func (sp *StreamingParser) parseStylesXML(file *zip.File, doc *types.Document) {
	rc, err := file.Open()
	if err != nil {
		return
	}
	defer rc.Close()

	reader := bufio.NewReaderSize(rc, sp.bufferSize)
	decoder := xml.NewDecoder(reader)

	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			continue
		}

		switch t := token.(type) {
		case xml.StartElement:
			sp.handleStyleElement(t, doc)
		}
	}
}

// handleStyleElement 处理样式XML元素
func (sp *StreamingParser) handleStyleElement(element xml.StartElement, doc *types.Document) {
	switch element.Name.Local {
	case "style":
		style := types.ParagraphStyle{
			ID:   sp.extractAttribute(element, "styleId"),
			Name: sp.extractAttribute(element, "name"),
		}
		doc.Styles.ParagraphStyles = append(doc.Styles.ParagraphStyles, style)
	}
}

// parseDocStream 流式解析DOC文件
func (sp *StreamingParser) parseDocStream(filePath string, resultChan chan<- StreamingResult) {
	doc := sp.pool.Get().(*types.Document)
	defer sp.pool.Put(doc)

	// 重置文档对象
	*doc = types.Document{}

	file, err := os.Open(filePath)
	if err != nil {
		resultChan <- StreamingResult{Error: fmt.Errorf("failed to open DOC: %w", err)}
		return
	}
	defer file.Close()

	// 使用缓冲读取器
	reader := bufio.NewReaderSize(file, sp.bufferSize)

	// 读取文件头
	header := make([]byte, 64)
	_, err = reader.Read(header)
	if err != nil {
		resultChan <- StreamingResult{Error: fmt.Errorf("failed to read header: %w", err)}
		return
	}

	// 解析元数据
	sp.parseDocMetadata(reader, doc)

	// 解析内容
	sp.parseDocContent(reader, doc)

	resultChan <- StreamingResult{
		Document: doc,
		Error:    nil,
		Duration: time.Since(time.Now()),
	}
}

// parseDocMetadata 流式解析DOC元数据
func (sp *StreamingParser) parseDocMetadata(reader *bufio.Reader, doc *types.Document) {
	// 跳过文件头
	reader.Discard(512)

	// 读取元数据区域
	metadata := make([]byte, 256)
	reader.Read(metadata)

	// 提取标题
	if len(metadata) > 32 {
		title := bytes.TrimRight(metadata[32:64], "\x00")
		doc.Metadata.Title = string(title)
	}

	// 提取作者
	if len(metadata) > 64 {
		author := bytes.TrimRight(metadata[64:96], "\x00")
		doc.Metadata.Author = string(author)
	}

	// 设置默认值
	doc.Metadata.Created = time.Now()
	doc.Metadata.Modified = time.Now()
	doc.Metadata.Version = "8.0"
}

// parseDocContent 流式解析DOC内容
func (sp *StreamingParser) parseDocContent(reader *bufio.Reader, doc *types.Document) {
	// 跳过元数据区域
	reader.Discard(512)

	// 分块读取内容
	buffer := make([]byte, sp.bufferSize)

	for {
		n, err := reader.Read(buffer)
		if err == io.EOF {
			break
		}
		if err != nil {
			continue
		}

		// 处理内容块
		sp.processContentChunk(buffer[:n], doc)
	}
}

// processContentChunk 处理内容块
func (sp *StreamingParser) processContentChunk(chunk []byte, doc *types.Document) {
	// 查找段落分隔符
	text := string(chunk)
	lines := bytes.Split([]byte(text), []byte("\r\n"))

	for _, line := range lines {
		if len(line) > 0 {
			paragraph := types.Paragraph{
				ID: generateID(),
				Runs: []types.TextRun{
					{
						Text: string(line),
						Font: types.Font{Name: "Times New Roman", Size: 12.0},
					},
				},
				Style: types.ParagraphStyle{Name: "Normal"},
			}
			doc.Content.Paragraphs = append(doc.Content.Paragraphs, paragraph)
		}
	}
}

// parseRtfStream 流式解析RTF文件
func (sp *StreamingParser) parseRtfStream(filePath string, resultChan chan<- StreamingResult) {
	doc := sp.pool.Get().(*types.Document)
	defer sp.pool.Put(doc)

	// 重置文档对象
	*doc = types.Document{}

	file, err := os.Open(filePath)
	if err != nil {
		resultChan <- StreamingResult{Error: fmt.Errorf("failed to open RTF: %w", err)}
		return
	}
	defer file.Close()

	reader := bufio.NewReaderSize(file, sp.bufferSize)

	// 流式解析RTF
	sp.parseRtfContent(reader, doc)

	resultChan <- StreamingResult{
		Document: doc,
		Error:    nil,
		Duration: time.Since(time.Now()),
	}
}

// parseRtfContent 流式解析RTF内容
func (sp *StreamingParser) parseRtfContent(reader *bufio.Reader, doc *types.Document) {
	// 设置默认元数据
	doc.Metadata.Title = "RTF Document"
	doc.Metadata.Author = "Unknown"
	doc.Metadata.Created = time.Now()
	doc.Metadata.Modified = time.Now()
	doc.Metadata.Version = "1.0"

	// 分块读取RTF内容
	buffer := make([]byte, sp.bufferSize)

	for {
		n, err := reader.Read(buffer)
		if err == io.EOF {
			break
		}
		if err != nil {
			continue
		}

		// 处理RTF内容块
		sp.processRtfChunk(buffer[:n], doc)
	}
}

// processRtfChunk 处理RTF内容块
func (sp *StreamingParser) processRtfChunk(chunk []byte, doc *types.Document) {
	// 简化的RTF解析
	text := string(chunk)

	// 移除RTF控制字符
	cleanText := sp.cleanRtfText(text)

	if len(cleanText) > 0 {
		paragraph := types.Paragraph{
			ID: generateID(),
			Runs: []types.TextRun{
				{
					Text: cleanText,
					Font: types.Font{Name: "Arial", Size: 12.0},
				},
			},
			Style: types.ParagraphStyle{Name: "Normal"},
		}
		doc.Content.Paragraphs = append(doc.Content.Paragraphs, paragraph)
	}
}

// cleanRtfText 清理RTF文本
func (sp *StreamingParser) cleanRtfText(text string) string {
	// 移除RTF控制字符的简化实现
	// 实际实现需要更复杂的RTF解析
	return text
}

// parseLegacyStream 流式解析历史格式
func (sp *StreamingParser) parseLegacyStream(filePath string, resultChan chan<- StreamingResult) {
	doc := sp.pool.Get().(*types.Document)
	defer sp.pool.Put(doc)

	// 重置文档对象
	*doc = types.Document{}

	// 使用传统解析器作为后备
	legacyParser := NewParserFactory()
	legacyParser.RegisterParser("docx", &DefaultParser{})
	legacyDoc, err := legacyParser.ParseDocument(filePath)
	if err != nil {
		resultChan <- StreamingResult{Error: err}
		return
	}

	*doc = *legacyDoc

	resultChan <- StreamingResult{
		Document: doc,
		Error:    nil,
		Duration: time.Since(time.Now()),
	}
}

// 辅助函数

func getFileExtensionForStreaming(filePath string) string {
	// 简化实现，实际需要更复杂的扩展名检测
	if len(filePath) > 4 {
		return filePath[len(filePath)-4:]
	}
	return ""
}

func generateID() string {
	return fmt.Sprintf("id_%d", time.Now().UnixNano())
}

func (sp *StreamingParser) extractText(element xml.StartElement) string {
	// 简化实现，实际需要解析元素内容
	return ""
}

func (sp *StreamingParser) extractAttribute(element xml.StartElement, name string) string {
	for _, attr := range element.Attr {
		if attr.Name.Local == name {
			return attr.Value
		}
	}
	return ""
}

func (sp *StreamingParser) parseTime(timeStr string) time.Time {
	// 简化实现，实际需要解析各种时间格式
	return time.Now()
}
