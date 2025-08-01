package formats

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"docs-parser/internal/core/parser"
	"docs-parser/internal/core/types"
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
	// 验证文件
	if err := dp.ValidateFile(filePath); err != nil {
		return nil, err
	}

	// 打开zip文件
	reader, err := zip.OpenReader(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open docx file: %w", err)
	}
	defer reader.Close()

	doc := &types.Document{}

	// 解析元数据
	metadata, err := dp.parseMetadata(reader)
	if err != nil {
		return nil, fmt.Errorf("failed to parse metadata: %w", err)
	}
	doc.Metadata = *metadata

	// 解析内容
	content, err := dp.parseContent(reader)
	if err != nil {
		return nil, fmt.Errorf("failed to parse content: %w", err)
	}
	doc.Content = *content

	// 解析样式
	styles, err := dp.parseStyles(reader)
	if err != nil {
		return nil, fmt.Errorf("failed to parse styles: %w", err)
	}
	doc.Styles = *styles

	// 解析格式规则
	formatRules, err := dp.parseFormatRules(reader)
	if err != nil {
		return nil, fmt.Errorf("failed to parse format rules: %w", err)
	}
	doc.FormatRules = *formatRules

	return doc, nil
}

// ParseMetadata 解析元数据
func (dp *DocxParser) ParseMetadata(filePath string) (*types.DocumentMetadata, error) {
	reader, err := zip.OpenReader(filePath)
	if err != nil {
		return nil, err
	}
	defer reader.Close()

	return dp.parseMetadata(reader)
}

// ParseContent 解析内容
func (dp *DocxParser) ParseContent(filePath string) (*types.DocumentContent, error) {
	reader, err := zip.OpenReader(filePath)
	if err != nil {
		return nil, err
	}
	defer reader.Close()

	return dp.parseContent(reader)
}

// ParseStyles 解析样式
func (dp *DocxParser) ParseStyles(filePath string) (*types.DocumentStyles, error) {
	reader, err := zip.OpenReader(filePath)
	if err != nil {
		return nil, err
	}
	defer reader.Close()

	return dp.parseStyles(reader)
}

// ParseFormatRules 解析格式规则
func (dp *DocxParser) ParseFormatRules(filePath string) (*types.FormatRules, error) {
	reader, err := zip.OpenReader(filePath)
	if err != nil {
		return nil, err
	}
	defer reader.Close()

	return dp.parseFormatRules(reader)
}

// GetSupportedFormats 获取支持的格式
func (dp *DocxParser) GetSupportedFormats() []string {
	return []string{"docx"}
}

// ValidateFile 验证文件格式
func (dp *DocxParser) ValidateFile(filePath string) error {
	// 检查文件是否存在
	if _, err := os.Stat(filePath); os.IsNotExist(err) {
		return parser.ErrFileNotFound
	}

	// 检查文件扩展名
	ext := filepath.Ext(filePath)
	if ext != ".docx" {
		return parser.ErrUnsupportedFormat
	}

	// 检查是否为有效的zip文件
	reader, err := zip.OpenReader(filePath)
	if err != nil {
		return parser.ErrInvalidFile
	}
	defer reader.Close()

	// 检查是否包含必要的docx文件
	requiredFiles := []string{"word/document.xml", "word/styles.xml"}
	for _, requiredFile := range requiredFiles {
		found := false
		for _, file := range reader.File {
			if file.Name == requiredFile {
				found = true
				break
			}
		}
		if !found {
			return parser.ErrInvalidFile
		}
	}

	return nil
}

// parseMetadata 解析元数据
func (dp *DocxParser) parseMetadata(reader *zip.ReadCloser) (*types.DocumentMetadata, error) {
	metadata := &types.DocumentMetadata{}

	// 解析core.xml
	if err := dp.parseCoreProperties(reader, metadata); err != nil {
		return nil, err
	}

	// 解析app.xml
	if err := dp.parseAppProperties(reader, metadata); err != nil {
		return nil, err
	}

	return metadata, nil
}

// parseCoreProperties 解析核心属性
func (dp *DocxParser) parseCoreProperties(reader *zip.ReadCloser, metadata *types.DocumentMetadata) error {
	coreFile := dp.findFile(reader, "docProps/core.xml")
	if coreFile == nil {
		return nil
	}

	rc, err := coreFile.Open()
	if err != nil {
		return err
	}
	defer rc.Close()

	var coreProps struct {
		Title       string `xml:"title"`
		Subject     string `xml:"subject"`
		Creator     string `xml:"creator"`
		Keywords    string `xml:"keywords"`
		Created     string `xml:"created"`
		Modified    string `xml:"modified"`
		LastSavedBy string `xml:"lastSavedBy"`
		Revision    string `xml:"revision"`
		Version     string `xml:"version"`
	}

	if err := xml.NewDecoder(rc).Decode(&coreProps); err != nil {
		return err
	}

	metadata.Title = coreProps.Title
	metadata.Subject = coreProps.Subject
	metadata.Author = coreProps.Creator
	metadata.LastSavedBy = coreProps.LastSavedBy
	metadata.Version = coreProps.Version

	if coreProps.Keywords != "" {
		metadata.Keywords = strings.Split(coreProps.Keywords, ",")
	}

	if revision, err := strconv.Atoi(coreProps.Revision); err == nil {
		metadata.Revision = revision
	}

	// 解析时间
	if coreProps.Created != "" {
		if t, err := time.Parse(time.RFC3339, coreProps.Created); err == nil {
			metadata.Created = t
		}
	}
	if coreProps.Modified != "" {
		if t, err := time.Parse(time.RFC3339, coreProps.Modified); err == nil {
			metadata.Modified = t
		}
	}

	return nil
}

// parseAppProperties 解析应用属性
func (dp *DocxParser) parseAppProperties(reader *zip.ReadCloser, metadata *types.DocumentMetadata) error {
	appFile := dp.findFile(reader, "docProps/app.xml")
	if appFile == nil {
		return nil
	}

	rc, err := appFile.Open()
	if err != nil {
		return err
	}
	defer rc.Close()

	var appProps struct {
		Pages   string `xml:"Pages"`
		Words   string `xml:"Words"`
		Company string `xml:"Company"`
	}

	if err := xml.NewDecoder(rc).Decode(&appProps); err != nil {
		return err
	}

	if pages, err := strconv.Atoi(appProps.Pages); err == nil {
		metadata.PageCount = pages
	}
	if words, err := strconv.Atoi(appProps.Words); err == nil {
		metadata.WordCount = words
	}

	return nil
}

// parseContent 解析文档内容
func (dp *DocxParser) parseContent(reader *zip.ReadCloser) (*types.DocumentContent, error) {
	content := &types.DocumentContent{}

	// 解析主文档
	if err := dp.parseMainDocument(reader, content); err != nil {
		return nil, err
	}

	// 解析页眉页脚
	if err := dp.parseHeadersFooters(reader, content); err != nil {
		return nil, err
	}

	// 解析注释
	if err := dp.parseComments(reader, content); err != nil {
		return nil, err
	}

	return content, nil
}

// parseMainDocument 解析主文档
func (dp *DocxParser) parseMainDocument(reader *zip.ReadCloser, content *types.DocumentContent) error {
	docFile := dp.findFile(reader, "word/document.xml")
	if docFile == nil {
		return fmt.Errorf("document.xml not found")
	}

	rc, err := docFile.Open()
	if err != nil {
		return err
	}
	defer rc.Close()

	// 使用更详细的XML结构定义
	var document struct {
		Body struct {
			Paragraphs []struct {
				ID    string `xml:"id,attr"`
				Text  string `xml:"text"`
				Style struct {
					Name string `xml:"name,attr"`
				} `xml:"style"`
				Runs []struct {
					ID   string `xml:"id,attr"`
					Text string `xml:"text"`
					Font struct {
						Name   string  `xml:"name,attr"`
						Size   float64 `xml:"size,attr"`
						Bold   bool    `xml:"bold,attr"`
						Italic bool    `xml:"italic,attr"`
					} `xml:"font"`
				} `xml:"run"`
			} `xml:"paragraph"`
			Tables []struct {
				ID   string `xml:"id,attr"`
				Rows []struct {
					ID    string `xml:"id,attr"`
					Cells []struct {
						ID      string `xml:"id,attr"`
						Content string `xml:"content"`
					} `xml:"cell"`
				} `xml:"row"`
			} `xml:"table"`
		} `xml:"body"`
	}

	if err := xml.NewDecoder(rc).Decode(&document); err != nil {
		return err
	}

	// 解析段落
	for _, p := range document.Body.Paragraphs {
		paragraph := types.Paragraph{
			ID:   p.ID,
			Text: p.Text,
			Style: types.ParagraphStyle{
				Name: p.Style.Name,
			},
		}

		// 解析文本运行
		for _, r := range p.Runs {
			run := types.TextRun{
				ID:   r.ID,
				Text: r.Text,
				Font: types.Font{
					Name:   r.Font.Name,
					Size:   r.Font.Size,
					Bold:   r.Font.Bold,
					Italic: r.Font.Italic,
				},
			}
			paragraph.Runs = append(paragraph.Runs, run)
		}

		content.Paragraphs = append(content.Paragraphs, paragraph)
	}

	// 解析表格
	for _, t := range document.Body.Tables {
		table := types.Table{
			ID: t.ID,
		}

		for _, row := range t.Rows {
			tableRow := types.TableRow{
				ID: row.ID,
			}

			for _, cell := range row.Cells {
				tableCell := types.TableCell{
					ID:      cell.ID,
					Content: []types.Paragraph{{Text: cell.Content}},
				}
				tableRow.Cells = append(tableRow.Cells, tableCell)
			}

			table.Rows = append(table.Rows, tableRow)
		}

		content.Tables = append(content.Tables, table)
	}

	return nil
}

// parseHeadersFooters 解析页眉页脚
func (dp *DocxParser) parseHeadersFooters(reader *zip.ReadCloser, content *types.DocumentContent) error {
	// 解析页眉
	headerFiles := dp.findFiles(reader, "word/header")
	for _, headerFile := range headerFiles {
		header := types.Header{
			ID: headerFile.Name,
		}
		content.Headers = append(content.Headers, header)
	}

	// 解析页脚
	footerFiles := dp.findFiles(reader, "word/footer")
	for _, footerFile := range footerFiles {
		footer := types.Footer{
			ID: footerFile.Name,
		}
		content.Footers = append(content.Footers, footer)
	}

	return nil
}

// parseComments 解析注释
func (dp *DocxParser) parseComments(reader *zip.ReadCloser, content *types.DocumentContent) error {
	commentsFile := dp.findFile(reader, "word/comments.xml")
	if commentsFile == nil {
		return nil
	}

	rc, err := commentsFile.Open()
	if err != nil {
		return err
	}
	defer rc.Close()

	var comments struct {
		Comments []struct {
			ID     string `xml:"id,attr"`
			Author string `xml:"author,attr"`
			Date   string `xml:"date,attr"`
			Text   string `xml:"text"`
		} `xml:"comment"`
	}

	if err := xml.NewDecoder(rc).Decode(&comments); err != nil {
		return err
	}

	for _, c := range comments.Comments {
		comment := types.Comment{
			ID:     c.ID,
			Author: c.Author,
			Text:   c.Text,
		}

		if c.Date != "" {
			if t, err := time.Parse(time.RFC3339, c.Date); err == nil {
				comment.Date = t
			}
		}

		content.Comments = append(content.Comments, comment)
	}

	return nil
}

// parseStyles 解析样式
func (dp *DocxParser) parseStyles(reader *zip.ReadCloser) (*types.DocumentStyles, error) {
	styles := &types.DocumentStyles{}

	stylesFile := dp.findFile(reader, "word/styles.xml")
	if stylesFile == nil {
		return styles, nil
	}

	rc, err := stylesFile.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()

	var stylesDoc struct {
		Styles []struct {
			ID   string `xml:"id,attr"`
			Name string `xml:"name,attr"`
			Type string `xml:"type,attr"`
		} `xml:"style"`
	}

	if err := xml.NewDecoder(rc).Decode(&stylesDoc); err != nil {
		return nil, err
	}

	for _, s := range stylesDoc.Styles {
		switch s.Type {
		case "paragraph":
			styles.ParagraphStyles = append(styles.ParagraphStyles, types.ParagraphStyle{
				ID:   s.ID,
				Name: s.Name,
			})
		case "character":
			styles.CharacterStyles = append(styles.CharacterStyles, types.CharacterStyle{
				ID:   s.ID,
				Name: s.Name,
			})
		case "table":
			styles.TableStyles = append(styles.TableStyles, types.TableStyle{
				ID:   s.ID,
				Name: s.Name,
			})
		}
	}

	return styles, nil
}

// parseFormatRules 解析格式规则
func (dp *DocxParser) parseFormatRules(reader *zip.ReadCloser) (*types.FormatRules, error) {
	formatRules := &types.FormatRules{}

	// 解析字体规则
	if err := dp.parseFontRules(reader, formatRules); err != nil {
		return nil, err
	}

	// 解析段落规则
	if err := dp.parseParagraphRules(reader, formatRules); err != nil {
		return nil, err
	}

	// 解析表格规则
	if err := dp.parseTableRules(reader, formatRules); err != nil {
		return nil, err
	}

	// 解析页面规则
	if err := dp.parsePageRules(reader, formatRules); err != nil {
		return nil, err
	}

	return formatRules, nil
}

// parseFontRules 解析字体规则
func (dp *DocxParser) parseFontRules(reader *zip.ReadCloser, formatRules *types.FormatRules) error {
	// 从样式中提取字体规则
	stylesFile := dp.findFile(reader, "word/styles.xml")
	if stylesFile == nil {
		return nil
	}

	rc, err := stylesFile.Open()
	if err != nil {
		return err
	}
	defer rc.Close()

	var stylesDoc struct {
		Styles []struct {
			ID   string `xml:"id,attr"`
			Name string `xml:"name,attr"`
			Font struct {
				Name   string  `xml:"name,attr"`
				Size   float64 `xml:"size,attr"`
				Bold   bool    `xml:"bold,attr"`
				Italic bool    `xml:"italic,attr"`
			} `xml:"font"`
		} `xml:"style"`
	}

	if err := xml.NewDecoder(rc).Decode(&stylesDoc); err != nil {
		return err
	}

	for _, s := range stylesDoc.Styles {
		fontRule := types.FontRule{
			ID:     s.ID,
			Name:   s.Name,
			Size:   s.Font.Size,
			Color:  types.Color{},
			Bold:   s.Font.Bold,
			Italic: s.Font.Italic,
		}
		formatRules.FontRules = append(formatRules.FontRules, fontRule)
	}

	return nil
}

// parseParagraphRules 解析段落规则
func (dp *DocxParser) parseParagraphRules(reader *zip.ReadCloser, formatRules *types.FormatRules) error {
	// 从样式中提取段落规则
	stylesFile := dp.findFile(reader, "word/styles.xml")
	if stylesFile == nil {
		return nil
	}

	rc, err := stylesFile.Open()
	if err != nil {
		return err
	}
	defer rc.Close()

	var stylesDoc struct {
		Styles []struct {
			ID        string `xml:"id,attr"`
			Name      string `xml:"name,attr"`
			Paragraph struct {
				Alignment string `xml:"alignment,attr"`
				Indent    struct {
					Left    float64 `xml:"left,attr"`
					Right   float64 `xml:"right,attr"`
					First   float64 `xml:"first,attr"`
					Hanging float64 `xml:"hanging,attr"`
				} `xml:"indent"`
				Spacing struct {
					Before float64 `xml:"before,attr"`
					After  float64 `xml:"after,attr"`
					Line   float64 `xml:"line,attr"`
				} `xml:"spacing"`
			} `xml:"paragraph"`
		} `xml:"style"`
	}

	if err := xml.NewDecoder(rc).Decode(&stylesDoc); err != nil {
		return err
	}

	for _, s := range stylesDoc.Styles {
		paragraphRule := types.ParagraphRule{
			ID:        s.ID,
			Name:      s.Name,
			Alignment: types.Alignment(s.Paragraph.Alignment),
			Indentation: types.Indentation{
				Left:    s.Paragraph.Indent.Left,
				Right:   s.Paragraph.Indent.Right,
				First:   s.Paragraph.Indent.First,
				Hanging: s.Paragraph.Indent.Hanging,
			},
			Spacing: types.Spacing{
				Before: s.Paragraph.Spacing.Before,
				After:  s.Paragraph.Spacing.After,
				Line:   s.Paragraph.Spacing.Line,
			},
		}
		formatRules.ParagraphRules = append(formatRules.ParagraphRules, paragraphRule)
	}

	return nil
}

// parseTableRules 解析表格规则
func (dp *DocxParser) parseTableRules(reader *zip.ReadCloser, formatRules *types.FormatRules) error {
	// 从样式中提取表格规则
	stylesFile := dp.findFile(reader, "word/styles.xml")
	if stylesFile == nil {
		return nil
	}

	rc, err := stylesFile.Open()
	if err != nil {
		return err
	}
	defer rc.Close()

	var stylesDoc struct {
		Styles []struct {
			ID    string `xml:"id,attr"`
			Name  string `xml:"name,attr"`
			Table struct {
				Width     float64 `xml:"width,attr"`
				Alignment string  `xml:"alignment,attr"`
				Borders   struct {
					Top    string `xml:"top,attr"`
					Bottom string `xml:"bottom,attr"`
					Left   string `xml:"left,attr"`
					Right  string `xml:"right,attr"`
				} `xml:"borders"`
			} `xml:"table"`
		} `xml:"style"`
	}

	if err := xml.NewDecoder(rc).Decode(&stylesDoc); err != nil {
		return err
	}

	for _, s := range stylesDoc.Styles {
		tableRule := types.TableRule{
			ID:        s.ID,
			Name:      s.Name,
			Width:     s.Table.Width,
			Alignment: types.Alignment(s.Table.Alignment),
		}
		formatRules.TableRules = append(formatRules.TableRules, tableRule)
	}

	return nil
}

// parsePageRules 解析页面规则
func (dp *DocxParser) parsePageRules(reader *zip.ReadCloser, formatRules *types.FormatRules) error {
	// 从样式中提取页面规则
	stylesFile := dp.findFile(reader, "word/styles.xml")
	if stylesFile == nil {
		return nil
	}

	rc, err := stylesFile.Open()
	if err != nil {
		return err
	}
	defer rc.Close()

	var stylesDoc struct {
		Styles []struct {
			ID   string `xml:"id,attr"`
			Name string `xml:"name,attr"`
			Page struct {
				Size struct {
					Width  float64 `xml:"width,attr"`
					Height float64 `xml:"height,attr"`
				} `xml:"size"`
				Margins struct {
					Top    float64 `xml:"top,attr"`
					Bottom float64 `xml:"bottom,attr"`
					Left   float64 `xml:"left,attr"`
					Right  float64 `xml:"right,attr"`
				} `xml:"margins"`
			} `xml:"page"`
		} `xml:"style"`
	}

	if err := xml.NewDecoder(rc).Decode(&stylesDoc); err != nil {
		return err
	}

	for _, s := range stylesDoc.Styles {
		pageRule := types.PageRule{
			ID:   s.ID,
			Name: s.Name,
			PageSize: types.PageSize{
				Width:  s.Page.Size.Width,
				Height: s.Page.Size.Height,
			},
			PageMargins: types.PageMargins{
				Top:    s.Page.Margins.Top,
				Bottom: s.Page.Margins.Bottom,
				Left:   s.Page.Margins.Left,
				Right:  s.Page.Margins.Right,
			},
		}
		formatRules.PageRules = append(formatRules.PageRules, pageRule)
	}

	return nil
}

// findFile 查找文件
func (dp *DocxParser) findFile(reader *zip.ReadCloser, name string) *zip.File {
	for _, file := range reader.File {
		if file.Name == name {
			return file
		}
	}
	return nil
}

// findFiles 查找匹配的文件
func (dp *DocxParser) findFiles(reader *zip.ReadCloser, prefix string) []*zip.File {
	var files []*zip.File
	for _, file := range reader.File {
		if strings.HasPrefix(file.Name, prefix) {
			files = append(files, file)
		}
	}
	return files
}
