package documents

import (
	"encoding/xml"
	"fmt"
	"strconv"
	"strings"
	"time"

	"docs-parser/internal/core/types"
	"docs-parser/internal/packaging"
)

// WordprocessingDocument 表示Word文档
type WordprocessingDocument struct {
	Container *packaging.OPCContainer
	Document  *types.Document
	Parts     map[string]*DocumentPart
}

// DocumentPart 表示文档部分
type DocumentPart struct {
	Name     string
	Content  []byte
	Type     string
	Modified time.Time
}

// NewWordprocessingDocument 创建新的Word文档
func NewWordprocessingDocument(path string) *WordprocessingDocument {
	return &WordprocessingDocument{
		Container: packaging.NewOPCContainer(path),
		Parts:     make(map[string]*DocumentPart),
	}
}

// Open 打开Word文档
func (wd *WordprocessingDocument) Open() error {
	// 打开OPC容器
	if err := wd.Container.Open(); err != nil {
		return fmt.Errorf("failed to open OPC container: %w", err)
	}

	// 验证容器
	if err := wd.Container.Validate(); err != nil {
		return fmt.Errorf("invalid OPC container: %w", err)
	}

	// 加载文档部分
	if err := wd.loadParts(); err != nil {
		return fmt.Errorf("failed to load document parts: %w", err)
	}

	return nil
}

// loadParts 加载文档部分
func (wd *WordprocessingDocument) loadParts() error {
	// 加载主文档
	if err := wd.loadMainDocument(); err != nil {
		return fmt.Errorf("failed to load main document: %w", err)
	}

	// 加载样式
	if err := wd.loadStyles(); err != nil {
		return fmt.Errorf("failed to load styles: %w", err)
	}

	// 加载字体表
	if err := wd.loadFontTable(); err != nil {
		return fmt.Errorf("failed to load font table: %w", err)
	}

	// 加载设置
	if err := wd.loadSettings(); err != nil {
		return fmt.Errorf("failed to load settings: %w", err)
	}

	return nil
}

// loadMainDocument 加载主文档
func (wd *WordprocessingDocument) loadMainDocument() error {
	content, err := wd.Container.ReadFile("word/document.xml")
	if err != nil {
		return fmt.Errorf("failed to read main document: %w", err)
	}

	wd.Parts["document.xml"] = &DocumentPart{
		Name:    "document.xml",
		Content: content,
		Type:    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
	}

	return nil
}

// loadStyles 加载样式
func (wd *WordprocessingDocument) loadStyles() error {
	if wd.Container.HasFile("word/styles.xml") {
		content, err := wd.Container.ReadFile("word/styles.xml")
		if err != nil {
			return fmt.Errorf("failed to read styles: %w", err)
		}

		wd.Parts["styles.xml"] = &DocumentPart{
			Name:    "styles.xml",
			Content: content,
			Type:    "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
		}
	}

	return nil
}

// loadFontTable 加载字体表
func (wd *WordprocessingDocument) loadFontTable() error {
	if wd.Container.HasFile("word/fontTable.xml") {
		content, err := wd.Container.ReadFile("word/fontTable.xml")
		if err != nil {
			return fmt.Errorf("failed to read font table: %w", err)
		}

		wd.Parts["fontTable.xml"] = &DocumentPart{
			Name:    "fontTable.xml",
			Content: content,
			Type:    "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
		}
	}

	return nil
}

// loadSettings 加载设置
func (wd *WordprocessingDocument) loadSettings() error {
	if wd.Container.HasFile("word/settings.xml") {
		content, err := wd.Container.ReadFile("word/settings.xml")
		if err != nil {
			return fmt.Errorf("failed to read settings: %w", err)
		}

		wd.Parts["settings.xml"] = &DocumentPart{
			Name:    "settings.xml",
			Content: content,
			Type:    "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
		}
	}

	return nil
}

// Parse 解析Word文档
func (wd *WordprocessingDocument) Parse() (*types.Document, error) {
	doc := &types.Document{}

	// 解析元数据
	if err := wd.parseMetadata(doc); err != nil {
		return nil, fmt.Errorf("failed to parse metadata: %w", err)
	}

	// 解析内容
	if err := wd.parseContent(doc); err != nil {
		return nil, fmt.Errorf("failed to parse content: %w", err)
	}

	// 解析样式
	if err := wd.parseStyles(doc); err != nil {
		return nil, fmt.Errorf("failed to parse styles: %w", err)
	}

	// 解析格式规则
	if err := wd.parseFormatRules(doc); err != nil {
		return nil, fmt.Errorf("failed to parse format rules: %w", err)
	}

	wd.Document = doc
	return doc, nil
}

// parseMetadata 解析元数据
func (wd *WordprocessingDocument) parseMetadata(doc *types.Document) error {
	// 解析核心属性
	if err := wd.parseCoreProperties(doc); err != nil {
		return err
	}

	// 解析应用属性
	if err := wd.parseAppProperties(doc); err != nil {
		return err
	}

	return nil
}

// parseCoreProperties 解析核心属性
func (wd *WordprocessingDocument) parseCoreProperties(doc *types.Document) error {
	if !wd.Container.HasFile("docProps/core.xml") {
		return nil
	}

	content, err := wd.Container.ReadFile("docProps/core.xml")
	if err != nil {
		return err
	}

	var coreProps struct {
		Title       string `xml:"title"`
		Subject     string `xml:"subject"`
		Creator     string `xml:"creator"`
		Keywords    string `xml:"keywords"`
		Description string `xml:"description"`
		Created     string `xml:"created"`
		Modified    string `xml:"modified"`
	}

	if err := xml.Unmarshal(content, &coreProps); err != nil {
		return err
	}

	doc.Metadata.Title = coreProps.Title
	doc.Metadata.Subject = coreProps.Subject
	doc.Metadata.Author = coreProps.Creator
	doc.Metadata.Keywords = strings.Split(coreProps.Keywords, ",")

	// 解析时间
	if coreProps.Created != "" {
		if t, err := time.Parse(time.RFC3339, coreProps.Created); err == nil {
			doc.Metadata.Created = t
		}
	}
	if coreProps.Modified != "" {
		if t, err := time.Parse(time.RFC3339, coreProps.Modified); err == nil {
			doc.Metadata.Modified = t
		}
	}

	return nil
}

// parseAppProperties 解析应用属性
func (wd *WordprocessingDocument) parseAppProperties(doc *types.Document) error {
	if !wd.Container.HasFile("docProps/app.xml") {
		return nil
	}

	content, err := wd.Container.ReadFile("docProps/app.xml")
	if err != nil {
		return err
	}

	var appProps struct {
		Application   string `xml:"Application"`
		DocSecurity   string `xml:"DocSecurity"`
		ScaleCrop     string `xml:"ScaleCrop"`
		LinksUpToDate string `xml:"LinksUpToDate"`
		Pages         string `xml:"Pages"`
		Words         string `xml:"Words"`
		Characters    string `xml:"Characters"`
		Lines         string `xml:"Lines"`
		Paragraphs    string `xml:"Paragraphs"`
	}

	if err := xml.Unmarshal(content, &appProps); err != nil {
		return err
	}

	// 解析页数和字数
	if appProps.Pages != "" {
		if pages, err := strconv.Atoi(appProps.Pages); err == nil {
			doc.Metadata.PageCount = pages
		}
	}
	if appProps.Words != "" {
		if words, err := strconv.Atoi(appProps.Words); err == nil {
			doc.Metadata.WordCount = words
		}
	}

	return nil
}

// parseContent 解析内容
func (wd *WordprocessingDocument) parseContent(doc *types.Document) error {
	part, exists := wd.Parts["document.xml"]
	if !exists {
		return fmt.Errorf("main document part not found")
	}

	// 解析主文档内容
	if err := wd.parseMainDocument(part.Content, doc); err != nil {
		return fmt.Errorf("failed to parse main document: %w", err)
	}

	return nil
}

// parseMainDocument 解析主文档
func (wd *WordprocessingDocument) parseMainDocument(content []byte, doc *types.Document) error {
	var document struct {
		XMLName xml.Name `xml:"document"`
		Body    struct {
			XMLName     xml.Name `xml:"body"`
			Paragraphs  []struct {
				XMLName xml.Name `xml:"p"`
				Properties struct {
					Style struct {
						Val string `xml:"val,attr"`
					} `xml:"pStyle"`
					Justification struct {
						Val string `xml:"val,attr"`
					} `xml:"jc"`
					Indentation struct {
						Left   string `xml:"left,attr"`
						Right  string `xml:"right,attr"`
						First  string `xml:"firstLine,attr"`
						Hanging string `xml:"hanging,attr"`
					} `xml:"ind"`
					Spacing struct {
						Before string `xml:"before,attr"`
						After  string `xml:"after,attr"`
						Line   string `xml:"line,attr"`
					} `xml:"spacing"`
				} `xml:"pPr"`
				Runs []struct {
					XMLName xml.Name `xml:"r"`
					Properties struct {
						Font struct {
							Val string `xml:"val,attr"`
						} `xml:"rFonts"`
						Size struct {
							Val string `xml:"val,attr"`
						} `xml:"sz"`
						Bold   bool `xml:"b"`
						Italic bool `xml:"i"`
						Color  struct {
							Val string `xml:"val,attr"`
						} `xml:"color"`
					} `xml:"rPr"`
					Text string `xml:"t"`
				} `xml:"r"`
			} `xml:"p"`
		} `xml:"body"`
	}

	if err := xml.Unmarshal(content, &document); err != nil {
		return fmt.Errorf("failed to unmarshal document: %w", err)
	}

	// 解析段落
	for i, p := range document.Body.Paragraphs {
		paragraph := types.Paragraph{
			ID: fmt.Sprintf("paragraph_%d", i+1),
			Style: types.ParagraphStyle{
				Name: p.Properties.Style.Val,
			},
		}

		// 解析对齐方式
		if p.Properties.Justification.Val != "" {
			paragraph.Alignment = types.Alignment(p.Properties.Justification.Val)
		}

		// 解析缩进
		if p.Properties.Indentation.Left != "" {
			if val, err := strconv.ParseFloat(p.Properties.Indentation.Left, 64); err == nil {
				paragraph.Indentation.Left = val / 20.0
			}
		}
		if p.Properties.Indentation.Right != "" {
			if val, err := strconv.ParseFloat(p.Properties.Indentation.Right, 64); err == nil {
				paragraph.Indentation.Right = val / 20.0
			}
		}
		if p.Properties.Indentation.First != "" {
			if val, err := strconv.ParseFloat(p.Properties.Indentation.First, 64); err == nil {
				paragraph.Indentation.First = val / 20.0
			}
		}

		// 解析间距
		if p.Properties.Spacing.Before != "" {
			if val, err := strconv.ParseFloat(p.Properties.Spacing.Before, 64); err == nil {
				paragraph.Spacing.Before = val / 20.0
			}
		}
		if p.Properties.Spacing.After != "" {
			if val, err := strconv.ParseFloat(p.Properties.Spacing.After, 64); err == nil {
				paragraph.Spacing.After = val / 20.0
			}
		}
		if p.Properties.Spacing.Line != "" {
			if val, err := strconv.ParseFloat(p.Properties.Spacing.Line, 64); err == nil {
				paragraph.Spacing.Line = val / 240.0
			}
		}

		// 解析文本运行
		var paragraphText strings.Builder
		for j, r := range p.Runs {
			run := types.TextRun{
				ID:     fmt.Sprintf("run_%d_%d", i+1, j+1),
				Text:   r.Text,
				Bold:   r.Properties.Bold,
				Italic: r.Properties.Italic,
			}

			// 解析字体
			if r.Properties.Font.Val != "" {
				run.Font.Name = r.Properties.Font.Val
			}

			// 解析字体大小
			if r.Properties.Size.Val != "" {
				if sz, err := strconv.ParseFloat(r.Properties.Size.Val, 64); err == nil {
					run.Font.Size = sz / 2.0
					run.Size = sz / 2.0
				}
			}

			// 解析颜色
			if r.Properties.Color.Val != "" {
				run.Font.Color.RGB = r.Properties.Color.Val
				run.Color.RGB = r.Properties.Color.Val
			}

			paragraph.Runs = append(paragraph.Runs, run)
			paragraphText.WriteString(r.Text)
		}

		paragraph.Text = paragraphText.String()
		doc.Content.Paragraphs = append(doc.Content.Paragraphs, paragraph)
	}

	return nil
}

// parseStyles 解析样式
func (wd *WordprocessingDocument) parseStyles(doc *types.Document) error {
	// 简化样式解析，实际应该解析styles.xml
	doc.Styles = types.DocumentStyles{
		ParagraphStyles: []types.ParagraphStyle{},
		CharacterStyles: []types.CharacterStyle{},
		TableStyles:     []types.TableStyle{},
	}

	return nil
}

// parseFormatRules 解析格式规则
func (wd *WordprocessingDocument) parseFormatRules(doc *types.Document) error {
	// 从内容中提取格式规则
	if err := wd.extractFontRules(doc); err != nil {
		return err
	}

	if err := wd.extractParagraphRules(doc); err != nil {
		return err
	}

	if err := wd.extractPageRules(doc); err != nil {
		return err
	}

	return nil
}

// extractFontRules 提取字体规则
func (wd *WordprocessingDocument) extractFontRules(doc *types.Document) error {
	usedFonts := make(map[string]*types.FontRule)

	for _, para := range doc.Content.Paragraphs {
		for _, run := range para.Runs {
			if run.Font.Name != "" {
				if _, exists := usedFonts[run.Font.Name]; !exists {
					fontRule := &types.FontRule{
						ID:     run.Font.Name,
						Name:   run.Font.Name,
						Size:   run.Font.Size,
						Color:  run.Font.Color,
						Bold:   run.Bold,
						Italic: run.Italic,
					}
					usedFonts[run.Font.Name] = fontRule
				}
			}
		}
	}

	for _, fontRule := range usedFonts {
		doc.FormatRules.FontRules = append(doc.FormatRules.FontRules, *fontRule)
	}

	return nil
}

// extractParagraphRules 提取段落规则
func (wd *WordprocessingDocument) extractParagraphRules(doc *types.Document) error {
	for i, para := range doc.Content.Paragraphs {
		paragraphRule := types.ParagraphRule{
			ID:          fmt.Sprintf("paragraph_%d", i+1),
			Name:        para.Style.Name,
			Alignment:   para.Alignment,
			Indentation: para.Indentation,
			Spacing:     para.Spacing,
		}
		doc.FormatRules.ParagraphRules = append(doc.FormatRules.ParagraphRules, paragraphRule)
	}

	return nil
}

// extractPageRules 提取页面规则
func (wd *WordprocessingDocument) extractPageRules(doc *types.Document) error {
	// 默认页面规则
	pageRule := types.PageRule{
		ID:   "page_1",
		Name: "Default Page",
		PageSize: types.PageSize{
			Width:  11906.0 / 20.0, // A4宽度
			Height: 16838.0 / 20.0, // A4高度
		},
		PageMargins: types.PageMargins{
			Top:    1440.0 / 20.0,
			Bottom: 1440.0 / 20.0,
			Left:   1800.0 / 20.0,
			Right:  1800.0 / 20.0,
			Header: 851.0 / 20.0,
			Footer: 992.0 / 20.0,
		},
		HeaderDistance: 851.0 / 20.0,
		FooterDistance: 992.0 / 20.0,
	}

	doc.FormatRules.PageRules = append(doc.FormatRules.PageRules, pageRule)
	return nil
}

// Close 关闭Word文档
func (wd *WordprocessingDocument) Close() error {
	if wd.Container != nil {
		return wd.Container.Close()
	}
	return nil
} 