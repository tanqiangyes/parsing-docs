package documents

import (
	"encoding/xml"
	"fmt"
	"strconv"
	"strings"
	"time"

	"docs-parser/internal/core/types"
	"docs-parser/internal/packaging"
	"docs-parser/internal/utils"
)

// WordprocessingDocument 表示Word文档
type WordprocessingDocument struct {
	Container *packaging.OPCContainer
	Document  *types.Document
	Parts     map[string]*DocumentPart
	Monitor   *utils.PerformanceMonitor
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
		Monitor:   utils.NewPerformanceMonitor(),
	}
}

// Open 打开Word文档
func (wd *WordprocessingDocument) Open() error {
	openStep := wd.Monitor.StartStep("打开OPC容器")
	defer openStep()

	// 打开OPC容器
	if err := wd.Container.Open(); err != nil {
		return fmt.Errorf("failed to open OPC container: %w", err)
	}

	// 验证容器
	if err := wd.Container.Validate(); err != nil {
		return fmt.Errorf("invalid OPC container: %w", err)
	}

	// 加载文档部分
	loadStep := wd.Monitor.StartStep("加载文档部分")
	defer loadStep()

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
	metadataStep := wd.Monitor.StartStep("解析元数据")
	if err := wd.parseMetadata(doc); err != nil {
		return nil, fmt.Errorf("failed to parse metadata: %w", err)
	}
	metadataStep()

	// 解析内容
	contentStep := wd.Monitor.StartStep("解析内容")
	if err := wd.parseContent(doc); err != nil {
		return nil, fmt.Errorf("failed to parse content: %w", err)
	}
	contentStep()

	// 解析样式
	styleStep := wd.Monitor.StartStep("解析样式")
	if err := wd.parseStyles(doc); err != nil {
		return nil, fmt.Errorf("failed to parse styles: %w", err)
	}
	styleStep()

	// 解析格式规则
	formatStep := wd.Monitor.StartStep("解析格式规则")
	if err := wd.parseFormatRules(doc); err != nil {
		return nil, fmt.Errorf("failed to parse format rules: %w", err)
	}
	formatStep()

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
			XMLName    xml.Name `xml:"body"`
			Paragraphs []struct {
				XMLName    xml.Name `xml:"p"`
				Properties struct {
					Style struct {
						Val string `xml:"val,attr"`
					} `xml:"pStyle"`
					Justification struct {
						Val string `xml:"val,attr"`
					} `xml:"jc"`
					Indentation struct {
						Left    string `xml:"left,attr"`
						Right   string `xml:"right,attr"`
						First   string `xml:"firstLine,attr"`
						Hanging string `xml:"hanging,attr"`
					} `xml:"ind"`
					Spacing struct {
						Before string `xml:"before,attr"`
						After  string `xml:"after,attr"`
						Line   string `xml:"line,attr"`
					} `xml:"spacing"`
				} `xml:"pPr"`
				Runs []struct {
					XMLName    xml.Name `xml:"r"`
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
			Tables []struct {
				XMLName    xml.Name `xml:"tbl"`
				Properties struct {
					Width struct {
						Val string `xml:"val,attr"`
					} `xml:"tblW"`
					Justification struct {
						Val string `xml:"val,attr"`
					} `xml:"jc"`
					Borders struct {
						Top struct {
							Val string `xml:"val,attr"`
						} `xml:"top"`
						Bottom struct {
							Val string `xml:"val,attr"`
						} `xml:"bottom"`
						Left struct {
							Val string `xml:"val,attr"`
						} `xml:"left"`
						Right struct {
							Val string `xml:"val,attr"`
						} `xml:"right"`
					} `xml:"tblBorders"`
				} `xml:"tblPr"`
				Rows []struct {
					XMLName    xml.Name `xml:"tr"`
					Properties struct {
						Height struct {
							Val string `xml:"val,attr"`
						} `xml:"trHeight"`
					} `xml:"trPr"`
					Cells []struct {
						XMLName    xml.Name `xml:"tc"`
						Properties struct {
							Width struct {
								Val string `xml:"val,attr"`
							} `xml:"tcW"`
							Borders struct {
								Top struct {
									Val string `xml:"val,attr"`
								} `xml:"top"`
								Bottom struct {
									Val string `xml:"val,attr"`
								} `xml:"bottom"`
								Left struct {
									Val string `xml:"val,attr"`
								} `xml:"left"`
								Right struct {
									Val string `xml:"val,attr"`
								} `xml:"right"`
							} `xml:"tcBorders"`
						} `xml:"tcPr"`
						Paragraphs []struct {
							Properties struct {
								Alignment struct {
									Val string `xml:"val,attr"`
								} `xml:"jc"`
							} `xml:"pPr"`
							Runs []struct {
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
					} `xml:"tc"`
				} `xml:"tr"`
			} `xml:"tbl"`
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

	// 解析表格
	for i, t := range document.Body.Tables {
		table := types.Table{
			ID: fmt.Sprintf("table_%d", i+1),
		}

		// 解析表格属性
		if t.Properties.Width.Val != "" {
			if val, err := strconv.ParseFloat(t.Properties.Width.Val, 64); err == nil {
				table.Width = val / 20.0
			}
		}

		if t.Properties.Justification.Val != "" {
			table.Alignment = types.Alignment(t.Properties.Justification.Val)
		}

		// 解析表格边框
		if t.Properties.Borders.Top.Val != "" {
			table.Borders.Top.Style = types.BorderStyle(t.Properties.Borders.Top.Val)
		}
		if t.Properties.Borders.Bottom.Val != "" {
			table.Borders.Bottom.Style = types.BorderStyle(t.Properties.Borders.Bottom.Val)
		}
		if t.Properties.Borders.Left.Val != "" {
			table.Borders.Left.Style = types.BorderStyle(t.Properties.Borders.Left.Val)
		}
		if t.Properties.Borders.Right.Val != "" {
			table.Borders.Right.Style = types.BorderStyle(t.Properties.Borders.Right.Val)
		}

		for j, row := range t.Rows {
			tableRow := types.TableRow{
				ID: fmt.Sprintf("row_%d_%d", i+1, j+1),
			}

			// 解析行高度
			if row.Properties.Height.Val != "" {
				if val, err := strconv.ParseFloat(row.Properties.Height.Val, 64); err == nil {
					tableRow.Height = val / 20.0
				}
			}

			for k, cell := range row.Cells {
				tableCell := types.TableCell{
					ID: fmt.Sprintf("cell_%d_%d_%d", i+1, j+1, k+1),
				}

				// 解析单元格宽度
				if cell.Properties.Width.Val != "" {
					if val, err := strconv.ParseFloat(cell.Properties.Width.Val, 64); err == nil {
						tableCell.Width = val / 20.0
					}
				}

				// 解析单元格边框
				if cell.Properties.Borders.Top.Val != "" {
					tableCell.Borders.Top.Style = types.BorderStyle(cell.Properties.Borders.Top.Val)
				}
				if cell.Properties.Borders.Bottom.Val != "" {
					tableCell.Borders.Bottom.Style = types.BorderStyle(cell.Properties.Borders.Bottom.Val)
				}
				if cell.Properties.Borders.Left.Val != "" {
					tableCell.Borders.Left.Style = types.BorderStyle(cell.Properties.Borders.Left.Val)
				}
				if cell.Properties.Borders.Right.Val != "" {
					tableCell.Borders.Right.Style = types.BorderStyle(cell.Properties.Borders.Right.Val)
				}

				// 解析单元格内容
				var cellText strings.Builder
				for _, para := range cell.Paragraphs {
					cellParagraph := types.Paragraph{
						ID: fmt.Sprintf("cell_para_%d_%d_%d", i+1, j+1, k+1),
					}

					// 解析段落对齐方式
					if para.Properties.Alignment.Val != "" {
						cellParagraph.Alignment = types.Alignment(para.Properties.Alignment.Val)
					}

					// 解析段落文本运行
					for _, run := range para.Runs {
						cellRun := types.TextRun{
							ID:     fmt.Sprintf("cell_run_%d_%d_%d", i+1, j+1, k+1),
							Text:   run.Text,
							Bold:   run.Properties.Bold,
							Italic: run.Properties.Italic,
						}

						// 解析字体
						if run.Properties.Font.Val != "" {
							cellRun.Font.Name = run.Properties.Font.Val
						}

						// 解析字体大小
						if run.Properties.Size.Val != "" {
							if sz, err := strconv.ParseFloat(run.Properties.Size.Val, 64); err == nil {
								cellRun.Font.Size = sz / 2.0
								cellRun.Size = sz / 2.0
							}
						}

						// 解析颜色
						if run.Properties.Color.Val != "" {
							cellRun.Font.Color.RGB = run.Properties.Color.Val
							cellRun.Color.RGB = run.Properties.Color.Val
						}

						cellParagraph.Runs = append(cellParagraph.Runs, cellRun)
						cellText.WriteString(run.Text)
					}

					cellParagraph.Text = cellText.String()
					tableCell.Content = append(tableCell.Content, cellParagraph)
				}

				tableRow.Cells = append(tableRow.Cells, tableCell)
			}

			table.Rows = append(table.Rows, tableRow)
		}

		doc.Content.Tables = append(doc.Content.Tables, table)
	}

	return nil
}

// parseStyles 解析样式
func (wd *WordprocessingDocument) parseStyles(doc *types.Document) error {
	// 初始化样式结构
	doc.Styles = types.DocumentStyles{
		ParagraphStyles: []types.ParagraphStyle{},
		CharacterStyles: []types.CharacterStyle{},
		TableStyles:     []types.TableStyle{},
	}

	// 尝试解析styles.xml
	if err := wd.parseStylesXML(doc); err != nil {
		// 如果styles.xml不存在或解析失败，从内联样式中提取
		if err := wd.extractInlineStyles(doc); err != nil {
			return fmt.Errorf("failed to parse styles: %w", err)
		}
	}

	return nil
}

// parseStylesXML 解析styles.xml文件
func (wd *WordprocessingDocument) parseStylesXML(doc *types.Document) error {
	part, exists := wd.Parts["styles.xml"]
	if !exists {
		return fmt.Errorf("styles.xml not found")
	}

	var stylesDoc struct {
		XMLName xml.Name `xml:"styles"`
		Styles  []struct {
			XMLName xml.Name `xml:"style"`
			ID      string   `xml:"styleId,attr"`
			Name    string   `xml:"name,attr"`
			Type    string   `xml:"type,attr"`
			BasedOn struct {
				Val string `xml:"val,attr"`
			} `xml:"basedOn"`
			Next struct {
				Val string `xml:"val,attr"`
			} `xml:"next"`
			Linked struct {
				Val string `xml:"val,attr"`
			} `xml:"link"`
			Properties struct {
				Font struct {
					Name string `xml:"name,attr"`
					Size string `xml:"size,attr"`
				} `xml:"rFonts"`
				Paragraph struct {
					Alignment struct {
						Val string `xml:"val,attr"`
					} `xml:"jc"`
					Indentation struct {
						Left    string `xml:"left,attr"`
						Right   string `xml:"right,attr"`
						First   string `xml:"firstLine,attr"`
						Hanging string `xml:"hanging,attr"`
					} `xml:"ind"`
					Spacing struct {
						Before string `xml:"before,attr"`
						After  string `xml:"after,attr"`
						Line   string `xml:"line,attr"`
					} `xml:"spacing"`
				} `xml:"pPr"`
			} `xml:"rPr"`
		} `xml:"style"`
	}

	if err := xml.Unmarshal(part.Content, &stylesDoc); err != nil {
		return fmt.Errorf("failed to unmarshal styles.xml: %w", err)
	}

	// 解析样式
	for _, style := range stylesDoc.Styles {
		switch style.Type {
		case "paragraph":
			paraStyle := types.ParagraphStyle{
				ID:   style.ID,
				Name: style.Name,
			}
			doc.Styles.ParagraphStyles = append(doc.Styles.ParagraphStyles, paraStyle)

		case "character":
			charStyle := types.CharacterStyle{
				ID:   style.ID,
				Name: style.Name,
			}
			doc.Styles.CharacterStyles = append(doc.Styles.CharacterStyles, charStyle)

		case "table":
			tableStyle := types.TableStyle{
				ID:   style.ID,
				Name: style.Name,
			}
			doc.Styles.TableStyles = append(doc.Styles.TableStyles, tableStyle)
		}
	}

	return nil
}

// extractInlineStyles 从内联样式中提取样式信息
func (wd *WordprocessingDocument) extractInlineStyles(doc *types.Document) error {
	// 从文档内容中提取使用的样式
	usedStyles := make(map[string]bool)

	// 从段落中提取样式
	for _, para := range doc.Content.Paragraphs {
		if para.Style.Name != "" {
			usedStyles[para.Style.Name] = true
		}
	}

	// 从文本运行中提取样式
	for _, para := range doc.Content.Paragraphs {
		for _, run := range para.Runs {
			if run.Font.Name != "" {
				usedStyles[run.Font.Name] = true
			}
		}
	}

	// 创建样式对象
	for styleName := range usedStyles {
		// 创建段落样式
		paraStyle := types.ParagraphStyle{
			ID:   styleName,
			Name: styleName,
		}
		doc.Styles.ParagraphStyles = append(doc.Styles.ParagraphStyles, paraStyle)

		// 创建字符样式
		charStyle := types.CharacterStyle{
			ID:   styleName,
			Name: styleName,
		}
		doc.Styles.CharacterStyles = append(doc.Styles.CharacterStyles, charStyle)
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
	// 首先尝试从fontTable.xml获取字体信息
	fontMap := make(map[string]*types.FontRule)

	// 尝试解析fontTable.xml
	if err := wd.parseFontTable(fontMap); err != nil {
		// 如果fontTable.xml不存在，从内联样式中提取
		if err := wd.extractInlineFonts(doc, fontMap); err != nil {
			return fmt.Errorf("failed to extract font rules: %w", err)
		}
	}

	// 将字体规则添加到文档中
	for _, fontRule := range fontMap {
		doc.FormatRules.FontRules = append(doc.FormatRules.FontRules, *fontRule)
	}

	return nil
}

// parseFontTable 解析fontTable.xml
func (wd *WordprocessingDocument) parseFontTable(fontMap map[string]*types.FontRule) error {
	part, exists := wd.Parts["fontTable.xml"]
	if !exists {
		return fmt.Errorf("fontTable.xml not found")
	}

	var fontTable struct {
		XMLName xml.Name `xml:"fontTable"`
		Fonts   []struct {
			XMLName xml.Name `xml:"font"`
			Name    string   `xml:"name,attr"`
			Family  struct {
				Val string `xml:"val,attr"`
			} `xml:"family"`
			Pitch struct {
				Val string `xml:"val,attr"`
			} `xml:"pitch"`
		} `xml:"font"`
	}

	if err := xml.Unmarshal(part.Content, &fontTable); err != nil {
		return fmt.Errorf("failed to unmarshal fontTable.xml: %w", err)
	}

	// 创建字体映射
	for _, font := range fontTable.Fonts {
		fontRule := &types.FontRule{
			ID:    font.Name,
			Name:  font.Name,
			Size:  12.0,                       // 默认大小
			Color: types.Color{RGB: "000000"}, // 默认黑色
		}
		fontMap[font.Name] = fontRule
	}

	return nil
}

// extractInlineFonts 从内联样式中提取字体信息
func (wd *WordprocessingDocument) extractInlineFonts(doc *types.Document, fontMap map[string]*types.FontRule) error {
	// 从文档内容中提取使用的字体
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
				} else {
					// 更新现有字体规则，合并属性
					existing := usedFonts[run.Font.Name]
					if run.Font.Size > 0 {
						existing.Size = run.Font.Size
					}
					if run.Font.Color.RGB != "" {
						existing.Color = run.Font.Color
					}
					existing.Bold = existing.Bold || run.Bold
					existing.Italic = existing.Italic || run.Italic
				}
			}
		}
	}

	// 将提取的字体规则复制到fontMap
	for name, rule := range usedFonts {
		fontMap[name] = rule
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
