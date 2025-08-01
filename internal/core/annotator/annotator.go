package annotator

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"

	"docs-parser/internal/core/types"
	"docs-parser/internal/formats"
	"docs-parser/internal/utils"
)

// Annotator 文档标注器
type Annotator struct {
	wordParser *formats.WordParser
}

// NewAnnotator 创建新的标注器
func NewAnnotator() *Annotator {
	return &Annotator{
		wordParser: formats.NewWordParser(),
	}
}

// AnnotateDocument 为文档添加格式标注
func (a *Annotator) AnnotateDocument(sourcePath, outputPath string) error {
	// 1. 复制源文件
	if err := utils.CopyFile(sourcePath, outputPath); err != nil {
		return fmt.Errorf("failed to copy file: %w", err)
	}

	// 2. 解析文档
	doc, err := a.wordParser.ParseDocument(sourcePath)
	if err != nil {
		return fmt.Errorf("failed to parse document: %w", err)
	}

	// 3. 分析格式问题
	issues, err := a.analyzeFormatIssues(doc)
	if err != nil {
		return fmt.Errorf("failed to analyze format issues: %w", err)
	}

	// 4. 在复制的文件上添加标注
	if err := a.addAnnotations(outputPath, issues); err != nil {
		return fmt.Errorf("failed to add annotations: %w", err)
	}

	return nil
}

// analyzeFormatIssues 分析格式问题
func (a *Annotator) analyzeFormatIssues(doc *types.Document) ([]FormatIssue, error) {
	var issues []FormatIssue

	// 分析字体问题
	fontIssues := a.analyzeFontIssues(doc)
	issues = append(issues, fontIssues...)

	// 分析段落问题
	paragraphIssues := a.analyzeParagraphIssues(doc)
	issues = append(issues, paragraphIssues...)

	// 分析表格问题
	tableIssues := a.analyzeTableIssues(doc)
	issues = append(issues, tableIssues...)

	// 分析页面问题
	pageIssues := a.analyzePageIssues(doc)
	issues = append(issues, pageIssues...)

	return issues, nil
}

// analyzeFontIssues 分析字体问题
func (a *Annotator) analyzeFontIssues(doc *types.Document) []FormatIssue {
	var issues []FormatIssue

	// 检查每个段落的字体
	for _, paragraph := range doc.Content.Paragraphs {
		for _, run := range paragraph.Runs {
			// 检查字体大小
			if run.Font.Size < 10.0 {
				issues = append(issues, FormatIssue{
					ID:          fmt.Sprintf("font_size_%s", run.ID),
					Type:        "font",
					Severity:    "medium",
					Location:    fmt.Sprintf("paragraph_%s", paragraph.ID),
					Description: "字体大小过小",
					Current:     run.Font.Size,
					Expected:    ">= 10.0",
					Rule:        "font_size_minimum",
					Suggestions: []string{"将字体大小调整为10.0或更大"},
				})
			}

			// 检查字体名称
			if run.Font.Name == "" {
				issues = append(issues, FormatIssue{
					ID:          fmt.Sprintf("font_name_%s", run.ID),
					Type:        "font",
					Severity:    "high",
					Location:    fmt.Sprintf("paragraph_%s", paragraph.ID),
					Description: "字体名称未设置",
					Current:     run.Font.Name,
					Expected:    "标准字体名称",
					Rule:        "font_name_required",
					Suggestions: []string{"设置标准字体名称，如宋体、黑体等"},
				})
			}
		}
	}

	return issues
}

// analyzeParagraphIssues 分析段落问题
func (a *Annotator) analyzeParagraphIssues(doc *types.Document) []FormatIssue {
	var issues []FormatIssue

	for _, paragraph := range doc.Content.Paragraphs {
		// 检查段落对齐方式
		if paragraph.Alignment == "" {
			issues = append(issues, FormatIssue{
				ID:          fmt.Sprintf("paragraph_alignment_%s", paragraph.ID),
				Type:        "paragraph",
				Severity:    "medium",
				Location:    fmt.Sprintf("paragraph_%s", paragraph.ID),
				Description: "段落对齐方式未设置",
				Current:     paragraph.Alignment,
				Expected:    "left/center/right/justify",
				Rule:        "paragraph_alignment_required",
				Suggestions: []string{"设置段落对齐方式"},
			})
		}

		// 检查段落间距
		if paragraph.Spacing.Before < 0 || paragraph.Spacing.After < 0 {
			issues = append(issues, FormatIssue{
				ID:          fmt.Sprintf("paragraph_spacing_%s", paragraph.ID),
				Type:        "paragraph",
				Severity:    "low",
				Location:    fmt.Sprintf("paragraph_%s", paragraph.ID),
				Description: "段落间距设置不当",
				Current:     paragraph.Spacing,
				Expected:    ">= 0",
				Rule:        "paragraph_spacing_valid",
				Suggestions: []string{"调整段落间距为正值"},
			})
		}
	}

	return issues
}

// analyzeTableIssues 分析表格问题
func (a *Annotator) analyzeTableIssues(doc *types.Document) []FormatIssue {
	var issues []FormatIssue

	for _, table := range doc.Content.Tables {
		// 检查表格边框
		if table.Borders.Top.Style == "" {
			issues = append(issues, FormatIssue{
				ID:          fmt.Sprintf("table_border_%s", table.ID),
				Type:        "table",
				Severity:    "medium",
				Location:    fmt.Sprintf("table_%s", table.ID),
				Description: "表格边框未设置",
				Current:     table.Borders,
				Expected:    "完整的边框设置",
				Rule:        "table_border_required",
				Suggestions: []string{"为表格设置完整的边框"},
			})
		}

		// 检查表格宽度
		if table.Width <= 0 {
			issues = append(issues, FormatIssue{
				ID:          fmt.Sprintf("table_width_%s", table.ID),
				Type:        "table",
				Severity:    "low",
				Location:    fmt.Sprintf("table_%s", table.ID),
				Description: "表格宽度未设置",
				Current:     table.Width,
				Expected:    "> 0",
				Rule:        "table_width_required",
				Suggestions: []string{"设置表格宽度"},
			})
		}
	}

	return issues
}

// analyzePageIssues 分析页面问题
func (a *Annotator) analyzePageIssues(doc *types.Document) []FormatIssue {
	var issues []FormatIssue

	for _, section := range doc.Content.Sections {
		// 检查页面边距
		if section.PageMargins.Top < 0 || section.PageMargins.Bottom < 0 ||
			section.PageMargins.Left < 0 || section.PageMargins.Right < 0 {
			issues = append(issues, FormatIssue{
				ID:          fmt.Sprintf("page_margins_%s", section.ID),
				Type:        "page",
				Severity:    "medium",
				Location:    fmt.Sprintf("section_%s", section.ID),
				Description: "页面边距设置不当",
				Current:     section.PageMargins,
				Expected:    "所有边距 >= 0",
				Rule:        "page_margins_valid",
				Suggestions: []string{"调整页面边距为正值"},
			})
		}

		// 检查页面大小
		if section.PageSize.Width <= 0 || section.PageSize.Height <= 0 {
			issues = append(issues, FormatIssue{
				ID:          fmt.Sprintf("page_size_%s", section.ID),
				Type:        "page",
				Severity:    "high",
				Location:    fmt.Sprintf("section_%s", section.ID),
				Description: "页面大小设置不当",
				Current:     section.PageSize,
				Expected:    "宽度和高度 > 0",
				Rule:        "page_size_valid",
				Suggestions: []string{"设置正确的页面大小"},
			})
		}
	}

	return issues
}

// addAnnotations 添加标注
func (a *Annotator) addAnnotations(filePath string, issues []FormatIssue) error {
	ext := filepath.Ext(filePath)

	switch ext {
	case ".docx":
		return a.addDocxAnnotations(filePath, issues)
	case ".doc":
		return a.addDocAnnotations(filePath, issues)
	case ".rtf":
		return a.addRtfAnnotations(filePath, issues)
	default:
		return fmt.Errorf("unsupported file format: %s", ext)
	}
}

// addDocxAnnotations 为.docx文件添加标注
func (a *Annotator) addDocxAnnotations(filePath string, issues []FormatIssue) error {
	// 打开zip文件
	reader, err := zip.OpenReader(filePath)
	if err != nil {
		return err
	}
	defer reader.Close()

	// 创建临时文件
	tempFile, err := os.CreateTemp("", "annotated_*.docx")
	if err != nil {
		return err
	}
	defer os.Remove(tempFile.Name())
	defer tempFile.Close()

	// 创建新的zip文件
	writer := zip.NewWriter(tempFile)
	defer writer.Close()

	// 复制所有文件并添加标注
	for _, file := range reader.File {
		rc, err := file.Open()
		if err != nil {
			return err
		}

		// 创建新文件
		newFile, err := writer.Create(file.Name)
		if err != nil {
			rc.Close()
			return err
		}

		// 如果是文档内容文件，添加标注
		if file.Name == "word/document.xml" {
			if err := a.addAnnotationsToDocument(rc, newFile, issues); err != nil {
				rc.Close()
				return err
			}
		} else {
			// 直接复制其他文件
			if _, err := io.Copy(newFile, rc); err != nil {
				rc.Close()
				return err
			}
		}

		rc.Close()
	}

	// 替换原文件
	if err := os.Rename(tempFile.Name(), filePath); err != nil {
		return err
	}

	return nil
}

// addAnnotationsToDocument 为文档内容添加标注
func (a *Annotator) addAnnotationsToDocument(reader io.Reader, writer io.Writer, issues []FormatIssue) error {
	// 读取XML内容
	content, err := io.ReadAll(reader)
	if err != nil {
		return err
	}

	// 解析XML
	var doc DocumentXML
	if err := xml.Unmarshal(content, &doc); err != nil {
		return err
	}

	// 添加标注
	a.addAnnotationsToXML(&doc, issues)

	// 重新序列化XML
	output, err := xml.MarshalIndent(doc, "", "  ")
	if err != nil {
		return err
	}

	// 写入XML声明
	if _, err := writer.Write([]byte(xml.Header)); err != nil {
		return err
	}

	// 写入内容
	if _, err := writer.Write(output); err != nil {
		return err
	}

	return nil
}

// addAnnotationsToXML 为XML文档添加标注
func (a *Annotator) addAnnotationsToXML(doc *DocumentXML, issues []FormatIssue) {
	// 为每个问题添加标注
	for _, issue := range issues {
		// 根据问题类型添加不同的标注
		switch issue.Type {
		case "font":
			a.addFontAnnotation(doc, issue)
		case "paragraph":
			a.addParagraphAnnotation(doc, issue)
		case "table":
			a.addTableAnnotation(doc, issue)
		case "page":
			a.addPageAnnotation(doc, issue)
		}
	}
}

// addFontAnnotation 添加字体标注
func (a *Annotator) addFontAnnotation(doc *DocumentXML, issue FormatIssue) {
	// 在文档中添加字体问题标注
	comment := fmt.Sprintf("格式问题: %s - %s", issue.Description, strings.Join(issue.Suggestions, "; "))

	// 查找对应的段落并添加标注
	for _, paragraph := range doc.Body.Paragraphs {
		if strings.Contains(paragraph.ID, issue.Location) {
			// 添加注释
			paragraph.Comments = append(paragraph.Comments, Comment{
				ID:   issue.ID,
				Text: comment,
			})
			break
		}
	}
}

// addParagraphAnnotation 添加段落标注
func (a *Annotator) addParagraphAnnotation(doc *DocumentXML, issue FormatIssue) {
	comment := fmt.Sprintf("段落格式问题: %s - %s", issue.Description, strings.Join(issue.Suggestions, "; "))

	for _, paragraph := range doc.Body.Paragraphs {
		if strings.Contains(paragraph.ID, issue.Location) {
			paragraph.Comments = append(paragraph.Comments, Comment{
				ID:   issue.ID,
				Text: comment,
			})
			break
		}
	}
}

// addTableAnnotation 添加表格标注
func (a *Annotator) addTableAnnotation(doc *DocumentXML, issue FormatIssue) {
	comment := fmt.Sprintf("表格格式问题: %s - %s", issue.Description, strings.Join(issue.Suggestions, "; "))

	for _, table := range doc.Body.Tables {
		if strings.Contains(table.ID, issue.Location) {
			table.Comments = append(table.Comments, Comment{
				ID:   issue.ID,
				Text: comment,
			})
			break
		}
	}
}

// addPageAnnotation 添加页面标注
func (a *Annotator) addPageAnnotation(doc *DocumentXML, issue FormatIssue) {
	comment := fmt.Sprintf("页面格式问题: %s - %s", issue.Description, strings.Join(issue.Suggestions, "; "))

	// 在文档级别添加页面问题标注
	doc.Body.PageComments = append(doc.Body.PageComments, Comment{
		ID:   issue.ID,
		Text: comment,
	})
}

// addDocAnnotations 为.doc文件添加标注
func (a *Annotator) addDocAnnotations(filePath string, issues []FormatIssue) error {
	// 读取文件内容
	content, err := os.ReadFile(filePath)
	if err != nil {
		return err
	}

	// 添加标注到.doc文件
	annotatedContent := a.addAnnotationsToDocContent(content, issues)

	// 写回文件
	return os.WriteFile(filePath, annotatedContent, 0644)
}

// addAnnotationsToDocContent 为.doc文件内容添加标注
func (a *Annotator) addAnnotationsToDocContent(content []byte, issues []FormatIssue) []byte {
	// 创建标注信息
	var annotations []string
	for _, issue := range issues {
		annotation := fmt.Sprintf("/* 格式问题: %s - %s */", issue.Description, strings.Join(issue.Suggestions, "; "))
		annotations = append(annotations, annotation)
	}

	// 在文件开头添加标注
	header := fmt.Sprintf("/* 文档格式标注 - 共发现 %d 个问题 */\n", len(issues))
	for _, annotation := range annotations {
		header += annotation + "\n"
	}
	header += "\n"

	// 组合新内容
	return append([]byte(header), content...)
}

// addRtfAnnotations 为.rtf文件添加标注
func (a *Annotator) addRtfAnnotations(filePath string, issues []FormatIssue) error {
	// 读取文件内容
	content, err := os.ReadFile(filePath)
	if err != nil {
		return err
	}

	// 添加标注到.rtf文件
	annotatedContent := a.addAnnotationsToRtfContent(content, issues)

	// 写回文件
	return os.WriteFile(filePath, annotatedContent, 0644)
}

// addAnnotationsToRtfContent 为.rtf文件内容添加标注
func (a *Annotator) addAnnotationsToRtfContent(content []byte, issues []FormatIssue) []byte {
	// 创建RTF格式的标注
	var annotations []string
	for _, issue := range issues {
		annotation := fmt.Sprintf("\\par \\cf0 \\fs16 \\b 格式问题: %s\\b0 \\par \\cf1 %s",
			issue.Description, strings.Join(issue.Suggestions, "; "))
		annotations = append(annotations, annotation)
	}

	// 在文档开头添加标注
	header := fmt.Sprintf("{\\rtf1\\ansi\\deff0 {\\fonttbl {\\f0 Times New Roman;}}\n")
	header += fmt.Sprintf("\\cf0 \\fs24 \\b 文档格式标注 - 共发现 %d 个问题\\b0 \\par \\par\n", len(issues))
	for _, annotation := range annotations {
		header += annotation + "\\par \\par\n"
	}
	header += "\\cf0 \\fs24 \\b 原始文档内容:\\b0 \\par \\par\n"

	// 组合新内容
	return append([]byte(header), content...)
}

// FormatIssue 格式问题
type FormatIssue struct {
	ID          string      `json:"id"`
	Type        string      `json:"type"`
	Severity    string      `json:"severity"`
	Location    string      `json:"location"`
	Description string      `json:"description"`
	Current     interface{} `json:"current"`
	Expected    interface{} `json:"expected"`
	Rule        string      `json:"rule"`
	Suggestions []string    `json:"suggestions"`
}

// XML文档结构（简化版）
type DocumentXML struct {
	XMLName xml.Name `xml:"w:document"`
	Body    BodyXML  `xml:"w:body"`
}

type BodyXML struct {
	Paragraphs   []ParagraphXML `xml:"w:p"`
	Tables       []TableXML     `xml:"w:tbl"`
	PageComments []Comment      `xml:"w:comment"`
}

type ParagraphXML struct {
	ID       string    `xml:"id,attr"`
	Runs     []RunXML  `xml:"w:r"`
	Comments []Comment `xml:"w:comment"`
}

type RunXML struct {
	ID   string  `xml:"id,attr"`
	Text string  `xml:"w:t"`
	Font FontXML `xml:"w:rPr"`
}

type FontXML struct {
	Name string  `xml:"w:rFonts,attr"`
	Size float64 `xml:"w:sz,attr"`
}

type TableXML struct {
	ID       string    `xml:"id,attr"`
	Rows     []RowXML  `xml:"w:tr"`
	Comments []Comment `xml:"w:comment"`
}

type RowXML struct {
	Cells []CellXML `xml:"w:tc"`
}

type CellXML struct {
	Content string `xml:"w:p"`
}

type Comment struct {
	ID   string `xml:"id,attr"`
	Text string `xml:"w:t"`
}
