package templates

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"time"

	"docs-parser/internal/core/types"
	"docs-parser/internal/formats"
)

// Template 文档格式模板
type Template struct {
	ID          string            `json:"id"`
	Name        string            `json:"name"`
	Description string            `json:"description"`
	Version     string            `json:"version"`
	FormatRules types.FormatRules `json:"format_rules"`
	Metadata    TemplateMetadata  `json:"metadata"`
	SourcePath  string            `json:"source_path"` // Word文档路径
}

// TemplateMetadata 模板元数据
type TemplateMetadata struct {
	Author      string   `json:"author"`
	Created     string   `json:"created"`
	LastUpdated string   `json:"last_updated"`
	Category    string   `json:"category"`
	Tags        []string `json:"tags"`
}

// TemplateManager 模板管理器
type TemplateManager struct {
	templates  map[string]*Template
	basePath   string
	wordParser *formats.WordParser
}

// NewTemplateManager 创建模板管理器
func NewTemplateManager(basePath string) *TemplateManager {
	return &TemplateManager{
		templates:  make(map[string]*Template),
		basePath:   basePath,
		wordParser: formats.NewWordParser(),
	}
}

// LoadTemplate 从Word文档加载模板
func (tm *TemplateManager) LoadTemplate(templatePath string) (*Template, error) {
	// 验证文件路径
	if templatePath == "" {
		return nil, fmt.Errorf("template path is empty")
	}

	// 检查文件扩展名
	ext := strings.ToLower(filepath.Ext(templatePath))
	if !tm.isSupportedWordFormat(ext) {
		return nil, fmt.Errorf("unsupported template format: %s, only Word documents (.docx, .doc, .dot, .dotx) are supported", ext)
	}

	// 解析Word文档
	doc, err := tm.wordParser.ParseDocument(templatePath)
	if err != nil {
		return nil, fmt.Errorf("failed to parse Word template: %w", err)
	}

	// 从Word文档创建模板
	template, err := tm.createTemplateFromDocument(doc, templatePath)
	if err != nil {
		return nil, fmt.Errorf("failed to create template from document: %w", err)
	}

	// 验证模板
	if err := tm.validateTemplate(template); err != nil {
		return nil, fmt.Errorf("template validation failed: %w", err)
	}

	tm.templates[template.ID] = template
	return template, nil
}

// isSupportedWordFormat 检查是否为支持的Word格式
func (tm *TemplateManager) isSupportedWordFormat(ext string) bool {
	supportedFormats := []string{".docx", ".doc", ".dot", ".dotx"}
	for _, format := range supportedFormats {
		if ext == format {
			return true
		}
	}
	return false
}

// createTemplateFromDocument 从Word文档创建模板
func (tm *TemplateManager) createTemplateFromDocument(doc *types.Document, templatePath string) (*Template, error) {
	// 生成模板ID和名称
	templateID := tm.generateTemplateID(templatePath)
	templateName := tm.generateTemplateName(templatePath)

	// 提取格式规则
	formatRules := tm.extractFormatRules(doc)

	// 创建模板元数据
	metadata := tm.createTemplateMetadata(templatePath, doc)

	// 创建模板
	template := &Template{
		ID:          templateID,
		Name:        templateName,
		Description: fmt.Sprintf("Template created from Word document: %s", filepath.Base(templatePath)),
		Version:     "1.0",
		FormatRules: formatRules,
		Metadata:    metadata,
		SourcePath:  templatePath,
	}

	return template, nil
}

// generateTemplateID 生成模板ID
func (tm *TemplateManager) generateTemplateID(templatePath string) string {
	baseName := filepath.Base(templatePath)
	ext := filepath.Ext(baseName)
	nameWithoutExt := baseName[:len(baseName)-len(ext)]

	// 清理文件名，使其适合作为ID
	cleanName := strings.ReplaceAll(nameWithoutExt, " ", "_")
	cleanName = strings.ReplaceAll(cleanName, "-", "_")
	cleanName = strings.ToLower(cleanName)

	return fmt.Sprintf("word_template_%s", cleanName)
}

// generateTemplateName 生成模板名称
func (tm *TemplateManager) generateTemplateName(templatePath string) string {
	baseName := filepath.Base(templatePath)
	ext := filepath.Ext(baseName)
	nameWithoutExt := baseName[:len(baseName)-len(ext)]

	return fmt.Sprintf("Word Template: %s", nameWithoutExt)
}

// extractFormatRules 从Word文档提取格式规则
func (tm *TemplateManager) extractFormatRules(doc *types.Document) types.FormatRules {
	formatRules := types.FormatRules{
		FontRules:      []types.FontRule{},
		ParagraphRules: []types.ParagraphRule{},
		TableRules:     []types.TableRule{},
		PageRules:      []types.PageRule{},
	}

	// 提取字体规则
	formatRules.FontRules = tm.extractFontRules(doc)

	// 提取段落规则
	formatRules.ParagraphRules = tm.extractParagraphRules(doc)

	// 提取表格规则
	formatRules.TableRules = tm.extractTableRules(doc)

	// 提取页面规则
	formatRules.PageRules = tm.extractPageRules(doc)

	return formatRules
}

// extractFontRules 提取字体规则
func (tm *TemplateManager) extractFontRules(doc *types.Document) []types.FontRule {
	var fontRules []types.FontRule
	fontMap := make(map[string]bool)

	// 从段落中提取字体信息
	for _, paragraph := range doc.Content.Paragraphs {
		for _, run := range paragraph.Runs {
			fontKey := fmt.Sprintf("%s_%.1f", run.Font.Name, run.Font.Size)
			if !fontMap[fontKey] {
				fontRule := types.FontRule{
					ID:    fmt.Sprintf("font_%d", len(fontRules)+1),
					Name:  run.Font.Name,
					Size:  run.Font.Size,
					Color: run.Font.Color,
				}
				fontRules = append(fontRules, fontRule)
				fontMap[fontKey] = true
			}
		}
	}

	// 从样式中提取字体信息（CharacterStyle没有Font字段，跳过）
	// 如果需要从样式中提取字体信息，需要额外的样式解析逻辑

	return fontRules
}

// extractParagraphRules 提取段落规则
func (tm *TemplateManager) extractParagraphRules(doc *types.Document) []types.ParagraphRule {
	var paragraphRules []types.ParagraphRule
	paragraphMap := make(map[string]bool)

	// 从段落中提取段落格式信息
	for _, paragraph := range doc.Content.Paragraphs {
		paraKey := fmt.Sprintf("%s_%.1f_%.1f_%.1f", paragraph.Alignment, paragraph.Spacing.Before, paragraph.Spacing.After, paragraph.Spacing.Line)
		if !paragraphMap[paraKey] {
			paragraphRule := types.ParagraphRule{
				ID:        fmt.Sprintf("paragraph_%d", len(paragraphRules)+1),
				Name:      fmt.Sprintf("Paragraph Style %d", len(paragraphRules)+1),
				Alignment: paragraph.Alignment,
				Spacing:   paragraph.Spacing,
			}
			paragraphRules = append(paragraphRules, paragraphRule)
			paragraphMap[paraKey] = true
		}
	}

	// 从样式中提取段落格式信息（ParagraphStyle没有Format字段，跳过）
	// 如果需要从样式中提取段落格式信息，需要额外的样式解析逻辑

	return paragraphRules
}

// extractTableRules 提取表格规则
func (tm *TemplateManager) extractTableRules(doc *types.Document) []types.TableRule {
	var tableRules []types.TableRule

	// 从表格中提取表格格式信息
	for _, table := range doc.Content.Tables {
		tableRule := types.TableRule{
			ID:        fmt.Sprintf("table_%d", len(tableRules)+1),
			Name:      fmt.Sprintf("Table Style %d", len(tableRules)+1),
			Width:     table.Width,
			Alignment: table.Alignment,
		}
		tableRules = append(tableRules, tableRule)
	}

	return tableRules
}

// extractPageRules 提取页面规则
func (tm *TemplateManager) extractPageRules(doc *types.Document) []types.PageRule {
	var pageRules []types.PageRule

	// 从文档的节中提取页面信息
	for _, section := range doc.Content.Sections {
		pageRule := types.PageRule{
			ID:             fmt.Sprintf("page_%d", len(pageRules)+1),
			Name:           fmt.Sprintf("Page Style %d", len(pageRules)+1),
			PageSize:       section.PageSize,
			PageMargins:    section.PageMargins,
			HeaderDistance: section.HeaderDistance,
			FooterDistance: section.FooterDistance,
			Columns:        section.Columns,
			PageNumbering:  section.PageNumbering,
			LineNumbering:  section.LineNumbering,
		}
		pageRules = append(pageRules, pageRule)
	}

	// 如果没有节，创建默认页面规则
	if len(pageRules) == 0 {
		pageRule := types.PageRule{
			ID:   "page_default",
			Name: "Default Page",
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
		pageRules = append(pageRules, pageRule)
	}

	return pageRules
}

// createTemplateMetadata 创建模板元数据
func (tm *TemplateManager) createTemplateMetadata(templatePath string, doc *types.Document) TemplateMetadata {
	now := time.Now().Format("2006-01-02")

	metadata := TemplateMetadata{
		Author:      "System",
		Created:     now,
		LastUpdated: now,
		Category:    "Word Template",
		Tags:        []string{"word", "template", "document"},
	}

	// 如果文档有元数据信息，使用文档的元数据
	if doc.Metadata.Author != "" {
		metadata.Author = doc.Metadata.Author
	}
	if !doc.Metadata.Created.IsZero() {
		metadata.Created = doc.Metadata.Created.Format("2006-01-02")
	}
	if !doc.Metadata.Modified.IsZero() {
		metadata.LastUpdated = doc.Metadata.Modified.Format("2006-01-02")
	}

	return metadata
}

// LoadTemplatesFromDirectory 从目录加载所有Word文档模板
func (tm *TemplateManager) LoadTemplatesFromDirectory(dirPath string) error {
	entries, err := os.ReadDir(dirPath)
	if err != nil {
		return fmt.Errorf("failed to read directory: %w", err)
	}

	for _, entry := range entries {
		if entry.IsDir() {
			continue
		}

		// 只加载Word文档格式
		ext := strings.ToLower(filepath.Ext(entry.Name()))
		if !tm.isSupportedWordFormat(ext) {
			continue
		}

		templatePath := filepath.Join(dirPath, entry.Name())
		if _, err := tm.LoadTemplate(templatePath); err != nil {
			return fmt.Errorf("failed to load template %s: %w", entry.Name(), err)
		}
	}

	return nil
}

// GetTemplate 根据ID获取模板
func (tm *TemplateManager) GetTemplate(templateID string) (*Template, error) {
	template, exists := tm.templates[templateID]
	if !exists {
		return nil, fmt.Errorf("template not found: %s", templateID)
	}
	return template, nil
}

// ListTemplates 列出所有可用模板
func (tm *TemplateManager) ListTemplates() []*Template {
	templates := make([]*Template, 0, len(tm.templates))
	for _, template := range tm.templates {
		templates = append(templates, template)
	}
	return templates
}

// validateTemplate 验证模板格式
func (tm *TemplateManager) validateTemplate(template *Template) error {
	if template.ID == "" {
		return fmt.Errorf("template ID is required")
	}
	if template.Name == "" {
		return fmt.Errorf("template name is required")
	}
	if template.Version == "" {
		return fmt.Errorf("template version is required")
	}
	if template.SourcePath == "" {
		return fmt.Errorf("template source path is required")
	}

	// 验证格式规则
	if err := tm.validateFormatRules(&template.FormatRules); err != nil {
		return fmt.Errorf("format rules validation failed: %w", err)
	}

	return nil
}

// ValidateTemplate 验证模板（公共方法）
func (tm *TemplateManager) ValidateTemplate(template *Template) error {
	return tm.validateTemplate(template)
}

// validateFormatRules 验证格式规则
func (tm *TemplateManager) validateFormatRules(rules *types.FormatRules) error {
	// 验证字体规则
	for i, rule := range rules.FontRules {
		if rule.ID == "" {
			return fmt.Errorf("font rule %d: ID is required", i)
		}
		if rule.Name == "" {
			return fmt.Errorf("font rule %d: name is required", i)
		}
		if rule.Size <= 0 {
			return fmt.Errorf("font rule %d: size must be positive", i)
		}
	}

	// 验证段落规则
	for i, rule := range rules.ParagraphRules {
		if rule.ID == "" {
			return fmt.Errorf("paragraph rule %d: ID is required", i)
		}
		if rule.Name == "" {
			return fmt.Errorf("paragraph rule %d: name is required", i)
		}
	}

	// 验证表格规则
	for i, rule := range rules.TableRules {
		if rule.ID == "" {
			return fmt.Errorf("table rule %d: ID is required", i)
		}
		if rule.Name == "" {
			return fmt.Errorf("table rule %d: name is required", i)
		}
		if rule.Width <= 0 {
			return fmt.Errorf("table rule %d: width must be positive", i)
		}
	}

	// 验证页面规则
	for i, rule := range rules.PageRules {
		if rule.ID == "" {
			return fmt.Errorf("page rule %d: ID is required", i)
		}
		if rule.Name == "" {
			return fmt.Errorf("page rule %d: name is required", i)
		}
		if rule.PageSize.Width <= 0 || rule.PageSize.Height <= 0 {
			return fmt.Errorf("page rule %d: page size must be positive", i)
		}
	}

	return nil
}

// ValidateWordTemplate 验证Word模板
func (tm *TemplateManager) ValidateWordTemplate(templatePath string) error {
	// 检查文件是否存在
	if templatePath == "" {
		return fmt.Errorf("template path is empty")
	}

	// 检查文件扩展名
	ext := strings.ToLower(filepath.Ext(templatePath))
	if !tm.isSupportedWordFormat(ext) {
		return fmt.Errorf("unsupported template format: %s, only Word documents (.docx, .doc, .dot, .dotx) are supported", ext)
	}

	// 尝试解析文档
	_, err := tm.LoadTemplate(templatePath)
	if err != nil {
		return fmt.Errorf("invalid Word template: %w", err)
	}

	return nil
}
