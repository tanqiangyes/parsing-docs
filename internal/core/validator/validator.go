package validator

import (
	"docs-parser/internal/core/types"
	"docs-parser/internal/formats"
	"docs-parser/internal/utils"
	"fmt"
)

// Validator 文档验证器
type Validator struct {
	wordParser *formats.WordParser
	rules      []ValidationRule
}

// NewValidator 创建新的验证器
func NewValidator() *Validator {
	validator := &Validator{
		wordParser: formats.NewWordParser(),
	}
	validator.loadDefaultRules()
	return validator
}

// ValidateDocument 验证文档
func (v *Validator) ValidateDocument(filePath string) (*ValidationResult, error) {
	// 验证文件存在
	if err := utils.ValidateFile(filePath); err != nil {
		return nil, fmt.Errorf("file validation failed: %w", err)
	}

	// 解析文档
	doc, err := v.wordParser.ParseDocument(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to parse document: %w", err)
	}

	// 执行验证
	result := v.validateDocument(doc)

	return result, nil
}

// validateDocument 验证文档内容
func (v *Validator) validateDocument(doc *types.Document) *ValidationResult {
	result := &ValidationResult{
		ComplianceRate:  0.0,
		Issues:          []ValidationIssue{},
		Recommendations: []Recommendation{},
	}

	// 验证字体规则
	fontIssues := v.validateFontRules(doc)
	result.Issues = append(result.Issues, fontIssues...)

	// 验证段落规则
	paragraphIssues := v.validateParagraphRules(doc)
	result.Issues = append(result.Issues, paragraphIssues...)

	// 验证表格规则
	tableIssues := v.validateTableRules(doc)
	result.Issues = append(result.Issues, tableIssues...)

	// 验证页面规则
	pageIssues := v.validatePageRules(doc)
	result.Issues = append(result.Issues, pageIssues...)

	// 验证样式规则
	styleIssues := v.validateStyleRules(doc)
	result.Issues = append(result.Issues, styleIssues...)

	// 计算合规率
	result.ComplianceRate = v.calculateComplianceRate(result.Issues)

	// 生成建议
	result.Recommendations = v.generateRecommendations(result.Issues)

	return result
}

// validateFontRules 验证字体规则
func (v *Validator) validateFontRules(doc *types.Document) []ValidationIssue {
	var issues []ValidationIssue

	for _, paragraph := range doc.Content.Paragraphs {
		for _, run := range paragraph.Runs {
			// 验证字体大小
			if run.Font.Size < 10.0 {
				issues = append(issues, ValidationIssue{
					ID:          "font_size_minimum",
					Type:        "font",
					Severity:    "medium",
					Location:    run.ID,
					Description: "字体大小不符合最小要求",
					Current:     run.Font.Size,
					Expected:    ">= 10.0",
					Rule:        "font_size_minimum",
				})
			}

			// 验证字体名称
			if run.Font.Name == "" {
				issues = append(issues, ValidationIssue{
					ID:          "font_name_required",
					Type:        "font",
					Severity:    "high",
					Location:    run.ID,
					Description: "字体名称未设置",
					Current:     run.Font.Name,
					Expected:    "标准字体名称",
					Rule:        "font_name_required",
				})
			}

			// 验证字体颜色
			if run.Font.Color.RGB == "" {
				issues = append(issues, ValidationIssue{
					ID:          "font_color_required",
					Type:        "font",
					Severity:    "low",
					Location:    run.ID,
					Description: "字体颜色未设置",
					Current:     run.Font.Color,
					Expected:    "有效的颜色值",
					Rule:        "font_color_required",
				})
			}

			// 验证字体粗细
			if run.Font.Bold && run.Font.Size < 12.0 {
				issues = append(issues, ValidationIssue{
					ID:          "bold_font_size",
					Type:        "font",
					Severity:    "low",
					Location:    run.ID,
					Description: "粗体字体大小过小",
					Current:     run.Font.Size,
					Expected:    ">= 12.0",
					Rule:        "bold_font_size_minimum",
				})
			}
		}
	}

	return issues
}

// validateParagraphRules 验证段落规则
func (v *Validator) validateParagraphRules(doc *types.Document) []ValidationIssue {
	var issues []ValidationIssue

	for _, paragraph := range doc.Content.Paragraphs {
		// 验证段落对齐方式
		if paragraph.Alignment == "" {
			issues = append(issues, ValidationIssue{
				ID:          "paragraph_alignment_required",
				Type:        "paragraph",
				Severity:    "medium",
				Location:    paragraph.ID,
				Description: "段落对齐方式未设置",
				Current:     paragraph.Alignment,
				Expected:    "left/center/right/justify",
				Rule:        "paragraph_alignment_required",
			})
		}

		// 验证段落间距
		if paragraph.Spacing.Before < 0 || paragraph.Spacing.After < 0 {
			issues = append(issues, ValidationIssue{
				ID:          "paragraph_spacing_valid",
				Type:        "paragraph",
				Severity:    "low",
				Location:    paragraph.ID,
				Description: "段落间距设置不当",
				Current:     paragraph.Spacing,
				Expected:    ">= 0",
				Rule:        "paragraph_spacing_valid",
			})
		}

		// 验证段落缩进
		if paragraph.Indentation.Left < 0 || paragraph.Indentation.Right < 0 {
			issues = append(issues, ValidationIssue{
				ID:          "paragraph_indentation_valid",
				Type:        "paragraph",
				Severity:    "low",
				Location:    paragraph.ID,
				Description: "段落缩进设置不当",
				Current:     paragraph.Indentation,
				Expected:    ">= 0",
				Rule:        "paragraph_indentation_valid",
			})
		}

		// 验证段落行距
		if paragraph.Spacing.Line < 1.0 {
			issues = append(issues, ValidationIssue{
				ID:          "paragraph_line_spacing",
				Type:        "paragraph",
				Severity:    "low",
				Location:    paragraph.ID,
				Description: "段落行距设置不当",
				Current:     paragraph.Spacing.Line,
				Expected:    ">= 1.0",
				Rule:        "paragraph_line_spacing_minimum",
			})
		}
	}

	return issues
}

// validateTableRules 验证表格规则
func (v *Validator) validateTableRules(doc *types.Document) []ValidationIssue {
	var issues []ValidationIssue

	for _, table := range doc.Content.Tables {
		// 验证表格边框
		if table.Borders.Top.Style == "" {
			issues = append(issues, ValidationIssue{
				ID:          "table_border_required",
				Type:        "table",
				Severity:    "medium",
				Location:    table.ID,
				Description: "表格边框未设置",
				Current:     table.Borders,
				Expected:    "完整的边框设置",
				Rule:        "table_border_required",
			})
		}

		// 验证表格宽度
		if table.Width <= 0 {
			issues = append(issues, ValidationIssue{
				ID:          "table_width_required",
				Type:        "table",
				Severity:    "low",
				Location:    table.ID,
				Description: "表格宽度未设置",
				Current:     table.Width,
				Expected:    "> 0",
				Rule:        "table_width_required",
			})
		}

		// 验证表格行
		if len(table.Rows) == 0 {
			issues = append(issues, ValidationIssue{
				ID:          "table_rows_required",
				Type:        "table",
				Severity:    "high",
				Location:    table.ID,
				Description: "表格没有行",
				Current:     len(table.Rows),
				Expected:    "> 0",
				Rule:        "table_rows_required",
			})
		}

		// 验证表格单元格
		for _, row := range table.Rows {
			if len(row.Cells) == 0 {
				issues = append(issues, ValidationIssue{
					ID:          "table_cells_required",
					Type:        "table",
					Severity:    "high",
					Location:    fmt.Sprintf("%s_row_%s", table.ID, row.ID),
					Description: "表格行没有单元格",
					Current:     len(row.Cells),
					Expected:    "> 0",
					Rule:        "table_cells_required",
				})
			}
		}
	}

	return issues
}

// validatePageRules 验证页面规则
func (v *Validator) validatePageRules(doc *types.Document) []ValidationIssue {
	var issues []ValidationIssue

	for _, section := range doc.Content.Sections {
		// 验证页面边距
		if section.PageMargins.Top < 0 || section.PageMargins.Bottom < 0 ||
			section.PageMargins.Left < 0 || section.PageMargins.Right < 0 {
			issues = append(issues, ValidationIssue{
				ID:          "page_margins_valid",
				Type:        "page",
				Severity:    "medium",
				Location:    section.ID,
				Description: "页面边距设置不当",
				Current:     section.PageMargins,
				Expected:    "所有边距 >= 0",
				Rule:        "page_margins_valid",
			})
		}

		// 验证页面大小
		if section.PageSize.Width <= 0 || section.PageSize.Height <= 0 {
			issues = append(issues, ValidationIssue{
				ID:          "page_size_valid",
				Type:        "page",
				Severity:    "high",
				Location:    section.ID,
				Description: "页面大小设置不当",
				Current:     section.PageSize,
				Expected:    "宽度和高度 > 0",
				Rule:        "page_size_valid",
			})
		}

		// 验证页面边距
		if section.PageMargins.Top < 0 || section.PageMargins.Bottom < 0 ||
			section.PageMargins.Left < 0 || section.PageMargins.Right < 0 {
			issues = append(issues, ValidationIssue{
				ID:          "page_margins_valid",
				Type:        "page",
				Severity:    "medium",
				Location:    section.ID,
				Description: "页面边距设置不当",
				Current:     section.PageMargins,
				Expected:    "所有边距 >= 0",
				Rule:        "page_margins_valid",
			})
		}

		// 验证页面大小
		if section.PageSize.Width <= 0 || section.PageSize.Height <= 0 {
			issues = append(issues, ValidationIssue{
				ID:          "page_size_valid",
				Type:        "page",
				Severity:    "high",
				Location:    section.ID,
				Description: "页面大小设置不当",
				Current:     section.PageSize,
				Expected:    "宽度和高度 > 0",
				Rule:        "page_size_valid",
			})
		}
	}

	return issues
}

// validateStyleRules 验证样式规则
func (v *Validator) validateStyleRules(doc *types.Document) []ValidationIssue {
	var issues []ValidationIssue

	// 验证段落样式
	for _, style := range doc.Styles.ParagraphStyles {
		if style.Name == "" {
			issues = append(issues, ValidationIssue{
				ID:          "style_name_required",
				Type:        "style",
				Severity:    "medium",
				Location:    style.ID,
				Description: "样式名称未设置",
				Current:     style.Name,
				Expected:    "有效的样式名称",
				Rule:        "style_name_required",
			})
		}
	}

	// 验证字符样式
	for _, style := range doc.Styles.CharacterStyles {
		if style.Name == "" {
			issues = append(issues, ValidationIssue{
				ID:          "character_style_name_required",
				Type:        "style",
				Severity:    "medium",
				Location:    style.ID,
				Description: "字符样式名称未设置",
				Current:     style.Name,
				Expected:    "有效的样式名称",
				Rule:        "character_style_name_required",
			})
		}
	}

	return issues
}

// calculateComplianceRate 计算合规率
func (v *Validator) calculateComplianceRate(issues []ValidationIssue) float64 {
	if len(issues) == 0 {
		return 100.0
	}

	// 根据严重程度计算权重
	totalWeight := 0.0
	issueWeight := 0.0

	for _, issue := range issues {
		weight := v.getSeverityWeight(issue.Severity)
		totalWeight += weight
		issueWeight += weight
	}

	if totalWeight == 0 {
		return 100.0
	}

	complianceRate := (totalWeight - issueWeight) / totalWeight * 100.0
	if complianceRate < 0 {
		complianceRate = 0.0
	}

	return complianceRate
}

// getSeverityWeight 获取严重程度权重
func (v *Validator) getSeverityWeight(severity string) float64 {
	switch severity {
	case "critical":
		return 4.0
	case "high":
		return 3.0
	case "medium":
		return 2.0
	case "low":
		return 1.0
	default:
		return 1.0
	}
}

// generateRecommendations 生成建议
func (v *Validator) generateRecommendations(issues []ValidationIssue) []Recommendation {
	var recommendations []Recommendation

	// 按类型分组问题
	fontIssues := v.filterIssuesByType(issues, "font")
	paragraphIssues := v.filterIssuesByType(issues, "paragraph")
	tableIssues := v.filterIssuesByType(issues, "table")
	pageIssues := v.filterIssuesByType(issues, "page")
	styleIssues := v.filterIssuesByType(issues, "style")

	// 生成字体建议
	if len(fontIssues) > 0 {
		recommendations = append(recommendations, Recommendation{
			ID:          "font_improvements",
			Type:        "font",
			Priority:    v.getPriorityByIssues(fontIssues),
			Description: "字体格式需要改进",
			Actions:     v.generateFontActions(fontIssues),
			Impact:      "提高文档可读性和专业性",
		})
	}

	// 生成段落建议
	if len(paragraphIssues) > 0 {
		recommendations = append(recommendations, Recommendation{
			ID:          "paragraph_improvements",
			Type:        "paragraph",
			Priority:    v.getPriorityByIssues(paragraphIssues),
			Description: "段落格式需要改进",
			Actions:     v.generateParagraphActions(paragraphIssues),
			Impact:      "提高文档结构和可读性",
		})
	}

	// 生成表格建议
	if len(tableIssues) > 0 {
		recommendations = append(recommendations, Recommendation{
			ID:          "table_improvements",
			Type:        "table",
			Priority:    v.getPriorityByIssues(tableIssues),
			Description: "表格格式需要改进",
			Actions:     v.generateTableActions(tableIssues),
			Impact:      "提高表格的可读性和专业性",
		})
	}

	// 生成页面建议
	if len(pageIssues) > 0 {
		recommendations = append(recommendations, Recommendation{
			ID:          "page_improvements",
			Type:        "page",
			Priority:    v.getPriorityByIssues(pageIssues),
			Description: "页面设置需要改进",
			Actions:     v.generatePageActions(pageIssues),
			Impact:      "确保文档打印和显示效果",
		})
	}

	// 生成样式建议
	if len(styleIssues) > 0 {
		recommendations = append(recommendations, Recommendation{
			ID:          "style_improvements",
			Type:        "style",
			Priority:    v.getPriorityByIssues(styleIssues),
			Description: "样式设置需要改进",
			Actions:     v.generateStyleActions(styleIssues),
			Impact:      "提高文档样式一致性",
		})
	}

	return recommendations
}

// filterIssuesByType 按类型过滤问题
func (v *Validator) filterIssuesByType(issues []ValidationIssue, issueType string) []ValidationIssue {
	var filtered []ValidationIssue
	for _, issue := range issues {
		if issue.Type == issueType {
			filtered = append(filtered, issue)
		}
	}
	return filtered
}

// getPriorityByIssues 根据问题确定优先级
func (v *Validator) getPriorityByIssues(issues []ValidationIssue) string {
	hasCritical := false
	hasHigh := false

	for _, issue := range issues {
		if issue.Severity == "critical" {
			hasCritical = true
		}
		if issue.Severity == "high" {
			hasHigh = true
		}
	}

	if hasCritical {
		return "urgent"
	}
	if hasHigh {
		return "high"
	}
	return "medium"
}

// generateFontActions 生成字体操作建议
func (v *Validator) generateFontActions(issues []ValidationIssue) []Action {
	var actions []Action

	// 字体大小建议
	if v.hasIssue(issues, "font_size_minimum") {
		actions = append(actions, Action{
			Type:        "font_size",
			Description: "调整字体大小",
			Steps: []Step{
				{Order: 1, Description: "选择需要调整的文本", Details: "选中字体大小过小的文本"},
				{Order: 2, Description: "设置字体大小", Details: "将字体大小调整为10.0或更大"},
			},
		})
	}

	// 字体名称建议
	if v.hasIssue(issues, "font_name_required") {
		actions = append(actions, Action{
			Type:        "font_name",
			Description: "设置字体名称",
			Steps: []Step{
				{Order: 1, Description: "选择需要调整的文本", Details: "选中未设置字体的文本"},
				{Order: 2, Description: "设置标准字体", Details: "选择宋体、黑体等标准字体"},
			},
		})
	}

	// 粗体字体大小建议
	if v.hasIssue(issues, "bold_font_size_minimum") {
		actions = append(actions, Action{
			Type:        "bold_font_size",
			Description: "调整粗体字体大小",
			Steps: []Step{
				{Order: 1, Description: "选择粗体文本", Details: "选中粗体字体大小过小的文本"},
				{Order: 2, Description: "调整字体大小", Details: "将粗体字体大小调整为12.0或更大"},
			},
		})
	}

	return actions
}

// generateParagraphActions 生成段落操作建议
func (v *Validator) generateParagraphActions(issues []ValidationIssue) []Action {
	var actions []Action

	// 段落对齐建议
	if v.hasIssue(issues, "paragraph_alignment_required") {
		actions = append(actions, Action{
			Type:        "paragraph_alignment",
			Description: "设置段落对齐方式",
			Steps: []Step{
				{Order: 1, Description: "选择段落", Details: "选中需要设置对齐方式的段落"},
				{Order: 2, Description: "设置对齐", Details: "选择左对齐、居中、右对齐或两端对齐"},
			},
		})
	}

	// 段落间距建议
	if v.hasIssue(issues, "paragraph_spacing_valid") {
		actions = append(actions, Action{
			Type:        "paragraph_spacing",
			Description: "调整段落间距",
			Steps: []Step{
				{Order: 1, Description: "选择段落", Details: "选中需要调整间距的段落"},
				{Order: 2, Description: "设置间距", Details: "将段落间距调整为正值"},
			},
		})
	}

	// 行距建议
	if v.hasIssue(issues, "paragraph_line_spacing_minimum") {
		actions = append(actions, Action{
			Type:        "paragraph_line_spacing",
			Description: "调整段落行距",
			Steps: []Step{
				{Order: 1, Description: "选择段落", Details: "选中需要调整行距的段落"},
				{Order: 2, Description: "设置行距", Details: "将行距调整为1.0或更大"},
			},
		})
	}

	return actions
}

// generateTableActions 生成表格操作建议
func (v *Validator) generateTableActions(issues []ValidationIssue) []Action {
	var actions []Action

	// 表格边框建议
	if v.hasIssue(issues, "table_border_required") {
		actions = append(actions, Action{
			Type:        "table_border",
			Description: "设置表格边框",
			Steps: []Step{
				{Order: 1, Description: "选择表格", Details: "选中需要设置边框的表格"},
				{Order: 2, Description: "设置边框", Details: "为表格添加完整的边框线"},
			},
		})
	}

	// 表格宽度建议
	if v.hasIssue(issues, "table_width_required") {
		actions = append(actions, Action{
			Type:        "table_width",
			Description: "设置表格宽度",
			Steps: []Step{
				{Order: 1, Description: "选择表格", Details: "选中需要设置宽度的表格"},
				{Order: 2, Description: "设置宽度", Details: "为表格设置合适的宽度"},
			},
		})
	}

	return actions
}

// generatePageActions 生成页面操作建议
func (v *Validator) generatePageActions(issues []ValidationIssue) []Action {
	var actions []Action

	// 页面边距建议
	if v.hasIssue(issues, "page_margins_valid") {
		actions = append(actions, Action{
			Type:        "page_margins",
			Description: "调整页面边距",
			Steps: []Step{
				{Order: 1, Description: "打开页面设置", Details: "进入页面布局设置"},
				{Order: 2, Description: "设置边距", Details: "将页面边距调整为正值"},
			},
		})
	}

	// 页面大小建议
	if v.hasIssue(issues, "page_size_valid") {
		actions = append(actions, Action{
			Type:        "page_size",
			Description: "设置页面大小",
			Steps: []Step{
				{Order: 1, Description: "打开页面设置", Details: "进入页面布局设置"},
				{Order: 2, Description: "设置页面大小", Details: "选择A4、Letter等标准页面大小"},
			},
		})
	}

	return actions
}

// generateStyleActions 生成样式操作建议
func (v *Validator) generateStyleActions(issues []ValidationIssue) []Action {
	var actions []Action

	// 样式名称建议
	if v.hasIssue(issues, "style_name_required") || v.hasIssue(issues, "character_style_name_required") {
		actions = append(actions, Action{
			Type:        "style_name",
			Description: "设置样式名称",
			Steps: []Step{
				{Order: 1, Description: "选择样式", Details: "选中未设置名称的样式"},
				{Order: 2, Description: "设置名称", Details: "为样式设置有效的名称"},
			},
		})
	}

	return actions
}

// hasIssue 检查是否有特定问题
func (v *Validator) hasIssue(issues []ValidationIssue, ruleID string) bool {
	for _, issue := range issues {
		if issue.Rule == ruleID {
			return true
		}
	}
	return false
}

// loadDefaultRules 加载默认规则
func (v *Validator) loadDefaultRules() {
	v.rules = []ValidationRule{
		{
			ID:          "font_size_minimum",
			Name:        "字体大小最小值",
			Type:        "font",
			Description: "字体大小必须大于等于10.0",
			Severity:    "medium",
			Enabled:     true,
		},
		{
			ID:          "font_name_required",
			Name:        "字体名称必需",
			Type:        "font",
			Description: "字体必须设置名称",
			Severity:    "high",
			Enabled:     true,
		},
		{
			ID:          "paragraph_alignment_required",
			Name:        "段落对齐必需",
			Type:        "paragraph",
			Description: "段落必须设置对齐方式",
			Severity:    "medium",
			Enabled:     true,
		},
		{
			ID:          "table_border_required",
			Name:        "表格边框必需",
			Type:        "table",
			Description: "表格必须设置边框",
			Severity:    "medium",
			Enabled:     true,
		},
		{
			ID:          "page_margins_valid",
			Name:        "页面边距有效",
			Type:        "page",
			Description: "页面边距必须为正值",
			Severity:    "medium",
			Enabled:     true,
		},
	}
}

// ValidationResult 验证结果
type ValidationResult struct {
	ComplianceRate  float64           `json:"compliance_rate"`
	Issues          []ValidationIssue `json:"issues"`
	Recommendations []Recommendation  `json:"recommendations"`
}

// ValidationIssue 验证问题
type ValidationIssue struct {
	ID          string      `json:"id"`
	Type        string      `json:"type"`
	Severity    string      `json:"severity"`
	Location    string      `json:"location"`
	Description string      `json:"description"`
	Current     interface{} `json:"current"`
	Expected    interface{} `json:"expected"`
	Rule        string      `json:"rule"`
}

// Recommendation 建议
type Recommendation struct {
	ID          string   `json:"id"`
	Type        string   `json:"type"`
	Priority    string   `json:"priority"`
	Description string   `json:"description"`
	Actions     []Action `json:"actions"`
	Impact      string   `json:"impact"`
}

// Action 操作
type Action struct {
	Type        string `json:"type"`
	Description string `json:"description"`
	Steps       []Step `json:"steps"`
}

// Step 步骤
type Step struct {
	Order       int    `json:"order"`
	Description string `json:"description"`
	Details     string `json:"details"`
}

// ValidationRule 验证规则
type ValidationRule struct {
	ID          string `json:"id"`
	Name        string `json:"name"`
	Type        string `json:"type"`
	Description string `json:"description"`
	Severity    string `json:"severity"`
	Enabled     bool   `json:"enabled"`
}
