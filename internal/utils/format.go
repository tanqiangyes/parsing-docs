package utils

import (
	"fmt"
	"math"
	"strings"

	"docs-parser/internal/core/types"
)

// FormatComparison 格式比较结果
type FormatComparison struct {
	IsIdentical bool                    `json:"is_identical"`
	Score       float64                 `json:"score"`
	Issues      []FormatIssue           `json:"issues"`
	Details     FormatComparisonDetails `json:"details"`
}

// FormatIssue 格式问题
type FormatIssue struct {
	Type        string `json:"type"`
	Severity    string `json:"severity"`
	Description string `json:"description"`
	Location    string `json:"location"`
	Suggestion  string `json:"suggestion"`
}

// FormatComparisonDetails 格式比较详情
type FormatComparisonDetails struct {
	FontIssues      []FormatIssue `json:"font_issues"`
	ParagraphIssues []FormatIssue `json:"paragraph_issues"`
	TableIssues     []FormatIssue `json:"table_issues"`
	PageIssues      []FormatIssue `json:"page_issues"`
}

// CompareFormatRules 比较格式规则
func CompareFormatRules(docRules, templateRules *types.FormatRules) *FormatComparison {
	comparison := &FormatComparison{
		Issues:  []FormatIssue{},
		Details: FormatComparisonDetails{},
	}

	// 比较字体规则
	fontIssues := compareFontRules(docRules.FontRules, templateRules.FontRules)
	comparison.Details.FontIssues = fontIssues
	comparison.Issues = append(comparison.Issues, fontIssues...)

	// 比较段落规则
	paragraphIssues := compareParagraphRules(docRules.ParagraphRules, templateRules.ParagraphRules)
	comparison.Details.ParagraphIssues = paragraphIssues
	comparison.Issues = append(comparison.Issues, paragraphIssues...)

	// 比较表格规则
	tableIssues := compareTableRules(docRules.TableRules, templateRules.TableRules)
	comparison.Details.TableIssues = tableIssues
	comparison.Issues = append(comparison.Issues, tableIssues...)

	// 比较页面规则
	pageIssues := comparePageRules(docRules.PageRules, templateRules.PageRules)
	comparison.Details.PageIssues = pageIssues
	comparison.Issues = append(comparison.Issues, pageIssues...)

	// 计算相似度分数
	comparison.Score = calculateSimilarityScore(comparison.Issues)
	comparison.IsIdentical = len(comparison.Issues) == 0

	return comparison
}

// compareFontRules 比较字体规则
func compareFontRules(docFonts, templateFonts []types.FontRule) []FormatIssue {
	var issues []FormatIssue

	// 检查字体名称
	for _, templateFont := range templateFonts {
		found := false
		for _, docFont := range docFonts {
			if strings.EqualFold(docFont.Name, templateFont.Name) {
				found = true
				// 检查字体大小
				if !isFloatEqual(docFont.Size, templateFont.Size, 0.1) {
					issues = append(issues, FormatIssue{
						Type:        "font_size",
						Severity:    "warning",
						Description: fmt.Sprintf("字体大小不匹配: 期望 %.1f, 实际 %.1f", templateFont.Size, docFont.Size),
						Location:    fmt.Sprintf("字体: %s", docFont.Name),
						Suggestion:  fmt.Sprintf("将字体大小调整为 %.1f", templateFont.Size),
					})
				}
				break
			}
		}
		if !found {
			issues = append(issues, FormatIssue{
				Type:        "font_missing",
				Severity:    "error",
				Description: fmt.Sprintf("缺少必需的字体: %s", templateFont.Name),
				Location:    "文档全局",
				Suggestion:  fmt.Sprintf("添加字体: %s", templateFont.Name),
			})
		}
	}

	return issues
}

// compareParagraphRules 比较段落规则
func compareParagraphRules(docParagraphs, templateParagraphs []types.ParagraphRule) []FormatIssue {
	var issues []FormatIssue

	for _, templatePara := range templateParagraphs {
		found := false
		for _, docPara := range docParagraphs {
			if strings.EqualFold(docPara.Name, templatePara.Name) {
				found = true
				// 检查对齐方式
				if docPara.Alignment != templatePara.Alignment {
					issues = append(issues, FormatIssue{
						Type:        "paragraph_alignment",
						Severity:    "warning",
						Description: fmt.Sprintf("段落对齐方式不匹配: 期望 %s, 实际 %s", templatePara.Alignment, docPara.Alignment),
						Location:    fmt.Sprintf("段落样式: %s", docPara.Name),
						Suggestion:  fmt.Sprintf("将对齐方式调整为 %s", templatePara.Alignment),
					})
				}
				break
			}
		}
		if !found {
			issues = append(issues, FormatIssue{
				Type:        "paragraph_missing",
				Severity:    "error",
				Description: fmt.Sprintf("缺少必需的段落样式: %s", templatePara.Name),
				Location:    "文档全局",
				Suggestion:  fmt.Sprintf("添加段落样式: %s", templatePara.Name),
			})
		}
	}

	return issues
}

// compareTableRules 比较表格规则
func compareTableRules(docTables, templateTables []types.TableRule) []FormatIssue {
	var issues []FormatIssue

	for _, templateTable := range templateTables {
		found := false
		for _, docTable := range docTables {
			if strings.EqualFold(docTable.Name, templateTable.Name) {
				found = true
				// 检查表格宽度
				if !isFloatEqual(docTable.Width, templateTable.Width, 1.0) {
					issues = append(issues, FormatIssue{
						Type:        "table_width",
						Severity:    "warning",
						Description: fmt.Sprintf("表格宽度不匹配: 期望 %.1f%%, 实际 %.1f%%", templateTable.Width, docTable.Width),
						Location:    fmt.Sprintf("表格样式: %s", docTable.Name),
						Suggestion:  fmt.Sprintf("将表格宽度调整为 %.1f%%", templateTable.Width),
					})
				}
				break
			}
		}
		if !found {
			issues = append(issues, FormatIssue{
				Type:        "table_missing",
				Severity:    "error",
				Description: fmt.Sprintf("缺少必需的表格样式: %s", templateTable.Name),
				Location:    "文档全局",
				Suggestion:  fmt.Sprintf("添加表格样式: %s", templateTable.Name),
			})
		}
	}

	return issues
}

// comparePageRules 比较页面规则
func comparePageRules(docPages, templatePages []types.PageRule) []FormatIssue {
	var issues []FormatIssue

	for _, templatePage := range templatePages {
		found := false
		for _, docPage := range docPages {
			if strings.EqualFold(docPage.Name, templatePage.Name) {
				found = true
				// 检查页面大小
				if !isFloatEqual(docPage.PageSize.Width, templatePage.PageSize.Width, 1.0) ||
					!isFloatEqual(docPage.PageSize.Height, templatePage.PageSize.Height, 1.0) {
					issues = append(issues, FormatIssue{
						Type:     "page_size",
						Severity: "warning",
						Description: fmt.Sprintf("页面大小不匹配: 期望 %.1fx%.1f, 实际 %.1fx%.1f",
							templatePage.PageSize.Width, templatePage.PageSize.Height,
							docPage.PageSize.Width, docPage.PageSize.Height),
						Location:   fmt.Sprintf("页面样式: %s", docPage.Name),
						Suggestion: fmt.Sprintf("将页面大小调整为 %.1fx%.1f", templatePage.PageSize.Width, templatePage.PageSize.Height),
					})
				}
				break
			}
		}
		if !found {
			issues = append(issues, FormatIssue{
				Type:        "page_missing",
				Severity:    "error",
				Description: fmt.Sprintf("缺少必需的页面样式: %s", templatePage.Name),
				Location:    "文档全局",
				Suggestion:  fmt.Sprintf("添加页面样式: %s", templatePage.Name),
			})
		}
	}

	return issues
}

// isFloatEqual 比较两个浮点数是否相等（允许误差）
func isFloatEqual(a, b, tolerance float64) bool {
	return math.Abs(a-b) <= tolerance
}

// calculateSimilarityScore 计算相似度分数
func calculateSimilarityScore(issues []FormatIssue) float64 {
	if len(issues) == 0 {
		return 100.0
	}

	errorCount := 0
	warningCount := 0

	for _, issue := range issues {
		switch issue.Severity {
		case "error":
			errorCount++
		case "warning":
			warningCount++
		}
	}

	// 错误权重更高
	totalPenalty := float64(errorCount)*10.0 + float64(warningCount)*2.0
	score := 100.0 - totalPenalty

	if score < 0 {
		score = 0
	}

	return score
}

// GenerateModificationSuggestions 生成修改建议
func GenerateModificationSuggestions(comparison *FormatComparison) []string {
	var suggestions []string

	for _, issue := range comparison.Issues {
		if issue.Suggestion != "" {
			suggestions = append(suggestions, issue.Suggestion)
		}
	}

	return suggestions
}
