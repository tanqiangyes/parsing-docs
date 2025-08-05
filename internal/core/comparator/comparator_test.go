package comparator

import (
	"testing"
	"docs-parser/internal/core/types"
)

// TestNewDocumentComparator 测试创建新的文档比较器
func TestNewDocumentComparator(t *testing.T) {
	comparator := NewDocumentComparator()
	if comparator == nil {
		t.Error("期望返回非空的文档比较器")
	}
}

// TestDocumentComparator_CompareDocuments 测试文档比较功能
func TestDocumentComparator_CompareDocuments(t *testing.T) {
	comparator := NewDocumentComparator()

	// 测试比较不存在的文档
	_, err := comparator.CompareDocuments("nonexistent1.docx", "nonexistent2.docx")
	if err == nil {
		t.Error("期望比较不存在的文档返回错误")
	}
}

// TestDocumentComparator_CompareWithTemplate 测试与模板比较功能
func TestDocumentComparator_CompareWithTemplate(t *testing.T) {
	comparator := NewDocumentComparator()

	// 测试与不存在的模板比较
	_, err := comparator.CompareWithTemplate("nonexistent.docx", "nonexistent_template.docx")
	if err == nil {
		t.Error("期望与不存在的模板比较返回错误")
	}
}

// TestDocumentComparator_CompareFormatRules 测试格式规则比较功能
func TestDocumentComparator_CompareFormatRules(t *testing.T) {
	comparator := NewDocumentComparator()

	// 创建测试用的格式规则
	docRules := &types.FormatRules{
		FontRules: []types.FontRule{
			{
				ID:       "font1",
				Name:     "宋体",
				Size:     11.0,
				Color:    types.Color{RGB: "000000"},
				Bold:     false,
				Italic:   false,
			},
		},
		ParagraphRules: []types.ParagraphRule{
			{
				ID:        "para1",
				Alignment: "left",
				Indentation: types.Indentation{
					Left:   0,
					Right:  0,
					First:  0,
				},
				Spacing: types.Spacing{
					Before: 0,
					After:  0,
					Line:   1.0,
				},
			},
		},
	}

	templateRules := &types.FormatRules{
		FontRules: []types.FontRule{
			{
				ID:       "font1",
				Name:     "黑体",
				Size:     12.0,
				Color:    types.Color{RGB: "000000"},
				Bold:     true,
				Italic:   false,
			},
		},
		ParagraphRules: []types.ParagraphRule{
			{
				ID:        "para1",
				Alignment: "center",
				Indentation: types.Indentation{
					Left:   0,
					Right:  0,
					First:  0,
				},
				Spacing: types.Spacing{
					Before: 0,
					After:  0,
					Line:   1.0,
				},
			},
		},
	}

	// 测试格式规则比较
	comparison, err := comparator.CompareFormatRules(docRules, templateRules, nil, nil)
	if err != nil {
		t.Errorf("格式规则比较失败: %v", err)
	}

	if comparison == nil {
		t.Error("期望返回比较结果")
	}

	// 验证比较结果
	if len(comparison.Issues) == 0 {
		t.Error("期望发现格式差异")
	}
}

// TestDocumentComparator_CompareContent 测试内容比较功能
func TestDocumentComparator_CompareContent(t *testing.T) {
	comparator := NewDocumentComparator()

	// 创建测试用的文档内容
	docContent := &types.DocumentContent{
		Paragraphs: []types.Paragraph{
			{
				Text: "测试段落1",
				Runs: []types.TextRun{
					{
						Text: "测试文本",
						Font: types.Font{
							Name:  "宋体",
							Size:  11.0,
							Color: types.Color{RGB: "000000"},
							Bold:  false,
						},
					},
				},
			},
		},
	}

	templateContent := &types.DocumentContent{
		Paragraphs: []types.Paragraph{
			{
				Text: "测试段落1",
				Runs: []types.TextRun{
					{
						Text: "测试文本",
						Font: types.Font{
							Name:  "黑体",
							Size:  12.0,
							Color: types.Color{RGB: "000000"},
							Bold:  true,
						},
					},
				},
			},
		},
	}

	// 测试内容比较
	comparison, err := comparator.CompareContent(docContent, templateContent)
	if err != nil {
		t.Errorf("内容比较失败: %v", err)
	}

	if comparison == nil {
		t.Error("期望返回比较结果")
	}
}

// TestDocumentComparator_CompareStyles 测试样式比较功能
func TestDocumentComparator_CompareStyles(t *testing.T) {
	comparator := NewDocumentComparator()

	// 创建测试用的文档样式
	docStyles := &types.DocumentStyles{
		ParagraphStyles: []types.ParagraphStyle{
			{
				ID:   "style1",
				Name: "标题1",
				Font: types.Font{
					Name:  "宋体",
					Size:  16.0,
					Color: types.Color{RGB: "000000"},
					Bold:  false,
				},
				Alignment: "left",
			},
		},
	}

	templateStyles := &types.DocumentStyles{
		ParagraphStyles: []types.ParagraphStyle{
			{
				ID:   "style1",
				Name: "标题1",
				Font: types.Font{
					Name:  "黑体",
					Size:  18.0,
					Color: types.Color{RGB: "000000"},
					Bold:  true,
				},
				Alignment: "center",
			},
		},
	}

	// 测试样式比较
	comparison, err := comparator.CompareStyles(docStyles, templateStyles)
	if err != nil {
		t.Errorf("样式比较失败: %v", err)
	}

	if comparison == nil {
		t.Error("期望返回比较结果")
	}
}

// TestFormatIssue_Validation 测试格式问题验证
func TestFormatIssue_Validation(t *testing.T) {
	// 测试有效的格式问题
	validIssue := types.FormatIssue{
		ID:          "test_issue_1",
		Type:        "font",
		Severity:    "medium",
		Location:    "第1段第1个文本",
		Description: "测试格式问题",
		Current:     map[string]interface{}{"font": "宋体"},
		Expected:    map[string]interface{}{"font": "黑体"},
		Rule:        "font_format",
		Suggestions: []string{"调整字体格式"},
	}

	if validIssue.ID == "" {
		t.Error("格式问题应该有ID")
	}

	if validIssue.Type == "" {
		t.Error("格式问题应该有类型")
	}

	if validIssue.Severity == "" {
		t.Error("格式问题应该有严重程度")
	}

	if len(validIssue.Suggestions) == 0 {
		t.Error("格式问题应该有建议")
	}
}

// TestComparisonReport_Validation 测试比较报告验证
func TestComparisonReport_Validation(t *testing.T) {
	// 测试有效的比较报告
	validReport := &ComparisonReport{
		DocumentPath: "test.docx",
		TemplatePath: "template.docx",
		Issues: []types.FormatIssue{
			{
				ID:          "test_issue_1",
				Type:        "font",
				Severity:    "medium",
				Location:    "第1段第1个文本",
				Description: "测试格式问题",
				Current:     map[string]interface{}{"font": "宋体"},
				Expected:    map[string]interface{}{"font": "黑体"},
				Rule:        "font_format",
				Suggestions: []string{"调整字体格式"},
			},
		},
		FormatComparison: &FormatComparison{
			Issues: []types.FormatIssue{},
		},
		ContentComparison: &ContentComparison{
			Issues: []types.FormatIssue{},
		},
		StyleComparison: &StyleComparison{
			Issues: []types.FormatIssue{},
		},
	}

	if validReport.DocumentPath == "" {
		t.Error("比较报告应该有文档路径")
	}

	if validReport.TemplatePath == "" {
		t.Error("比较报告应该有模板路径")
	}

	if validReport.FormatComparison == nil {
		t.Error("比较报告应该有格式比较结果")
	}

	if validReport.ContentComparison == nil {
		t.Error("比较报告应该有内容比较结果")
	}

	if validReport.StyleComparison == nil {
		t.Error("比较报告应该有样式比较结果")
	}
} 