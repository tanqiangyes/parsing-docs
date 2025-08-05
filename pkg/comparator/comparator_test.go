package comparator

import (
	"fmt"
	"testing"
	"docs-parser/internal/core/types"
)

// TestNewComparator 测试创建新的比较器
func TestNewComparator(t *testing.T) {
	comparator := NewComparator()
	if comparator == nil {
		t.Error("期望返回非空的比较器")
	}
}

// TestComparator_CompareWithTemplate 测试与模板比较功能
func TestComparator_CompareWithTemplate(t *testing.T) {
	comparator := NewComparator()

	// 测试与不存在的模板比较
	_, err := comparator.CompareWithTemplate("nonexistent.docx", "nonexistent_template.docx")
	if err == nil {
		t.Error("期望与不存在的模板比较返回错误")
	}

	// 测试空路径
	_, err = comparator.CompareWithTemplate("", "template.docx")
	if err == nil {
		t.Error("期望空文档路径返回错误")
	}

	_, err = comparator.CompareWithTemplate("document.docx", "")
	if err == nil {
		t.Error("期望空模板路径返回错误")
	}
}

// TestComparator_CompareFormatRules 测试格式规则比较功能
func TestComparator_CompareFormatRules(t *testing.T) {
	comparator := NewComparator()

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
	comparison, err := comparator.CompareFormatRules(docRules, templateRules)
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

// TestComparator_CompareFormatRules_NilRules 测试空规则比较
func TestComparator_CompareFormatRules_NilRules(t *testing.T) {
	comparator := NewComparator()

	// 测试nil规则
	_, err := comparator.CompareFormatRules(nil, nil)
	if err == nil {
		t.Error("期望nil规则返回错误")
	}

	// 测试部分nil规则
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
	}

	_, err = comparator.CompareFormatRules(docRules, nil)
	if err == nil {
		t.Error("期望nil模板规则返回错误")
	}

	_, err = comparator.CompareFormatRules(nil, docRules)
	if err == nil {
		t.Error("期望nil文档规则返回错误")
	}
}

// TestComparator_CompareFormatRules_EmptyRules 测试空规则比较
func TestComparator_CompareFormatRules_EmptyRules(t *testing.T) {
	comparator := NewComparator()

	// 创建空的格式规则
	emptyRules := &types.FormatRules{
		FontRules:      []types.FontRule{},
		ParagraphRules: []types.ParagraphRule{},
	}

	// 测试空规则比较
	comparison, err := comparator.CompareFormatRules(emptyRules, emptyRules)
	if err != nil {
		t.Errorf("空规则比较失败: %v", err)
	}

	if comparison == nil {
		t.Error("期望返回比较结果")
	}

	// 空规则比较应该没有差异
	if len(comparison.Issues) > 0 {
		t.Error("期望空规则比较没有差异")
	}
}

// TestComparator_CompareFormatRules_DifferentRules 测试不同规则比较
func TestComparator_CompareFormatRules_DifferentRules(t *testing.T) {
	comparator := NewComparator()

	// 创建不同的格式规则
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
				ID:       "font2", // 不同的ID
				Name:     "黑体",
				Size:     12.0,
				Color:    types.Color{RGB: "000000"},
				Bold:     true,
				Italic:   false,
			},
		},
		ParagraphRules: []types.ParagraphRule{
			{
				ID:        "para2", // 不同的ID
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

	// 测试不同规则比较
	comparison, err := comparator.CompareFormatRules(docRules, templateRules)
	if err != nil {
		t.Errorf("不同规则比较失败: %v", err)
	}

	if comparison == nil {
		t.Error("期望返回比较结果")
	}

	// 不同规则比较应该有差异
	if len(comparison.Issues) == 0 {
		t.Error("期望发现格式差异")
	}
}

// TestComparator_CompareFormatRules_SameRules 测试相同规则比较
func TestComparator_CompareFormatRules_SameRules(t *testing.T) {
	comparator := NewComparator()

	// 创建相同的格式规则
	sameRules := &types.FormatRules{
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

	// 测试相同规则比较
	comparison, err := comparator.CompareFormatRules(sameRules, sameRules)
	if err != nil {
		t.Errorf("相同规则比较失败: %v", err)
	}

	if comparison == nil {
		t.Error("期望返回比较结果")
	}

	// 相同规则比较应该没有差异
	if len(comparison.Issues) > 0 {
		t.Error("期望相同规则比较没有差异")
	}
}

// TestFormatComparison_Validation 测试格式比较结果验证
func TestFormatComparison_Validation(t *testing.T) {
	// 测试有效的格式比较结果
	validComparison := &types.ComparisonReport{
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
	}

	if validComparison == nil {
		t.Error("期望格式比较结果不为nil")
	}

	if len(validComparison.Issues) == 0 {
		t.Error("期望格式比较结果包含问题")
	}

	// 验证问题内容
	issue := validComparison.Issues[0]
	if issue.ID == "" {
		t.Error("期望问题有ID")
	}

	if issue.Type == "" {
		t.Error("期望问题有类型")
	}

	if issue.Severity == "" {
		t.Error("期望问题有严重程度")
	}

	if issue.Location == "" {
		t.Error("期望问题有位置")
	}

	if issue.Description == "" {
		t.Error("期望问题有描述")
	}

	if len(issue.Suggestions) == 0 {
		t.Error("期望问题有建议")
	}
}

// TestComparator_ErrorHandling 测试错误处理
func TestComparator_ErrorHandling(t *testing.T) {
	comparator := NewComparator()

	// 测试无效的文件路径
	_, err := comparator.CompareWithTemplate("invalid/path/with/invalid/chars/*", "template.docx")
	if err == nil {
		t.Error("期望无效文件路径返回错误")
	}

	_, err = comparator.CompareWithTemplate("document.docx", "invalid/path/with/invalid/chars/*")
	if err == nil {
		t.Error("期望无效模板路径返回错误")
	}

	// 测试不存在的文件
	_, err = comparator.CompareWithTemplate("nonexistent.docx", "nonexistent_template.docx")
	if err == nil {
		t.Error("期望不存在的文件返回错误")
	}
}

// TestComparator_Performance 测试性能
func TestComparator_Performance(t *testing.T) {
	comparator := NewComparator()

	// 创建大量规则进行性能测试
	docRules := &types.FormatRules{
		FontRules:      make([]types.FontRule, 100),
		ParagraphRules: make([]types.ParagraphRule, 100),
	}

	templateRules := &types.FormatRules{
		FontRules:      make([]types.FontRule, 100),
		ParagraphRules: make([]types.ParagraphRule, 100),
	}

	// 填充规则数据
	for i := 0; i < 100; i++ {
		docRules.FontRules[i] = types.FontRule{
			ID:       fmt.Sprintf("font%d", i),
			Name:     "宋体",
			Size:     11.0,
			Color:    types.Color{RGB: "000000"},
			Bold:     false,
			Italic:   false,
		}

		templateRules.FontRules[i] = types.FontRule{
			ID:       fmt.Sprintf("font%d", i),
			Name:     "黑体",
			Size:     12.0,
			Color:    types.Color{RGB: "000000"},
			Bold:     true,
			Italic:   false,
		}

		docRules.ParagraphRules[i] = types.ParagraphRule{
			ID:        fmt.Sprintf("para%d", i),
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
		}

		templateRules.ParagraphRules[i] = types.ParagraphRule{
			ID:        fmt.Sprintf("para%d", i),
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
		}
	}

	// 测试性能
	comparison, err := comparator.CompareFormatRules(docRules, templateRules)
	if err != nil {
		t.Errorf("大量规则比较失败: %v", err)
	}

	if comparison == nil {
		t.Error("期望返回比较结果")
	}

	// 应该有大量差异
	if len(comparison.Issues) == 0 {
		t.Error("期望发现大量格式差异")
	}
} 