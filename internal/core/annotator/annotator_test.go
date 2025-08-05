package annotator

import (
	"testing"
	"docs-parser/internal/core/types"
)

// TestNewAnnotator 测试创建新的标注器
func TestNewAnnotator(t *testing.T) {
	annotator := NewAnnotator()
	if annotator == nil {
		t.Error("期望返回非空的标注器")
	}
}

// TestAnnotator_AnnotateDocument 测试文档标注功能
func TestAnnotator_AnnotateDocument(t *testing.T) {
	annotator := NewAnnotator()

	// 创建测试用的格式问题
	testIssues := []types.FormatIssue{
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
	}

	// 测试不存在的源文档
	err := annotator.AnnotateDocument("nonexistent.docx", "output.docx", testIssues)
	if err == nil {
		t.Error("期望不存在的源文档返回错误")
	}

	// 测试无效的输出路径
	err = annotator.AnnotateDocument("document.docx", "", testIssues)
	if err == nil {
		t.Error("期望空的输出路径返回错误")
	}

	// 测试空的问题列表
	err = annotator.AnnotateDocument("document.docx", "output.docx", []types.FormatIssue{})
	if err == nil {
		t.Error("期望不存在的源文档返回错误")
	}
}

// TestAnnotator_CopyDocument 测试文档复制功能
func TestAnnotator_CopyDocument(t *testing.T) {
	annotator := NewAnnotator()

	// 测试复制不存在的文件
	err := annotator.copyDocument("nonexistent.docx", "output.docx")
	if err == nil {
		t.Error("期望复制不存在的文件返回错误")
	}
}

// TestAnnotator_AddAnnotations 测试添加批注功能
func TestAnnotator_AddAnnotations(t *testing.T) {
	annotator := NewAnnotator()

	// 测试添加批注到不存在的文档
	testIssues := []types.FormatIssue{
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
	}

	err := annotator.addAnnotations("nonexistent.docx", testIssues)
	if err == nil {
		t.Error("期望添加批注到不存在的文档返回错误")
	}
}

// TestAnnotator_GenerateSpecificComments 测试生成具体批注功能
func TestAnnotator_GenerateSpecificComments(t *testing.T) {
	annotator := NewAnnotator()

	testIssue := types.FormatIssue{
		ID:          "test_issue_1",
		Type:        "font",
		Severity:    "medium",
		Location:    "第1段第1个文本",
		Description: "测试格式问题",
		Current:     map[string]interface{}{"font": "宋体", "size": 11.0},
		Expected:    map[string]interface{}{"font": "黑体", "size": 12.0},
		Rule:        "font_format",
		Suggestions: []string{"调整字体格式"},
	}

	comments := annotator.generateSpecificComments(testIssue, 0)
	if len(comments) == 0 {
		t.Error("期望生成至少一个批注")
	}

	// 验证批注内容
	comment := comments[0]
	if comment.id != 0 {
		t.Errorf("期望批注ID为0，实际为%d", comment.id)
	}
	if comment.problem == "" {
		t.Error("期望批注包含问题描述")
	}
}

// TestAnnotator_ExtractCurrentFormat 测试提取当前格式功能
func TestAnnotator_ExtractCurrentFormat(t *testing.T) {
	annotator := NewAnnotator()

	testIssue := types.FormatIssue{
		Current: map[string]interface{}{
			"font":  "宋体",
			"size":  11.0,
			"color": "000000",
		},
	}

	format := annotator.extractCurrentFormat(testIssue)
	if format == "" {
		t.Error("期望提取到格式信息")
	}
}

// TestAnnotator_ExtractExpectedFormat 测试提取期望格式功能
func TestAnnotator_ExtractExpectedFormat(t *testing.T) {
	annotator := NewAnnotator()

	testIssue := types.FormatIssue{
		Expected: map[string]interface{}{
			"font":  "黑体",
			"size":  12.0,
			"color": "000000",
		},
	}

	format := annotator.extractExpectedFormat(testIssue)
	if format == "" {
		t.Error("期望提取到格式信息")
	}
} 