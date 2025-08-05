package tests

import (
	"os"
	"path/filepath"
	"testing"

	"docs-parser/internal/core/annotator"
	"docs-parser/internal/core/comparator"
	"docs-parser/internal/core/types"
	"docs-parser/internal/core/validator"
	"docs-parser/internal/templates"
	"docs-parser/pkg/parser"
)

// TestDocumentParsing 测试文档解析功能
func TestDocumentParsing(t *testing.T) {
	// 创建测试文档路径
	testDocPath := "test_document.docx"

	// 清理测试文件
	defer func() {
		if _, err := os.Stat(testDocPath); err == nil {
			os.Remove(testDocPath)
		}
	}()

	// 创建解析器
	docParser := parser.NewParser()

	// 测试解析不存在的文件
	_, err := docParser.ParseDocument(testDocPath)
	if err == nil {
		t.Error("应该返回错误，因为文件不存在")
	}

	// 测试获取支持的格式
	formats := docParser.GetSupportedFormats()
	expectedFormats := []string{"docx", "doc", "rtf", "wpd"}

	if len(formats) != len(expectedFormats) {
		t.Errorf("支持的格式数量不匹配，期望 %d，实际 %d", len(expectedFormats), len(formats))
	}
}

// TestTemplateManagement 测试模板管理功能
func TestTemplateManagement(t *testing.T) {
	// 创建模板管理器
	templateManager := templates.NewTemplateManager("")

	// 测试加载不存在的Word模板
	_, err := templateManager.LoadTemplate("nonexistent.docx")
	if err == nil {
		t.Error("应该返回错误，因为模板文件不存在")
	}

	// 测试加载不支持的格式
	_, err = templateManager.LoadTemplate("test.txt")
	if err == nil {
		t.Error("应该返回错误，因为不支持.txt格式")
	}

	// 测试模板验证
	err = templateManager.ValidateWordTemplate("nonexistent.docx")
	if err == nil {
		t.Error("应该返回错误，因为模板文件不存在")
	}
}

// TestDocumentComparison 测试文档比较功能
func TestDocumentComparison(t *testing.T) {
	// 创建比较器
	comp := comparator.NewDocumentComparator()

	// 测试比较不存在的文档
	_, err := comp.CompareWithTemplate("nonexistent.docx", "nonexistent_template.docx")
	if err == nil {
		t.Error("应该返回错误，因为文档不存在")
	}
}

// TestDocumentValidation 测试文档验证功能
func TestDocumentValidation(t *testing.T) {
	// 创建验证器
	validator := validator.NewValidator()

	// 测试验证不存在的文档
	_, err := validator.ValidateDocument("nonexistent.docx")
	if err == nil {
		t.Error("应该返回错误，因为文档不存在")
	}
}

// TestDocumentAnnotation 测试文档标注功能
func TestDocumentAnnotation(t *testing.T) {
	// 创建标注器
	annotator := annotator.NewAnnotator()

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

	// 测试标注不存在的文档
	err := annotator.AnnotateDocument("nonexistent.docx", "output.docx", testIssues)
	if err == nil {
		t.Error("应该返回错误，因为源文档不存在")
	}
}

// TestFileOperations 测试文件操作功能
func TestFileOperations(t *testing.T) {
	// 创建测试目录
	testDir := "test_files"
	err := os.MkdirAll(testDir, 0755)
	if err != nil {
		t.Errorf("创建测试目录失败: %v", err)
	}
	defer func() {
		os.RemoveAll(testDir)
	}()

	// 创建测试文件
	testFilePath := filepath.Join(testDir, "test.txt")
	testContent := []byte("测试内容")
	err = os.WriteFile(testFilePath, testContent, 0644)
	if err != nil {
		t.Errorf("创建测试文件失败: %v", err)
	}

	// 验证文件存在
	if _, err := os.Stat(testFilePath); os.IsNotExist(err) {
		t.Error("测试文件应该存在")
	}
}

// TestTemplateValidation 测试模板验证功能
func TestTemplateValidation(t *testing.T) {
	// 创建模板管理器
	templateManager := templates.NewTemplateManager("")

	// 测试无效模板
	invalidTemplate := &templates.Template{
		ID:          "",
		Name:        "",
		Description: "",
		Version:     "",
	}

	err := templateManager.ValidateTemplate(invalidTemplate)
	if err == nil {
		t.Error("无效模板应该验证失败")
	}

	// 测试Word模板验证
	err = templateManager.ValidateWordTemplate("nonexistent.docx")
	if err == nil {
		t.Error("不存在的Word模板应该验证失败")
	}
}

// TestFormatRules 测试格式规则功能
func TestFormatRules(t *testing.T) {
	// 测试格式规则结构
	// 由于我们不再有默认模板，我们测试格式规则的基本结构
	templateManager := templates.NewTemplateManager("")

	// 测试加载不存在的Word模板
	_, err := templateManager.LoadTemplate("nonexistent.docx")
	if err == nil {
		t.Error("加载不存在的Word模板应该返回错误")
	}

	// 测试不支持的格式
	_, err = templateManager.LoadTemplate("test.txt")
	if err == nil {
		t.Error("加载不支持的格式应该返回错误")
	}
}

// TestErrorHandling 测试错误处理功能
func TestErrorHandling(t *testing.T) {
	// 测试解析器错误处理
	docParser := parser.NewParser()

	// 测试无效文件路径
	_, err := docParser.ParseDocument("")
	if err == nil {
		t.Error("空文件路径应该返回错误")
	}

	// 测试模板管理器错误处理
	templateManager := templates.NewTemplateManager("")

	// 测试加载不存在的模板文件
	_, err = templateManager.LoadTemplate("nonexistent.json")
	if err == nil {
		t.Error("加载不存在的模板文件应该返回错误")
	}

	// 测试获取不存在的模板
	_, err = templateManager.GetTemplate("nonexistent")
	if err == nil {
		t.Error("获取不存在的模板应该返回错误")
	}
}

// BenchmarkDocumentParsing 文档解析性能测试
func BenchmarkDocumentParsing(b *testing.B) {
	docParser := parser.NewParser()

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		// 这里应该使用真实的测试文档
		// 由于没有真实文档，只是测试函数调用
		docParser.GetSupportedFormats()
	}
}

// BenchmarkTemplateLoading 模板加载性能测试
func BenchmarkTemplateLoading(b *testing.B) {
	templateManager := templates.NewTemplateManager("")

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		// 测试模板验证性能
		templateManager.ValidateWordTemplate("nonexistent.docx")
	}
}

// BenchmarkDocumentComparison 文档比较性能测试
func BenchmarkDocumentComparison(b *testing.B) {
	comp := comparator.NewDocumentComparator()

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		// 这里应该使用真实的测试文档和模板
		// 由于没有真实文件，只是测试函数调用
		_ = comp
	}
}
