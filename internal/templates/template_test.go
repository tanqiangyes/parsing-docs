package templates

import (
	"testing"
)

func TestNewTemplateManager(t *testing.T) {
	manager := NewTemplateManager("test_templates")
	if manager == nil {
		t.Error("期望返回非空的模板管理器")
	}
}

func TestTemplateManager_LoadTemplate(t *testing.T) {
	manager := NewTemplateManager("test_templates")

	// 测试不存在的模板文件
	_, err := manager.LoadTemplate("nonexistent.docx")
	if err == nil {
		t.Error("期望不存在的模板文件返回错误")
	}

	// 测试空路径
	_, err = manager.LoadTemplate("")
	if err == nil {
		t.Error("期望空路径返回错误")
	}
}

func TestTemplateManager_LoadTemplatesFromDirectory(t *testing.T) {
	manager := NewTemplateManager("test_templates")

	// 测试不存在的目录
	err := manager.LoadTemplatesFromDirectory("nonexistent_directory")
	if err == nil {
		t.Error("期望不存在的目录返回错误")
	}
}

func TestTemplateManager_ValidateWordTemplate(t *testing.T) {
	manager := NewTemplateManager("test_templates")

	// 测试不存在的模板文件
	err := manager.ValidateWordTemplate("nonexistent.docx")
	if err == nil {
		t.Error("期望不存在的模板文件返回错误")
	}
}

func TestTemplateManager_IsSupportedWordFormat(t *testing.T) {
	manager := NewTemplateManager("test_templates")

	// 测试支持的格式
	supportedFormats := []string{".docx", ".doc", ".dot", ".dotx"}
	for _, format := range supportedFormats {
		if !manager.isSupportedWordFormat(format) {
			t.Errorf("期望 %s 是支持的格式", format)
		}
	}

	// 测试不支持的格式
	unsupportedFormats := []string{".txt", ".pdf", ".json", ".xml"}
	for _, format := range unsupportedFormats {
		if manager.isSupportedWordFormat(format) {
			t.Errorf("期望 %s 不是支持的格式", format)
		}
	}
} 