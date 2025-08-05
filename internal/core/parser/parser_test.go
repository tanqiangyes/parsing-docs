package parser

import (
	"testing"
)

func TestDefaultParser_ValidateFile(t *testing.T) {
	parser := &DefaultParser{}

	// 测试空文件路径
	err := parser.ValidateFile("")
	if err == nil {
		t.Error("期望空文件路径返回错误")
	}

	// 测试不存在的文件
	err = parser.ValidateFile("nonexistent_file.docx")
	if err == nil {
		t.Error("期望不存在的文件返回错误")
	}
}

func TestDefaultParser_GetSupportedFormats(t *testing.T) {
	parser := &DefaultParser{}
	formats := parser.GetSupportedFormats()

	expectedFormats := []string{"docx", "doc", "rtf", "wpd"}
	for _, expected := range expectedFormats {
		found := false
		for _, format := range formats {
			if format == expected {
				found = true
				break
			}
		}
		if !found {
			t.Errorf("期望支持的格式包含 %s", expected)
		}
	}
} 