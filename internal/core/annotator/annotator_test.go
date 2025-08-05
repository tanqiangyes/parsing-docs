package annotator

import (
	"testing"
)

func TestNewAnnotator(t *testing.T) {
	annotator := NewAnnotator()
	if annotator == nil {
		t.Error("期望返回非空的标注器")
	}
}

func TestAnnotator_AnnotateDocument(t *testing.T) {
	annotator := NewAnnotator()

	// 测试不存在的源文档
	err := annotator.AnnotateDocument("nonexistent.docx", "output.docx")
	if err == nil {
		t.Error("期望不存在的源文档返回错误")
	}

	// 测试无效的输出路径
	err = annotator.AnnotateDocument("document.docx", "")
	if err == nil {
		t.Error("期望空的输出路径返回错误")
	}
} 