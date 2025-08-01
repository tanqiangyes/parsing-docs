package comparator

import (
	"testing"
)

func TestDocumentComparator_CompareWithTemplate(t *testing.T) {
	comparator := NewDocumentComparator()

	// 测试不存在的文档
	_, err := comparator.CompareWithTemplate("nonexistent.docx", "template.docx")
	if err == nil {
		t.Error("期望不存在的文档返回错误")
	}

	// 测试不存在的模板
	_, err = comparator.CompareWithTemplate("document.docx", "nonexistent.docx")
	if err == nil {
		t.Error("期望不存在的模板返回错误")
	}
}

func TestDocumentComparator_CompareDocuments(t *testing.T) {
	comparator := NewDocumentComparator()

	// 测试不存在的文档
	_, err := comparator.CompareDocuments("nonexistent1.docx", "nonexistent2.docx")
	if err == nil {
		t.Error("期望不存在的文档返回错误")
	}
}

func TestNewDocumentComparator(t *testing.T) {
	comparator := NewDocumentComparator()
	if comparator == nil {
		t.Error("期望返回非空的比较器")
	}
} 