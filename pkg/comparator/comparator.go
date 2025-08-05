package comparator

import (
	"fmt"
	"docs-parser/internal/core/comparator"
	"docs-parser/internal/core/types"
)

// Comparator 文档对比器
type Comparator struct {
	factory *comparator.ComparatorFactory
}

// NewComparator 创建新的对比器
func NewComparator() *Comparator {
	factory := comparator.NewComparatorFactory()

	// 注册对比器实现
	factory.RegisterComparator("default", comparator.NewDocumentComparator())

	return &Comparator{
		factory: factory,
	}
}

// CompareWithTemplate 与模板进行对比
func (c *Comparator) CompareWithTemplate(docPath, templatePath string) (*comparator.ComparisonReport, error) {
	fmt.Printf("DEBUG: pkg/comparator 开始比较文档\n")
	
	// 获取默认对比器
	comparator, err := c.factory.GetComparator("default")
	if err != nil {
		return nil, err
	}

	// 执行模板对比
	report, err := comparator.CompareWithTemplate(docPath, templatePath)
	if err != nil {
		return nil, err
	}
	
	fmt.Printf("DEBUG: pkg/comparator 比较完成，发现 %d 个问题\n", len(report.Issues))
	return report, nil
}

// CompareDocuments 对比两个文档
func (c *Comparator) CompareDocuments(doc1Path, doc2Path string) (*comparator.ComparisonReport, error) {
	// 获取默认对比器
	comparator, err := c.factory.GetComparator("default")
	if err != nil {
		return nil, err
	}

	// 执行文档对比
	return comparator.CompareDocuments(doc1Path, doc2Path)
}

// CompareFormatRules 对比格式规则
func (c *Comparator) CompareFormatRules(docRules, templateRules *types.FormatRules) (*comparator.FormatComparison, error) {
	// 获取默认对比器
	comparator, err := c.factory.GetComparator("default")
	if err != nil {
		return nil, err
	}

	// 执行格式规则对比
	return comparator.CompareFormatRules(docRules, templateRules, nil, nil)
}

// CompareContent 对比内容
func (c *Comparator) CompareContent(docContent, templateContent *types.DocumentContent) (*comparator.ContentComparison, error) {
	// 获取默认对比器
	comparator, err := c.factory.GetComparator("default")
	if err != nil {
		return nil, err
	}

	// 执行内容对比
	return comparator.CompareContent(docContent, templateContent)
}

// CompareStyles 对比样式
func (c *Comparator) CompareStyles(docStyles, templateStyles *types.DocumentStyles) (*comparator.StyleComparison, error) {
	// 获取默认对比器
	comparator, err := c.factory.GetComparator("default")
	if err != nil {
		return nil, err
	}

	// 执行样式对比
	return comparator.CompareStyles(docStyles, templateStyles)
}
