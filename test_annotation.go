package main

import (
	"fmt"
	"docs-parser/internal/core/comparator"
)

func main() {
	fmt.Println("测试标注功能...")
	
	// 创建比较器
	dc := comparator.NewDocumentComparator()
	
	// 执行比较
	report, err := dc.CompareWithTemplate("1.docx", "2.docx")
	if err != nil {
		fmt.Printf("比较失败: %v\n", err)
		return
	}
	
	fmt.Printf("比较完成，发现 %d 个问题\n", len(report.Issues))
	
	if len(report.Issues) > 0 {
		fmt.Printf("标注文档路径: %s\n", report.AnnotatedDocumentPath)
	}
} 