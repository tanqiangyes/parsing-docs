package examples

import (
	"fmt"

	"docs-parser/pkg/comparator"
	"docs-parser/pkg/parser"
)

func BasicUsageExample() {
	fmt.Println("=== 文档解析库示例 ===")

	// 创建解析器
	docParser := parser.NewParser()

	// 显示支持的格式
	fmt.Printf("支持的格式: %v\n", docParser.GetSupportedFormats())

	// 示例：解析文档（需要实际的.docx文件）
	fmt.Println("\n=== 解析文档示例 ===")
	fmt.Println("注意：需要提供实际的.docx文件路径")

	// 示例：验证文件格式
	fmt.Println("\n=== 验证文件格式示例 ===")
	fmt.Println("注意：需要提供实际的.docx文件路径")

	// 示例：对比文档
	fmt.Println("\n=== 对比文档示例 ===")
	fmt.Println("注意：需要提供实际的.docx文件和模板文件路径")

	// 创建对比器（用于演示）
	_ = comparator.NewComparator()

	fmt.Println("\n=== 使用说明 ===")
	fmt.Println("1. 解析文档：")
	fmt.Println("   docParser := parser.NewParser()")
	fmt.Println("   doc, err := docParser.ParseDocument(\"document.docx\")")

	fmt.Println("\n2. 对比文档：")
	fmt.Println("   docComparator := comparator.NewComparator()")
	fmt.Println("   report, err := docComparator.CompareWithTemplate(\"document.docx\", \"template.docx\")")

	fmt.Println("\n3. 命令行使用：")
	fmt.Println("   go run cmd/main.go parse document.docx")
	fmt.Println("   go run cmd/main.go compare document.docx template.docx")
	fmt.Println("   go run cmd/main.go validate document.docx")
	fmt.Println("   go run cmd/main.go annotate document.docx annotated_document.docx")

	fmt.Println("\n=== 功能特点 ===")
	fmt.Println("✅ 精确解析Word文档的所有格式规则")
	fmt.Println("✅ 支持.docx、.doc、.rtf、.wpd等格式")
	fmt.Println("✅ 全面的格式对比分析")
	fmt.Println("✅ 详细的修改建议")
	fmt.Println("✅ 在文档上直接添加格式标注")
	fmt.Println("✅ 模块化设计，解析包和对比包独立")

	fmt.Println("\n=== 项目状态 ===")
	fmt.Println("🟢 第一阶段：核心解析功能 - 已完成")
	fmt.Println("🟡 第二阶段：对比功能 - 开发中")
	fmt.Println("🟡 第三阶段：验证和建议 - 开发中")
	fmt.Println("🔴 第四阶段：扩展和优化 - 计划中")

	fmt.Println("\n项目已成功构建！可以开始使用。")
}
