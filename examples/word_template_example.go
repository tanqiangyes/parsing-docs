package examples

import (
	"fmt"
	"log"

	"docs-parser/internal/templates"
)

// WordTemplateExample 演示Word文档模板功能
func WordTemplateExample() {
	fmt.Println("=== Word文档模板功能示例 ===")

	// 创建模板管理器
	templateManager := templates.NewTemplateManager("")

	// 示例：从Word文档创建模板
	templatePath := "example_template.docx"

	fmt.Printf("正在从Word文档创建模板: %s\n", templatePath)

	// 加载Word文档作为模板
	template, err := templateManager.LoadTemplate(templatePath)
	if err != nil {
		log.Printf("加载模板失败: %v", err)
		return
	}

	// 显示模板信息
	displayTemplateInfo(template)

	// 验证模板
	fmt.Println("\n验证模板...")
	if err := templateManager.ValidateWordTemplate(templatePath); err != nil {
		fmt.Printf("模板验证失败: %v\n", err)
	} else {
		fmt.Println("模板验证成功")
	}

	// 列出所有模板
	fmt.Println("\n当前加载的模板:")
	templates := templateManager.ListTemplates()
	for i, t := range templates {
		fmt.Printf("  模板 %d: %s (%s)\n", i+1, t.Name, t.ID)
	}
}

// displayTemplateInfo 显示模板信息
func displayTemplateInfo(template *templates.Template) {
	fmt.Printf("\n模板信息:\n")
	fmt.Printf("  ID: %s\n", template.ID)
	fmt.Printf("  名称: %s\n", template.Name)
	fmt.Printf("  描述: %s\n", template.Description)
	fmt.Printf("  版本: %s\n", template.Version)
	fmt.Printf("  源文件: %s\n", template.SourcePath)

	fmt.Printf("  元数据:\n")
	fmt.Printf("    作者: %s\n", template.Metadata.Author)
	fmt.Printf("    创建时间: %s\n", template.Metadata.Created)
	fmt.Printf("    最后更新: %s\n", template.Metadata.LastUpdated)
	fmt.Printf("    分类: %s\n", template.Metadata.Category)
	fmt.Printf("    标签: %v\n", template.Metadata.Tags)

	fmt.Printf("  格式规则:\n")
	fmt.Printf("    字体规则数量: %d\n", len(template.FormatRules.FontRules))
	fmt.Printf("    段落规则数量: %d\n", len(template.FormatRules.ParagraphRules))
	fmt.Printf("    表格规则数量: %d\n", len(template.FormatRules.TableRules))
	fmt.Printf("    页面规则数量: %d\n", len(template.FormatRules.PageRules))

	// 显示字体规则
	if len(template.FormatRules.FontRules) > 0 {
		fmt.Printf("    字体规则:\n")
		for i, rule := range template.FormatRules.FontRules {
			fmt.Printf("      %d. %s (%.1fpt, %s)\n", i+1, rule.Name, rule.Size, rule.Color.RGB)
		}
	}

	// 显示段落规则
	if len(template.FormatRules.ParagraphRules) > 0 {
		fmt.Printf("    段落规则:\n")
		for i, rule := range template.FormatRules.ParagraphRules {
			fmt.Printf("      %d. %s (对齐: %s)\n", i+1, rule.Name, rule.Alignment)
		}
	}

	// 显示表格规则
	if len(template.FormatRules.TableRules) > 0 {
		fmt.Printf("    表格规则:\n")
		for i, rule := range template.FormatRules.TableRules {
			fmt.Printf("      %d. %s (宽度: %.1f%%)\n", i+1, rule.Name, rule.Width)
		}
	}

	// 显示页面规则
	if len(template.FormatRules.PageRules) > 0 {
		fmt.Printf("    页面规则:\n")
		for i, rule := range template.FormatRules.PageRules {
			fmt.Printf("      %d. %s (%.1fx%.1f)\n", i+1, rule.Name, rule.PageSize.Width, rule.PageSize.Height)
		}
	}
}

// CompareWithWordTemplateExample 演示使用Word文档模板进行比较
func CompareWithWordTemplateExample() {
	fmt.Println("\n=== Word文档模板比较示例 ===")

	// 示例文档和模板路径
	documentPath := "example_document.docx"
	templatePath := "example_template.docx"

	fmt.Printf("正在比较文档: %s 与模板: %s\n", documentPath, templatePath)

	// 这里应该调用比较器进行比较
	// 由于比较器已经更新为支持Word文档模板，可以直接使用
	fmt.Println("比较功能已集成到主程序中")
	fmt.Println("使用方法: docs-parser compare <文档路径> <模板路径>")
	fmt.Printf("示例: docs-parser compare %s %s\n", documentPath, templatePath)
}
