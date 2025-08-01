package main

import (
	"fmt"
	"os"
	"path/filepath"

	"docs-parser/internal/core/annotator"
	"docs-parser/internal/templates"
	pkgcomparator "docs-parser/pkg/comparator"
	pkgparser "docs-parser/pkg/parser"

	"github.com/spf13/cobra"
)

var rootCmd = &cobra.Command{
	Use:   "docs-parser",
	Short: "精确解析Word文档格式并提供修改建议",
	Long: `文档解析库 - 一个用于精确解析Word文档格式的工具
支持解析.docx、.doc、.rtf等格式，并提供详细的格式对比和修改建议。`,
}

var parseCmd = &cobra.Command{
	Use:   "parse [文件路径]",
	Short: "解析Word文档",
	Args:  cobra.ExactArgs(1),
	Run: func(cmd *cobra.Command, args []string) {
		filePath := args[0]
		fmt.Printf("正在解析文件: %s\n", filePath)

		// 使用解析包
		docParser := pkgparser.NewParser()
		doc, err := docParser.ParseDocument(filePath)
		if err != nil {
			fmt.Printf("解析失败: %v\n", err)
			os.Exit(1)
		}

		fmt.Printf("解析成功，文档包含 %d 个段落\n", len(doc.Content.Paragraphs))
	},
}

var compareCmd = &cobra.Command{
	Use:   "compare [文档路径] [模板路径]",
	Short: "对比文档与Word文档模板",
	Long: `对比文档与Word文档模板，支持.docx、.doc、.dot、.dotx格式的模板文件。
模板应该是Word文档，包含所需的格式规则和样式。`,
	Args: cobra.ExactArgs(2),
	Run: func(cmd *cobra.Command, args []string) {
		docPath := args[0]
		templatePath := args[1]

		fmt.Printf("正在对比文档: %s 与Word模板: %s\n", docPath, templatePath)

		// 使用对比包
		docComparator := pkgcomparator.NewComparator()
		report, err := docComparator.CompareWithTemplate(docPath, templatePath)
		if err != nil {
			fmt.Printf("对比失败: %v\n", err)
			os.Exit(1)
		}

		// 检查是否有格式问题
		if len(report.Issues) == 0 {
			fmt.Println("格式相同")
			return
		}

		// 如果有格式问题，自动生成标注文档
		fmt.Printf("发现 %d 个格式问题，正在生成标注文档...\n", len(report.Issues))

		// 生成输出文件路径
		ext := filepath.Ext(docPath)
		baseName := docPath[:len(docPath)-len(ext)]
		outputPath := baseName + "_annotated" + ext

		// 使用标注包生成标注文档
		docAnnotator := annotator.NewAnnotator()
		err = docAnnotator.AnnotateDocument(docPath, outputPath)
		if err != nil {
			fmt.Printf("生成标注文档失败: %v\n", err)
			os.Exit(1)
		}

		fmt.Printf("标注文档已生成: %s\n", outputPath)
		fmt.Printf("对比完成，发现 %d 个格式问题\n", len(report.Issues))
	},
}

var templateCmd = &cobra.Command{
	Use:   "template [模板路径]",
	Short: "显示Word文档模板信息",
	Long: `解析并显示Word文档模板的详细信息，包括格式规则、样式等。
支持.docx、.doc、.dot、.dotx格式的模板文件。`,
	Args: cobra.ExactArgs(1),
	Run: func(cmd *cobra.Command, args []string) {
		templatePath := args[0]

		fmt.Printf("正在解析Word文档模板: %s\n", templatePath)

		// 使用模板管理器
		templateManager := templates.NewTemplateManager("")
		template, err := templateManager.LoadTemplate(templatePath)
		if err != nil {
			fmt.Printf("解析模板失败: %v\n", err)
			os.Exit(1)
		}

		// 显示模板信息
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

		fmt.Println("\n模板解析完成")
	},
}

func init() {
	rootCmd.AddCommand(parseCmd)
	rootCmd.AddCommand(compareCmd)
	rootCmd.AddCommand(templateCmd)
}

func main() {
	if err := rootCmd.Execute(); err != nil {
		fmt.Println(err)
		os.Exit(1)
	}
}
