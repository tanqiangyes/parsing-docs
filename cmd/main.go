package main

import (
	"fmt"
	"os"
	"path/filepath"

	"docs-parser/internal/core/annotator"
	"docs-parser/internal/utils"
	pkgcomparator "docs-parser/pkg/comparator"

	"github.com/spf13/cobra"
)

var rootCmd = &cobra.Command{
	Use:   "docs-parser",
	Short: "Word文档格式解析与比较工具",
	Long: `基于Open XML SDK设计原则的Go语言文档解析库，
支持Word文档格式解析、比较和标注功能。`,
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

		// 如果有格式问题，显示问题详情
		fmt.Printf("发现 %d 个格式问题\n", len(report.Issues))

		// 自动生成标注文档
		if len(report.Issues) > 0 {
			fmt.Println("正在生成标注文档...")

			// 生成输出路径
			ext := filepath.Ext(docPath)
			baseName := docPath[:len(docPath)-len(ext)]
			outputPath := baseName + "_annotated" + ext

			// 使用标注器生成标注文档
			docAnnotator := annotator.NewAnnotator()
			err = docAnnotator.AnnotateDocument(docPath, outputPath, report.Issues)
			if err != nil {
				fmt.Printf("警告: 生成标注文档失败: %v\n", err)
			} else {
				fmt.Printf("已生成标注文档: %s\n", outputPath)
			}
		}

		fmt.Printf("对比完成，发现 %d 个格式问题\n", len(report.Issues))
	},
}

var validateCmd = &cobra.Command{
	Use:   "validate [文档路径]",
	Short: "验证文档格式",
	Long:  `验证Word文档的格式是否符合标准。`,
	Args:  cobra.ExactArgs(1),
	Run: func(cmd *cobra.Command, args []string) {
		docPath := args[0]
		fmt.Printf("正在验证文档: %s\n", docPath)
		fmt.Println("验证功能正在开发中...")
	},
}

var annotateCmd = &cobra.Command{
	Use:   "annotate [文档路径]",
	Short: "标注文档",
	Long:  `为Word文档添加格式标注。`,
	Args:  cobra.ExactArgs(1),
	Run: func(cmd *cobra.Command, args []string) {
		docPath := args[0]
		fmt.Printf("正在标注文档: %s\n", docPath)
		fmt.Println("标注功能正在开发中...")
	},
}

var configCmd = &cobra.Command{
	Use:   "config",
	Short: "配置管理",
	Long:  `管理docs-parser的配置选项。`,
}

var configShowCmd = &cobra.Command{
	Use:   "show",
	Short: "显示当前配置",
	Run: func(cmd *cobra.Command, args []string) {
		configPath := utils.GetConfigPath()
		config, err := utils.LoadConfig(configPath)
		if err != nil {
			fmt.Printf("加载配置失败: %v\n", err)
			os.Exit(1)
		}

		fmt.Println("当前配置:")
		fmt.Printf("  性能监控: %v\n", config.ParseOptions.EnablePerformanceMonitoring)
		fmt.Printf("  详细输出: %v\n", config.ParseOptions.EnableDetailedOutput)
		fmt.Printf("  样式解析: %v\n", config.ParseOptions.EnableStyleParsing)
		fmt.Printf("  表格解析: %v\n", config.ParseOptions.EnableTableParsing)
		fmt.Printf("  严格模式: %v\n", config.CompareOptions.StrictMode)
		fmt.Printf("  忽略大小写: %v\n", config.CompareOptions.IgnoreCase)
		fmt.Printf("  缓存启用: %v\n", config.PerformanceOptions.EnableCaching)
		fmt.Printf("  缓存大小: %d\n", config.PerformanceOptions.CacheSize)
	},
}

var configResetCmd = &cobra.Command{
	Use:   "reset",
	Short: "重置为默认配置",
	Run: func(cmd *cobra.Command, args []string) {
		configPath := utils.GetConfigPath()
		config := utils.DefaultConfig()

		if err := utils.SaveConfig(configPath, config); err != nil {
			fmt.Printf("保存配置失败: %v\n", err)
			os.Exit(1)
		}

		fmt.Println("配置已重置为默认值")
	},
}

var configPathCmd = &cobra.Command{
	Use:   "path",
	Short: "显示配置文件路径",
	Run: func(cmd *cobra.Command, args []string) {
		configPath := utils.GetConfigPath()
		fmt.Printf("配置文件路径: %s\n", configPath)
	},
}

func init() {
	rootCmd.AddCommand(compareCmd)
	rootCmd.AddCommand(validateCmd)
	rootCmd.AddCommand(annotateCmd)

	// 配置命令
	configCmd.AddCommand(configShowCmd)
	configCmd.AddCommand(configResetCmd)
	configCmd.AddCommand(configPathCmd)
	rootCmd.AddCommand(configCmd)
}

func main() {
	if err := rootCmd.Execute(); err != nil {
		fmt.Println(err)
		os.Exit(1)
	}
}
