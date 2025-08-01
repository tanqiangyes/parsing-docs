# Docs Parser - Word文档格式解析与比较工具

一个用Go语言开发的模块化Word文档解析库，支持多种Word格式的精确解析、格式比较和自动标注功能。

## 🚀 功能特性

### 支持的文档格式
- **现代格式**: `.docx` (Word 2007+)
- **传统格式**: `.doc` (Word 97-2003)
- **富文本格式**: `.rtf` (Rich Text Format)
- **WordPerfect格式**: `.wpd`
- **模板格式**: `.dot`, `.dotx`
- **历史版本**: Word 1.0-6.0, Word 95-2003, Word 365

### 核心功能
- **精确解析**: 深度解析Word文档的所有格式规则和内容结构
- **格式比较**: 与模板或参考文档进行详细的格式对比
- **自动标注**: 在复制的文档中直接标注格式问题
- **修改建议**: 提供具体的格式修改建议和操作步骤
- **合规检查**: 验证文档是否符合指定的格式标准

## 📦 安装

### 环境要求
- Go 1.21+
- Windows/Linux/macOS

### 安装步骤

1. **克隆项目**
```bash
git clone https://github.com/your-username/docs-parser.git
cd docs-parser
```

2. **安装依赖**
```bash
go mod tidy
```

3. **编译项目**
```bash
go build -o docs-parser.exe cmd/main.go
```

## 🛠️ 使用方法

### 命令行工具

#### 比较文档与模板
```bash
# 比较文档与模板
./docs-parser.exe compare --document sample.docx --template template.json

# 比较两个文档
./docs-parser.exe compare --document1 doc1.docx --document2 doc2.docx
```

#### 验证文档格式
```bash
# 验证文档格式
./docs-parser.exe validate --document sample.docx
```

#### 为文档添加标注
```bash
# 为文档添加格式标注
./docs-parser.exe annotate --input sample.docx --output annotated_sample.docx
```

### 编程接口

#### 基本使用示例

```go
package main

import (
    "fmt"
    "log"
    
    "docs-parser/pkg/parser"
    "docs-parser/pkg/comparator"
)

func main() {
    // 解析文档
    doc, err := parser.ParseDocument("sample.docx")
    if err != nil {
        log.Fatal(err)
    }
    
    // 比较文档与模板
    result, err := comparator.CompareWithTemplate("sample.docx", "template.json")
    if err != nil {
        log.Fatal(err)
    }
    
    // 输出比较结果
    fmt.Printf("合规率: %.2f%%\n", result.ComplianceRate)
    fmt.Printf("发现问题: %d个\n", len(result.Issues))
}
```

#### 高级使用示例

```go
package main

import (
    "fmt"
    "log"
    
    "docs-parser/internal/core/comparator"
    "docs-parser/internal/core/annotator"
    "docs-parser/internal/core/validator"
)

func main() {
    // 创建比较器
    comp := comparator.NewDocumentComparator()
    
    // 比较文档
    report, err := comp.CompareWithTemplate("document.docx", "template.json")
    if err != nil {
        log.Fatal(err)
    }
    
    // 检查是否有格式差异
    if report.OverallScore < 100.0 {
        fmt.Println("发现格式差异，生成标注文档...")
        
        // 创建标注器
        annotator := annotator.NewAnnotator()
        
        // 生成标注文档
        err = annotator.AnnotateDocument("document.docx", "document_annotated.docx")
        if err != nil {
            log.Fatal(err)
        }
        
        fmt.Println("标注文档已生成: document_annotated.docx")
    } else {
        fmt.Println("格式相同")
    }
    
    // 验证文档
    validator := validator.NewValidator()
    validationResult, err := validator.ValidateDocument("document.docx")
    if err != nil {
        log.Fatal(err)
    }
    
    fmt.Printf("文档验证结果: 合规率 %.2f%%\n", validationResult.ComplianceRate)
}
```

## 📁 项目结构

```
docs-parser/
├── cmd/                    # 命令行入口
│   └── main.go
├── internal/              # 内部包
│   ├── core/             # 核心功能
│   │   ├── types/        # 数据类型定义
│   │   ├── parser/       # 解析器接口
│   │   ├── comparator/   # 比较器实现
│   │   ├── annotator/    # 标注器实现
│   │   └── validator/    # 验证器实现
│   ├── formats/          # 格式解析器
│   │   ├── docx.go       # DOCX解析器
│   │   ├── doc.go        # DOC解析器
│   │   ├── rtf.go        # RTF解析器
│   │   ├── wpd.go        # WPD解析器
│   │   ├── legacy.go     # 历史版本解析器
│   │   └── word.go       # 通用Word解析器
│   ├── templates/        # 模板管理
│   │   └── template.go
│   └── utils/            # 工具函数
│       ├── file.go
│       └── format.go
├── pkg/                  # 公共API
│   ├── parser/           # 解析器API
│   └── comparator/       # 比较器API
├── examples/             # 使用示例
│   └── basic_usage.go
├── docs/                 # 文档
├── go.mod
├── go.sum
└── README.md
```

## 🔧 API文档

### 解析器 (Parser)

#### 解析文档
```go
doc, err := parser.ParseDocument(filePath string) (*types.Document, error)
```

#### 支持的格式
```go
formats := parser.GetSupportedFormats() []string
```

### 比较器 (Comparator)

#### 与模板比较
```go
result, err := comparator.CompareWithTemplate(docPath, templatePath string) (*ComparisonReport, error)
```

#### 文档间比较
```go
result, err := comparator.CompareDocuments(doc1Path, doc2Path string) (*ComparisonReport, error)
```

### 标注器 (Annotator)

#### 添加标注
```go
err := annotator.AnnotateDocument(sourcePath, outputPath string) error
```

### 验证器 (Validator)

#### 验证文档
```go
result, err := validator.ValidateDocument(filePath string) (*ValidationResult, error)
```

## 📊 数据类型

### Document (文档)
```go
type Document struct {
    Metadata    DocumentMetadata
    Content     DocumentContent
    Styles      DocumentStyles
    FormatRules FormatRules
}
```

### ComparisonReport (比较报告)
```go
type ComparisonReport struct {
    DocumentPath      string
    TemplatePath      string
    OverallScore      float64
    ComplianceRate    float64
    Issues            []FormatIssue
    FormatComparison  *FormatComparison
    ContentComparison *ContentComparison
    StyleComparison   *StyleComparison
    Recommendations   []Recommendation
    Summary           ComparisonSummary
}
```

### ValidationResult (验证结果)
```go
type ValidationResult struct {
    ComplianceRate  float64
    Issues          []ValidationIssue
    Recommendations []Recommendation
}
```

## 🎯 使用场景

### 1. 文档格式标准化
- 确保所有文档符合公司格式标准
- 自动检测格式不一致的地方
- 提供具体的修改建议

### 2. 模板验证
- 验证文档是否按照模板格式编写
- 检查字体、段落、表格等格式要求
- 生成详细的合规报告

### 3. 文档质量检查
- 检查文档的格式完整性
- 验证页面设置和样式
- 提供质量改进建议

### 4. 批量文档处理
- 批量验证多个文档
- 自动生成标注版本
- 统计格式合规情况

## 🔍 格式检查项目

### 字体格式
- 字体名称设置
- 字体大小范围
- 字体颜色配置
- 粗体/斜体设置

### 段落格式
- 段落对齐方式
- 段落间距设置
- 段落缩进配置
- 行距设置

### 表格格式
- 表格边框设置
- 表格宽度配置
- 单元格内容检查
- 表格样式验证

### 页面格式
- 页面大小设置
- 页面边距配置
- 页眉页脚设置
- 分页符检查

## 🚧 开发状态

### 已完成功能 ✅
- [x] 基础架构设计
- [x] 数据类型定义
- [x] DOCX格式解析
- [x] 文档比较功能
- [x] 格式验证功能
- [x] 文档标注功能
- [x] 命令行工具
- [x] 模板管理系统

### 开发中功能 🚧
- [ ] 完整的历史格式支持
- [ ] 高级样式解析
- [ ] 批量处理优化
- [ ] 性能优化

### 计划功能 📋
- [ ] 图形和图片解析
- [ ] 宏和脚本检测
- [ ] 加密文档支持
- [ ] Web界面
- [ ] 插件系统

## 🤝 贡献指南

### 开发环境设置
1. Fork项目
2. 创建功能分支: `git checkout -b feature/new-feature`
3. 提交更改: `git commit -am 'Add new feature'`
4. 推送分支: `git push origin feature/new-feature`
5. 创建Pull Request

### 代码规范
- 遵循Go语言官方代码规范
- 添加适当的注释和文档
- 编写单元测试
- 确保代码通过lint检查

### 测试
```bash
# 运行所有测试
go test ./...

# 运行特定包的测试
go test ./internal/core/comparator

# 运行基准测试
go test -bench=.
```

## 📄 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情。

## 📞 联系方式

- 项目主页: https://github.com/your-username/docs-parser
- 问题反馈: https://github.com/your-username/docs-parser/issues
- 邮箱: your-email@example.com

## 🙏 致谢

感谢所有为这个项目做出贡献的开发者和用户！

---

**注意**: 本项目仍在积极开发中，API可能会有变化。建议在生产环境中使用前进行充分测试。 