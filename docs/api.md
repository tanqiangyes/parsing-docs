# Docs Parser API 文档

## 概述

Docs Parser 提供了完整的 Word 文档解析、比较和标注功能的 API。本文档详细介绍了所有可用的 API 接口和使用方法。

## 核心 API

### 文档解析器 (Document Parser)

#### 创建解析器

```go
import "docs-parser/internal/formats"

// 创建 Word 文档解析器
parser := formats.NewWordParser()
```

#### 解析文档

```go
// 解析单个文档
doc, err := parser.ParseDocument("document.docx")
if err != nil {
    log.Fatal(err)
}

// 使用解析结果
fmt.Printf("文档包含 %d 个段落\n", len(doc.Content.Paragraphs))
fmt.Printf("文档包含 %d 个表格\n", len(doc.Content.Tables))
```

#### 解析文档结构

```go
type Document struct {
    Metadata    DocumentMetadata
    Content     DocumentContent
    Styles      DocumentStyles
    FormatRules FormatRules
}

type DocumentContent struct {
    Paragraphs []Paragraph
    Tables     []Table
    Images     []Image
}

type Paragraph struct {
    TextRuns []TextRun
    Style    ParagraphStyle
}

type TextRun struct {
    Text string
    Font FontProperties
}
```

### 文档比较器 (Document Comparator)

#### 创建比较器

```go
import "docs-parser/internal/core/comparator"

// 创建文档比较器
comparator := comparator.NewDocumentComparator()
```

#### 比较文档与模板

```go
// 比较文档与模板
report, err := comparator.CompareWithTemplate("document.docx", "template.docx")
if err != nil {
    log.Fatal(err)
}

// 查看比较结果
fmt.Printf("发现 %d 个格式问题\n", len(report.Issues))
for _, issue := range report.Issues {
    fmt.Printf("问题: %s\n", issue.Description)
    fmt.Printf("位置: %s\n", issue.Location)
    fmt.Printf("建议: %s\n", issue.Suggestions[0])
}
```

#### 比较结果结构

```go
type ComparisonReport struct {
    Issues            []FormatIssue
    ContentComparison []FormatIssue
    StyleComparison   []FormatIssue
    AnnotatedPath     string
}

type FormatIssue struct {
    ID          string
    Type        string
    Severity    string
    Location    string
    Description string
    Current     map[string]interface{}
    Expected    map[string]interface{}
    Rule        string
    Suggestions []string
}
```

### 批注生成器 (Document Annotator)

#### 创建批注器

```go
import "docs-parser/internal/core/annotator"

// 创建批注生成器
annotator := annotator.NewAnnotator()
```

#### 生成标注文档

```go
// 为文档添加批注
annotatedPath, err := annotator.AnnotateDocumentWithIssues("document.docx", issues)
if err != nil {
    log.Fatal(err)
}

fmt.Printf("已生成标注文档: %s\n", annotatedPath)
```

#### 批注功能特性

- **自动合并**: 同一文本的多个格式问题自动合并为一个批注
- **精确定位**: 批注插入到具体的问题位置
- **详细内容**: 批注包含当前格式vs期望格式的详细对比
- **Word兼容**: 生成的批注完全兼容Word/WPS

### 文档处理器 (Document Processor)

#### 创建文档处理器

```go
import "docs-parser/internal/documents"

// 创建 Word 文档处理器
wordDoc := documents.NewWordprocessingDocument("document.docx")
defer wordDoc.Close()
```

#### 打开和解析文档

```go
// 打开文档
if err := wordDoc.Open(); err != nil {
    log.Fatal(err)
}

// 解析文档
doc, err := wordDoc.Parse()
if err != nil {
    log.Fatal(err)
}

// 使用解析结果
fmt.Printf("文档包含 %d 个段落\n", len(doc.Content.Paragraphs))
fmt.Printf("文档包含 %d 个表格\n", len(doc.Content.Tables))
fmt.Printf("文档包含 %d 个字体规则\n", len(doc.FormatRules.FontRules))
```

#### 性能监控

```go
// 获取性能监控报告
wordDoc.Monitor.PrintReport()

// 输出示例:
// === 性能监控报告 ===
// 总耗时: 2.1801ms
// 各步骤耗时:
//   - 打开OPC容器: 515.2µs
//   - 加载文档部分: 515.2µs
//   - 解析元数据: 48µs
//   - 解析内容: 559.3µs
//   - 解析样式: 538.4µs
//   - 解析格式规则: 510.6µs
// ==================
```

## 高级功能 API

### 样式解析

#### 样式继承处理

```go
import "docs-parser/internal/core/styles"

// 创建样式处理器
styleProcessor := styles.NewStyleProcessor()

// 处理样式继承
processedStyles, err := styleProcessor.ProcessInheritance(doc.Styles)
if err != nil {
    log.Fatal(err)
}
```

#### 主题样式支持

```go
// 应用主题样式
themeStyles, err := styleProcessor.ApplyThemeStyles(doc.Styles, "Office Theme")
if err != nil {
    log.Fatal(err)
}
```

### 图形元素处理

#### 图片解析

```go
import "docs-parser/internal/core/graphics"

// 创建图形处理器
graphicsProcessor := graphics.NewGraphicsProcessor()

// 解析文档中的图片
images, err := graphicsProcessor.ExtractImages(doc)
if err != nil {
    log.Fatal(err)
}

for _, image := range images {
    fmt.Printf("图片: %s, 大小: %dx%d\n", image.Name, image.Width, image.Height)
}
```

#### 形状和图表

```go
// 解析形状
shapes, err := graphicsProcessor.ExtractShapes(doc)
if err != nil {
    log.Fatal(err)
}

// 解析图表
charts, err := graphicsProcessor.ExtractCharts(doc)
if err != nil {
    log.Fatal(err)
}
```

### 配置管理

#### 加载配置

```go
import "docs-parser/internal/utils"

// 加载配置文件
config, err := utils.LoadConfig("config.json")
if err != nil {
    log.Fatal(err)
}

// 使用配置
fmt.Printf("默认模板路径: %s\n", config.DefaultTemplatePath)
fmt.Printf("输出目录: %s\n", config.OutputDirectory)
```

#### 保存配置

```go
// 创建新配置
config := &utils.Config{
    DefaultTemplatePath: "templates/default.docx",
    OutputDirectory:     "output",
    EnableDebug:         true,
}

// 保存配置
err := utils.SaveConfig(config, "config.json")
if err != nil {
    log.Fatal(err)
}
```

## 错误处理

### 常见错误类型

```go
// 文档不存在
if errors.Is(err, os.ErrNotExist) {
    fmt.Println("文档文件不存在")
}

// 格式不支持
if errors.Is(err, formats.ErrUnsupportedFormat) {
    fmt.Println("不支持的文档格式")
}

// 解析错误
if errors.Is(err, documents.ErrParseFailed) {
    fmt.Println("文档解析失败")
}
```

### 错误处理最佳实践

```go
// 使用错误包装
if err != nil {
    return fmt.Errorf("解析文档失败: %w", err)
}

// 检查特定错误
var parseErr *documents.ParseError
if errors.As(err, &parseErr) {
    fmt.Printf("解析错误: %s, 位置: %s\n", parseErr.Message, parseErr.Location)
}
```

## 性能优化

### 流式处理

```go
// 使用流式解析器处理大文件
streamParser := formats.NewStreamingParser()
doc, err := streamParser.ParseStream("large_document.docx")
if err != nil {
    log.Fatal(err)
}
```

### 并发处理

```go
// 并发解析多个文档
docs := []string{"doc1.docx", "doc2.docx", "doc3.docx"}
results := make(chan *types.Document, len(docs))

for _, docPath := range docs {
    go func(path string) {
        doc, err := parser.ParseDocument(path)
        if err != nil {
            log.Printf("解析失败: %s", err)
            return
        }
        results <- doc
    }(docPath)
}

// 收集结果
for i := 0; i < len(docs); i++ {
    doc := <-results
    fmt.Printf("解析完成: %s\n", doc.Metadata.Title)
}
```

### 内存优化

```go
// 使用内存池
pool := utils.NewDocumentPool()

// 从池中获取文档对象
doc := pool.Get()
defer pool.Put(doc)

// 使用文档对象
err := parser.ParseInto(doc, "document.docx")
```

## 示例代码

### 完整的文档比较流程

```go
package main

import (
    "fmt"
    "log"
    "docs-parser/internal/core/comparator"
    "docs-parser/internal/core/annotator"
    "docs-parser/internal/formats"
)

func main() {
    // 1. 创建解析器
    parser := formats.NewWordParser()
    
    // 2. 解析文档和模板
    doc, err := parser.ParseDocument("document.docx")
    if err != nil {
        log.Fatal(err)
    }
    
    template, err := parser.ParseDocument("template.docx")
    if err != nil {
        log.Fatal(err)
    }
    
    // 3. 创建比较器
    comparator := comparator.NewDocumentComparator()
    
    // 4. 比较文档
    report, err := comparator.CompareWithTemplate("document.docx", "template.docx")
    if err != nil {
        log.Fatal(err)
    }
    
    // 5. 生成标注文档
    if len(report.Issues) > 0 {
        annotator := annotator.NewAnnotator()
        annotatedPath, err := annotator.AnnotateDocumentWithIssues("document.docx", report.Issues)
        if err != nil {
            log.Fatal(err)
        }
        
        fmt.Printf("发现 %d 个格式问题，已生成标注文档: %s\n", len(report.Issues), annotatedPath)
    } else {
        fmt.Println("文档格式符合模板要求")
    }
}
```

### 批量文档处理

```go
package main

import (
    "fmt"
    "log"
    "path/filepath"
    "docs-parser/internal/core/comparator"
    "docs-parser/internal/formats"
)

func main() {
    // 获取所有文档文件
    documents, err := filepath.Glob("documents/*.docx")
    if err != nil {
        log.Fatal(err)
    }
    
    // 创建比较器
    comparator := comparator.NewDocumentComparator()
    
    // 批量比较
    for _, docPath := range documents {
        report, err := comparator.CompareWithTemplate(docPath, "template.docx")
        if err != nil {
            log.Printf("比较失败 %s: %v", docPath, err)
            continue
        }
        
        fmt.Printf("%s: 发现 %d 个问题\n", filepath.Base(docPath), len(report.Issues))
    }
}
```

## 版本兼容性

### API 版本

- **v1.0.0**: 基础功能实现
- **v1.1.0**: 文本格式比较优化，智能批注功能

### 向后兼容性

- 所有公共 API 都保持向后兼容
- 废弃的 API 会提供迁移指南
- 新版本会添加新功能而不破坏现有代码

## 许可证

本项目采用 MIT 许可证。详情请参见 [LICENSE](../LICENSE) 文件。 