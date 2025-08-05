# Docs Parser - Go 文档解析库

基于 Open XML SDK 设计原则的 Go 语言文档解析库，提供高性能、类型安全的 Word 文档解析、比较和标注功能。

## 🚀 核心特性

### 智能文档比较
- **文本格式比较**: 精确比较字体名称、大小、颜色、粗体、斜体等属性
- **段落格式比较**: 对比对齐方式、缩进、间距等段落属性
- **问题合并**: 同一文本的多个格式问题自动合并为一个批注
- **详细报告**: 提供当前格式vs期望格式的详细对比

### 智能批注功能
- **自动生成**: 检测到格式差异时自动生成标注文档
- **精确定位**: 批注插入到具体的问题位置
- **详细内容**: 批注包含具体的格式差异和建议
- **Word兼容**: 生成的批注完全兼容Word/WPS

### 分层架构设计
- **Packaging Layer**: OPC 容器处理，支持 Open Packaging Convention
- **Document Layer**: Word 文档处理，管理文档部分和解析流程
- **Part Layer**: 文档部分处理，XML 内容加载和缓存
- **Element Layer**: XML 元素处理，强类型系统

### 高性能解析
- **流式处理**: 支持大文件的内存高效处理
- **延迟加载**: 按需解析文档部分
- **并发处理**: 并行解析文档部分
- **性能监控**: 内置性能监控和报告

### 完整的格式支持
- **Word 格式**: 支持 .docx, .doc, .dot, .dotx
- **样式解析**: 完整的样式继承和主题支持
- **表格处理**: 完整的表格结构和内容解析
- **图形元素**: 图片、形状、图表支持

## 📦 安装

```bash
# 克隆仓库
git clone https://github.com/tanqiangyes/parsing-docs.git
cd parsing-docs

# 构建项目
go build -o docs-parser cmd/main.go

# 或者使用 go install
go install ./cmd/main.go
```

## 🛠️ 基本使用

### 命令行工具

```bash
# 对比文档与模板（推荐使用）
./docs-parser compare document.docx template.docx

# 验证文档格式
./docs-parser validate document.docx

# 标注文档
./docs-parser annotate document.docx

# 配置管理
./docs-parser config show
./docs-parser config reset
./docs-parser config path
```

### 文档比较示例

```bash
# 比较两个文档的格式差异
./docs-parser compare 1.docx 2.docx
```

**输出示例**:
```
正在对比文档: 1.docx 与Word模板: 2.docx
DEBUG: 开始对比段落格式，文档段落数: 8, 模板段落数: 9
DEBUG: 段落格式对比完成，发现问题数: 0
DEBUG: 开始对比内容字体，文档段落数: 8, 模板段落数: 9
DEBUG: 发现字体名称问题: 文档=宋体, 模板=黑体
DEBUG: 发现字体大小问题: 文档=22.0, 模板=16.0
DEBUG: 格式对比问题数量: 1
DEBUG: 发现 1 个问题，准备生成标注文档
已生成标注文档: 1_annotated.docx
对比完成，发现 1 个格式问题
```

### API 使用

```go
package main

import (
    "fmt"
    "docs-parser/internal/documents"
    "docs-parser/internal/utils"
)

func main() {
    // 创建 Word 文档处理器
    wordDoc := documents.NewWordprocessingDocument("document.docx")
    defer wordDoc.Close()

    // 打开文档
    if err := wordDoc.Open(); err != nil {
        panic(err)
    }

    // 解析文档
    doc, err := wordDoc.Parse()
    if err != nil {
        panic(err)
    }

    // 使用解析结果
    fmt.Printf("文档包含 %d 个段落\n", len(doc.Content.Paragraphs))
    fmt.Printf("文档包含 %d 个表格\n", len(doc.Content.Tables))
    fmt.Printf("文档包含 %d 个字体规则\n", len(doc.FormatRules.FontRules))

    // 性能监控
    wordDoc.Monitor.PrintReport()
}
```

## 🎯 核心功能详解

### 文本格式比较

系统能够精确比较每个文本运行的格式属性：

- **字体名称**: 宋体 vs 黑体
- **字体大小**: 22.0pt vs 16.0pt  
- **字体颜色**: RGB颜色值比较
- **粗体**: true/false比较
- **斜体**: true/false比较

### 段落格式比较

系统会检查每个段落的格式属性：

- **对齐方式**: left, center, right, justify
- **段落间距**: 段前距、段后距
- **行间距**: 行高倍数
- **缩进**: 左缩进、右缩进、首行缩进

### 智能批注生成

当检测到格式差异时，系统会：

1. **自动生成标注文档**: 复制原文档并添加批注
2. **精确定位**: 将批注插入到具体的问题位置
3. **详细内容**: 批注包含当前格式vs期望格式的详细对比
4. **修改建议**: 提供具体的格式调整建议

**批注内容示例**:
```
字体格式不符合模板要求
当前: 宋体, 22.0pt
期望: 黑体, 16.0pt
建议: 调整字体格式: 字体名称: 文档=宋体, 模板=黑体; 字体大小: 文档=22.0, 模板=16.0
```

## 🏗️ 架构设计

### 分层架构
```
┌─────────────────────────────────────────────────────────────┐
│                    CLI Layer (命令行层)                      │
├─────────────────────────────────────────────────────────────┤
│                 Core Layer (核心功能层)                     │
│  ┌─────────────┐ ┌─────────────┐ ┌─────────────┐         │
│  │ Comparator  │ │  Annotator  │ │   Parser    │         │
│  │   (比较器)   │ │   (批注器)   │ │   (解析器)   │         │
│  └─────────────┘ └─────────────┘ └─────────────┘         │
├─────────────────────────────────────────────────────────────┤
│              Document Layer (文档处理层)                    │
│  ┌─────────────┐ ┌─────────────┐ ┌─────────────┐         │
│  │   Word      │ │   Styles    │ │  Graphics   │         │
│  │ Processing  │ │   Parser    │ │   Parser    │         │
│  └─────────────┘ └─────────────┘ └─────────────┘         │
├─────────────────────────────────────────────────────────────┤
│              Packaging Layer (容器处理层)                   │
│  ┌─────────────┐ ┌─────────────┐ ┌─────────────┐         │
│  │    OPC      │ │   ZIP       │ │   XML       │         │
│  │  Container  │ │  Handler    │ │   Parser    │         │
│  └─────────────┘ └─────────────┘ └─────────────┘         │
└─────────────────────────────────────────────────────────────┘
```

### 核心组件

#### 文档比较器 (Comparator)
```go
// 智能文档比较
comparator := comparator.NewDocumentComparator()
report, err := comparator.CompareWithTemplate("doc.docx", "template.docx")

// 比较结果包含：
// - 文本格式差异（字体名称、大小、颜色等）
// - 段落格式差异（对齐、缩进、间距等）
// - 自动生成的标注文档
```

#### 批注生成器 (Annotator)
```go
// 智能批注生成
annotator := annotator.NewAnnotator()
annotatedPath, err := annotator.AnnotateDocumentWithIssues("doc.docx", issues)

// 批注功能：
// - 自动合并同一文本的多个问题
// - 精确定位到问题位置
// - 生成详细的格式对比信息
```

#### 文档解析器 (Parser)
```go
// 高性能文档解析
wordParser := formats.NewWordParser()
doc, err := wordParser.ParseDocument("document.docx")

// 解析功能：
// - 流式处理大文件
// - 并发解析文档部分
// - 完整的样式继承支持
```

### OPC 容器层
```go
// 处理 Open Packaging Convention 容器
container := packaging.NewOPCContainer("document.docx")
container.Open()
defer container.Close()

// 访问文档部分
content, err := container.ReadFile("word/document.xml")
```

### 文档层
```go
// Word 文档处理
wordDoc := documents.NewWordprocessingDocument("document.docx")
wordDoc.Open()
defer wordDoc.Close()

// 解析文档
doc, err := wordDoc.Parse()
```

## 📊 性能优化

### 流式处理
- 支持大文件的内存高效处理
- 延迟加载文档部分
- 按需解析 XML 内容

### 缓存策略
- 文档部分缓存
- 解析结果缓存
- 配置缓存

### 并发处理
- 并行解析文档部分
- 异步 I/O 操作
- 工作池管理

### 性能监控
内置性能监控功能，提供详细的解析性能报告：

```
=== 性能监控报告 ===
总耗时: 2.1801ms
各步骤耗时:
  - 打开OPC容器: 515.2µs
  - 加载文档部分: 515.2µs
  - 解析元数据: 48µs
  - 解析内容: 559.3µs
  - 解析样式: 538.4µs
  - 解析格式规则: 510.6µs
==================
```

## 📁 项目结构

```
docs-parser/
├── cmd/                    # 命令行工具
│   └── main.go
├── internal/               # 内部包
│   ├── core/              # 核心功能
│   │   ├── comparator/    # 文档比较器
│   │   │   ├── comparator.go    # 智能比较逻辑
│   │   │   └── interface.go     # 比较器接口
│   │   ├── annotator/     # 批注生成器
│   │   │   └── annotator.go     # 智能批注生成
│   │   ├── types/         # 类型定义
│   │   │   ├── document.go      # 文档结构
│   │   │   ├── graphics.go      # 图形元素
│   │   │   └── styles.go        # 样式定义
│   │   └── utils/         # 工具函数
│   ├── documents/         # 文档处理层
│   │   └── wordprocessing.go    # Word文档处理
│   ├── packaging/         # OPC 容器层
│   │   └── opc.go
│   ├── formats/           # 格式解析器
│   │   ├── docx.go       # DOCX格式解析
│   │   └── doc.go        # DOC格式解析
│   └── utils/             # 工具包
│       ├── performance.go # 性能监控
│       └── config.go      # 配置管理
├── pkg/                   # 公共包
│   ├── parser/            # 解析器
│   └── comparator/        # 比较器
├── tests/                 # 测试文件
├── examples/              # 示例代码
├── docs/                  # 文档
├── README.md
└── go.mod
```

## 🧪 测试

```bash
# 运行所有测试
go test ./...

# 运行特定包的测试
go test ./internal/documents

# 运行基准测试
go test -bench=. ./...

# 生成测试覆盖率报告
go test -coverprofile=coverage.out ./...
go tool cover -html=coverage.out
```

## 🔧 开发

### 环境要求
- Go 1.21+
- Git

### 开发流程
```bash
# 克隆项目
git clone https://github.com/tanqiangyes/parsing-docs.git
cd parsing-docs

# 安装依赖
go mod tidy

# 运行测试
go test ./...

# 构建项目
go build -o docs-parser cmd/main.go

# 运行示例
./docs-parser compare examples/doc1.docx examples/template.docx
```

### 代码规范
- 使用 `gofmt` 格式化代码
- 遵循 Go 官方代码规范
- 添加适当的注释和文档
- 编写单元测试

## 📚 文档

- [API 文档](docs/api.md)
- [架构设计](docs/architecture.md)
- [性能优化](docs/performance.md)
- [配置指南](docs/configuration.md)
- [示例代码](examples/)

## 🤝 贡献

欢迎贡献代码！请遵循以下步骤：

1. Fork 项目
2. 创建功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 打开 Pull Request

## 📄 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情。

## 🙏 致谢

- [Microsoft Open XML SDK](https://github.com/dotnet/Open-XML-SDK) - 架构设计参考
- [Open XML Specification](https://docs.microsoft.com/en-us/office/open-xml/) - 规范文档
- [OPC Specification](https://docs.microsoft.com/en-us/office/open-xml/opc) - 容器规范

## 📞 联系方式

- 项目主页: [GitHub Repository](https://github.com/tanqiangyes/parsing-docs)
- 问题反馈: [Issues](https://github.com/tanqiangyes/parsing-docs/issues)
- 功能请求: [Feature Requests](https://github.com/tanqiangyes/parsing-docs/issues/new)

## 📋 更新日志

### v1.1.0 (2025-08-04) - 文本格式比较优化
- ✅ **智能文本格式比较**: 精确比较字体名称、大小、颜色、粗体、斜体等属性
- ✅ **段落格式比较**: 对比对齐方式、缩进、间距等段落属性
- ✅ **问题合并**: 同一文本的多个格式问题自动合并为一个批注
- ✅ **智能批注生成**: 检测到格式差异时自动生成标注文档
- ✅ **精确定位**: 批注插入到具体的问题位置
- ✅ **详细内容**: 批注包含当前格式vs期望格式的详细对比
- ✅ **Word兼容**: 生成的批注完全兼容Word/WPS
- ✅ **性能优化**: 改进解析性能和内存使用

### v1.0.0 (2025-08-05) - 基础功能实现
- ✅ 基于 Open XML SDK 重构架构
- ✅ 实现分层架构设计
- ✅ 添加性能监控功能
- ✅ 完善样式解析
- ✅ 添加表格内容解析
- ✅ 实现配置管理功能
- ✅ 优化命令行工具
- ✅ 添加详细输出功能
- ✅ 改进错误处理
- ✅ 完善文档和示例

## 🎯 使用场景

### 文档格式标准化
- 确保文档符合公司模板要求
- 自动检测格式差异
- 生成详细的修改建议

### 文档质量检查
- 批量检查文档格式
- 生成格式问题报告
- 提供具体的修改指导

### 文档模板验证
- 验证文档是否符合模板规范
- 检测格式不一致的地方
- 生成标注文档便于修改

---

**Docs Parser** - 让 Word 文档解析和格式检查变得简单高效！ 🚀 