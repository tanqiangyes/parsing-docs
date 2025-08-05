# Docs Parser - Go 文档解析库

基于 Open XML SDK 设计原则的 Go 语言文档解析库，提供高性能、类型安全的 Word 文档解析、比较和标注功能。

## 🚀 核心特性

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

### 智能比较功能
- **格式对比**: 精确的格式规则比较
- **差异检测**: 自动识别格式差异
- **问题报告**: 详细的格式问题报告
- **修改建议**: 智能的格式修改建议

## 📦 安装

```bash
# 克隆仓库
git clone https://github.com/your-repo/docs-parser.git
cd docs-parser

# 构建项目
go build -o docs-parser cmd/main.go

# 或者使用 go install
go install ./cmd/main.go
```

## 🛠️ 基本使用

### 命令行工具

```bash
# 对比文档与模板
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

## 🏗️ 架构设计

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

### 解析层
```go
// 格式解析器
docxParser := formats.NewDocxParser()
doc, err := docxParser.ParseDocument("document.docx")
```

### 核心层
```go
// 文档比较
comparator := comparator.NewComparator()
report, err := comparator.CompareWithTemplate("doc.docx", "template.docx")
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

## 📈 性能监控

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
│   │   ├── comparator/    # 文档比较
│   │   ├── types/         # 类型定义
│   │   └── utils/         # 工具函数
│   ├── documents/         # 文档处理层
│   │   └── wordprocessing.go
│   ├── packaging/         # OPC 容器层
│   │   └── opc.go
│   ├── formats/           # 格式解析器
│   │   ├── docx.go
│   │   └── doc.go
│   └── utils/             # 工具包
│       ├── performance.go
│       └── config.go
├── pkg/                   # 公共包
│   ├── parser/            # 解析器
│   └── comparator/        # 比较器
├── tests/                 # 测试文件
├── examples/              # 示例代码
├── docs/                  # 文档
├── README.md
└── go.mod
```

## 🔧 开发

### 环境要求
- Go 1.21+
- Git

### 开发流程
```bash
# 克隆项目
git clone https://github.com/your-repo/docs-parser.git
cd docs-parser

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

- 项目主页: [GitHub Repository](https://github.com/your-repo/docs-parser)
- 问题反馈: [Issues](https://github.com/your-repo/docs-parser/issues)
- 功能请求: [Feature Requests](https://github.com/your-repo/docs-parser/issues/new)

## 📋 更新日志

### v1.0.0 (2024-01-XX)
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

---

**Docs Parser** - 让 Word 文档解析变得简单高效！ 🚀 