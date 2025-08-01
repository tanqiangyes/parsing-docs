# Docs Parser - Go 文档解析库

基于 Open XML SDK 设计原则的 Go 语言文档解析库，支持 Word 文档格式解析、比较和标注。

## 特性

### 🏗️ 分层架构
- **Packaging Layer**: OPC (Open Packaging Convention) 容器处理
- **Document Layer**: 特定文档类型处理 (Word, Excel, PowerPoint)
- **Part Layer**: 文档内各个部分处理 (document.xml, styles.xml 等)
- **Element Layer**: XML 元素和属性处理

### 📄 支持的格式
- **Word 文档**: `.docx`, `.doc`, `.dot`, `.dotx`
- **历史版本**: Word 1.0-6.0, 95-2003, 2007-2019, 365
- **模板文件**: Word 文档模板

### 🔍 核心功能
- **文档解析**: 完整的 WordprocessingML 解析
- **格式比较**: 精确的格式规则比较
- **文档标注**: 自动生成格式标注文档
- **模板验证**: 基于 Word 文档模板的验证

### ⚡ 性能优化
- **流式处理**: 内存高效的文档处理
- **延迟加载**: 按需解析文档部分
- **并发支持**: 并行处理大型文档

## 快速开始

### 安装

```bash
git clone https://github.com/your-repo/docs-parser.git
cd docs-parser
go mod tidy
go build -o main.exe cmd/main.go
```

### 基本使用

```bash
# 比较两个文档
./main.exe compare document1.docx document2.docx

# 验证文档格式
./main.exe validate document.docx

# 标注文档
./main.exe annotate document.docx
```

## 架构设计

### 分层架构

```
internal/
├── packaging/     # OPC 容器处理
├── documents/     # 文档类型处理
├── parts/         # 文档部分处理
├── elements/      # XML 元素处理
└── schemas/       # XML 模式定义
```

### 核心组件

#### 1. OPC 容器层 (`internal/packaging/`)
- 处理 Open Packaging Convention 容器
- 文件索引和访问
- 内容类型映射
- 关系文件处理

#### 2. 文档层 (`internal/documents/`)
- Word 文档处理
- 文档部分加载
- 元数据解析
- 内容提取

#### 3. 解析层 (`internal/formats/`)
- 格式特定解析器
- XML 结构定义
- 数据转换

#### 4. 核心层 (`internal/core/`)
- 类型定义
- 比较算法
- 验证逻辑
- 标注功能

## API 使用

### 文档解析

```go
import "docs-parser/internal/documents"

// 创建 Word 文档
wordDoc := documents.NewWordprocessingDocument("document.docx")
defer wordDoc.Close()

// 打开文档
if err := wordDoc.Open(); err != nil {
    log.Fatal(err)
}

// 解析文档
doc, err := wordDoc.Parse()
if err != nil {
    log.Fatal(err)
}

// 访问解析结果
fmt.Printf("段落数量: %d\n", len(doc.Content.Paragraphs))
fmt.Printf("字体规则: %d\n", len(doc.FormatRules.FontRules))
```

### 格式比较

```go
import "docs-parser/internal/core/comparator"

// 创建比较器
comparator := comparator.NewComparator()

// 比较文档与模板
report, err := comparator.CompareWithTemplate("document.docx", "template.docx")
if err != nil {
    log.Fatal(err)
}

// 检查格式问题
if len(report.Issues) == 0 {
    fmt.Println("格式相同")
} else {
    fmt.Printf("发现 %d 个格式问题\n", len(report.Issues))
}
```

## 命令行工具

### 比较命令

```bash
# 比较两个文档
./main.exe compare document.docx template.docx

# 输出示例
正在对比文档: document.docx 与Word模板: template.docx
文档解析完成: document.docx
  - 段落数量: 8
  - 字体规则数量: 2
  - 段落规则数量: 8
  - 页面规则数量: 1
发现 2 个格式问题
对比完成，发现 2 个格式问题
```

### 验证命令

```bash
# 验证文档格式
./main.exe validate document.docx
```

### 标注命令

```bash
# 生成标注文档
./main.exe annotate document.docx
```

## 开发指南

### 项目结构

```
docs-parser/
├── cmd/                    # 命令行入口
├── internal/               # 内部包
│   ├── packaging/         # OPC 容器处理
│   ├── documents/         # 文档类型处理
│   ├── formats/           # 格式解析器
│   ├── core/              # 核心功能
│   │   ├── types/         # 类型定义
│   │   ├── parser/        # 解析器工厂
│   │   ├── comparator/    # 比较器
│   │   ├── validator/     # 验证器
│   │   ├── annotator/     # 标注器
│   │   └── styles/        # 样式处理
│   └── utils/             # 工具函数
├── pkg/                   # 公共包
├── tests/                 # 测试文件
├── examples/              # 示例代码
└── docs/                  # 文档
```

### 添加新格式支持

1. **创建格式解析器**
```go
// internal/formats/newformat.go
type NewFormatParser struct{}

func (nfp *NewFormatParser) ParseDocument(filePath string) (*types.Document, error) {
    // 实现解析逻辑
}
```

2. **注册解析器**
```go
// internal/core/parser/factory.go
func (pf *ParserFactory) RegisterParser(format string, parser Parser) {
    pf.parsers[format] = parser
}
```

3. **添加测试**
```go
// tests/newformat_test.go
func TestNewFormatParser(t *testing.T) {
    // 测试实现
}
```

## 性能优化

### 内存管理
- 流式读取大文件
- 延迟解析文档部分
- 内存池使用

### 并发处理
- 并行解析文档部分
- 异步 I/O 操作
- 工作池模式

### 缓存策略
- 解析结果缓存
- 格式规则缓存
- 模板缓存

## 测试

```bash
# 运行所有测试
go test ./...

# 运行特定测试
go test ./internal/core/comparator

# 运行基准测试
go test -bench=. ./...
```

## 贡献指南

1. Fork 项目
2. 创建功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 打开 Pull Request

## 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情。

## 致谢

- 参考 [Microsoft Open XML SDK](https://github.com/dotnet/Open-XML-SDK) 的设计原则
- 基于 [Open XML 规范](https://docs.microsoft.com/en-us/office/open-xml/open-xml-sdk)
- 遵循 [OPC 规范](https://docs.microsoft.com/en-us/office/open-xml/opc)

## 更新日志

### v1.0.0 (2024-08-01)
- ✅ 基于 Open XML SDK 的分层架构
- ✅ 完整的 WordprocessingML 解析
- ✅ 精确的格式比较算法
- ✅ 文档标注功能
- ✅ 命令行工具
- ✅ 单元测试覆盖
- ✅ 性能优化

## 联系方式

- 项目主页: [GitHub Repository](https://github.com/your-repo/docs-parser)
- 问题反馈: [Issues](https://github.com/your-repo/docs-parser/issues)
- 功能请求: [Feature Requests](https://github.com/your-repo/docs-parser/issues/new) 